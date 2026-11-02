import * as XmlUtils from "@/utils/xmlUtils";
import { RelManager } from "@/core/PartsManagers/RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import { Run as RunInterface } from "@/core/files/paragraph/types";
import AdmZip from "adm-zip";
import { wordRootAttributes } from "@/utils/rootNamespaces";

const FOOTNOTES_PATH = "word/footnotes.xml";
const FOOTNOTE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const FOOTNOTE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const FOOTNOTE_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface FootnoteEntry {
  id: number;
  text: string;
}

export class FootnoteManager {
  private zip: AdmZip;
  private rels: RelManager;
  private contentTypes: ContentTypesManager;

  constructor(zip: AdmZip) {
    this.zip = zip;
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
  }

  // ─── Internal read/write ──────────────────────────────────────────────────

  private async readFootnotes(): Promise<any> {
    const xml = this.zip.readAsText(FOOTNOTES_PATH);
    if (!xml) return this.emptyFootnotesDoc();
    return XmlUtils.parseXml(xml);
  }

  private async writeFootnotes(obj: any): Promise<void> {
    const xml = XmlUtils.buildXml(obj["w:footnotes"], {
      rootName: "w:footnotes",
      headless: false,
      pretty: true,
    });
    this.zip.addFile(FOOTNOTES_PATH, Buffer.from(xml, "utf-8"));
  }

  private emptyFootnotesDoc(): any {
    return {
      "w:footnotes": {
        // The standard Word set: footnotes copied from a modern source carry
        // r:id, w14:paraId, mc:AlternateContent… all of which need a binding.
        $: { ...wordRootAttributes(), "xmlns:w": FOOTNOTE_NS },
        "w:footnote": [
          // Required separator footnotes that Word expects
          {
            $: { "w:type": "separator", "w:id": "-1" },
            "w:p": {
              "w:r": { "w:separator": {} },
            },
          },
          {
            $: { "w:type": "continuationSeparator", "w:id": "0" },
            "w:p": {
              "w:r": { "w:continuationSeparator": {} },
            },
          },
        ],
      },
    };
  }

  private normalizeArray(obj: any): any[] {
    const raw = obj?.["w:footnotes"]?.["w:footnote"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  private nextId(footnotes: any[]): number {
    const ids = footnotes
      .map((f) => parseInt(f.$?.["w:id"] ?? "0", 10))
      .filter((n) => n > 0);
    return ids.length ? Math.max(...ids) + 1 : 1;
  }

  // ─── Registration (called once when footnotes are first added) ────────────

  private async ensureRegistered(): Promise<void> {
    const alreadyExists = this.zip.getEntry(FOOTNOTES_PATH);
    if (!alreadyExists) {
      await this.writeFootnotes(this.emptyFootnotesDoc());
      await this.contentTypes.addOverride(`/${FOOTNOTES_PATH}`, FOOTNOTE_CONTENT_TYPE);
      await this.rels.addRelationship(
        await this.rels.genId(),
        FOOTNOTE_REL_TYPE,
        "footnotes.xml",
      );
    }
  }

  // ─── Public API ───────────────────────────────────────────────────────────

  /**
   * Returns all user-defined footnotes (excludes separator entries).
   */
  public async getFootnotes(): Promise<FootnoteEntry[]> {
    const obj = await this.readFootnotes();
    return this.normalizeArray(obj)
      .filter((f) => !f.$?.["w:type"])
      .map((f) => ({
        id: parseInt(f.$?.["w:id"] ?? "0", 10),
        text: this.extractText(f),
      }));
  }

  /**
   * Adds a new footnote to footnotes.xml and returns a run object
   * that should be inserted into the paragraph at the desired position.
   *
   * @param text   The footnote body text.
   * @returns      The footnote id and a run object containing the reference mark.
   */
  public async addFootnote(text: string): Promise<{ id: number; run: RunInterface }> {
    await this.ensureRegistered();

    const obj = await this.readFootnotes();
    const footnotes = this.normalizeArray(obj);
    const id = this.nextId(footnotes);

    footnotes.push({
      $: { "w:id": String(id) },
      "w:p": {
        "w:pPr": { "w:pStyle": { $: { "w:val": "FootnoteText" } } },
        "w:r": [
          {
            "w:rPr": { "w:rStyle": { $: { "w:val": "FootnoteReference" } } },
            "w:footnoteRef": {},
          },
          {
            "w:t": { _: ` ${text}`, $: { "xml:space": "preserve" } },
          },
        ],
      },
    });

    obj["w:footnotes"]["w:footnote"] = footnotes;
    await this.writeFootnotes(obj);

    const run: RunInterface = {
      "w:rPr": { "w:rStyle": { $: { "w:val": "FootnoteReference" } } },
      "w:t": undefined,
      ...({"w:footnoteReference": { $: { "w:id": String(id) } }} as any),
    };

    return { id, run };
  }

  /**
   * Copy footnote ELEMENTS verbatim from another document's footnotes.xml string,
   * preserving their rich content (runs, rPr, hyperlinks). Each needed footnote is
   * appended under a fresh, collision-free id; returns source-id → new-id.
   *
   * Byte-faithful string injection (no xml2js round-trip) so footnote content can
   * never be mis-nested. NOTE: footnote-internal media (r:embed) / numbering are
   * NOT remapped here — footnotes embedding images are a rare deferred edge.
   */
  public async copyFootnotesVerbatim(
    sourceFootnotesXml: string | null,
    neededIds: Set<string>,
    rewrite?: (footnoteXml: string) => Promise<string>,
  ): Promise<Record<string, string>> {
    if (!sourceFootnotesXml || neededIds.size === 0) return {};

    // Extract the needed user footnote elements (skip separators) from the source.
    const elements: { id: string; xml: string }[] = [];
    const re = /<w:footnote\b[^>]*\bw:id="([^"]+)"[^>]*?(?:\/>|>[\s\S]*?<\/w:footnote>)/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(sourceFootnotesXml)) !== null) {
      const openTag = m[0].slice(0, m[0].indexOf(">"));
      if (/\bw:type=/.test(openTag)) continue; // separator / continuationSeparator
      if (!neededIds.has(m[1])) continue;
      elements.push({ id: m[1], xml: m[0] });
    }
    if (elements.length === 0) return {};

    await this.ensureRegistered();
    let targetXml = this.zip.readAsText(FOOTNOTES_PATH);
    if (!targetXml) return {};

    // Next id = max existing footnote id + 1 (separators -1/0 never raise it).
    let nextId = 1;
    const idRe = /<w:footnote\b[^>]*\bw:id="(-?\d+)"/g;
    let im: RegExpExecArray | null;
    while ((im = idRe.exec(targetXml)) !== null) {
      const n = parseInt(im[1], 10);
      if (n >= nextId) nextId = n + 1;
    }

    const idMap: Record<string, string> = {};
    const injected: string[] = [];
    for (const el of elements) {
      const newId = String(nextId++);
      const xml = rewrite ? await rewrite(el.xml) : el.xml;
      const gt = xml.indexOf(">");
      const open = xml.slice(0, gt).replace(/\bw:id="[^"]+"/, `w:id="${newId}"`);
      injected.push(open + xml.slice(gt));
      idMap[el.id] = newId;
    }

    const close = targetXml.lastIndexOf("</w:footnotes>");
    if (close === -1) return {};
    targetXml = targetXml.slice(0, close) + injected.join("") + targetXml.slice(close);
    this.zip.addFile(FOOTNOTES_PATH, Buffer.from(targetXml, "utf-8"));
    return idMap;
  }

  /**
   * Removes a footnote by id from footnotes.xml.
   * You must also manually remove the <w:footnoteReference> run from the document.
   */
  public async removeFootnote(id: number): Promise<void> {
    const obj = await this.readFootnotes();
    obj["w:footnotes"]["w:footnote"] = this.normalizeArray(obj).filter(
      (f) => parseInt(f.$?.["w:id"] ?? "0", 10) !== id,
    );
    await this.writeFootnotes(obj);
  }

  /**
   * Returns a run object (w:r) with a footnote reference mark for inline insertion.
   * Use after addFootnote() to get the inline anchor run.
   */
  public createFootnoteRun(footnoteId: number): RunInterface {
    return {
      "w:rPr": { "w:rStyle": { $: { "w:val": "FootnoteReference" } } },
      ...({ "w:footnoteReference": { $: { "w:id": String(footnoteId) } } } as any),
    };
  }

  private extractText(footnote: any): string {
    const p = footnote["w:p"];
    if (!p) return "";
    const runs = p["w:r"];
    if (!runs) return "";
    const arr = Array.isArray(runs) ? runs : [runs];
    return arr
      .map((r: any) => {
        const t = r["w:t"];
        if (!t) return "";
        return typeof t === "string" ? t : t._ ?? "";
      })
      .join("")
      .trim();
  }
}
