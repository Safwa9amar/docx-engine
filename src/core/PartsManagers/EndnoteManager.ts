import * as XmlUtils from "@/utils/xmlUtils";
import { RelManager } from "@/core/PartsManagers/RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import { SEPARATOR_NOTE_TYPE } from "./FootnoteManager";
import { Run as RunInterface } from "@/core/files/paragraph/types";
import AdmZip from "adm-zip";
import { wordRootAttributes } from "@/utils/rootNamespaces";

const ENDNOTES_PATH = "word/endnotes.xml";
const ENDNOTE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
const ENDNOTE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const ENDNOTE_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export interface EndnoteEntry {
  id: number;
  text: string;
}

export class EndnoteManager {
  private zip: AdmZip;
  private rels: RelManager;
  private contentTypes: ContentTypesManager;

  constructor(zip: AdmZip) {
    this.zip = zip;
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
  }

  // ─── Internal read/write ──────────────────────────────────────────────────

  private async readEndnotes(): Promise<any> {
    const xml = this.zip.readAsText(ENDNOTES_PATH);
    if (!xml) return this.emptyEndnotesDoc();
    return XmlUtils.parseXml(xml);
  }

  private async writeEndnotes(obj: any): Promise<void> {
    const xml = XmlUtils.buildXml(obj["w:endnotes"], {
      rootName: "w:endnotes",
      headless: false,
      pretty: true,
    });
    this.zip.addFile(ENDNOTES_PATH, Buffer.from(xml, "utf-8"));
  }

  private emptyEndnotesDoc(): any {
    return {
      "w:endnotes": {
        // The standard Word set: endnotes copied from a modern source carry
        // r:id, w14:paraId, mc:AlternateContent… all of which need a binding.
        $: { ...wordRootAttributes(), "xmlns:w": ENDNOTE_NS },
        "w:endnote": [
          {
            $: { "w:type": "separator", "w:id": "-1" },
            "w:p": { "w:r": { "w:separator": {} } },
          },
          {
            $: { "w:type": "continuationSeparator", "w:id": "0" },
            "w:p": { "w:r": { "w:continuationSeparator": {} } },
          },
        ],
      },
    };
  }

  private normalizeArray(obj: any): any[] {
    const raw = obj?.["w:endnotes"]?.["w:endnote"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  private nextId(endnotes: any[]): number {
    const ids = endnotes
      .map((e) => parseInt(e.$?.["w:id"] ?? "0", 10))
      .filter((n) => n > 0);
    return ids.length ? Math.max(...ids) + 1 : 1;
  }

  // ─── Registration ─────────────────────────────────────────────────────────

  public async ensureRegistered(): Promise<void> {
    if (!this.zip.getEntry(ENDNOTES_PATH)) {
      await this.writeEndnotes(this.emptyEndnotesDoc());
      await this.contentTypes.addOverride(`/${ENDNOTES_PATH}`, ENDNOTE_CONTENT_TYPE);
      await this.rels.addRelationship(
        await this.rels.genId(),
        ENDNOTE_REL_TYPE,
        "endnotes.xml",
      );
    }
  }

  // ─── Public API ───────────────────────────────────────────────────────────

  /**
   * Returns all user-defined endnotes (excludes separator entries).
   */
  public async getEndnotes(): Promise<EndnoteEntry[]> {
    const obj = await this.readEndnotes();
    return this.normalizeArray(obj)
      .filter((e) => !/^(?:separator|continuationSeparator|continuationNotice)$/.test(e.$?.["w:type"] ?? ""))
      .map((e) => ({
        id: parseInt(e.$?.["w:id"] ?? "0", 10),
        text: this.extractText(e),
      }));
  }

  /**
   * Adds a new endnote to endnotes.xml.
   * @param text   The endnote body text.
   * @returns      The endnote id and a run object to insert inline.
   */
  public async addEndnote(text: string): Promise<{ id: number; run: RunInterface }> {
    await this.ensureRegistered();

    const obj = await this.readEndnotes();
    const endnotes = this.normalizeArray(obj);
    const id = this.nextId(endnotes);

    endnotes.push({
      $: { "w:id": String(id) },
      "w:p": {
        "w:pPr": { "w:pStyle": { $: { "w:val": "EndnoteText" } } },
        "w:r": [
          {
            "w:rPr": { "w:rStyle": { $: { "w:val": "EndnoteReference" } } },
            "w:endnoteRef": {},
          },
          {
            "w:t": { _: ` ${text}`, $: { "xml:space": "preserve" } },
          },
        ],
      },
    });

    obj["w:endnotes"]["w:endnote"] = endnotes;
    await this.writeEndnotes(obj);

    return { id, run: this.createEndnoteRun(id) };
  }

  /**
   * Copy endnote ELEMENTS verbatim from another document's endnotes.xml string,
   * preserving their rich content (runs, rPr, hyperlinks). Each needed endnote is
   * appended under a fresh, collision-free id; returns source-id → new-id.
   *
   * Byte-faithful string injection, the same as
   * `FootnoteManager.copyFootnotesVerbatim`: `rewrite` is where the caller fixes
   * what a note references (relationships, numbering).
   */
  public async copyEndnotesVerbatim(
    sourceEndnotesXml: string | null,
    neededIds: Set<string>,
    rewrite?: (endnoteXml: string) => Promise<string>,
  ): Promise<Record<string, string>> {
    if (!sourceEndnotesXml || neededIds.size === 0) return {};

    // Extract the needed user endnote elements (skip separators) from the source.
    const elements: { id: string; xml: string }[] = [];
    const re = /<w:endnote\b[^>]*\bw:id="([^"]+)"[^>]*?(?:\/>|>[\s\S]*?<\/w:endnote>)/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(sourceEndnotesXml)) !== null) {
      const openTag = m[0].slice(0, m[0].indexOf(">"));
      if (SEPARATOR_NOTE_TYPE.test(openTag)) continue;
      if (!neededIds.has(m[1])) continue;
      elements.push({ id: m[1], xml: m[0] });
    }
    if (elements.length === 0) return {};

    await this.ensureRegistered();
    let targetXml = this.zip.readAsText(ENDNOTES_PATH);
    if (!targetXml) return {};

    // Next id = max existing endnote id + 1 (separators -1/0 never raise it).
    let nextId = 1;
    const idRe = /<w:endnote\b[^>]*\bw:id="(-?\d+)"/g;
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

    const close = targetXml.lastIndexOf("</w:endnotes>");
    if (close === -1) return {};
    targetXml = targetXml.slice(0, close) + injected.join("") + targetXml.slice(close);
    this.zip.addFile(ENDNOTES_PATH, Buffer.from(targetXml, "utf-8"));
    return idMap;
  }

  /**
   * Removes an endnote by id.
   */
  public async removeEndnote(id: number): Promise<void> {
    const obj = await this.readEndnotes();
    obj["w:endnotes"]["w:endnote"] = this.normalizeArray(obj).filter(
      (e) => parseInt(e.$?.["w:id"] ?? "0", 10) !== id,
    );
    await this.writeEndnotes(obj);
  }

  /**
   * Returns a run object with an endnote reference mark for inline insertion.
   */
  public createEndnoteRun(endnoteId: number): RunInterface {
    return {
      "w:rPr": { "w:rStyle": { $: { "w:val": "EndnoteReference" } } },
      ...({ "w:endnoteReference": { $: { "w:id": String(endnoteId) } } } as any),
    };
  }

  private extractText(endnote: any): string {
    const p = endnote["w:p"];
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
