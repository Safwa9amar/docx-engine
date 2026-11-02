import { wordRootAttributes } from "@/utils/rootNamespaces";
import * as XmlUtils from "@/utils/xmlUtils";
import Paragraph from "@/core/files/paragraph/index";
import { RelManager } from "@/core/PartsManagers/RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import AdmZip from "adm-zip";

const NUMBERING_PATH = "word/numbering.xml";
const NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

export interface NumberingDefinition {
  abstractNumId: string;
  numId: string;
}

/** Raw parsed numbering parts (xml2js objects), for cross-document copy. */
export interface RawNumbering {
  abstractNums: any[];
  nums: any[];
}

const W_MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export class NumberingManager {
  private zip: AdmZip;

  constructor(zip: AdmZip) {
    this.zip = zip;
  }

  private async readNumbering(): Promise<any> {
    const xml = this.zip.readAsText(NUMBERING_PATH);
    if (!xml) return { "w:numbering": {} };
    return XmlUtils.parseXml(xml);
  }

  private async writeNumbering(obj: any): Promise<void> {
    const isNew = !this.zip.getEntry(NUMBERING_PATH);
    const xml = XmlUtils.buildXml(obj["w:numbering"], {
      rootName: "w:numbering",
      headless: false,
      pretty: true,
    });
    this.zip.addFile(NUMBERING_PATH, Buffer.from(xml, "utf-8"));
    // A part with no relationship and no content type is one Word never reads:
    // every list pointing into it would lose its bullet or number.
    if (isNew) {
      const rels = new RelManager(this.zip);
      await new ContentTypesManager(this.zip).addOverride(`/${NUMBERING_PATH}`, NUMBERING_CONTENT_TYPE);
      await rels.addRelationship(await rels.genId(), NUMBERING_REL_TYPE, "numbering.xml");
    }
  }

  /**
   * Returns all concrete numbering definitions (numId → abstractNumId).
   */
  public async getNumberingDefinitions(): Promise<NumberingDefinition[]> {
    const obj = await this.readNumbering();
    const raw = obj?.["w:numbering"]?.["w:num"];
    if (!raw) return [];
    const arr = Array.isArray(raw) ? raw : [raw];

    return arr.map((n: any) => ({
      numId: n.$?.["w:numId"] ?? "",
      abstractNumId: n["w:abstractNumId"]?.$?.["w:val"] ?? "",
    }));
  }

  /**
   * Applies numbering to a paragraph by setting <w:numPr> on its <w:pPr>.
   * @param paragraph  The Paragraph instance to modify.
   * @param numId      The concrete numId from numbering.xml.
   * @param ilvl       Indent level (0-based). Defaults to 0.
   */
  public applyNumbering(paragraph: Paragraph, numId: string, ilvl = 0): void {
    if (!paragraph.paragraph["w:pPr"]) {
      paragraph.paragraph["w:pPr"] = {};
    }
    paragraph.paragraph["w:pPr"]["w:numPr"] = {
      "w:ilvl": { $: { "w:val": String(ilvl) } },
      "w:numId": { $: { "w:val": numId } },
    };
  }

  /**
   * Removes numbering from a paragraph.
   */
  public removeNumbering(paragraph: Paragraph): void {
    if (paragraph.paragraph["w:pPr"]) {
      delete paragraph.paragraph["w:pPr"]["w:numPr"];
    }
  }

  /**
   * Adds a new concrete <w:num> definition referencing an abstract numbering ID.
   * @param numId         The new numId to use.
   * @param abstractNumId The abstractNumId to reference.
   */
  public async addNumberingDefinition(numId: string, abstractNumId: string): Promise<void> {
    const obj = await this.readNumbering();
    const numbering = obj["w:numbering"];

    if (!numbering["w:num"]) {
      numbering["w:num"] = [];
    } else if (!Array.isArray(numbering["w:num"])) {
      numbering["w:num"] = [numbering["w:num"]];
    }

    numbering["w:num"].push({
      $: { "w:numId": numId },
      "w:abstractNumId": { $: { "w:val": abstractNumId } },
    });

    await this.writeNumbering(obj);
  }

  // ── Cross-document copy primitives (used by MergeManager) ──────────────────

  /** Raw `<w:abstractNum>` + `<w:num>` arrays from this document's numbering.xml. */
  public async getRawDefinitions(): Promise<RawNumbering> {
    const obj = await this.readNumbering();
    const a = obj?.["w:numbering"]?.["w:abstractNum"];
    const n = obj?.["w:numbering"]?.["w:num"];
    return {
      abstractNums: Array.isArray(a) ? a : a ? [a] : [],
      nums: Array.isArray(n) ? n : n ? [n] : [],
    };
  }

  /** Highest existing abstractNumId / numId (for collision-free allocation). */
  public async maxIds(): Promise<{ absMax: number; numMax: number }> {
    const { abstractNums, nums } = await this.getRawDefinitions();
    const absMax = abstractNums.reduce(
      (m, x) => Math.max(m, parseInt(x?.$?.["w:abstractNumId"] ?? "0", 10) || 0),
      0,
    );
    const numMax = nums.reduce((m, x) => Math.max(m, parseInt(x?.$?.["w:numId"] ?? "0", 10) || 0), 0);
    return { absMax, numMax };
  }

  /**
   * Append already-id-rewritten `<w:abstractNum>` + `<w:num>` nodes. Keeps the
   * schema order (all abstractNum before all num) because they are separate keys.
   */
  public async appendRawDefinitions(abstractNums: any[], nums: any[]): Promise<void> {
    if (abstractNums.length === 0 && nums.length === 0) return;
    const obj = await this.readNumbering();
    if (!obj["w:numbering"]) obj["w:numbering"] = {};
    const root = obj["w:numbering"];
    if (!root.$) root.$ = { ...wordRootAttributes(), "xmlns:w": W_MAIN_NS };
    const ensureArr = (k: string) => {
      root[k] = Array.isArray(root[k]) ? root[k] : root[k] ? [root[k]] : [];
    };
    ensureArr("w:abstractNum");
    ensureArr("w:num");
    root["w:abstractNum"].push(...abstractNums);
    root["w:num"].push(...nums);
    await this.writeNumbering(obj);
  }
}
