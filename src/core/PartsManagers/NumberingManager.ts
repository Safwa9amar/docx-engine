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
  /** `<w:numPicBullet>`: the picture a level's `w:lvlPicBulletId` names. */
  picBullets: any[];
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
    // A fresh root declares its namespaces, or the part is one Word refuses to read.
    if (!obj["w:numbering"].$) obj["w:numbering"].$ = { ...wordRootAttributes(), "xmlns:w": W_MAIN_NS };
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

  /** Raw `<w:numPicBullet>`, `<w:abstractNum>` + `<w:num>` arrays from this document's numbering.xml. */
  public async getRawDefinitions(): Promise<RawNumbering> {
    const obj = await this.readNumbering();
    const list = (k: string) => {
      const v = obj?.["w:numbering"]?.[k];
      return Array.isArray(v) ? v : v ? [v] : [];
    };
    return { picBullets: list("w:numPicBullet"), abstractNums: list("w:abstractNum"), nums: list("w:num") };
  }

  /** Highest existing id of each kind (for collision-free allocation). */
  public async maxIds(): Promise<{ absMax: number; numMax: number; picMax: number }> {
    const { picBullets, abstractNums, nums } = await this.getRawDefinitions();
    const max = (xs: any[], attr: string) => xs.reduce((m, x) => Math.max(m, parseInt(x?.$?.[attr] ?? "0", 10) || 0), -1);
    return {
      absMax: Math.max(0, max(abstractNums, "w:abstractNumId")),
      numMax: Math.max(0, max(nums, "w:numId")),
      picMax: max(picBullets, "w:numPicBulletId"),
    };
  }

  /**
   * Append already-id-rewritten nodes. `CT_Numbering` is a sequence (picture bullets,
   * abstractNum, num, the rest), and the builder writes keys in insertion order, so the
   * root is rebuilt in that order: a key added last would land after `w:num`.
   */
  public async appendRawDefinitions(abstractNums: any[], nums: any[], picBullets: any[] = []): Promise<void> {
    if (abstractNums.length === 0 && nums.length === 0 && picBullets.length === 0) return;
    const obj = await this.readNumbering();
    const root = obj["w:numbering"] ?? {};
    const merged = (k: string, more: any[]) => {
      const had = Array.isArray(root[k]) ? root[k] : root[k] ? [root[k]] : [];
      return [...had, ...more];
    };
    const { $, "w:numPicBullet": _pics, "w:abstractNum": _abs, "w:num": _nums, ...rest } = root;
    const pics = merged("w:numPicBullet", picBullets);
    obj["w:numbering"] = {
      ...($ ? { $ } : {}),
      ...(pics.length ? { "w:numPicBullet": pics } : {}),
      "w:abstractNum": merged("w:abstractNum", abstractNums),
      "w:num": merged("w:num", nums),
      ...rest,
    };
    await this.writeNumbering(obj);
  }
}
