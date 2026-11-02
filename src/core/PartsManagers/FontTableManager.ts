// src/core/PartsManagers/FontTableManager.ts
// Embeds TrueType fonts into a .docx the way Word does: an obfuscated
// word/fonts/fontN.odttf per variant, a relationship from fontTable.xml, a
// w:embedRegular|Bold|Italic|BoldItalic element on the font's w:font entry, the
// odttf Default content type, and w:embedTrueTypeFonts in settings.xml.
// Fonts whose licence forbids embedding are refused, and nothing is written.
import type { ZipManager } from "../../utils/ZipManager";
import type { ContentTypesManager } from "./ContentTypesManager";
import { readFsType, embeddingPermission, canEmbed, type EmbeddingPermission } from "../fonts/fsType";
import { obfuscateFont, newFontKey } from "../fonts/obfuscate";

const FONT_TABLE = "word/fontTable.xml";
const FONT_TABLE_RELS = "word/_rels/fontTable.xml.rels";
const DOC_RELS = "word/_rels/document.xml.rels";
const SETTINGS = "word/settings.xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const FONT_REL = `${R_NS}/font`;
const FONT_TABLE_REL = `${R_NS}/fontTable`;
const ODTTF_CT = "application/vnd.openxmlformats-officedocument.obfuscatedFont";
const FONT_TABLE_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export type FontVariant = "regular" | "bold" | "italic" | "boldItalic";
const VARIANTS: FontVariant[] = ["regular", "bold", "italic", "boldItalic"];
const EMBED_TAG: Record<FontVariant, string> = {
  regular: "embedRegular", bold: "embedBold", italic: "embedItalic", boldItalic: "embedBoldItalic",
};
/** Children of w:settings that must come BEFORE w:embedTrueTypeFonts (CT_Settings). */
const SETTINGS_BEFORE = [
  "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
  "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
  "printFractionalCharacterWidth", "printFormsData",
];

export interface EmbedFontInput { name: string; regular: Buffer; bold?: Buffer; italic?: Buffer; boldItalic?: Buffer }
export interface EmbedFontResult {
  name: string;
  embedded: boolean;
  reason?: EmbeddingPermission;
  variants: FontVariant[];
  settingsUpdated: boolean;
}

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const unescAttr = (s: string) => s.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&amp;/g, "&");
const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

export class FontTableManager {
  /** Replaceable in tests so font keys are deterministic. */
  newKey: () => string = () => newFontKey();

  constructor(private zip: ZipManager, private contentTypes: ContentTypesManager) {}

  /** A part's text, or "" when the part does not exist. */
  private read(path: string): string {
    const entry = this.zip.getEntry(path);
    return entry ? entry.getData().toString("utf-8") : "";
  }

  /** Every w:font name in fontTable.xml, in document order. */
  listFonts(): string[] {
    const xml = this.read(FONT_TABLE);
    return [...xml.matchAll(/<w:font\b[^>]*\bw:name="([^"]*)"/g)].map((m) => unescAttr(m[1]));
  }

  /** Names of fonts that carry at least one embedded variant. */
  embeddedFonts(): string[] {
    const xml = this.read(FONT_TABLE);
    return [...xml.matchAll(/<w:font\b[^>]*\bw:name="([^"]*)"[^>]*>([\s\S]*?)<\/w:font>/g)]
      .filter((m) => /<w:embed(Regular|Bold|Italic|BoldItalic)\b/.test(m[2]))
      .map((m) => unescAttr(m[1]));
  }

  async embedFont(input: EmbedFontInput): Promise<EmbedFontResult> {
    const present = VARIANTS.filter((v) => input[v]);
    for (const v of present) {
      const p = embeddingPermission(readFsType(input[v] as Buffer));
      if (!canEmbed(p)) return { name: input.name, embedded: false, reason: p, variants: [], settingsUpdated: false };
    }
    await this.ensureFontTable();
    const embeds: string[] = [];
    for (const v of present) {
      const key = this.newKey();
      const target = this.addFontPart(obfuscateFont(input[v] as Buffer, key));
      const relId = this.addFontRel(target);
      embeds.push(`<w:${EMBED_TAG[v]} r:id="${relId}" w:fontKey="${key}"/>`);
    }
    await this.contentTypes.addDefault("odttf", ODTTF_CT);
    this.writeFontEntry(input.name, embeds.join(""));
    const settingsUpdated = this.ensureEmbedSetting();
    return { name: input.name, embedded: true, variants: present, settingsUpdated };
  }

  private async ensureFontTable(): Promise<void> {
    if (!this.zip.getEntry(FONT_TABLE)) {
      this.zip.addFile(FONT_TABLE, Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${W_NS}" xmlns:r="${R_NS}"></w:fonts>`, "utf-8"));
      await this.contentTypes.addOverride(`/${FONT_TABLE}`, FONT_TABLE_CT);
      const rels = this.read(DOC_RELS);
      if (!/Target="fontTable\.xml"/.test(rels)) {
        this.zip.addFile(DOC_RELS, Buffer.from(appendRel(rels, FONT_TABLE_REL, "fontTable.xml").xml, "utf-8"));
      }
      return;
    }
    const xml = this.read(FONT_TABLE);
    if (!/<w:fonts\b[^>]*\bxmlns:r=/.test(xml)) {
      this.zip.addFile(FONT_TABLE, Buffer.from(xml.replace(/<w:fonts\b/, `<w:fonts xmlns:r="${R_NS}"`), "utf-8"));
    }
  }

  private addFontPart(bytes: Buffer): string {
    let k = 1;
    while (this.zip.getEntry(`word/fonts/font${k}.odttf`)) k++;
    this.zip.addFile(`word/fonts/font${k}.odttf`, bytes);
    return `fonts/font${k}.odttf`;
  }

  private addFontRel(target: string): string {
    const rels = this.read(FONT_TABLE_RELS) || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"></Relationships>`;
    const { xml, id } = appendRel(rels, FONT_REL, target);
    this.zip.addFile(FONT_TABLE_RELS, Buffer.from(xml, "utf-8"));
    return id;
  }

  private writeFontEntry(name: string, embeds: string): void {
    const xml = this.read(FONT_TABLE);
    const n = escRe(escAttr(name));
    const open = new RegExp(`<w:font\\b[^>]*\\bw:name="${n}"[^>]*?(/?)>`);
    const m = xml.match(open);
    let next: string;
    if (!m) {
      next = xml.replace("</w:fonts>", `<w:font w:name="${escAttr(name)}">${embeds}</w:font></w:fonts>`);
    } else if (m[1] === "/") {
      next = xml.replace(m[0], `${m[0].slice(0, -2)}>${embeds}</w:font>`);
    } else {
      const start = (m.index as number) + m[0].length;
      const end = xml.indexOf("</w:font>", start);
      const body = xml.slice(start, end).replace(/<w:embed(Regular|Bold|Italic|BoldItalic)\b[^>]*\/>/g, "");
      next = xml.slice(0, start) + body + embeds + xml.slice(end);
    }
    this.zip.addFile(FONT_TABLE, Buffer.from(next, "utf-8"));
  }

  /** Adds w:embedTrueTypeFonts in schema position. False when there is no settings part. */
  private ensureEmbedSetting(): boolean {
    const xml = this.read(SETTINGS);
    if (!xml) return false;
    if (xml.includes("<w:embedTrueTypeFonts")) return true;
    let at = -1;
    for (const tag of SETTINGS_BEFORE) {
      const re = new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`, "g");
      for (const m of xml.matchAll(re)) at = Math.max(at, (m.index as number) + m[0].length);
    }
    if (at < 0) {
      const open = xml.match(/<w:settings\b[^>]*>/) as RegExpMatchArray;
      at = (open.index as number) + open[0].length;
    }
    this.zip.addFile(SETTINGS, Buffer.from(`${xml.slice(0, at)}<w:embedTrueTypeFonts/>${xml.slice(at)}`, "utf-8"));
    return true;
  }
}

/** Appends a Relationship with the next free rIdN. */
export function appendRel(relsXml: string, type: string, target: string): { xml: string; id: string } {
  const ids = [...relsXml.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
  const rel = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  return { xml: relsXml.replace("</Relationships>", `${rel}</Relationships>`), id };
}
