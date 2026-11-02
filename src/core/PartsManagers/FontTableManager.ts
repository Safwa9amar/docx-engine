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
    return [...xml.matchAll(/<w:font\b[^>]*?\bw:name="([^"]*)"[^>]*?(?:\/>|>([\s\S]*?)<\/w:font>)/g)]
      .filter((m) => m[2] !== undefined && /<w:embed(Regular|Bold|Italic|BoldItalic)\b/.test(m[2]))
      .map((m) => unescAttr(m[1]));
  }

  async embedFont(input: EmbedFontInput): Promise<EmbedFontResult> {
    const present = VARIANTS.filter((v) => input[v]);
    for (const v of present) {
      const p = embeddingPermission(readFsType(input[v] as Buffer));
      if (!canEmbed(p)) return { name: input.name, embedded: false, reason: p, variants: [], settingsUpdated: false };
    }

    // Resolve everything in memory first; nothing is written until all of it is known.
    const hadTable = !!this.zip.getEntry(FONT_TABLE);
    let table = hadTable
      ? this.read(FONT_TABLE)
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${W_NS}" xmlns:r="${R_NS}"></w:fonts>`;
    if (!/<w:fonts\b[^>]*\bxmlns:r=/.test(table)) {
      table = table.replace(/<w:fonts\b/, () => `<w:fonts xmlns:r="${R_NS}"`);
    }
    let rels = this.read(FONT_TABLE_RELS) || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}"></Relationships>`;

    const { next: tableWithoutOld, oldIds } = stripOldEmbeds(table, input.name);
    const orphanParts: string[] = [];
    for (const id of oldIds) {
      const rel = rels.match(new RegExp(`<Relationship\\b[^>]*\\bId="${escRe(id)}"[^>]*?/>`));
      if (!rel) continue;
      const target = rel[0].match(/\bTarget="([^"]*)"/);
      if (target) orphanParts.push(`word/${target[1]}`);
      rels = rels.replace(rel[0], () => "");
    }

    const taken = new Set<string>();
    const parts: Array<{ path: string; bytes: Buffer }> = [];
    const embeds: string[] = [];
    for (const v of present) {
      const key = this.newKey();
      let k = 1;
      while ((this.zip.getEntry(`word/fonts/font${k}.odttf`) && !orphanParts.includes(`word/fonts/font${k}.odttf`)) || taken.has(`font${k}`)) k++;
      taken.add(`font${k}`);
      const target = `fonts/font${k}.odttf`;
      parts.push({ path: `word/${target}`, bytes: obfuscateFont(input[v] as Buffer, key) });
      const added = appendRel(rels, FONT_REL, target);
      rels = added.xml;
      embeds.push(`<w:${EMBED_TAG[v]} r:id="${added.id}" w:fontKey="${key}"/>`);
    }
    table = insertFontEntry(tableWithoutOld, input.name, embeds.join(""));
    const settings = this.resolveSettings();

    // Writes.
    for (const path of orphanParts) {
      if (!parts.some((x) => x.path === path)) this.zip.deleteFile(path);
    }
    for (const part of parts) this.zip.addFile(part.path, part.bytes);
    this.zip.addFile(FONT_TABLE_RELS, Buffer.from(rels, "utf-8"));
    this.zip.addFile(FONT_TABLE, Buffer.from(table, "utf-8"));
    if (!hadTable) {
      await this.contentTypes.addOverride(`/${FONT_TABLE}`, FONT_TABLE_CT);
      const docRels = this.read(DOC_RELS);
      if (!/Target="fontTable\.xml"/.test(docRels)) {
        this.zip.addFile(DOC_RELS, Buffer.from(appendRel(docRels, FONT_TABLE_REL, "fontTable.xml").xml, "utf-8"));
      }
    }
    await this.contentTypes.addDefault("odttf", ODTTF_CT);
    if (settings.xml !== null) this.zip.addFile(SETTINGS, Buffer.from(settings.xml, "utf-8"));
    return { name: input.name, embedded: true, variants: present, settingsUpdated: settings.ok };
  }

  /** The settings.xml text with w:embedTrueTypeFonts switched on (null = leave untouched). */
  private resolveSettings(): { ok: boolean; xml: string | null } {
    const xml = this.read(SETTINGS);
    const existing = xml.match(/<w:embedTrueTypeFonts\b[^>]*?(?:\/>|>[\s\S]*?<\/w:embedTrueTypeFonts>)/);
    if (existing) {
      const off = /\bw:val="(0|false|off)"/.test(existing[0]);
      return { ok: true, xml: off ? xml.replace(existing[0], () => "<w:embedTrueTypeFonts/>") : null };
    }
    const open = xml.match(/<w:settings\b[^>]*>/);
    if (!open) return { ok: false, xml: null };
    let at = (open.index as number) + open[0].length;
    for (const tag of SETTINGS_BEFORE) {
      const re = new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`, "g");
      for (const m of xml.matchAll(re)) at = Math.max(at, (m.index as number) + m[0].length);
    }
    return { ok: true, xml: `${xml.slice(0, at)}<w:embedTrueTypeFonts/>${xml.slice(at)}` };
  }
}

const EMBED_ANY = /<w:embed(Regular|Bold|Italic|BoldItalic)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:embed\1>)/g;

/** Removes the named font's existing embed elements; returns their relationship ids. */
function stripOldEmbeds(xml: string, name: string): { next: string; oldIds: string[] } {
  const m = xml.match(new RegExp(`<w:font\\b[^>]*\\bw:name="${escRe(escAttr(name))}"[^>]*?(/?)>`));
  if (!m || m[1] === "/") return { next: xml, oldIds: [] };
  const start = (m.index as number) + m[0].length;
  const end = xml.indexOf("</w:font>", start);
  const oldIds: string[] = [];
  const body = xml.slice(start, end).replace(EMBED_ANY, (all) => {
    const id = all.match(/\br:id="([^"]*)"/);
    if (id) oldIds.push(id[1]);
    return "";
  });
  return { next: xml.slice(0, start) + body + xml.slice(end), oldIds };
}

/** Appends the embeds to the named w:font (creating the entry when absent). */
function insertFontEntry(xml: string, name: string, embeds: string): string {
  const m = xml.match(new RegExp(`<w:font\\b[^>]*\\bw:name="${escRe(escAttr(name))}"[^>]*?(/?)>`));
  if (!m) return xml.replace("</w:fonts>", () => `<w:font w:name="${escAttr(name)}">${embeds}</w:font></w:fonts>`);
  if (m[1] === "/") return xml.replace(m[0], () => `${m[0].slice(0, -2)}>${embeds}</w:font>`);
  const end = xml.indexOf("</w:font>", (m.index as number) + m[0].length);
  return xml.slice(0, end) + embeds + xml.slice(end);
}

/** Appends a Relationship with the next free rIdN. */
export function appendRel(relsXml: string, type: string, target: string): { xml: string; id: string } {
  const ids = [...relsXml.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
  const rel = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  return { xml: relsXml.replace("</Relationships>", () => `${rel}</Relationships>`), id };
}
