import { StylesManager } from "@/core/PartsManagers/StylesManager";
import DocumentManager from "@/core/PartsManagers/DocumentManager";
import {
  escapeXmlText,
  decodeXmlText,
  stripInvalidXmlChars,
  paragraphText,
  paragraphHeadingLevel,
  paragraphPropsXml,
  paragraphRunPropsXml,
  type BodyBlock,
} from "@/core/files/body/OrderedBody";
import AdmZip from "adm-zip";

const DOCUMENT_PATH = "word/document.xml";
const CAPTION_STYLE_ID = "Caption";

/** Word's "Table of Figures" paragraph style for list entries. */
const TABLE_OF_FIGURES_STYLE = {
  $: { "w:type": "paragraph", "w:styleId": "TableofFigures" },
  "w:name":       { $: { "w:val": "table of figures" } },
  "w:basedOn":    { $: { "w:val": "Normal" } },
  "w:next":       { $: { "w:val": "Normal" } },
  "w:uiPriority": { $: { "w:val": "99" } },
  "w:unhideWhenUsed": {},
  // `w:pPr` is an ORDERED sequence — tabs (11) before spacing (22).
  "w:pPr": {
    "w:tabs": {
      "w:tab": { $: { "w:val": "right", "w:leader": "dot", "w:pos": "8748" } },
    },
    "w:spacing": { $: { "w:before": "0", "w:after": "200" } },
  },
  "w:rPr": {
    "w:sz":   { $: { "w:val": "20" } },
    "w:szCs": { $: { "w:val": "20" } },
  },
};

// ─── Style definitions ────────────────────────────────────────────────────────

/**
 * Word's built-in "Caption" paragraph style.
 * Bold, 9 pt, dark blue-grey (#44546A), 6 pt spacing before and after.
 *
 * Deliberately carries NO `w:jc`: Word's own Caption style doesn't, and forcing
 * `left` would mis-align every caption in an RTL thesis.
 */
export const CAPTION_PARA_STYLE = {
  $: { "w:type": "paragraph", "w:styleId": CAPTION_STYLE_ID },
  "w:name":      { $: { "w:val": "caption" } },
  "w:basedOn":   { $: { "w:val": "Normal" } },
  "w:next":      { $: { "w:val": "Normal" } },
  "w:uiPriority":{ $: { "w:val": "99" } },
  "w:qFormat":   {},
  "w:pPr": {
    "w:spacing": { $: { "w:before": "120", "w:after": "120" } },
  },
  "w:rPr": {
    "w:b":    {},
    "w:bCs":  {},
    "w:color":{ $: { "w:val": "44546A", "w:themeColor": "text2" } },
    "w:sz":   { $: { "w:val": "18" } },
    "w:szCs": { $: { "w:val": "18" } },
  },
};

/**
 * "CaptionChar" character style — kept registered for documents that already
 * reference it (older captions this engine wrote used it on every run).
 */
const CAPTION_CHAR_STYLE = {
  $: { "w:type": "character", "w:styleId": "CaptionChar", "w:customStyle": "1" },
  "w:name":      { $: { "w:val": "caption char" } },
  "w:basedOn":   { $: { "w:val": "DefaultParagraphFont" } },
  "w:uiPriority":{ $: { "w:val": "99" } },
  "w:rPr": {
    "w:b":    {},
    "w:bCs":  {},
    "w:color":{ $: { "w:val": "44546A", "w:themeColor": "text2" } },
    "w:sz":   { $: { "w:val": "18" } },
    "w:szCs": { $: { "w:val": "18" } },
  },
};

// ─── Types ────────────────────────────────────────────────────────────────────

/** Numbering format matching Word's Caption Numbering dialog */
export type CaptionNumberFormat =
  | "arabic"       // 1, 2, 3, …       (default)
  | "ROMAN"        // I, II, III, …
  | "roman"        // i, ii, iii, …
  | "ALPHABETIC"   // A, B, C, …
  | "alphabetic";  // a, b, c, …

/** Chapter-number separator matching Word's "Use separator" dropdown */
export type ChapterSeparator = "-" | "." | ":" | "–" | "—"; // hyphen, period, colon, en-dash, em-dash

export type CaptionPosition = "below" | "above";

export interface CaptionNumberingOptions {
  /** Numbering format (default: "arabic") */
  format?: CaptionNumberFormat;
  /** Include the chapter number (e.g. "Figure 1-1") */
  includeChapterNumber?: boolean;
  /** Heading style that starts each chapter (default: "Heading1") */
  chapterStyle?: string;
  /** Separator between chapter number and caption number (default: "-") */
  chapterSeparator?: ChapterSeparator;
  /**
   * Format of the CHAPTER half of the number — "Figure I-1" instead of
   * "Figure 1-1" (default: follow the heading's own numbering).
   *
   * Word's `STYLEREF n \s` echoes whatever the chapter heading's list format
   * renders, so a thesis whose Heading 1 is numbered "Chapitre I" already reads
   * "I-1". A thesis whose headings are plain text or arabic-numbered does not,
   * and Word's Caption dialog offers no way to fix that — the `\*` switch does:
   * ` STYLEREF 1 \s \* ROMAN ` converts the result on the fly.
   */
  chapterFormat?: CaptionNumberFormat;
}

export interface CaptionOptions {
  /** Built-in or custom label: "Figure", "Table", "Equation", or any string */
  label?: string;
  /** Caption body text after the label + number */
  text?: string;
  /** Where to insert relative to the block at insertIndex (default: "below") */
  position?: CaptionPosition;
  /** Omit the label prefix, e.g. just "1" instead of "Figure 1" */
  excludeLabel?: boolean;
  /** Numbering configuration */
  numbering?: CaptionNumberingOptions;
  /** Write the caption right-to-left (Arabic thesis): `w:bidi` + `w:rtl`. */
  rtl?: boolean;
  /**
   * What sits between the number and the text (default: a single space, the way
   * Word's own dialog writes it). Conversion passes the separator the thesis
   * already types — " : " in French, " - ", ". " — so a converted caption still
   * READS exactly as the student wrote it.
   */
  textSeparator?: string;
}

export interface ConvertTextCaptionsOptions {
  /** First block to scan (default 0). */
  fromIndex?: number;
  /** Last block to scan, inclusive (default: the last block). */
  toIndex?: number;
  /** Restrict to one kind; "all" (the default) converts figures, tables and equations. */
  kind?: CaptionKind | "all";
  /** Force the label to write, e.g. "Tableau". Only honoured for a single `kind`. */
  label?: string;
  /**
   * Numbering, as in Word's dialog. Omit for plain sequential numbers
   * ("Figure 1", "Figure 2", …); set `includeChapterNumber` for per-chapter
   * numbering that restarts at each chapter ("Figure I-1", "Figure II-1", …).
   */
  numbering?: CaptionNumberingOptions;
  /** Force RTL/LTR; omit to follow each paragraph's own direction. */
  rtl?: boolean;
  /**
   * What to write between the number and the wording of EVERY converted caption
   * (": ", " : "). Omit it and each caption keeps the punctuation its author
   * typed, which is the right default for a tool editing one caption; a caller
   * that rebuilds a whole thesis passes it so the list of tables does not show
   * "Tableau 1 : …" beside "Tableau 2 - …".
   */
  textSeparator?: string;
  /** Report what WOULD convert without writing anything. */
  dryRun?: boolean;
}

/** Why a paragraph that reads like a caption was left alone. */
export type SkipReason =
  | "already-a-caption"
  | "is-a-heading"
  | "contains-image"
  | "contains-field";

export interface ConvertedCaption {
  blockIndex: number;
  kind: CaptionKind;
  /** The label written, after unification (e.g. "Figure"). */
  label: string;
  /** The paragraph's text before conversion. */
  before: string;
  /** Its text after — the same wording, now with a live number. */
  after: string;
  /** The number the student had typed ("" if none); the live field replaces it. */
  typedNumber: string;
  /** The live number now shown by the field. */
  number: string;
}

export interface SkippedTextCaption {
  blockIndex: number;
  reason: SkipReason;
  text: string;
}

export interface ConvertTextCaptionsResult {
  converted: ConvertedCaption[];
  skipped: SkippedTextCaption[];
  /** The label chosen per kind. */
  labels: Partial<Record<CaptionKind, string>>;
  /** True when nothing was written. */
  dryRun: boolean;
}

/** Why {@link CaptionManager.addGeneratedCaption} left a block alone. */
export type GeneratedCaptionRefusal =
  | "no-such-block"
  | "not-a-paragraph"
  | "already-a-caption"
  | "is-a-heading"
  | "empty-text"
  | "contains-image";

const DEFAULT_LABEL: Record<CaptionKind, string> = { figure: "Figure", table: "Table", equation: "Equation" };

export interface CaptionEntry {
  /** The SEQ identifier the caption is numbered under (e.g. "Figure"). */
  label: string;
  /** The label as it READS in the document (e.g. "الشكل رقم"), "" if excluded. */
  displayLabel: string;
  /** The number currently shown by the field (Word recomputes on update). */
  number: string;
  /** The caption text after the label and number. */
  text: string;
  /** The caption's full visible text ("Figure 1 Site plan"). */
  fullText: string;
  /** Editable BLOCK index (the same index space as `document.getBlocks()`). */
  blockIndex: number;
  /** @deprecated Use {@link blockIndex}. Kept as an alias for older callers. */
  paragraphIndex: number;
  /** Bookmark anchoring this caption, if it has one. */
  bookmark: string | null;
  /**
   * True when the caption lives INSIDE the block at `blockIndex` (a table cell)
   * rather than being that block itself — `blockIndex` then addresses the table.
   */
  inTable: boolean;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

const FORMAT_MAP: Record<CaptionNumberFormat, string> = {
  arabic:     "ARABIC",
  ROMAN:      "ROMAN",
  roman:      "roman",
  ALPHABETIC: "ALPHABETIC",
  alphabetic: "alphabetic",
};

/**
 * The SEQ identifier for a caption label. A SEQ name is bookmark-like, so it
 * cannot contain whitespace — "الشكل رقم" numbers under "الشكل_رقم" while the
 * document still READS "الشكل رقم". The Table-of-Figures `\c` switch uses the
 * same identifier, so the two always agree.
 */
export function seqIdentifier(label: string): string {
  return label.trim().replace(/\s+/g, "_");
}

/**
 * Escape for ELEMENT TEXT (`<w:t>`, `<w:instrText>`), where `"` and `'` are
 * legal literals — Word writes `TOC \c "Figure"` with a bare quote, and
 * `escapeXmlText`'s attribute-grade escaping would turn it into `&quot;`.
 * Attribute values (`w:instr="…"`) still use the full escape.
 */
function escapeXmlContent(text: string): string {
  return stripInvalidXmlChars(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const ROMAN_PAIRS: Array<[number, string]> = [
  [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
  [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
];

function toRoman(n: number): string {
  if (n <= 0) return String(n);
  let out = "";
  let rest = n;
  for (const [value, sym] of ROMAN_PAIRS) {
    while (rest >= value) { out += sym; rest -= value; }
  }
  return out;
}

/** Word's `\* ALPHABETIC`: A…Z, then AA, BB, CC … (the letter repeats). */
function toAlpha(n: number): string {
  if (n <= 0) return String(n);
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

/** Render a sequence number the way Word's `\*` format switch would. */
export function formatSeqNumber(n: number, fmt: string): string {
  switch (fmt) {
    case "ROMAN":      return toRoman(n);
    case "roman":      return toRoman(n).toLowerCase();
    case "ALPHABETIC": return toAlpha(n);
    case "alphabetic": return toAlpha(n).toLowerCase();
    default:           return String(n); // ARABIC / CARDTEXT / unknown
  }
}

/** "Heading1" / "Heading 1" / "Titre2" → 1..9 (default 1). */
function chapterLevel(chapterStyle?: string): number {
  const n = parseInt(String(chapterStyle ?? "").replace(/\D/g, ""), 10);
  return n >= 1 && n <= 9 ? n : 1;
}

// ─── Field scanning (string-scoped, byte-safe) ────────────────────────────────

interface FieldHit {
  /** The field code, e.g. ` SEQ Figure \* ARABIC `. */
  instr: string;
  /** Index of the first character of the whole field. */
  start: number;
  /** Index just past the field's last character. */
  end: number;
  /** Index where the field RESULT begins (-1 when the field has no result). */
  resultStart: number;
  /** Index just past the field result. */
  resultEnd: number;
}

/**
 * Every field in a paragraph, in document order — both forms Word writes:
 * `<w:fldSimple w:instr="…">result</w:fldSimple>` and the run-split complex form
 * (`fldChar begin` → `instrText` → `fldChar separate` → result → `fldChar end`).
 * Operates on the paragraph's own XML substring only, so it can never touch a
 * sibling block.
 */
function findFields(xml: string): FieldHit[] {
  const hits: FieldHit[] = [];

  // Simple fields
  const simpleRe = /<w:fldSimple\b[^>]*?>/g;
  let m: RegExpExecArray | null;
  while ((m = simpleRe.exec(xml)) !== null) {
    const openTag = m[0];
    const instrM = /\bw:instr="([^"]*)"/.exec(openTag);
    const instr = instrM ? decodeXmlText(instrM[1]) : "";
    if (openTag.endsWith("/>")) {
      hits.push({ instr, start: m.index, end: m.index + openTag.length, resultStart: -1, resultEnd: -1 });
      continue;
    }
    const bodyStart = m.index + openTag.length;
    const closeIdx = xml.indexOf("</w:fldSimple>", bodyStart);
    if (closeIdx === -1) continue;
    hits.push({
      instr,
      start: m.index,
      end: closeIdx + "</w:fldSimple>".length,
      resultStart: bodyStart,
      resultEnd: closeIdx,
    });
  }

  // Complex fields — a small state machine over fldChar / instrText tokens.
  const tokenRe =
    /<w:fldChar\b[^>]*?w:fldCharType="(begin|separate|end)"[^>]*?>|<w:instrText\b[^>]*?>([\s\S]*?)<\/w:instrText>|<w:instrText\b[^>]*?\/>/g;
  const stack: Array<{ instr: string; start: number; sepEnd: number | null }> = [];
  let t: RegExpExecArray | null;
  while ((t = tokenRe.exec(xml)) !== null) {
    const kind = t[1];
    if (kind === "begin") {
      stack.push({ instr: "", start: t.index, sepEnd: null });
    } else if (kind === "separate") {
      if (stack.length) stack[stack.length - 1].sepEnd = t.index + t[0].length;
    } else if (kind === "end") {
      const f = stack.pop();
      if (!f) continue;
      hits.push({
        instr: decodeXmlText(f.instr),
        start: f.start,
        end: t.index + t[0].length,
        resultStart: f.sepEnd ?? -1,
        resultEnd: f.sepEnd == null ? -1 : t.index,
      });
    } else if (t[2] !== undefined && stack.length) {
      stack[stack.length - 1].instr += t[2];
    }
  }

  return hits.sort((a, b) => a.start - b.start);
}

/**
 * Every outermost `<w:p>…</w:p>` inside a block's XML, in document order.
 *
 * A body block is not always a paragraph: theses routinely park a figure and its
 * caption inside a one-cell table to centre them, so caption paragraphs live
 * inside `w:tbl` blocks too. `<w:pPr>` / `<w:pict>` are not matched — `\b` needs
 * a non-word character after the `p`.
 */
function eachParagraph(xml: string): Array<{ start: number; end: number; xml: string }> {
  const out: Array<{ start: number; end: number; xml: string }> = [];
  const re = /<w:p\b[^>]*?(\/?)>|<\/w:p>/g;
  const stack: number[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[0] === "</w:p>") {
      const start = stack.pop();
      if (start !== undefined && stack.length === 0) {
        out.push({ start, end: re.lastIndex, xml: xml.slice(start, re.lastIndex) });
      }
    } else if (m[1] !== "/") {
      stack.push(m.index);
    }
  }
  return out;
}

const TEXT_ELEMENT_RE = /<w:t(?:\s[^>]*?)?>[\s\S]*?<\/w:t>|<w:t(?:\s[^>]*?)?\/>/;

/**
 * Replace the FIRST `<w:t>` inside a field-result region with `text`.
 * Returns null when the region holds no text element — a dirty field Word will
 * fill in on update; rewriting it would mean inventing runs.
 */
function replaceFirstText(region: string, text: string): string | null {
  const m = TEXT_ELEMENT_RE.exec(region);
  if (!m) return null;
  return (
    region.slice(0, m.index) +
    `<w:t xml:space="preserve">${escapeXmlContent(text)}</w:t>` +
    region.slice(m.index + m[0].length)
  );
}

/** ` SEQ Figure \* ROMAN \s 1 ` → { id, format, resetLevel } */
function parseSeqInstr(instr: string): { id: string; format: string; resetLevel: number } | null {
  const m = /^\s*SEQ\s+("[^"]+"|\S+)/.exec(instr);
  if (!m) return null;
  const fmtM = /\\\*\s*(\w+)/.exec(instr);
  const resetM = /\\s\s*(\d)/.exec(instr);
  return {
    id: m[1].replace(/^"|"$/g, ""),
    format: fmtM ? fmtM[1] : "ARABIC",
    resetLevel: resetM ? Number(resetM[1]) : 0,
  };
}

/** ` STYLEREF 1 \s \* ROMAN ` → { level, format }; null when not a STYLEREF. */
function parseStyleRef(instr: string): { level: number; format: string } | null {
  const m = /^\s*STYLEREF\s+("[^"]+"|\S+)/.exec(instr);
  if (!m) return null;
  const fmtM = /\\\*\s*(\w+)/.exec(instr);
  return {
    level: chapterLevel(m[1].replace(/^"|"$/g, "")),
    format: fmtM ? fmtM[1] : "ARABIC",
  };
}

// ─── Plain-text caption detection ─────────────────────────────────────────────
//
// A thesis that was never written in Word's Caption dialog still HAS captions —
// they are ordinary paragraphs the student typed ("Figure 1 : Organigramme de
// l'entreprise", "الجدول رقم 3: توزيع العينة"). Converting them into real
// captions is the only way a List of Figures, a cross-reference or automatic
// renumbering can ever work, so the detector has to recognise how the three
// languages of the corpus actually write one.

export type CaptionKind = "figure" | "table" | "equation";

/**
 * Label spellings that plainly open a caption, per kind, in en/fr/ar.
 * Matched longest-first, so "Tableau" wins over "Table" and "الشكل رقم" over
 * "الشكل". A multi-word entry matches any run of whitespace between its words.
 */
const KIND_LABELS: Record<CaptionKind, string[]> = {
  figure: [
    "figure", "fig.", "fig", "illustration", "illus.", "schéma", "schema",
    "graphique", "graph.", "photo", "photographie", "image", "carte", "diagramme",
    "الشكل رقم", "شكل رقم", "الشكل البياني", "شكل بياني", "الشكل", "شكل", "الصورة", "صورة",
    "الرسم البياني", "الرسم", "رسم", "الخريطة", "خريطة",
  ],
  table: [
    "tableau", "table", "tabl.", "tabl", "tab.", "tbl",
    "الجدول رقم", "جدول رقم", "الجدول", "جدول",
  ],
  equation: [
    "équation", "equation", "éq.", "eq.", "formule",
    "المعادلة رقم", "معادلة رقم", "المعادلة", "معادلة",
  ],
};

/** Normalised label ("الشكل رقم") → kind. */
const LABEL_TO_KIND = new Map<string, CaptionKind>();
for (const kind of Object.keys(KIND_LABELS) as CaptionKind[]) {
  for (const label of KIND_LABELS[kind]) LABEL_TO_KIND.set(label.toLowerCase(), kind);
}

/** Collapse NBSP + runs of whitespace so " الشكل  رقم " keys the vocab map. */
function normalizeLabel(raw: string): string {
  return raw.replace(/[\s\u00A0]+/g, " ").trim();
}

const RX_META = /[.*+?^${}()|[\]\\]/g;

/** One alternation over every label, longest-first, spaces → `\s+`. */
const LABEL_ALT = [...LABEL_TO_KIND.keys()]
  .sort((a, b) => b.length - a.length)
  .map((l) => l.replace(RX_META, "\\$&").replace(/ /g, "[\\s\\u00A0]+"))
  .join("|");

// Arabic-Indic digits count too — an Arabic thesis writes "الجدول رقم ٣".
const NUM = "[0-9\\u0660-\\u0669IVXLCMivxlcm]+";

/**
 * `Figure n° I-1 : Organigramme` →
 *   1 label-with-number-word ("Figure n°")  2 number ("I-1")
 *   3 separator (" : ")                     4 text ("Organigramme")
 */
const TEXT_CAPTION_RE = new RegExp(
  "^[\\s\\u00A0]*" +
  `((?:${LABEL_ALT})(?:[\\s\\u00A0]*(?:n[°ºo]\\.?|nº|رقم|#))?)` +
  "[\\s\\u00A0]*[([]?[\\s\\u00A0]*" +
  `(${NUM}(?:[\\s\\u00A0]*[.\\-–—/][\\s\\u00A0]*${NUM})*)?` +
  // Only eat whitespace here when a closing bracket follows it — otherwise the
  // space in "Figure 1 : A" is swallowed and the caption reads "Figure 1: A".
  "(?:[\\s\\u00A0]*[)\\]])?" +
  "([\\s\\u00A0]*[:：.\\-–—،,][\\s\\u00A0]*|[\\s\\u00A0]+|$)" +
  "([\\s\\S]*)$",
  "i",
);

/**
 * A list-style number typed IN FRONT of the label: `2- جدول …`, `3) Figure : x`,
 * `١. شكل س`. It needs a closing mark (`-`, `.`, `)`, …) so a section number
 * like `2.1 Tableau …` never reads as one. 1 = the number as typed.
 */
const LEAD_DIGITS = "[0-9\\u0660-\\u0669]+"; // digits only: "civil- Tableau" is not numbered
const LEAD_NUMBER_RE = new RegExp(
  `^[\\s\\u00A0]*[([]?[\\s\\u00A0]*(${LEAD_DIGITS}(?:[\\s\\u00A0]*[.\\-–—/][\\s\\u00A0]*${LEAD_DIGITS})*)` +
  "[\\s\\u00A0]*[)\\]]?[\\s\\u00A0]*[.\\-–—:：)/][\\s\\u00A0]*",
);

/** True for a separator that PUNCTUATES — the strongest "this is a caption" signal. */
const PUNCTUATED = /[:：.\-–—،,]/;

export interface TextCaptionMatch {
  kind: CaptionKind;
  /** The label exactly as typed, e.g. "Fig." or "الشكل رقم". */
  rawLabel: string;
  /** The number exactly as typed, e.g. "I-1" or "٣"; "" when unnumbered. */
  rawNumber: string;
  /** What sat between the number and the text, e.g. " : ". */
  separator: string;
  /** The caption wording after the label, number and separator. */
  text: string;
  /**
   * A list-style number typed BEFORE the label ("2- " in "2- جدول …"); "" when
   * there is none. When the label carries no number of its own, `rawNumber`
   * holds this one.
   */
  leadingNumber: string;
}

/**
 * Parse a paragraph's plain text as a hand-typed caption, or null.
 *
 * Deliberately conservative — a false positive rewrites a body paragraph into a
 * caption, which is far worse than missing one the student can point at. A hit
 * needs a caption label at the very START of the paragraph plus either a number
 * or punctuation after it, and prose that runs on ("Figure 1 shows that …")
 * is rejected: with no punctuation the remainder must be short and must not
 * open with a lowercase word.
 */
export function parseTextCaption(text: string): TextCaptionMatch | null {
  const plain = parseCaptionBody(text, false);
  if (plain) return plain;
  // "2- جدول البيان": a typed list number in front is a caption number too.
  const lead = LEAD_NUMBER_RE.exec(text);
  if (!lead) return null;
  const inner = parseCaptionBody(text.slice(lead[0].length), true);
  if (!inner) return null;
  const leadingNumber = lead[1].replace(/[\s\u00A0]+/g, "");
  return { ...inner, rawNumber: inner.rawNumber || leadingNumber, leadingNumber };
}

function parseCaptionBody(text: string, hasLead: boolean): TextCaptionMatch | null {
  const m = TEXT_CAPTION_RE.exec(text);
  if (!m) return null;

  const rawLabel  = m[1];
  const rawNumber = (m[2] ?? "").replace(/[\s\u00A0]+/g, "");
  const separator = m[3] ?? "";
  const body      = (m[4] ?? "").trim();

  const base = normalizeLabel(rawLabel).toLowerCase().replace(/[\s\u00A0]*(?:n[°ºo]\.?|nº|رقم|#)$/, "").trim();
  const kind = LABEL_TO_KIND.get(base);
  if (!kind) return null;

  const punctuated = PUNCTUATED.test(separator);
  // "Table of contents" / "Figure" alone: a label and nothing that marks a caption.
  if (!rawNumber && !punctuated && !hasLead) return null;
  // A typed list number in front of a bare label ("2- Figure") is not a caption.
  if (hasLead && !rawNumber && !punctuated && !body) return null;
  if (body.length > 300) return null;
  if (!punctuated && body && (body.length > 80 || /^\p{Ll}/u.test(body))) return null;

  return { kind, rawLabel: normalizeLabel(rawLabel), rawNumber, separator, text: body, leadingNumber: "" };
}

/** `<w:p …>` / `</w:p>`, normalising a self-closing `<w:p …/>` into a pair. */
function paragraphShell(xml: string): { head: string; tail: string } {
  const openEnd = xml.indexOf(">") + 1;
  const selfClosing = xml.slice(0, openEnd).endsWith("/>");
  if (selfClosing) return { head: xml.slice(0, openEnd).replace(/\s*\/>$/, ">"), tail: "</w:p>" };
  return { head: xml.slice(0, openEnd), tail: "</w:p>" };
}

/**
 * The paragraph's own `<w:pPr>` with its style switched to Caption.
 *
 * Keeping the student's `<w:pPr>` — alignment, indent, spacing, `w:bidi`, the
 * numbering they attached — is the product rule: converting a caption must not
 * restyle their thesis. Only `w:pStyle` changes, and it has to stay FIRST:
 * `w:pPr` is an ORDERED sequence and Word refuses a file that reorders it.
 */
function captionPPr(paragraphXml: string): string {
  const pPr = paragraphPropsXml(paragraphXml);
  const style = `<w:pStyle w:val="${CAPTION_STYLE_ID}"/>`;
  if (!pPr || /^<w:pPr\b[^>]*\/>$/.test(pPr)) return `<w:pPr>${style}</w:pPr>`;
  if (/<w:pStyle\b[^>]*\/>/.test(pPr)) return pPr.replace(/<w:pStyle\b[^>]*\/>/, style);
  return pPr.replace(/^<w:pPr\b[^>]*>/, (open) => open + style);
}

/** Bookmark starts/ends already in the paragraph — cross-references point at them. */
function existingBookmarks(xml: string): { starts: string; ends: string } {
  return {
    starts: (xml.match(/<w:bookmarkStart\b[^>]*\/>/g) ?? []).join(""),
    ends:   (xml.match(/<w:bookmarkEnd\b[^>]*\/>/g) ?? []).join(""),
  };
}

/** Name prefix of the hidden bookmark that wraps a GENERATED caption's separator. */
const SEP_BOOKMARK = "_CapSep";

/**
 * Drop the separator run `addGeneratedCaption` wrote ("`: `") from the text that
 * follows a caption's number, so `text` is the student's own words only. A
 * converted caption has no such marker — its separator IS part of its text.
 */
function stripGeneratedSeparator(region: string): string {
  return region.replace(
    new RegExp(`<w:bookmarkStart w:id="(\\d+)" w:name="${SEP_BOOKMARK}\\d*"/>[\\s\\S]*?<w:bookmarkEnd w:id="\\1"/>`),
    "",
  );
}

// ─── A typed label word that becomes the generated label ─────────────────────

/** Whitespace a student types between words; a `w:t` never holds a newline. */
const TYPED_SPACE = "[ \\t\\u00A0\\u202F]";
const rxEscape = (t: string) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * A paragraph's first run up to and including its first `<w:t>`:
 * 1 open tag  2 run properties (+ a rendering hint)  3 the `w:t` attributes  4 its raw text.
 * A run whose properties carry a revision record (`w:rPrChange`) never matches: its
 * `<w:rPr>` nests another, and cutting at the first close would split the revision.
 */
const FIRST_RUN_RE =
  /^(<w:r(?:\s[^>]*[^/>])?>)((?:<w:rPr\s*\/>|<w:rPr(?:\s[^>]*)?>(?:(?!<w:rPrChange)[\s\S])*?<\/w:rPr>)?(?:<w:lastRenderedPageBreak\s*\/>)?)<w:t((?:\s[^>]*)?)>([^<]*)<\/w:t>/;

/**
 * Cut a typed label WORD ("جدول") off the front of a paragraph, together with the
 * whitespace after it, so that a generated label can take its place instead of
 * repeating it ("جدول 1: جدول الميزانية" -> "جدول 1: الميزانية").
 *
 * Run-level surgery on the paragraph's FIRST text run and nothing else. It cuts only
 * when the word sits wholly at the start of that run's first `<w:t>`, is a whole word
 * (whitespace or the end of the run after it), and more wording follows. Otherwise —
 * the word split across runs, a different word that merely starts with it, a tab or
 * a field or a hyperlink or a revision in front, the word standing alone — it returns
 * the paragraph untouched with `dropped: false`. Every other run, bookmark and field
 * is copied through byte for byte. A run that held nothing but the word disappears;
 * a tatweel inside the word ("جـدول") does not stop it matching.
 *
 * Pure, so that a caller can ask what {@link CaptionManager.addGeneratedCaption} would
 * do with `dropLeadingWord` without touching a document.
 */
export function removeLeadingWord(paragraphXml: string, word: string): { xml: string; dropped: boolean } {
  const unchanged = { xml: paragraphXml, dropped: false };
  const words = word.trim().split(/\s+/).filter(Boolean);
  if (!words.length) return unchanged;

  const shell = /^(\s*<w:p(?:\s[^>]*[^/>])?>)([\s\S]*?)(<\/w:p>\s*)$/.exec(paragraphXml);
  if (!shell) return unchanged; // not a paragraph, or an empty `<w:p/>`
  const [, open, inner, close] = shell;
  const pPr = paragraphPropsXml(paragraphXml);
  if (/<w:pPrChange\b/.test(pPr)) return unchanged;
  const prefix = pPr ? inner.slice(0, inner.indexOf(pPr) + pPr.length) : "";
  const rest = inner.slice(prefix.length);

  // Only markers that carry no text may stand in front of the first run.
  const lead = /^(?:\s*<w:(?:bookmarkStart|bookmarkEnd|proofErr)\b[^>]*\/>)*\s*/.exec(rest)![0];
  const run = FIRST_RUN_RE.exec(rest.slice(lead.length));
  if (!run) return unchanged;
  const [head, runProps, tAttrs, rawText] = [run[1], run[2], run[3], run[4]];

  const wordPattern = words
    .map((w) => Array.from(w).map((ch) => rxEscape(escapeXmlContent(ch))).join("\\u0640*"))
    .join(`${TYPED_SPACE}+`);
  const matched = new RegExp(`^${TYPED_SPACE}*${wordPattern}`).exec(rawText);
  if (!matched) return unchanged;
  const afterWord = rawText.slice(matched[0].length);
  const spaces = new RegExp(`^${TYPED_SPACE}*`).exec(afterWord)![0];
  if (afterWord !== "" && spaces === "") return unchanged; // glued to a longer word
  const tail = afterWord.slice(spaces.length); // what stays of this run's text

  const afterText = lead.length + run[0].length; // `rest` offset just past </w:t>
  const runEnd = rest.indexOf("</w:r>", afterText);
  if (runEnd < 0) return unchanged;
  const runTail = rest.slice(afterText, runEnd); // the run's content after that w:t
  if (/<w:r[\s>/]/.test(runTail)) return unchanged; // a run nested in this one (an object, alternate content)
  const after = rest.slice(runEnd + "</w:r>".length);

  const following = paragraphText(runTail + after);
  if (tail === "" && spaces === "" && !/^[\s\u00A0\u202F]/.test(following)) return unchanged; // the word runs on in the next run
  if (tail.trim() === "" && following.trim() === "") return unchanged; // the word stands alone

  const newRun =
    tail !== "" ? `${head}${runProps}<w:t${tAttrs}>${tail}</w:t>${runTail}</w:r>`
    : runTail.trim() === "" ? ""
    : `${head}${runProps}${runTail}</w:r>`;
  return { xml: `${open}${prefix}${lead}${newRun}${after}${close}`, dropped: true };
}

const LEADING_SPACES_RUN = new RegExp(
  "^((?:\\s*<w:(?:bookmarkStart|bookmarkEnd|proofErr)\\b[^>]*/>)*\\s*<w:r(?:\\s[^>]*)?>(?:<w:rPr>(?:(?!</w:rPr>)[\\s\\S])*</w:rPr>|<w:rPr/>)?)" +
  `<w:t((?:\\s[^>]*)?)>(${TYPED_SPACE}+)([^<]*)</w:t>(\\s*</w:r>)`,
);

/**
 * Wording's run XML without the spaces typed in front of it. A run of nothing but spaces goes, and the
 * next one is read too. Stops at anything else: a tab, a field, a picture, a link stays where it stands.
 */
function withoutLeadingSpaces(rest: string): string {
  let out = rest;
  for (;;) {
    const m = LEADING_SPACES_RUN.exec(out);
    if (!m) return out;
    const [all, head, attrs, , text, close] = m;
    if (text !== "") return `${head}<w:t${attrs}>${text}</w:t>${close}${out.slice(all.length)}`;
    const lead = /^(?:\s*<w:(?:bookmarkStart|bookmarkEnd|proofErr)\b[^>]*\/>)*\s*/.exec(head)![0];
    out = lead + out.slice(all.length);
  }
}

/**
 * Read EVERY caption in one paragraph's XML — Word numbers each SEQ field it
 * finds, and real theses do put two figures on one line ("Figure 19: … Figure
 * 20: …"). Returns [] when the paragraph carries no caption field.
 *
 * Exported because reading a caption needs no document: the app's figure-caption
 * endpoint uses it to strip the label and number off a caption before handing the
 * wording to the model.
 */
export function parseCaptionParagraph(xml: string, blockIndex = 0): CaptionEntry[] {

  const fields = findFields(xml);
  const seqs = fields
    .map((f, i) => ({ f, i, seq: parseSeqInstr(f.instr) }))
    .filter((x): x is { f: FieldHit; i: number; seq: NonNullable<ReturnType<typeof parseSeqInstr>> } => x.seq !== null);
  if (!seqs.length) return [];

  const bm = /<w:bookmarkStart\b[^>]*\bw:name="([^"]+)"/.exec(xml);
  const out: CaptionEntry[] = [];

  for (let n = 0; n < seqs.length; n++) {
    const { f, i, seq } = seqs[n];
    const prev = n > 0 ? seqs[n - 1] : null;
    // The label sits between the previous caption and this one's FIRST field
    // (a chapter STYLEREF when "include chapter number" is on).
    const labelStart = prev ? prev.f.end : 0;
    const groupStart = fields[prev ? prev.i + 1 : 0]?.start ?? f.start;
    const next = seqs[n + 1] ?? null;
    const textEnd = next ? (fields[i + 1]?.start ?? next.f.start) : xml.length;

    out.push({
      label: seq.id,
      displayLabel: paragraphText(xml.slice(labelStart, groupStart)).trim(),
      number: f.resultStart >= 0 ? paragraphText(xml.slice(f.resultStart, f.resultEnd)).trim() : "",
      text: paragraphText(stripGeneratedSeparator(xml.slice(f.end, textEnd))).trim(),
      fullText: paragraphText(xml.slice(labelStart, textEnd)).trim(),
      blockIndex,
      paragraphIndex: blockIndex,
      bookmark: bm ? bm[1] : null,
      inTable: false,
    });
  }
  return out;
}

// ─── Manager ──────────────────────────────────────────────────────────────────

export class CaptionManager {
  private zip: AdmZip;
  private styles: StylesManager;
  private doc: DocumentManager;
  private stylesRegistered = false;

  constructor(zip: AdmZip) {
    this.zip    = zip;
    this.styles = new StylesManager(zip);
    // Every read/write goes through DocumentManager's ORDER-PRESERVING block API
    // (OrderedBody string splicing). Rebuilding document.xml through an XML
    // object model regroups the body by tag — hoisting every table away from its
    // paragraphs — and trims the whitespace-only runs Word uses between words.
    this.doc = new DocumentManager(zip);
  }

  /** Register Caption + CaptionChar styles into styles.xml the first time. */
  private async ensureStyles(): Promise<void> {
    if (this.stylesRegistered) return;
    const existing = await this.styles.listStyles();
    const ids      = new Set(existing.map((s) => s.id));
    if (!ids.has(CAPTION_STYLE_ID))  await this.styles.addStyle(CAPTION_PARA_STYLE);
    if (!ids.has("CaptionChar"))     await this.styles.addStyle(CAPTION_CHAR_STYLE);
    if (!ids.has("TableofFigures"))  await this.styles.addStyle(TABLE_OF_FIGURES_STYLE);
    this.stylesRegistered = true;
  }

  // ── Internal ────────────────────────────────────────────────────────────────

  /** Next free `w:id` for a bookmark, and a free `_Ref…` bookmark name. */
  private nextBookmark(): { id: number; name: string } {
    const xml = this.zip.readAsText(DOCUMENT_PATH) ?? "";
    let maxId = 0;
    for (const m of xml.matchAll(/<w:bookmark(?:Start|End)\b[^>]*\bw:id="(\d+)"/g)) {
      maxId = Math.max(maxId, Number(m[1]));
    }
    let maxRef = 100000000;
    for (const m of xml.matchAll(/w:name="_(?:Ref|Toc)(\d+)"/g)) {
      maxRef = Math.max(maxRef, Number(m[1]));
    }
    return { id: maxId + 1, name: `_Ref${maxRef + 1}` };
  }

  /**
   * A bookmark allocator that hands out a DISTINCT id/name on every call.
   *
   * {@link nextBookmark} re-reads document.xml each time, so a batch that mints
   * many bookmarks before saving (the converter) would stamp the same
   * `w:id`/`_Ref…` on all of them — duplicate bookmark names, and every
   * cross-reference and list-of-figures entry pointing at the first one.
   */
  private bookmarkSeries(): () => { id: number; name: string } {
    const first    = this.nextBookmark();
    const startRef = Number(first.name.slice("_Ref".length));
    let n = 0;
    return () => {
      const k = n++;
      return { id: first.id + k, name: `_Ref${startRef + k}` };
    };
  }

  /**
   * Build a caption paragraph the way Word writes one:
   *
   *   <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr>
   *     <w:bookmarkStart .../>            ← so a list of figures / cross-ref can point here
   *     <w:r><w:t>Figure</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>
   *     <w:fldSimple w:instr=" STYLEREF 1 \s ">…</w:fldSimple><w:r><w:t>-</w:t></w:r>
   *     <w:fldSimple w:instr=" SEQ Figure \* ARABIC \s 1 "><w:r><w:t>1</w:t></w:r></w:fldSimple>
   *     <w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>text</w:t></w:r>
   *   <w:bookmarkEnd .../></w:p>
   */
  private buildCaptionXml(
    opts: CaptionOptions,
    seqNumber: number,
    chapterNumber: number,
    bookmark: { id: number; name: string } | null,
  ): string {
    const rtl = opts.rtl === true;
    const pPr = `<w:pPr><w:pStyle w:val="${CAPTION_STYLE_ID}"/>${rtl ? "<w:bidi/>" : ""}</w:pPr>`;
    const rPr = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    return `<w:p>${pPr}${this.captionContent(opts, seqNumber, chapterNumber, bookmark, rPr)}</w:p>`;
  }

  /**
   * Everything a caption paragraph holds AFTER its `<w:pPr>`: the bookmark, the
   * label runs, the chapter STYLEREF, the SEQ field and the caption text.
   *
   * Split out of {@link buildCaptionXml} because {@link convertTextCaptions}
   * needs the same content inside a paragraph that ALREADY exists — it keeps the
   * student's own `<w:pPr>` and run properties rather than writing fresh ones.
   * `rPr` is the run-properties string every generated run carries.
   */
  private captionContent(
    opts: CaptionOptions,
    seqNumber: number,
    chapterNumber: number,
    bookmark: { id: number; name: string } | null,
    rPr: string,
  ): string {
    const label   = opts.label ?? "Figure";
    const seqId   = seqIdentifier(label);
    const num     = opts.numbering ?? {};
    const fmt     = FORMAT_MAP[num.format ?? "arabic"];
    const run     = (t: string) => `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(t)}</w:t></w:r>`;
    const level   = chapterLevel(num.chapterStyle);

    const parts: string[] = [];
    if (bookmark) parts.push(`<w:bookmarkStart w:id="${bookmark.id}" w:name="${bookmark.name}"/>`);

    if (!opts.excludeLabel) {
      parts.push(run(label));
      parts.push(run(" "));
    }

    if (num.includeChapterNumber) {
      // `\* ROMAN` on the STYLEREF is how "Figure I-1" is written when the
      // chapter heading itself is not roman-numbered; omitted otherwise so the
      // field keeps echoing the heading's own list format.
      const chFmt   = num.chapterFormat ? FORMAT_MAP[num.chapterFormat] : null;
      const chInstr = ` STYLEREF ${level} \\s ${chFmt ? `\\* ${chFmt} ` : ""}`;
      const chText  = formatSeqNumber(chapterNumber || 1, chFmt ?? "ARABIC");
      parts.push(`<w:fldSimple w:instr="${escapeXmlText(chInstr)}">${run(chText)}</w:fldSimple>`);
      parts.push(run(num.chapterSeparator ?? "-"));
    }

    const seqInstr = ` SEQ ${seqId} \\* ${fmt}${num.includeChapterNumber ? ` \\s ${level}` : ""} `;
    parts.push(`<w:fldSimple w:instr="${escapeXmlText(seqInstr)}">${run(formatSeqNumber(seqNumber, fmt))}</w:fldSimple>`);

    const text = (opts.text ?? "").trim();
    if (text) {
      parts.push(run(opts.textSeparator ?? " "));
      parts.push(run(text));
    }
    if (bookmark) parts.push(`<w:bookmarkEnd w:id="${bookmark.id}"/>`);

    return parts.join("");
  }

  /**
   * Recompute every SEQ (and caption STYLEREF) field RESULT in the body, in
   * DOCUMENT order, honouring each field's own `\*` format and `\s` chapter
   * reset — exactly what Word does when you press F9.
   *
   * This matters because nothing outside Word updates fields: the app's preview
   * and the OnlyOffice PDF both render the cached result. Without this pass a
   * caption inserted before an existing one keeps the higher number on screen.
   */
  public async renumber(): Promise<void> {
    const blocks = await this.doc.getBlocks();
    if (!this.renumberBlocks(blocks)) return;
    await this.doc.saveBlocks(blocks);
  }

  /** In-memory half of {@link renumber}. Returns true when something changed. */
  private renumberBlocks(blocks: BodyBlock[]): boolean {
    const headingCounts = new Array(10).fill(0) as number[];
    const counters = new Map<string, number>(); // `${seqId} ${resetLevel}` → n
    let changed = false;

    for (const b of blocks) {
      if (b.kind === "paragraph") {
        const level = paragraphHeadingLevel(b.xml);
        if (level > 0) {
          headingCounts[level]++;
          // A SEQ with `\s N` restarts at every level-N heading.
          for (const key of counters.keys()) {
            if (Number(key.split(" ")[1]) === level) counters.set(key, 0);
          }
        }
      }

      // Tables count too: theses routinely park a figure and its caption inside
      // a one-cell table to centre them, and Word numbers those SEQ fields in
      // the same sequence. Skipping them would drift every later number.
      if (!/\bSEQ\b/.test(b.xml)) continue;
      const next = this.renumberFieldsIn(b.xml, counters, headingCounts);
      if (next !== b.xml) { b.xml = next; changed = true; }
    }
    return changed;
  }

  /** Rewrite every SEQ / caption-STYLEREF field result inside one block's XML. */
  private renumberFieldsIn(
    xml: string,
    counters: Map<string, number>,
    headingCounts: number[],
  ): string {
    const fields = findFields(xml);
    const seqFields = fields.filter((f) => parseSeqInstr(f.instr));
    if (!seqFields.length) return xml;

    // Collect edits first, then apply back-to-front so earlier indices stay valid.
    const edits: Array<{ start: number; end: number; text: string }> = [];

    for (const f of fields) {
      if (f.resultStart < 0) continue;
      const seq = parseSeqInstr(f.instr);
      if (seq) {
        const key = `${seq.id} ${seq.resetLevel}`;
        const n = (counters.get(key) ?? 0) + 1;
        counters.set(key, n);
        const replaced = replaceFirstText(
          xml.slice(f.resultStart, f.resultEnd),
          formatSeqNumber(n, seq.format),
        );
        if (replaced !== null) edits.push({ start: f.resultStart, end: f.resultEnd, text: replaced });
        continue;
      }
      // Only chapter-number STYLEREFs that sit INSIDE a caption are ours to touch.
      const ref = parseStyleRef(f.instr);
      if (ref !== null) {
        const replaced = replaceFirstText(
          xml.slice(f.resultStart, f.resultEnd),
          formatSeqNumber(headingCounts[ref.level] || 1, ref.format),
        );
        if (replaced !== null) edits.push({ start: f.resultStart, end: f.resultEnd, text: replaced });
      }
    }

    let out = xml;
    for (const e of edits.sort((a, b) => b.start - a.start)) {
      out = out.slice(0, e.start) + e.text + out.slice(e.end);
    }
    return out;
  }

  /**
   * Read EVERY caption in one paragraph — Word numbers each SEQ field it finds,
   * and real theses do put two figures on one line ("Figure 19: … Figure 20: …").
   * Returns [] when the paragraph carries no caption field.
   */
  private readCaptions(xml: string, blockIndex: number): CaptionEntry[] {
    return parseCaptionParagraph(xml, blockIndex);
  }

  /** The first caption in a paragraph, or null. */
  private readCaption(xml: string, blockIndex: number): CaptionEntry | null {
    return this.readCaptions(xml, blockIndex)[0] ?? null;
  }

  // ── Public API ──────────────────────────────────────────────────────────────

  /**
   * Insert a caption paragraph next to the BLOCK at `nearIndex` — the same index
   * space as `document.getBlocks()`, so a table's index really does place the
   * caption against that table.
   *
   * `position: "below"` (default) inserts after `nearIndex`, `"above"` before it.
   *
   * @returns The BLOCK index at which the caption was inserted.
   */
  public async insertCaption(nearIndex: number, opts: CaptionOptions = {}): Promise<number> {
    await this.ensureStyles();

    const blocks   = await this.doc.getBlocks();
    const position = opts.position ?? "below";
    const at       = Math.max(0, Math.min(position === "below" ? nearIndex + 1 : nearIndex, blocks.length));

    // Cached field results, so the app preview and the exported PDF (neither of
    // which recalculates fields) show the same numbers Word will.
    const num     = opts.numbering ?? {};
    const level   = chapterLevel(num.chapterStyle);
    const seqId   = seqIdentifier(opts.label ?? "Figure");
    let seqNumber = 1;
    let chapter   = 0;
    for (let i = 0; i < at && i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === "paragraph" && paragraphHeadingLevel(b.xml) === level) {
        chapter++;
        if (num.includeChapterNumber) seqNumber = 1;
      }
      for (const para of eachParagraph(b.xml)) {
        for (const existing of this.readCaptions(para.xml, i)) {
          if (existing.label === seqId) seqNumber++;
        }
      }
    }

    const xml = this.buildCaptionXml(opts, seqNumber, chapter, this.nextBookmark());
    await this.doc.insertBlockAt({ kind: "paragraph", tag: "w:p", xml }, at);
    // One authoritative pass over the whole body — every caption of this label
    // (including any that now sit after the new one) gets its true number.
    await this.renumber();
    return at;
  }

  /**
   * Turn the thesis's HAND-TYPED captions into real Word captions.
   *
   * A student who never opened References → Insert Caption still typed
   * "Figure 1 : Organigramme" under each figure. Those are dead text: no List
   * of Figures collects them, no cross-reference points at them, and inserting
   * a figure in the middle means renumbering every one by hand. This rewrites
   * each such paragraph IN PLACE as the real thing — Caption style, a bookmark,
   * and a live SEQ field — keeping the student's own wording, separator,
   * alignment and run formatting.
   *
   * In place matters twice over: no block is added or removed, so **every block
   * index the caller holds stays valid**, and a caption that lives inside a
   * one-cell table (the usual way a thesis centres a figure) converts too.
   *
   * Numbering is the caller's choice, exactly as in Word's dialog: plain
   * sequential by default ("Figure 1", "Figure 2", …), or per-chapter when
   * `numbering.includeChapterNumber` is set ("Figure I-1", "Figure II-1", …
   * restarting at every chapter heading).
   *
   * Labels are UNIFIED per kind — a thesis that mixes "Fig. 2" and "Figure 3"
   * would otherwise end up with two competing SEQ sequences and two half-empty
   * lists of figures. The winner is `opts.label` when given, else the label the
   * document's existing REAL captions already use, else the spelling the typed
   * ones use most often.
   */
  public async convertTextCaptions(
    opts: ConvertTextCaptionsOptions = {},
  ): Promise<ConvertTextCaptionsResult> {
    const blocks = await this.doc.getBlocks();
    const last   = blocks.length - 1;
    const from   = Math.max(0, Math.min(opts.fromIndex ?? 0, last));
    const to     = Math.min(last, Math.max(opts.toIndex ?? last, from));
    const wantedKinds = new Set<CaptionKind>(
      !opts.kind || opts.kind === "all" ? ["figure", "table", "equation"] : [opts.kind],
    );

    interface Hit {
      blockIndex: number;
      /** Which paragraph inside that block (a table cell holds several). */
      paraOrdinal: number;
      start: number;
      end: number;
      xml: string;
      match: TextCaptionMatch;
    }
    const hits: Hit[] = [];
    const skipped: SkippedTextCaption[] = [];

    for (let i = from; i <= to; i++) {
      const block = blocks[i];
      if (block.kind === "sectPr" || block.kind === "other") continue;
      eachParagraph(block.xml).forEach((para, paraOrdinal) => {
        const text = paragraphText(para.xml).trim();
        if (!text) return;
        const match = parseTextCaption(text);
        if (!match || !wantedKinds.has(match.kind)) return;

        const note = (reason: SkipReason) => skipped.push({ blockIndex: i, reason, text });
        // Already a real caption — a second SEQ field would number it twice.
        if (this.readCaptions(para.xml, i).length) return note("already-a-caption");
        if (paragraphHeadingLevel(para.xml) > 0) return note("is-a-heading");
        // The figure itself lives in this paragraph: rebuilding the runs would
        // delete the picture along with the text.
        if (/<w:(?:drawing|pict|object)\b/.test(para.xml)) return note("contains-image");
        // A REF/footnote/PAGEREF here would be destroyed by the rewrite.
        if (findFields(para.xml).length) return note("contains-field");

        hits.push({ blockIndex: i, paraOrdinal, start: para.start, end: para.end, xml: para.xml, match });
      });
    }

    const labels = await this.resolveConvertLabels(hits.map((h) => h.match), opts.label, wantedKinds);
    // Registered only once there is something to convert, so a scan that finds
    // nothing leaves styles.xml — and the document's bytes — untouched.
    if (hits.length) await this.ensureStyles();

    // Back-to-front within each block: rewriting a paragraph changes the length
    // of the block's XML, so every earlier offset must still be untouched.
    const nextBookmark = this.bookmarkSeries();
    for (const hit of [...hits].sort((a, b) => b.blockIndex - a.blockIndex || b.start - a.start)) {
      const block = blocks[hit.blockIndex];
      const label = labels[hit.match.kind]!;
      const rtl   = opts.rtl ?? /<w:rtl\b|<w:bidi\b/.test(hit.xml);
      const rPr   = paragraphRunPropsXml(hit.xml) || (rtl ? "<w:rPr><w:rtl/></w:rPr>" : "");
      const bm    = nextBookmark();
      const keep  = existingBookmarks(hit.xml);
      const shell = paragraphShell(hit.xml);
      const content = this.captionContent(
        {
          label,
          text: hit.match.text,
          numbering: opts.numbering,
          rtl,
          // Keep the student's own punctuation (" : " in French, "- ", ". ") unless the
          // caller asked for one separator everywhere.
          textSeparator: opts.textSeparator ?? (hit.match.separator.replace(/[\s ]+/g, " ") || " "),
        },
        1, // placeholder — renumberBlocks below writes every true number
        1,
        bm,
        rPr,
      );

      block.xml =
        block.xml.slice(0, hit.start) +
        `${shell.head}${captionPPr(hit.xml)}${keep.starts}${content}${keep.ends}${shell.tail}` +
        block.xml.slice(hit.end);
    }

    // One authoritative pass: every SEQ and chapter STYLEREF in the document
    // (converted or pre-existing) gets its true number, in document order.
    if (hits.length) this.renumberBlocks(blocks);
    if (hits.length && !opts.dryRun) await this.doc.saveBlocks(blocks);

    const converted: ConvertedCaption[] = hits.map((hit) => ({
      blockIndex: hit.blockIndex,
      kind: hit.match.kind,
      label: labels[hit.match.kind]!,
      before: paragraphText(hit.xml).trim(),
      after: paragraphText(eachParagraph(blocks[hit.blockIndex].xml)[hit.paraOrdinal]?.xml ?? "").trim(),
      typedNumber: hit.match.rawNumber,
      number: parseCaptionParagraph(eachParagraph(blocks[hit.blockIndex].xml)[hit.paraOrdinal]?.xml ?? "")[0]?.number ?? "",
    }));

    return { converted, skipped, labels, dryRun: opts.dryRun === true };
  }

  /**
   * Give an UNNUMBERED caption a real label and number, leaving the student's
   * own words alone.
   *
   * A student who typed just "البيان المالي" under a table has a caption in
   * everything but name. This writes `label + SEQ field + separator` in FRONT of
   * their runs and switches the style to Caption — the label and number are
   * generated apparatus, like a page number. Unlike {@link convertTextCaptions}
   * it never rebuilds the runs from text: footnote references, hyperlinks,
   * fields, bookmarks and run formatting stay byte-identical and in order.
   *
   * Refuses, changing nothing, when the block is not a paragraph, already is a
   * real caption, is a heading, has no text, or carries a picture (the label
   * would land in front of the image, not the wording).
   */
  public async addGeneratedCaption(
    blockIndex: number,
    opts: {
      kind: CaptionKind; label?: string; rtl?: boolean; textSeparator?: string; numbering?: CaptionNumberingOptions;
      /**
       * The label WORD the student may have typed in front of their wording ("جدول"
       * in "جدول الميزانية"). When the paragraph starts with exactly that word, as a
       * whole word wholly inside its first text run, the word and the whitespace after
       * it are cut and the generated label takes its place: "جدول 1: الميزانية", not
       * "جدول 1: جدول الميزانية". When it cannot be cut cleanly (split across runs, not
       * at the start, standing alone, behind a field or a hyperlink) nothing is touched
       * and `droppedWord` is false. Only that first run changes; see {@link removeLeadingWord}.
       */
      dropLeadingWord?: string;
    },
  ): Promise<{ ok: true; label: string; number: string; droppedWord: boolean } | { ok: false; reason: GeneratedCaptionRefusal }> {
    const blocks = await this.doc.getBlocks();
    const block  = blocks[blockIndex];
    if (!block) return { ok: false, reason: "no-such-block" };
    if (block.kind !== "paragraph") return { ok: false, reason: "not-a-paragraph" };
    const xml = block.xml;
    if (this.readCaptions(xml, blockIndex).length) return { ok: false, reason: "already-a-caption" };
    if (paragraphHeadingLevel(xml) > 0) return { ok: false, reason: "is-a-heading" };
    if (!paragraphText(xml).trim()) return { ok: false, reason: "empty-text" };
    if (/<w:(?:drawing|pict|object)\b/.test(xml)) return { ok: false, reason: "contains-image" };

    await this.ensureStyles();

    const label = (opts.label ?? DEFAULT_LABEL[opts.kind]).trim();
    const rtl   = opts.rtl ?? /<w:rtl\b|<w:bidi\b/.test(xml);
    const rPr   = paragraphRunPropsXml(xml) || (rtl ? "<w:rPr><w:rtl/></w:rPr>" : "");
    const ids   = this.bookmarkSeries();
    const bm    = ids();
    const sepBm = ids();
    // Label + number only (no text), exactly as captionContent builds them.
    const prefix = this.captionContent({ label, numbering: opts.numbering, rtl }, 1, 1, bm, rPr);

    // The student's own label word, when asked for, is the one thing cut: the runs below come from `source`.
    const cut = opts.dropLeadingWord ? removeLeadingWord(xml, opts.dropLeadingWord) : null;
    const source = cut?.dropped ? cut.xml : xml;
    const shell = paragraphShell(source);
    const inner = source.slice(shell.head.length, source.endsWith("</w:p>") ? source.length - "</w:p>".length : source.length);
    const pPr   = paragraphPropsXml(source);
    const after = pPr ? inner.slice(inner.indexOf(pPr) + pPr.length) : inner;
    // The separator brings its own space: spaces typed in front of the wording would open a gap after it.
    const rest  = cut?.dropped ? after : withoutLeadingSpaces(after);

    let sep = opts.textSeparator ?? ": ";
    // The space the word left in the next run is the separator's own: do not write it twice.
    if (cut?.dropped && /^[\s\u00A0\u202F]/.test(paragraphText(rest))) sep = sep.replace(/[\s\u00A0\u202F]+$/, "");
    const sepRun =
      `<w:bookmarkStart w:id="${sepBm.id}" w:name="${SEP_BOOKMARK}${sepBm.id}"/>` +
      `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(sep)}</w:t></w:r>` +
      `<w:bookmarkEnd w:id="${sepBm.id}"/>`;
    block.xml = `${shell.head}${captionPPr(xml)}${prefix}${sepRun}${rest}${shell.tail}`;

    this.renumberBlocks(blocks);
    await this.doc.saveBlocks(blocks);

    const entry = parseCaptionParagraph(blocks[blockIndex].xml, blockIndex)[0];
    return { ok: true, label, number: entry?.number ?? "", droppedWord: cut?.dropped === true };
  }

  /**
   * The label each converted kind will carry. Existing REAL captions win — an
   * Arabic thesis whose figures already say "الشكل رقم" must not gain a second
   * "شكل" sequence — then the typed spelling used most often, then `override`.
   */
  private async resolveConvertLabels(
    matches: TextCaptionMatch[],
    override: string | undefined,
    wanted: Set<CaptionKind>,
  ): Promise<Partial<Record<CaptionKind, string>>> {
    const existing = await this.getCaptions().catch(() => [] as CaptionEntry[]);
    const out: Partial<Record<CaptionKind, string>> = {};

    for (const kind of wanted) {
      const mine = matches.filter((m) => m.kind === kind);
      if (!mine.length) continue;

      if (override && wanted.size === 1) { out[kind] = override.trim(); continue; }

      const real = existing.find((c) => {
        const base = normalizeLabel(c.displayLabel).toLowerCase();
        return base && LABEL_TO_KIND.get(base) === kind;
      });
      if (real) { out[kind] = normalizeLabel(real.displayLabel); continue; }

      const tally = new Map<string, number>();
      for (const m of mine) tally.set(m.rawLabel, (tally.get(m.rawLabel) ?? 0) + 1);
      out[kind] = [...tally.entries()].sort((a, b) => b[1] - a[1] || b[0].length - a[0].length)[0][0];
    }
    return out;
  }

  /**
   * Every caption in the document, in document order. `filterLabel` accepts
   * either the readable label ("الشكل رقم") or its SEQ identifier.
   */
  public async getCaptions(filterLabel?: string): Promise<CaptionEntry[]> {
    const blocks = await this.doc.getBlocks();
    const wanted = filterLabel ? seqIdentifier(filterLabel) : null;
    const out: CaptionEntry[] = [];
    blocks.forEach((b, i) => {
      for (const para of eachParagraph(b.xml)) {
        for (const entry of this.readCaptions(para.xml, i)) {
          if (wanted && entry.label !== wanted) continue;
          out.push({ ...entry, inTable: b.kind !== "paragraph" });
        }
      }
    });
    return out;
  }

  /** The distinct caption labels used in the document, most-used first. */
  public async listLabels(): Promise<Array<{ label: string; displayLabel: string; count: number }>> {
    const captions = await this.getCaptions();
    const byId = new Map<string, { label: string; displayLabel: string; count: number }>();
    for (const c of captions) {
      const hit = byId.get(c.label);
      if (hit) hit.count++;
      else byId.set(c.label, { label: c.label, displayLabel: c.displayLabel || c.label, count: 1 });
    }
    return [...byId.values()].sort((a, b) => b.count - a.count);
  }

  /**
   * Replace a caption's TEXT while keeping its label, number field and bookmark
   * intact — the edit Word makes when you retype a caption's wording.
   * Returns false when the block isn't a caption.
   */
  public async setCaptionText(blockIndex: number, text: string): Promise<boolean> {
    const blocks = await this.doc.getBlocks();
    const block  = blocks[blockIndex];
    if (!block) return false;
    for (const para of eachParagraph(block.xml)) {
      const next = this.replaceCaptionText(para.xml, text);
      if (next === null) continue;
      block.xml = block.xml.slice(0, para.start) + next + block.xml.slice(para.end);
      await this.doc.saveBlocks(blocks);
      return true;
    }
    return false;
  }

  /** String half of {@link setCaptionText}; null when `xml` isn't a caption. */
  private replaceCaptionText(xml: string, text: string): string | null {
    const fields = findFields(xml);
    let seqField: FieldHit | null = null;
    let seqIdx = -1;
    for (let i = 0; i < fields.length; i++) {
      if (parseSeqInstr(fields[i].instr)) { seqField = fields[i]; seqIdx = i; break; }
    }
    if (!seqField) return null;

    const closeIdx = xml.lastIndexOf("</w:p>");
    // Stop at the next field: a paragraph can hold a SECOND caption, and
    // swallowing everything to the end of the paragraph would delete it.
    const boundary = fields[seqIdx + 1]?.start ?? (closeIdx === -1 ? xml.length : closeIdx);
    const tail     = xml.slice(seqField.end, boundary);
    // Keep the bookmark that anchors this caption; drop the old text runs.
    const keep     = (tail.match(/<w:bookmarkEnd\b[^>]*\/>/g) ?? []).join("");
    const rtl      = /<w:rtl\b/.test(xml) || /<w:bidi\b/.test(xml);
    const rPr      = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    const clean    = text.trim();
    const runs     = clean
      ? `<w:r>${rPr}<w:t xml:space="preserve"> </w:t></w:r><w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(clean)}</w:t></w:r>`
      : "";

    return xml.slice(0, seqField.end) + runs + keep + xml.slice(boundary);
  }

  /**
   * Remove every caption for a label (e.g. all "Figure" captions), then
   * renumber what remains.
   */
  public async removeCaptions(label: string): Promise<void> {
    const wanted = seqIdentifier(label);
    const blocks = await this.doc.getBlocks();
    const kept = blocks.filter((b) => {
      // Only a caption that IS a body block can be removed; one inside a table
      // cell would take the table's layout with it.
      if (b.kind !== "paragraph") return true;
      const entry = this.readCaption(b.xml, 0);
      return !entry || entry.label !== wanted;
    });
    if (kept.length === blocks.length) return;
    this.renumberBlocks(kept);
    await this.doc.saveBlocks(kept);
  }

  /** Remove a single caption by BLOCK index, then renumber. */
  public async removeCaptionAt(blockIndex: number): Promise<void> {
    const blocks = await this.doc.getBlocks();
    if (blockIndex < 0 || blockIndex >= blocks.length) return;
    blocks.splice(blockIndex, 1);
    this.renumberBlocks(blocks);
    await this.doc.saveBlocks(blocks);
  }

  /** Convenience: insert a "Figure N — text" caption below a block. */
  public async insertFigureCaption(
    nearIndex: number,
    text: string,
    numbering?: CaptionNumberingOptions,
    opts: Omit<CaptionOptions, "text" | "numbering"> = {},
  ): Promise<number> {
    return this.insertCaption(nearIndex, { ...opts, label: opts.label ?? "Figure", text, numbering });
  }

  /** Convenience: insert a "Table N — text" caption above (Word default) a table. */
  public async insertTableCaption(
    nearIndex: number,
    text: string,
    position: CaptionPosition = "above",
    numbering?: CaptionNumberingOptions,
    opts: Omit<CaptionOptions, "text" | "numbering" | "position"> = {},
  ): Promise<number> {
    return this.insertCaption(nearIndex, { ...opts, label: opts.label ?? "Table", text, position, numbering });
  }

  /** Convenience: insert an "Equation N" caption. */
  public async insertEquationCaption(
    nearIndex: number,
    text = "",
    numbering?: CaptionNumberingOptions,
    opts: Omit<CaptionOptions, "text" | "numbering"> = {},
  ): Promise<number> {
    return this.insertCaption(nearIndex, { ...opts, label: opts.label ?? "Equation", text, numbering });
  }

  /** Insert a caption with a custom label (Word's "New Label"). */
  public async insertCustomCaption(
    nearIndex: number,
    label: string,
    text: string,
    opts: Omit<CaptionOptions, "label" | "text"> = {},
  ): Promise<number> {
    return this.insertCaption(nearIndex, { ...opts, label, text });
  }

  // ── Caption List (Table of Figures / Table of Tables) ──────────────────────

  /** One pre-populated entry: `<entry text><dot leader tab><page number>`. */
  private buildListEntryParagraph(entryText: string, bookmark: string | null, rtl: boolean): string {
    const rPr = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    const pageField = bookmark
      ? `<w:fldSimple w:instr="${escapeXmlText(` PAGEREF ${bookmark} \\h `)}"><w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>`
      : `<w:r>${rPr}<w:t>1</w:t></w:r>`;
    // `w:pPr` is an ORDERED sequence: tabs (11) BEFORE bidi (19). With bidi first
    // Word refuses the file — which only ever showed up on the Arabic theses.
    return (
      `<w:p><w:pPr><w:pStyle w:val="TableofFigures"/>` +
      `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8748"/></w:tabs>${rtl ? "<w:bidi/>" : ""}</w:pPr>` +
      `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(entryText)}</w:t></w:r>` +
      `<w:r>${rPr}<w:tab/></w:r>` +
      pageField +
      `</w:p>`
    );
  }

  /** Give a caption paragraph a bookmark if it has none, so PAGEREF can find it. */
  private ensureCaptionBookmark(xml: string, bm: { id: number; name: string }): string {
    if (/<w:bookmarkStart\b/.test(xml)) return xml;
    const pPrEnd = xml.indexOf("</w:pPr>");
    const insertAt = pPrEnd === -1 ? xml.indexOf(">") + 1 : pPrEnd + "</w:pPr>".length;
    const closeIdx = xml.lastIndexOf("</w:p>");
    return (
      xml.slice(0, insertAt) +
      `<w:bookmarkStart w:id="${bm.id}" w:name="${bm.name}"/>` +
      xml.slice(insertAt, closeIdx) +
      `<w:bookmarkEnd w:id="${bm.id}"/>` +
      xml.slice(closeIdx)
    );
  }

  /**
   * Insert a "List of <label>" — Word's References → Insert Table of Figures,
   * i.e. the field `TOC \h \z \c "Figure"`, pre-populated with one entry per
   * existing caption so the list reads correctly before Word ever updates it.
   *
   * @param label Caption label to collect ("Figure", "Table", "الشكل رقم", …).
   * @param title Heading shown above the list (pass "" to omit).
   * @param index BLOCK index to insert at (default: 0 = top of document).
   * @param rtl   Write the list right-to-left (Arabic thesis).
   */
  public async insertCaptionList(
    label: string,
    title?: string,
    index = 0,
    rtl = false,
  ): Promise<void> {
    await this.ensureStyles();

    const seqId  = seqIdentifier(label);
    const blocks = await this.doc.getBlocks();
    const bm     = this.nextBookmark();
    let bmSeq    = 0;

    // Collect the captions AND make sure each one is bookmarked, so the entries'
    // PAGEREF fields point at something real instead of "Error! Bookmark not
    // defined." on the first render.
    const entries: Array<{ text: string; bookmark: string | null }> = [];
    for (const b of blocks) {
      // Walk back-to-front so bookmarking one paragraph can't shift the offsets
      // of the ones still to be processed in the same block.
      const paras = eachParagraph(b.xml);
      const found: Array<{ text: string; bookmark: string | null }> = [];
      for (let k = paras.length - 1; k >= 0; k--) {
        const para = paras[k];
        const inPara = this.readCaptions(para.xml, 0).filter((e) => e.label === seqId);
        if (!inPara.length) continue;
        let name = inPara[0].bookmark;
        if (!name) {
          const next = { id: bm.id + bmSeq, name: `${bm.name}_${bmSeq}` };
          bmSeq++;
          b.xml = b.xml.slice(0, para.start) + this.ensureCaptionBookmark(para.xml, next) + b.xml.slice(para.end);
          name = next.name;
        }
        for (let e = inPara.length - 1; e >= 0; e--) found.unshift({ text: inPara[e].fullText, bookmark: name });
      }
      entries.push(...found);
    }

    const heading   = title !== undefined ? title : `List of ${label}s`;
    const instrText = ` TOC \\h \\z \\c "${seqId}" `;
    const rPr       = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    const bidi      = rtl ? "<w:bidi/>" : "";
    const newParas: BodyBlock[] = [];
    const para = (xml: string): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml });

    if (heading) {
      newParas.push(
        para(
          `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/>${bidi}</w:pPr>` +
            `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(heading)}</w:t></w:r></w:p>`,
        ),
      );
    }

    // Field begin + instruction + separate …
    newParas.push(
      para(
        `<w:p><w:pPr>${bidi}</w:pPr>` +
          `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
          `<w:r>${rPr}<w:instrText xml:space="preserve">${escapeXmlContent(instrText)}</w:instrText></w:r>` +
          `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r></w:p>`,
      ),
    );

    if (entries.length) {
      for (const e of entries) newParas.push(para(this.buildListEntryParagraph(e.text, e.bookmark, rtl)));
    } else {
      newParas.push(
        para(
          `<w:p><w:pPr><w:pStyle w:val="TableofFigures"/>${bidi}</w:pPr>` +
            `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(`No ${label} captions found`)}</w:t></w:r></w:p>`,
        ),
      );
    }

    // … field end.
    newParas.push(
      para(`<w:p><w:pPr>${bidi}</w:pPr><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r></w:p>`),
    );

    const insertAt = Math.max(0, Math.min(index, blocks.length));
    blocks.splice(insertAt, 0, ...newParas);
    await this.doc.saveBlocks(blocks);
  }

  /** Shorthand: insert List of Figures at the given BLOCK index. */
  public async insertListOfFigures(title = "List of Figures", index = 0, rtl = false): Promise<void> {
    return this.insertCaptionList("Figure", title, index, rtl);
  }

  /** Shorthand: insert List of Tables at the given BLOCK index. */
  public async insertListOfTables(title = "List of Tables", index = 0, rtl = false): Promise<void> {
    return this.insertCaptionList("Table", title, index, rtl);
  }

  /**
   * Remove the caption list (`TOC \c` field) for a given label — its heading,
   * the field, and every pre-populated entry paragraph.
   *
   * A TOC field SPANS paragraphs (begin … separate … entries … end), so this
   * walks from the paragraph holding the instruction to the one holding the
   * matching `fldChar end` rather than deleting a single paragraph.
   */
  public async removeCaptionList(label: string): Promise<void> {
    const seqId  = seqIdentifier(label);
    const blocks = await this.doc.getBlocks();
    const needle = `\\c "${seqId}"`;

    const instructionsIn = (xml: string): string =>
      [...xml.matchAll(/<w:instrText\b[^>]*?>([\s\S]*?)<\/w:instrText>/g)].map((m) => decodeXmlText(m[1])).join("") +
      [...xml.matchAll(/<w:fldSimple\b[^>]*?\bw:instr="([^"]*)"/g)].map((m) => decodeXmlText(m[1])).join("");

    const start = blocks.findIndex((b) => instructionsIn(b.xml).includes(needle));
    if (start === -1) return;

    let end = start;
    let depth = 0;
    for (let i = start; i < blocks.length; i++) {
      depth += (blocks[i].xml.match(/w:fldCharType="begin"/g) ?? []).length;
      depth -= (blocks[i].xml.match(/w:fldCharType="end"/g) ?? []).length;
      end = i;
      if (depth <= 0) break;
    }

    // The heading directly above ("List of Figures") belongs to the list.
    const from = start > 0 && /w:val="TOCHeading"/.test(blocks[start - 1].xml) ? start - 1 : start;
    blocks.splice(from, end - from + 1);
    await this.doc.saveBlocks(blocks);
  }
}
