/**
 * Doc — the human-friendly facade over mdocxengine.
 *
 * A flat, verb-first API for the 90% case: open a .docx, read its content as
 * plain data, edit it with simple verbs, and save. You never touch a zip entry,
 * a relationship, a part path, or raw OOXML. Blocks are addressed by their
 * zero-based index in document order (the same order `paragraphs()`/`tables()`
 * report), not by `paraId`/`relId`.
 *
 * The underlying `Mdocxengine` (and all its managers) stays reachable via
 * `doc.engine` for anything this facade doesn't cover.
 */
import { Mdocxengine } from "./index";
import { Table } from "./core/files/table/index";
import * as XmlUtils from "./utils/xmlUtils";
import {
  makeStyledParagraphNode,
  makeTableNode,
  makeDrawingParagraphNode,
  nextDrawingId,
  paragraphText,
  paragraphStyleId,
  paragraphHeadingLevel,
  paragraphAlignment,
  paragraphFontSizePt,
  paragraphIsBold,
  type BodyBlock,
  type StyledParagraphOptions,
} from "./core/files/body/OrderedBody";
import { pixelsToEmu, type InlineImage } from "./core/PartsManagers/MediaManager";
import { RelManager } from "./core/PartsManagers/RelManager";
import type { SectionEntry, SectionHeaderFooterRef } from "./core/PartsManagers/SectionManager";
import type { SectPrGeometry } from "./core/files/body/sectPr";
import type { SectPrPageBorderOptions } from "./core/files/body/sectPr";
import {
  TextStyleManager,
  expandTargets,
  type TextStyleTargetInput,
  type TargetReport,
} from "./core/PartsManagers/TextStyleManager";
import { TextStyleReader, type TextStyleInspection } from "./core/PartsManagers/TextStyleReader";
import type { RunProps } from "./core/ooxml/runProps";
import { toHalfPoints } from "./core/ooxml/hps";

const APPEND = Number.MAX_SAFE_INTEGER;

/** One body block as plain data. */
export interface BlockInfo {
  index: number;
  kind: "paragraph" | "table" | "image" | "other";
  /** Visible text (paragraph/cell text); "" for images. */
  text: string;
  /** Paragraph style id, if any. */
  styleId: string | null;
  /** Heading level 1–6, or 0 if not a heading. */
  headingLevel: number;
}

/** A node in the heading outline tree. */
export interface OutlineNode {
  index: number;
  level: number;
  title: string;
  children: OutlineNode[];
}

/** A table as a plain text grid. */
export interface TableInfo {
  index: number;
  rows: string[][];
}

/** An embedded image as plain data. */
export interface ImageInfo extends InlineImage {
  index: number;
}

/** A body block enriched with formatting signals — for inferring structure or AI labelling. */
export interface DetailedBlockInfo extends BlockInfo {
  /** Whole line is bold. */
  bold: boolean;
  /** Inline font size in points, or null if set via styles. */
  fontSizePt: number | null;
  /** Alignment (`center`, `right`, …) or null. */
  alignment: string | null;
  wordCount: number;
  /** Text looks like a figure/table caption (so NOT a heading). */
  looksLikeCaption: boolean;
}

/** A regex → heading-level rule used by {@link Doc.inferOutline}. */
export interface HeadingPattern {
  re: RegExp;
  level: number;
}

/** Tuning for {@link Doc.inferOutline}; all fields have sensible Arabic-academic defaults. */
export interface InferOutlineOptions {
  /** Text patterns that mark a heading at a given level (e.g. `^الفصل` → 1). */
  headingPatterns?: HeadingPattern[];
  /** Keywords that mark a top-level (level-1) section (e.g. `قائمة المراجع`). */
  level1Keywords?: string[];
  /** Keywords that mark a level-2 section (e.g. `تمهيد`). */
  sectionKeywords?: string[];
  /** Patterns that mark a caption (excluded from headings). */
  captionPatterns?: RegExp[];
  /** Max word count for a line to be considered a (format-based) heading. Default 14. */
  maxHeadingWords?: number;
}

/** A heading detected by {@link Doc.inferOutline}. */
export interface InferredHeading {
  index: number;
  level: number;
  title: string;
  /** `styled` = real heading markup; `high`/`medium` = inferred from text/format. */
  confidence: "styled" | "high" | "medium";
  reason: string;
}

const DEFAULT_CAPTION_PATTERNS: RegExp[] = [
  /^الجدول\s*رقم/,
  /^الشكل\s*رقم/,
  /^جدول\s*رقم/,
  /^شكل\s*رقم/,
  /^Table\s*\d/i,
  /^Figure\s*\d/i,
  /^Fig\.?\s*\d/i,
  /^Tableau\s*/i,
];

// NOTE: `\b` is an ASCII word boundary and does NOT trigger after Arabic
// letters — so heading patterns use a lookahead for whitespace/colon/end.
const DEFAULT_HEADING_PATTERNS: HeadingPattern[] = [
  { re: /^الفصل(?=[\s:：]|$)/, level: 1 },
  { re: /^الباب(?=[\s:：]|$)/, level: 1 },
  { re: /^Chapter\b/i, level: 1 },
  { re: /^المبحث(?=[\s:：]|$)/, level: 2 },
  { re: /^المطلب(?=[\s:：]|$)/, level: 3 },
];

const DEFAULT_LEVEL1_KEYWORDS = [
  "مقدمة", "المقدمة", "خاتمة", "الخاتمة", "قائمة المراجع", "المراجع",
  "الملاحق", "الاستنتاجات", "التوصيات", "التوصيات والاقتراحات",
  "قائمة الأشكال", "قائمة الجداول", "الفهرس",
];

const DEFAULT_SECTION_KEYWORDS = ["تمهيد", "مدخل"];

/** Trim surrounding whitespace and Arabic/Latin punctuation for keyword matching. */
function normalizeKey(s: string): string {
  return s.replace(/^[\s:：.)(»«\-]+/, "").replace(/[\s:：.)(»«\-]+$/, "").trim();
}

function matchesAny(text: string, patterns: RegExp[]): boolean {
  return patterns.some((re) => re.test(text));
}

/**
 * The section that owns paragraph `paragraphIndex`. Each intermediate section
 * ENDS at (and includes) its `paragraphIndex`; the final section catches the
 * rest. So the owner is the first section whose `paragraphIndex >= P`, else the
 * final one.
 */
function sectionIndexForParagraph(
  sections: Array<{ index: number; isFinal: boolean; paragraphIndex?: number }>,
  paragraphIndex: number,
): number {
  for (const s of sections) {
    if (s.isFinal) return s.index;
    if (s.paragraphIndex !== undefined && paragraphIndex <= s.paragraphIndex) return s.index;
  }
  return sections.length ? sections[sections.length - 1].index : 0;
}

/** Most frequent value (used to estimate the body font size). */
function mode(nums: number[]): number | null {
  if (!nums.length) return null;
  const counts = new Map<number, number>();
  let best = nums[0];
  let bestCount = 0;
  for (const n of nums) {
    const c = (counts.get(n) ?? 0) + 1;
    counts.set(n, c);
    if (c > bestCount) {
      bestCount = c;
      best = n;
    }
  }
  return best;
}

/** One image lifted out of a header/footer part that is about to be replaced,
 *  held as BYTES because `r:embed` ids are part-local and do not survive the move. */
interface CarriedChromeImage {
  oldRelId: string;
  bytes: Buffer;
  ext: string;
}

/** Artwork rescued from a header/footer part before it is overwritten. */
interface CarriedChromeDrawings {
  /** Whole `<w:p>` elements that carry a `<w:drawing>`/`<w:pict>`. */
  paragraphs: string[];
  images: CarriedChromeImage[];
}

/** Word's "Recolor" on a picture: the stored bytes are painted between two
 *  colours rather than shown as-is. `dark`/`light` are 6-hex when the file names
 *  a literal colour; `*Scheme` carries a theme slot ("accent4") to resolve
 *  against theme1.xml. `shade`/`satMod` are fractions (0.45 = 45%). */
export interface ChromeDuotone {
  dark: string | null;
  darkScheme: string | null;
  light: string | null;
  lightScheme: string | null;
  shade: number | null;
  satMod: number | null;
}

/** One placed axis of an anchored drawing. `offsetEmu` is signed — decorative
 *  full-page art is routinely offset NEGATIVELY so it overflows its anchor. */
export interface ChromeDrawingAxis {
  /** OOXML frame of reference: "page" | "margin" | "column" | "paragraph" | … */
  relativeTo: string;
  offsetEmu: number | null;
  /** Named alignment ("center", "right") when used instead of an offset. */
  align: string | null;
}

/** A picture placed in a header/footer part, with the geometry needed to paint
 *  it where Word paints it. See {@link extractChromeDrawings}. */
export interface ChromeDrawing {
  /** The part-local `r:embed` id — resolution detail, kept for debugging. */
  embedId: string;
  /** Media file name inside `word/media` (e.g. "image1.png"), null if unresolved. */
  image: string | null;
  extent: { cxEmu: number; cyEmu: number };
  /** Floating (`wp:anchor`) rather than in the text flow (`wp:inline`). */
  anchored: boolean;
  /** Word's "Behind Text" — paints under the body, not over it. */
  behindDoc: boolean;
  /** "none" | "square" | "tight" | "through" | "topAndBottom" | "inline" */
  wrap: string;
  posH: ChromeDrawingAxis;
  posV: ChromeDrawingAxis;
  duotone: ChromeDuotone | null;
  /** Alt text from `wp:docPr@descr`. */
  descr: string | null;
}

/** Extracted content of one header/footer part (internal). */
interface HeaderFooterContent {
  text: string;
  /** Tab-separated positioned parts of the running text (e.g. a right/left header
   *  renders as two segments). One entry when there are no tab stops. Empty ends
   *  trimmed. Lets a consumer lay them out positionally instead of concatenated. */
  segments: string[];
  /** The paragraph's BOTTOM border (Word's header rule), if any — `bottom` true when
   *  present, `color` the 6-hex rule colour (e.g. the brown thesis rule) or null. */
  border: { bottom: boolean; color: string | null };
  hasPage: boolean;
  /**
   * Page-number format in w:pgNumType vocabulary ("lowerRoman", …), derived
   * from the PAGE field's `\*` switch. null when there is no PAGE field or
   * its switch isn't recognized.
   */
  pageFormat: string | null;
  /** Pictures placed in the part — the decorative page frames among them. */
  drawings: ChromeDrawing[];
}

/**
 * Inverse of FooterManager's FORMAT_SWITCH: PAGE field `\*` switch token →
 * w:pgNumType format. Case-sensitive — "roman" and "ROMAN" differ.
 */
const SWITCH_TO_FORMAT: Record<string, string> = {
  ARABIC: "decimal",
  ROMAN: "upperRoman",
  roman: "lowerRoman",
  ALPHABETIC: "upperLetter",
  alphabetic: "lowerLetter",
};

/** The w:pgNumType-vocabulary format encoded in a PAGE field instruction, if any. */
function pageFormatFromInstr(instr: string): string | null {
  for (const m of instr.matchAll(/\\\*\s+([A-Za-z]+)/g)) {
    const fmt = SWITCH_TO_FORMAT[m[1]];
    if (fmt) return fmt;
  }
  return null;
}

/** Decode the five predefined XML entities in a `w:t` text node. */
function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * Plain text + PAGE-field flag of one header/footer part's XML.
 *
 * Field handling mirrors what Word RENDERS:
 *  • Complex non-page fields (e.g. a STYLEREF running header) KEEP their cached
 *    result between `separate` and `end` — that text is exactly what shows on
 *    the page.
 *  • PAGE/NUMPAGES fields — complex or `w:fldSimple` — DROP their cached
 *    result: it's a stale per-page number, not authored content ("Conf" plus a
 *    cached "1" would otherwise read back as "Conf1").
 *
 * Paragraph texts are joined with single spaces so multi-paragraph parts don't
 * concatenate ("Chapter 3Introduction"), and whitespace runs are collapsed.
 */
/** Split the running text into tab-separated positioned segments (the runs between
 *  <w:tab/> elements, in order). A two-part right/left header → two segments; no tab
 *  stops → one segment. Empty leading/trailing segments trimmed. */
function extractHeaderSegments(stripped: string): string[] {
  const textOf = (frag: string): string =>
    Array.from(frag.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g))
      .map((m) => decodeXmlEntities(m[1] ?? ""))
      .join("")
      .replace(/\s+/g, " ")
      .trim();
  // Table header (a very common two-part running head laid out with a 1-row table):
  // each cell (w:tc) is a positioned segment. Falls through to tab-splitting otherwise.
  if (/<w:tbl\b/.test(stripped)) {
    const cells = Array.from(stripped.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g))
      .map((m) => textOf(m[0]))
      .filter((s) => s.length > 0);
    if (cells.length) return cells;
  }
  // Otherwise split the runs on <w:tab/> (a tab-stopped header).
  const segs: string[] = [];
  let cur = "";
  for (const m of stripped.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\b[^>]*\/?>/g)) {
    if (m[1] !== undefined) cur += decodeXmlEntities(m[1]);
    else { segs.push(cur); cur = ""; }
  }
  segs.push(cur);
  const cleaned = segs.map((s) => s.replace(/\s+/g, " ").trim());
  while (cleaned.length && cleaned[0] === "") cleaned.shift();
  while (cleaned.length && cleaned[cleaned.length - 1] === "") cleaned.pop();
  return cleaned;
}

/** The header paragraph's BOTTOM border (Word's header rule) — from the first <w:pBdr>. */
function extractHeaderBorder(xml: string): { bottom: boolean; color: string | null } {
  // A header rule is a BOTTOM border with a real style — it may live in a paragraph
  // border (<w:pBdr>) OR a table/cell border (<w:tblBorders>/<w:tcBorders>). Scan for
  // the first <w:bottom> whose w:val isn't nil/none (skips tcMar spacing, which uses
  // w:w not w:val) and read its colour.
  for (const m of xml.matchAll(/<w:bottom\b([^>]*)>/g)) {
    const attrs = m[1] ?? "";
    const val = attrs.match(/w:val="([^"]*)"/)?.[1];
    if (!val || val === "nil" || val === "none") continue;
    const color = attrs.match(/w:color="([0-9A-Fa-f]{6})"/)?.[1] ?? null;
    return { bottom: true, color: color ? color.toUpperCase() : null };
  }
  return { bottom: false, color: null };
}

function extractHeaderFooterContent(xml: string): HeaderFooterContent {
  // The PAGE field's instruction — complex (<w:instrText>) or simple
  // (w:fldSimple/@w:instr) — drives both the flag and the format.
  const pageInstr =
    Array.from(xml.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g))
      .map((m) => m[1] ?? "")
      .find((t) => /\bPAGE\b/.test(t)) ??
    Array.from(xml.matchAll(/<w:fldSimple\b[^>]*w:instr="([^"]*)"/g))
      .map((m) => m[1] ?? "")
      .find((t) => /\bPAGE\b/.test(t)) ??
    null;
  const hasPage = pageInstr !== null;
  const pageFormat = pageInstr !== null ? pageFormatFromInstr(pageInstr) : null;

  // Complex fields: within each begin…end span, strip separate→end only when
  // the field instruction is PAGE/NUMPAGES; other fields stay untouched.
  let stripped = xml.replace(
    /<w:fldChar[^>]*w:fldCharType="begin"[^>]*\/>[\s\S]*?<w:fldChar[^>]*w:fldCharType="end"[^>]*\/>/g,
    (span) =>
      /<w:instrText[^>]*>[^<]*\b(?:PAGE|NUMPAGES)\b/.test(span)
        ? span.replace(/<w:fldChar[^>]*w:fldCharType="separate"[^>]*\/>[\s\S]*$/, "")
        : span,
  );
  // Simple fields: drop the whole element (cached digit included) when its
  // instruction is a page field; other fldSimple content stays extractable.
  stripped = stripped.replace(
    /<w:fldSimple\b[^>]*w:instr="[^"]*\b(?:PAGE|NUMPAGES)\b[^"]*"[^>]*>[\s\S]*?<\/w:fldSimple>/g,
    "",
  );

  const text = stripped
    .split("</w:p>")
    .map((chunk) =>
      Array.from(chunk.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g))
        .map((m) => decodeXmlEntities(m[1] ?? ""))
        .join(""),
    )
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();

  const segments = extractHeaderSegments(stripped);
  const border = extractHeaderBorder(xml);
  const drawings = extractChromeDrawings(xml);
  return { text, segments, border, hasPage, pageFormat, drawings };
}

/** EMU per inch — DrawingML's unit throughout `wp:extent` / `wp:posOffset`. */
export const EMU_PER_INCH = 914400;

/** Word's built-in Office colour scheme (Office 2013+). A package can omit the
 *  theme part entirely — Word then paints these, so resolving a `schemeClr`
 *  against them matches what the user sees rather than giving up. */
const OFFICE_DEFAULT_THEME_COLORS: Record<string, string> = {
  dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6",
  accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5",
  accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47",
  hlink: "0563C1", folHlink: "954F72",
};

/** `<a:srgbClr val>` / `<a:schemeClr val>` inside `frag` → a raw colour token.
 *  Scheme names stay symbolic here; {@link Doc.resolveThemeColor} maps them to hex
 *  once the theme part is available. */
function firstColorToken(frag: string): { srgb: string | null; scheme: string | null } {
  const srgb = frag.match(/<a:srgbClr[^>]*\bval="([0-9A-Fa-f]{6})"/)?.[1] ?? null;
  const scheme = frag.match(/<a:schemeClr[^>]*\bval="([A-Za-z0-9]+)"/)?.[1] ?? null;
  return { srgb: srgb ? srgb.toUpperCase() : null, scheme };
}

/**
 * Every picture in a header/footer part, as geometry a renderer can place.
 *
 * A thesis cover frame is one of these: a `<wp:anchor behindDoc="1">` picture,
 * deliberately larger than the header rectangle and negatively offset, so Word
 * paints it across the whole page behind the text. Reading only the header's
 * TEXT (which such a part often has none of) makes the page look empty.
 *
 * Offsets and extents stay in EMU — the unit the file uses — so no precision is
 * lost before the consumer knows its own scale.
 */
function extractChromeDrawings(xml: string): ChromeDrawing[] {
  const out: ChromeDrawing[] = [];
  // Anchors and inline pictures are the two placements; both wrap the same
  // <pic:pic>. Matching each frame whole keeps one drawing's geometry from
  // being read against another's blip.
  const frames = xml.matchAll(
    /<wp:(anchor|inline)\b([^>]*)>([\s\S]*?)<\/wp:(?:anchor|inline)>/g,
  );
  for (const f of frames) {
    const kind = f[1] as "anchor" | "inline";
    const attrs = f[2] ?? "";
    const body = f[3] ?? "";

    const embedId = body.match(/<a:blip[^>]*\br:embed="([^"]+)"/)?.[1] ?? null;
    if (!embedId) continue; // a shape/chart, not a picture — nothing to paint

    const cx = Number(body.match(/<wp:extent[^>]*\bcx="(\d+)"/)?.[1] ?? 0);
    const cy = Number(body.match(/<wp:extent[^>]*\bcy="(\d+)"/)?.[1] ?? 0);

    // Position: an explicit offset, or a named alignment ("center", "right"…).
    const axis = (name: "positionH" | "positionV") => {
      const m = body.match(new RegExp(`<wp:${name}[^>]*\\brelativeFrom="([^"]+)"[^>]*>([\\s\\S]*?)</wp:${name}>`));
      if (!m) return { relativeTo: name === "positionH" ? "column" : "paragraph", offsetEmu: null, align: null };
      const off = m[2]!.match(/<wp:posOffset>(-?\d+)<\/wp:posOffset>/)?.[1];
      const align = m[2]!.match(/<wp:align>([a-zA-Z]+)<\/wp:align>/)?.[1] ?? null;
      return { relativeTo: m[1]!, offsetEmu: off !== undefined ? Number(off) : null, align };
    };

    // Wrap mode. Inline pictures sit in the text flow and have no wrap element.
    const wrap =
      kind === "inline"
        ? "inline"
        : (body.match(/<wp:wrap(None|Square|Tight|Through|TopAndBottom)\b/)?.[1] ?? "None")
            .replace(/^./, (c) => c.toLowerCase());

    // Word's Recolor: the stored bytes are painted in two colours. The thesis
    // cover frame ships as a BLACK png recoloured to the theme accent — render
    // the raw bytes and the border comes out black instead of gold.
    const duo = body.match(/<a:duotone>([\s\S]*?)<\/a:duotone>/)?.[1] ?? null;
    let duotone: ChromeDuotone | null = null;
    if (duo) {
      // Two colour children, in order: shadow colour then highlight colour. Each
      // is either self-closing or wraps its own modifiers (<a:shade>, <a:satMod>)
      // — the backreference keeps a modifier's "/>" from ending the element early.
      const parts = Array.from(
        duo.matchAll(/<a:(srgbClr|schemeClr|prstClr|sysClr)\b[^>]*(?:\/>|>[\s\S]*?<\/a:\1>)/g),
      ).map((m) => m[0]);
      const darkFrag = parts[0] ?? "";
      const lightFrag = parts[1] ?? "";
      const dark = firstColorToken(darkFrag);
      const light = firstColorToken(lightFrag);
      const num = (frag: string, tag: string) => {
        const v = frag.match(new RegExp(`<a:${tag}[^>]*\\bval="(\\d+)"`))?.[1];
        return v === undefined ? null : Number(v) / 100000; // OOXML percentages are ×100000
      };
      duotone = {
        dark: dark.srgb,
        darkScheme: dark.scheme,
        light: light.srgb ?? (/<a:prstClr[^>]*val="white"/.test(lightFrag) ? "FFFFFF" : null),
        lightScheme: light.scheme,
        shade: num(darkFrag, "shade"),
        satMod: num(darkFrag, "satMod"),
      };
    }

    out.push({
      embedId,
      image: null, // resolved against the part's own _rels by the caller
      extent: { cxEmu: cx, cyEmu: cy },
      anchored: kind === "anchor",
      behindDoc: /\bbehindDoc="1"/.test(attrs),
      wrap,
      posH: axis("positionH"),
      posV: axis("positionV"),
      duotone,
      descr: body.match(/<wp:docPr[^>]*\bdescr="([^"]*)"/)?.[1] ?? null,
    });
  }
  return out;
}

/** Options accepted by paragraph/heading verbs. */
export type ParagraphFormat = Omit<StyledParagraphOptions, "outlineLevel" | "styleId"> & {
  styleId?: string;
};

/** Footer options (text and/or page numbers) for {@link Doc.setFooter}/{@link Doc.setSectionFooter}. */
export interface FooterOptions {
  text?: string;
  pageNumbers?: boolean;
  alignment?: "left" | "center" | "right";
  /** Text before the page number, e.g. "Page ". */
  prefix?: string;
  /** Render "current / total" (e.g. "3 / 40"). */
  includeTotalPages?: boolean;
}

/** Namespaces for a `<w:hdr>`/`<w:ftr>` part that may carry inline DrawingML
 *  images (logos). Includes the wordprocessing + relationship + DrawingML picture
 *  namespaces so compiled region bodies with `<w:drawing>` runs are valid. */
const CHROME_PART_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

/** One logo image embedded inside an applied header/footer part. `token` is a
 *  unique placeholder present in the region XML (e.g. inside `<a:blip r:embed>`),
 *  replaced with the part-local relationship id once the bytes are embedded. */
export type ChromeImage = { token: string; bytes: Buffer; ext: string };

/** A compiled header or footer to apply to a section. `xml` is the INNER region
 *  body (paragraphs/tables), NOT wrapped in `<w:hdr>`/`<w:ftr>` — {@link Doc.applySectionChrome}
 *  adds the wrapper + namespaces. */
export type ChromePart = { xml: string; images: ChromeImage[] };

/** Result of a section-scoped header/footer change. */
export interface SectionEditResult {
  sectionIndex: number;
  totalSections: number;
}

/** A section's page geometry in twips (1440 = 1 inch), inheritance resolved. */
export interface SectionPageGeometry {
  widthTwips: number;
  heightTwips: number;
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
    header: number;
    footer: number;
    gutter: number;
  };
}

const DEFAULT_MARGIN_TWIPS = 1440; // 1in
const DEFAULT_HF_TWIPS = 720;      // 0.5in — header/footer distance from the page edge
const DEFAULT_GUTTER_TWIPS = 0;    // no binding allowance unless the document asks for one

/**
 * Resolve one section's page geometry, falling back to the body sectPr's.
 *
 * This fallback is NOT an ECMA-376 rule — the spec does not make a section
 * inherit another section's w:pgSz/w:pgMar; an omitted one falls back to the
 * *application* default (this repo's own parseSectPr encodes that as Letter,
 * 12240x15840). The fallback exists because addSectionBreak writes a bare
 * `<w:sectPr><w:type/></w:sectPr>` with no geometry of its own — that gap is
 * ours, not the spec's, and addSectionBreak arguably ought to write full
 * geometry instead. Until then, treating the body sectPr's geometry as this
 * section's is what lets the app paginate against the real page size.
 */
export function resolveSectionPageGeometry(
  entry: SectionEntry,
  bodyEntry: SectionEntry | undefined,
): SectionPageGeometry | null {
  const size = entry.pageSize ?? bodyEntry?.pageSize;
  const mar = entry.margins ?? bodyEntry?.margins;
  if (!size) return null;
  return {
    widthTwips: size.width,
    heightTwips: size.height,
    margins: {
      top: mar?.top ?? DEFAULT_MARGIN_TWIPS,
      bottom: mar?.bottom ?? DEFAULT_MARGIN_TWIPS,
      left: mar?.left ?? DEFAULT_MARGIN_TWIPS,
      right: mar?.right ?? DEFAULT_MARGIN_TWIPS,
      header: mar?.header ?? DEFAULT_HF_TWIPS,
      footer: mar?.footer ?? DEFAULT_HF_TWIPS,
      gutter: mar?.gutter ?? DEFAULT_GUTTER_TWIPS,
    },
  };
}

export interface SectionInfo {
  /** Section position in document order (0-based; same order as getSections()). */
  index: number;
  /** Block index (document.getBlocks() order) of the section's first block. */
  startBlockIndex: number;
  /**
   * Effective running header text — the section's own default part, else the
   * previous section's (ECMA-376 inheritance). null = no header anywhere in
   * the chain; "" = an explicitly blank header part.
   */
  headerText: string | null;
  /** The effective header's tab-separated positioned segments (e.g. a right/left
   *  header → two entries) for faithful rendering. null when there's no header. */
  headerSegments: string[] | null;
  /** The effective header paragraph's bottom rule (Word's header line): `bottom`
   *  true when present + its 6-hex `color`. null when there's no header. */
  headerBorder: { bottom: boolean; color: string | null } | null;
  /** Pictures in the effective header part. A full-page decorative frame lives
   *  here — anchored, `behindDoc`, and usually the part's ONLY content, so a
   *  section can have artwork while `headerText` is "". */
  headerDrawings: ChromeDrawing[];
  /** Effective footer text (same inheritance rules). */
  footerText: string | null;
  /** Pictures in the effective footer part (same inheritance rules). */
  footerDrawings: ChromeDrawing[];
  /** True when the effective footer part contains a PAGE field. */
  footerHasPageNumbers: boolean;
  /**
   * Page-number format in w:pgNumType vocabulary ("decimal", "lowerRoman", …).
   * The section's own w:pgNumType format wins when set; otherwise the format
   * is derived from the effective footer's PAGE field `\*` switch (the normal
   * insertion path writes only the switch), which travels with the inherited
   * part. null when neither exists.
   */
  pageNumberFormat: string | null;
  /** This section's own w:pgNumType start value, if set. */
  pageNumberStart: number | null;
  /**
   * Page size + margins for this section, in twips, via {@link resolveSectionPageGeometry}.
   * A section's own w:sectPr wins; the fallback to the body sectPr's geometry
   * is not an ECMA-376 rule (an omitted w:pgSz's real fallback is the
   * application default, e.g. Letter) — it exists because addSectionBreak
   * writes a bare `<w:sectPr><w:type/></w:sectPr>` with no geometry of its
   * own, which is our own gap and arguably something addSectionBreak ought
   * to fix by writing full geometry. null only when the body sectPr declares
   * no page size at all.
   */
  page: SectionPageGeometry | null;
}

type BreakType = "nextPage" | "evenPage" | "oddPage";

/** A generated, always-accurate structural map of the document. */
export interface DocMap {
  title: string;
  wordCount: number;
  counts: { paragraphs: number; headings: number; tables: number; images: number; sections: number };
  page: { width: number; height: number; orientation: string };
  margins: { top: number; right: number; bottom: number; left: number };
  hasHeader: boolean;
  hasFooter: boolean;
  rtl: boolean;
  outline: OutlineNode[];
}

export class Doc {
  /** Escape hatch: the underlying engine + all its managers. */
  readonly engine: Mdocxengine;

  /** theme1.xml's colour scheme, parsed once per Doc (slot → 6-hex). */
  private themeColors: Map<string, string> | null = null;

  private constructor(engine: Mdocxengine) {
    this.engine = engine;
  }

  /** Open a document from a file path or an in-memory buffer. */
  static async open(source: string | Buffer): Promise<Doc> {
    const engine =
      typeof source === "string"
        ? await Mdocxengine.loadFromFile(source)
        : await Mdocxengine.loadFromBuffer(source);
    return new Doc(engine);
  }

  /** Wrap an already-loaded engine. */
  static from(engine: Mdocxengine): Doc {
    return new Doc(engine);
  }

  // ─── Read ──────────────────────────────────────────────────────────────────

  /** Whole-document plain text (blocks joined by `separator`, default newline). */
  async text(separator = "\n"): Promise<string> {
    return this.engine.document.getPlainText(separator);
  }

  /** Total word count of the body. */
  async wordCount(): Promise<number> {
    return this.engine.document.getWordCount();
  }

  /** Every body block as plain data, in document order. */
  async blocks(): Promise<BlockInfo[]> {
    const raw = await this.engine.document.getBlocks();
    return raw.map((b, index) => this.toBlockInfo(b, index));
  }

  private toBlockInfo(b: BodyBlock, index: number): BlockInfo {
    if (b.kind === "table") {
      return { index, kind: "table", text: "", styleId: null, headingLevel: 0 };
    }
    if (b.kind === "paragraph") {
      const isImage = b.xml.includes("<w:drawing>");
      return {
        index,
        kind: isImage ? "image" : "paragraph",
        text: isImage ? "" : paragraphText(b.xml),
        styleId: paragraphStyleId(b.xml),
        headingLevel: paragraphHeadingLevel(b.xml),
      };
    }
    return { index, kind: "other", text: "", styleId: null, headingLevel: 0 };
  }

  /** The heading outline as a nested tree (like a table of contents). */
  async outline(): Promise<OutlineNode[]> {
    const flat = (await this.blocks())
      .filter((b) => b.headingLevel > 0 && b.text.trim())
      .map((b) => ({ index: b.index, level: b.headingLevel, title: b.text.trim() }));

    const roots: OutlineNode[] = [];
    const stack: OutlineNode[] = [];
    for (const h of flat) {
      const node: OutlineNode = { ...h, children: [] };
      while (stack.length && stack[stack.length - 1].level >= h.level) stack.pop();
      (stack.length ? stack[stack.length - 1].children : roots).push(node);
      stack.push(node);
    }
    return roots;
  }

  /** Every table as a plain text grid. */
  async tables(): Promise<TableInfo[]> {
    const raw = await this.engine.document.getBlocks();
    const out: TableInfo[] = [];
    for (let i = 0; i < raw.length; i++) {
      if (raw[i].kind !== "table") continue;
      const t = await Table.fromXml(raw[i].xml);
      out.push({ index: i, rows: t.getAllCellText() });
    }
    return out;
  }

  /** Every embedded inline image as bytes + size + mime. */
  async images(): Promise<ImageInfo[]> {
    const raw = await this.engine.document.getBlocks();
    const out: ImageInfo[] = [];
    for (let i = 0; i < raw.length; i++) {
      if (!raw[i].xml.includes("<w:drawing>")) continue;
      const img = await this.engine.media.extractInlineImage(raw[i].xml);
      if (img) out.push({ index: i, ...img });
    }
    return out;
  }

  /**
   * Every body block enriched with formatting signals (bold, font size,
   * alignment, word count, caption flag). The raw material for inferring
   * structure when a document has no heading markup — feed it to a heuristic
   * (`inferOutline`) or to an LLM to label headings.
   */
  async blocksDetailed(captionPatterns: RegExp[] = DEFAULT_CAPTION_PATTERNS): Promise<DetailedBlockInfo[]> {
    const raw = await this.engine.document.getBlocks();
    return raw.map((b, index) => {
      const base = this.toBlockInfo(b, index);
      const wordCount = base.text ? base.text.split(/\s+/).filter(Boolean).length : 0;
      if (b.kind === "paragraph" && base.kind === "paragraph") {
        return {
          ...base,
          bold: paragraphIsBold(b.xml),
          fontSizePt: paragraphFontSizePt(b.xml),
          alignment: paragraphAlignment(b.xml),
          wordCount,
          looksLikeCaption: matchesAny(base.text.trim(), captionPatterns),
        };
      }
      return { ...base, bold: false, fontSizePt: null, alignment: null, wordCount, looksLikeCaption: false };
    });
  }

  /**
   * Infer a heading outline for documents whose titles are plain text (no
   * heading styles) — the common case for imported / copy-pasted theses.
   *
   * Real headings (styled or with an outline level) are reported as
   * `confidence: "styled"`. The rest are inferred from text patterns
   * (`^الفصل`, `^المبحث`…), section keywords (`تمهيد`, `قائمة المراجع`…), and
   * formatting (bold / larger-than-body font / centered + short line), while
   * figure/table captions are excluded. All rules are overridable.
   *
   * This is a best-effort heuristic; for messy documents, pass
   * `blocksDetailed()` to an LLM and call `setHeadingLevel()` with its labels.
   */
  async inferOutline(opts: InferOutlineOptions = {}): Promise<InferredHeading[]> {
    const captionPatterns = opts.captionPatterns ?? DEFAULT_CAPTION_PATTERNS;
    const headingPatterns = opts.headingPatterns ?? DEFAULT_HEADING_PATTERNS;
    const level1 = new Set((opts.level1Keywords ?? DEFAULT_LEVEL1_KEYWORDS).map(normalizeKey));
    const sections = new Set((opts.sectionKeywords ?? DEFAULT_SECTION_KEYWORDS).map(normalizeKey));
    const maxWords = opts.maxHeadingWords ?? 14;

    const detailed = await this.blocksDetailed(captionPatterns);
    // Estimate the body font as the most common inline size among long paragraphs.
    const bodyFont = mode(
      detailed.filter((b) => b.kind === "paragraph" && b.fontSizePt != null && b.wordCount > 20).map((b) => b.fontSizePt!),
    );

    const out: InferredHeading[] = [];
    for (const b of detailed) {
      if (b.kind !== "paragraph") continue;
      const title = b.text.trim();
      if (!title) continue;

      if (b.headingLevel > 0) {
        out.push({ index: b.index, level: b.headingLevel, title, confidence: "styled", reason: "heading style / outline level" });
        continue;
      }
      if (b.looksLikeCaption) continue;

      const pat = headingPatterns.find((p) => p.re.test(title));
      if (pat) {
        out.push({ index: b.index, level: pat.level, title, confidence: "high", reason: "matches heading pattern" });
        continue;
      }

      const key = normalizeKey(title);
      const firstWord = normalizeKey(title.split(/\s+/)[0] ?? "");
      if (level1.has(key) || level1.has(firstWord)) {
        out.push({ index: b.index, level: 1, title, confidence: "high", reason: "top-level keyword" });
        continue;
      }
      if (sections.has(key) || sections.has(firstWord)) {
        out.push({ index: b.index, level: 2, title, confidence: "high", reason: "section keyword" });
        continue;
      }

      const short = b.wordCount > 0 && b.wordCount <= maxWords;
      const larger = bodyFont != null && b.fontSizePt != null && b.fontSizePt > bodyFont;
      const centered = b.alignment === "center";
      if (short && (b.bold || larger)) {
        const strong = b.bold && (larger || centered);
        out.push({
          index: b.index,
          level: 2,
          title,
          confidence: strong ? "high" : "medium",
          reason: [b.bold && "bold", larger && "larger-font", centered && "centered", "short line"].filter(Boolean).join(", "),
        });
      }
    }
    return out;
  }

  // ─── Write (verbs) ───────────────────────────────────────────────────────────

  /** Append a paragraph (or insert at `at`). */
  async addParagraph(text: string, opts: ParagraphFormat = {}, at = APPEND): Promise<this> {
    await this.engine.document.insertBlockAt(makeStyledParagraphNode(text, opts), at);
    return this;
  }

  /** Append a heading at `level` (1–6) → maps to `Heading{level}`, bold, with outline level. */
  async addHeading(text: string, level = 1, opts: ParagraphFormat = {}, at = APPEND): Promise<this> {
    const lvl = Math.min(6, Math.max(1, level));
    const block = makeStyledParagraphNode(text, {
      styleId: `Heading${lvl}`,
      outlineLevel: lvl - 1,
      bold: true,
      ...opts,
    });
    await this.engine.document.insertBlockAt(block, at);
    return this;
  }

  /** Replace the text of the paragraph at `index` (preserves its formatting). */
  async editParagraph(index: number, text: string): Promise<this> {
    await this.engine.document.editParagraphText(index, text);
    return this;
  }

  /**
   * Promote (or demote) the paragraph at `index` to a real heading at `level`
   * (1–6): applies the `Heading{level}` style + matching outline level, keeping
   * the text and preserving RTL/alignment. This is how you turn a plain-text
   * title — e.g. one found by `inferOutline()` or labelled by an LLM — into a
   * structural heading the outline/TOC can see.
   */
  async setHeadingLevel(index: number, level: number): Promise<this> {
    const blocks = await this.engine.document.getBlocks();
    const b = blocks[index];
    if (!b || b.kind !== "paragraph" || b.xml.includes("<w:drawing>")) {
      throw new Error(`setHeadingLevel: no text paragraph at block index ${index}`);
    }
    const lvl = Math.min(6, Math.max(1, level));
    const rtl = /<w:bidi\b/.test(b.xml) || /<w:rtl\b/.test(b.xml);
    const alignment = paragraphAlignment(b.xml) as StyledParagraphOptions["alignment"] | null;
    blocks[index] = makeStyledParagraphNode(paragraphText(b.xml), {
      styleId: `Heading${lvl}`,
      outlineLevel: lvl - 1,
      rtl,
      ...(alignment ? { alignment } : {}),
    });
    await this.engine.document.saveBlocks(blocks);
    return this;
  }

  /**
   * Apply run-level formatting to one or more named PARTS of the document —
   * `body`, `headings` (or `heading1`…`heading6`), `title`, `captions`, `lists`,
   * `tables`, `footnotes`.
   *
   * Style-level with a strip: each target's Word style is ensured and patched,
   * then the named property is removed from that target's runs so the style
   * shows through (imported theses carry formatting on the RUNS, which would
   * otherwise win). Paragraphs that would not resolve to the target's patched
   * style get a direct write instead. Properties that were not named are never
   * touched.
   *
   * Returns one report per target, so a caller can tell a no-op from a change.
   *
   * @throws on an unknown target name, and on a malformed `styles.xml`. A
   * malformed individual RUN is skipped and counted, never aborting the pass.
   */
  async setTextStyle(
    targets: readonly TextStyleTargetInput[],
    props: RunProps,
  ): Promise<TargetReport[]> {
    const expanded = expandTargets(targets as readonly string[]);
    const infos = await this.blocks();
    return new TextStyleManager(this.engine.zip).apply(expanded, props, infos);
  }

  /**
   * READ the font / size / bold / italic / colour actually in force on the same
   * named PARTS `setTextStyle` writes to — `body`, `headings`, `title`,
   * `captions`, `lists`, `tables`, `footnotes`.
   *
   * Resolves the full OOXML cascade (docDefaults → paragraph style through its
   * `w:basedOn` chain → character style → direct `w:rPr`) and weighs every
   * answer by characters, so a thesis whose body is 96% Simplified Arabic 14
   * reports exactly that rather than whichever value happened to be first.
   * Latin and complex-script properties stay separate (`font`/`fontCs`,
   * `sizePt`/`sizeCsPt`): in an Arabic thesis it is `w:cs` that the reader sees.
   *
   * A target with `paragraphs: 0` is absent from the document, which is
   * information, not an error. Read-only — nothing is written.
   *
   * @throws on an unknown target name (same validation as `setTextStyle`).
   */
  async getTextStyle(
    targets: readonly TextStyleTargetInput[] = ["body", "headings", "title", "captions", "lists", "tables", "footnotes"],
  ): Promise<TextStyleInspection> {
    const expanded = expandTargets(targets as readonly string[]);
    const infos = await this.blocks();
    return new TextStyleReader(this.engine.zip).inspect(expanded, infos);
  }

  /** Delete the block at `index`. */
  async deleteBlock(index: number): Promise<this> {
    await this.engine.document.deleteBlockAt(index);
    return this;
  }

  /** Append a table from a row-major grid (or insert at `at`). */
  async addTable(
    rows: string[][],
    opts: { header?: boolean; rtl?: boolean } = {},
    at = APPEND,
  ): Promise<this> {
    const block = makeTableNode(rows, { headerRow: opts.header, rtl: opts.rtl });
    await this.engine.document.insertBlockAt(block, at);
    return this;
  }

  /** Set the text of one cell of the table at block `index`. */
  async editTableCell(index: number, row: number, col: number, value: string): Promise<this> {
    return this.mutateTable(index, (t) => t.setCellText(row, col, value));
  }

  // Shared load→mutate→save round-trip for the table at `index` (mirrors the old
  // editTableCell inline body). ANY Table mutation preserves per-cell formatting
  // because it edits the parsed TableObject, then re-serializes the whole <w:tbl>.
  private async mutateTable(index: number, fn: (t: Table) => void): Promise<this> {
    const blocks = await this.engine.document.getBlocks();
    const block = blocks[index];
    if (!block || block.kind !== "table") throw new Error(`No table at block index ${index}`);
    const table = await Table.fromXml(block.xml);
    fn(table);
    blocks[index] = {
      ...block,
      xml: XmlUtils.buildXml(table.toObject(), { rootName: "w:tbl", headless: true, pretty: false }),
    };
    await this.engine.document.saveBlocks(blocks);
    return this;
  }

  /**
   * Insert a row into the table at `index`: below row `at` (0-based), ABOVE it
   * when `before` is true (e.g. a new first row: at=0, before=true), or
   * appended when `at` is omitted.
   */
  async addTableRow(index: number, at?: number, before?: boolean): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (at != null && at >= 0 && at < t.getRowCount()) {
        if (before) t.insertRowAbove(at);
        else t.insertRowBelow(at);
      } else t.addRow();
    });
  }

  /** Remove row `row` (0-based) from the table at `index`. */
  async removeTableRow(index: number, row: number): Promise<this> {
    return this.mutateTable(index, (t) => t.removeRow(row));
  }

  /**
   * Insert a column into the table at `index`: to the right of column `at`
   * (0-based), to its LEFT when `before` is true (e.g. a new first column:
   * at=0, before=true), or appended when omitted.
   */
  async insertTableColumn(index: number, at?: number, before?: boolean): Promise<this> {
    return this.mutateTable(index, (t) => {
      const cols = t.getColumnCount();
      const target = at != null && at >= 0 && at < cols ? at : Math.max(0, cols - 1);
      if (before && at != null) t.insertColumnLeft(target);
      else t.insertColumnRight(target);
    });
  }

  /** Sort the table's data rows by column `col` (numeric when both parse). Header row 0 stays put unless includeHeader. */
  async sortTable(index: number, col: number, opts?: { desc?: boolean; includeHeader?: boolean }): Promise<this> {
    return this.mutateTable(index, (t) => {
      t.sortByColumn(col, opts?.desc ? "desc" : "asc", !(opts?.includeHeader ?? false));
    });
  }

  /**
   * Column/table widths: `columnsTwips[i]` sets column i's preferred width
   * (1440 twips = 1 inch; ~600 per cm); `autofit` switches the layout mode.
   */
  async setTableWidths(index: number, opts: { columnsTwips?: number[]; autofit?: "contents" | "window" }): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (opts.autofit === "contents") t.autoFitContents();
      else if (opts.autofit === "window") t.autoFitWindow();
      if (opts.columnsTwips) {
        for (let c = 0; c < opts.columnsTwips.length; c++) {
          const w = opts.columnsTwips[c];
          if (w && w > 0) t.setColumnWidth(c, Math.round(w));
        }
      }
    });
  }

  /**
   * Format a cell (row+col) or a whole row (row only, col omitted) — existing
   * text preserved: bold/italic, font size (points), font family, vertical
   * alignment inside the cell, uniform cell padding (twips).
   */
  async formatTableCellText(
    index: number,
    opts: {
      row?: number;
      col?: number;
      bold?: boolean;
      italic?: boolean;
      sizePt?: number;
      fontFamily?: string;
      vAlign?: "top" | "center" | "bottom";
      paddingTwips?: number;
    },
  ): Promise<this> {
    return this.mutateTable(index, (t) => {
      const row = opts.row ?? 0;
      const cols = t.getColumnCount();
      const targets = opts.col != null ? [opts.col] : Array.from({ length: cols }, (_, c) => c);
      for (const c of targets) {
        if (opts.bold !== undefined || opts.italic !== undefined || opts.sizePt !== undefined || opts.fontFamily !== undefined) {
          t.setCellTextFormat(row, c, {
            bold: opts.bold,
            italic: opts.italic,
            sizeHalfPoints: opts.sizePt != null ? toHalfPoints(opts.sizePt, "setTableCellFormat sizePt") : undefined,
            fontFamily: opts.fontFamily,
          });
        }
        if (opts.vAlign) t.setCellVerticalAlignment(row, c, opts.vAlign);
        if (opts.paddingTwips != null) {
          const p = Math.round(opts.paddingTwips);
          t.setCellMargins(row, c, { top: p, bottom: p, left: p, right: p });
        }
      }
    });
  }

  /** Merge cells: horizontal (row + startCol..endCol) or vertical (col + startRow..endRow). */
  async mergeTableCells(
    index: number,
    opts:
      | { direction: "horizontal"; row: number; start: number; end: number }
      | { direction: "vertical"; col: number; start: number; end: number },
  ): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (opts.direction === "horizontal") t.mergeCellsHorizontal(opts.row, opts.start, opts.end);
      else t.mergeCellsVertical(opts.col, opts.start, opts.end);
    });
  }

  /** Delete column `col` (0-based) from the table at `index`. */
  async deleteTableColumn(index: number, col: number): Promise<this> {
    return this.mutateTable(index, (t) => t.deleteColumn(col));
  }

  /**
   * Table-level layout + styling. All fields optional — pass what changes:
   * alignment/direction/header (row 0) as before; `borders` true/false for
   * simple single/none, or `{ style?, sizePt?, color?, sides? }` for custom
   * borders (sides defaults to all six); `widthPct` (10..100) sets the table
   * width as a page percentage; `indentTwips` indents from the margin;
   * `wrap` lets body text flow around the table; `styleId` applies a named
   * Word table style (must exist in the doc's styles.xml); `rowHeightTwips`
   * (+ optional `row`, default all rows) sets row height; `distributeRows` /
   * `distributeColumns` even out sizes; `allowRowBreaks` toggles rows
   * splitting across pages; `altTitle`/`altDescription` set accessibility
   * alt text.
   */
  async setTableLayout(
    index: number,
    opts: {
      alignment?: "left" | "center" | "right";
      direction?: "rtl" | "ltr";
      headerRow?: boolean;
      headerFill?: string;
      borders?: boolean | { style?: string; sizePt?: number; color?: string; sides?: ("top" | "bottom" | "left" | "right" | "insideH" | "insideV")[] };
      widthPct?: number;
      indentTwips?: number;
      wrap?: "none" | "around";
      styleId?: string;
      rowHeightTwips?: number;
      row?: number;
      distributeRows?: boolean;
      distributeColumns?: boolean;
      allowRowBreaks?: boolean;
      altTitle?: string;
      altDescription?: string;
    },
  ): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (opts.alignment) t.setTableAlignment(opts.alignment);
      if (opts.direction) t.setTableDirection(opts.direction === "rtl");
      // headerFill alone also marks+shades row 0 (e.g. "make the header orange").
      if (opts.headerRow || opts.headerFill) t.setHeaderRow(0, opts.headerFill);
      if (opts.borders != null) {
        if (typeof opts.borders === "boolean") {
          const side = opts.borders ? { style: "single", size: 4 } : { style: "none" };
          t.setTableBorders({ top: side, bottom: side, left: side, right: side, insideH: side, insideV: side });
        } else {
          const b = opts.borders;
          const side = {
            style: b.style ?? "single",
            size: Math.max(2, Math.round((b.sizePt ?? 0.5) * 8)),
            color: (b.color ?? "000000").replace("#", ""),
          };
          const sides = b.sides ?? ["top", "bottom", "left", "right", "insideH", "insideV"];
          const all: Record<string, typeof side> = {};
          for (const s of sides) all[s] = side;
          t.setTableBorders(all);
        }
      }
      if (opts.widthPct != null) t.setTableWidth(Math.round(Math.max(10, Math.min(100, opts.widthPct)) * 50), "pct");
      if (opts.indentTwips != null) t.setTableIndent(Math.round(opts.indentTwips));
      if (opts.wrap) t.setTextWrapping(opts.wrap);
      if (opts.styleId) t.setTableStyle(opts.styleId);
      if (opts.rowHeightTwips != null) {
        const rows = t.getRowCount();
        const targets = opts.row != null ? [opts.row] : Array.from({ length: rows }, (_, r) => r);
        for (const r of targets) t.setRowHeight(r, Math.round(opts.rowHeightTwips));
      }
      if (opts.distributeRows) t.distributeRows();
      if (opts.distributeColumns) t.distributeColumns();
      if (opts.allowRowBreaks != null) {
        for (let r = 0; r < t.getRowCount(); r++) t.setRowAllowBreak(r, opts.allowRowBreaks);
      }
      if (opts.altTitle != null) t.setAltText(opts.altTitle, opts.altDescription ?? "");
    });
  }

  /** Split a merged cell back apart: horizontal (gridSpan) or vertical (vMerge chain). */
  async splitTableCells(index: number, opts: { direction: "horizontal" | "vertical"; row: number; col: number }): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (opts.direction === "horizontal") t.splitCellHorizontal(opts.row, opts.col);
      else t.splitCellVertical(opts.row, opts.col);
    });
  }

  /** Move a row or a column from one position to another (0-based). */
  async moveTableLine(index: number, opts: { kind: "row" | "column"; from: number; to: number }): Promise<this> {
    return this.mutateTable(index, (t) => {
      if (opts.kind === "row") t.moveRow(opts.from, opts.to);
      else t.moveColumn(opts.from, opts.to);
    });
  }

  /**
   * Style one cell (row+col), a whole row (row only), or the header row
   * (neither): `fill` = 6-hex background, `textColor` = 6-hex font colour of
   * the existing text. Pass either or both.
   */
  async shadeTable(index: number, opts: { row?: number; col?: number; fill?: string; textColor?: string }): Promise<this> {
    return this.mutateTable(index, (t) => {
      const row = opts.row ?? 0;
      const cols = t.getColumnCount();
      const targets = opts.col != null ? [opts.col] : Array.from({ length: cols }, (_, c) => c);
      for (const c of targets) {
        if (opts.fill) t.setCellShading(row, c, opts.fill.replace("#", ""));
        if (opts.textColor) t.setCellTextColor(row, c, opts.textColor);
      }
    });
  }

  /**
   * Apply styling grids: fills[r][c] = 6-hex background, textColors[r][c] =
   * 6-hex font colour for that cell's existing text; null/undefined = leave
   * as-is. Either grid may be omitted.
   */
  async shadeTableCells(
    index: number,
    fills?: (string | null | undefined)[][] | null,
    textColors?: (string | null | undefined)[][] | null,
  ): Promise<this> {
    return this.mutateTable(index, (t) => {
      const rows = t.getRowCount();
      const cols = t.getColumnCount();
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const fill = fills?.[r]?.[c];
          if (fill) t.setCellShading(r, c, fill.replace("#", ""));
          const tc = textColors?.[r]?.[c];
          if (tc) t.setCellTextColor(r, c, tc);
        }
      }
    });
  }

  /** Append an image from bytes (or insert at `at`). Size in pixels @96dpi. */
  async addImage(
    bytes: Buffer,
    opts: { format?: string; width?: number; height?: number; at?: number } = {},
  ): Promise<this> {
    const format = (opts.format ?? "png").toLowerCase();
    const widthEmu = pixelsToEmu(opts.width ?? 400);
    const heightEmu = pixelsToEmu(opts.height ?? 300);
    const { relId } = await this.engine.media.insertImage(bytes, format);
    const docXml = this.engine.zip.getFileAsString("word/document.xml") ?? "";
    const drawingId = nextDrawingId(docXml);
    const block = makeDrawingParagraphNode(relId, widthEmu, heightEmu, drawingId, `Image ${drawingId}`);
    await this.engine.document.insertBlockAt(block, opts.at ?? APPEND);
    return this;
  }

  /** Render Markdown and append the result (headings, paragraphs, lists, tables). */
  async addMarkdown(blocks: BodyBlock[], at = APPEND): Promise<this> {
    // Accepts pre-built BodyBlocks (e.g. from a markdown renderer) and appends them.
    const target = await this.engine.document.getBlocks();
    const insertAt = at === APPEND ? target.length : at;
    await this.engine.document.saveBlocks([
      ...target.slice(0, insertAt),
      ...blocks,
      ...target.slice(insertAt),
    ]);
    return this;
  }

  /** Find-and-replace across the document body (e.g. fill `{{tokens}}`). */
  async replaceText(find: string | RegExp, replace: string): Promise<this> {
    await this.engine.document.findAndReplaceAll(find, replace);
    return this;
  }

  // ─── Page setup / layout (chainable) ─────────────────────────────────────────

  async setPageSize(preset: "A4" | "USLetter" | "Legal" | "A3" | "A5", orientation: "portrait" | "landscape" = "portrait"): Promise<this> {
    await this.engine.pageLayout.setPageSizePreset(preset as any, orientation);
    return this;
  }

  async setMargins(margins: { top?: number; right?: number; bottom?: number; left?: number }): Promise<this> {
    await this.engine.pageLayout.setMargins(margins as any);
    return this;
  }

  async setOrientation(orientation: "portrait" | "landscape"): Promise<this> {
    await this.engine.pageLayout.setOrientation(orientation);
    return this;
  }

  /** Page numbers in the footer (centered decimal by default). */
  async addPageNumbers(opts: { alignment?: "left" | "center" | "right"; format?: "decimal" | "lowerRoman" | "upperRoman" } = {}): Promise<this> {
    const f = await this.engine.footer.addFooter("", "default");
    await this.engine.footer.insertPageNumber(f.footerPath, {
      alignment: opts.alignment ?? "center",
      format: opts.format ?? "decimal",
    });
    return this;
  }

  /** Roman front-matter / arabic body footers, split at `bodyStartIndex`. */
  async frontMatterNumbering(bodyStartIndex: number): Promise<this> {
    await this.engine.applyFrontMatterNumbering({ bodyStartParaIndex: bodyStartIndex });
    return this;
  }

  /**
   * Set ONE document-wide page header (top of every page). Empty text removes it.
   * Replaces any existing header (no duplicates) and clears per-section headers.
   */
  async setHeader(text: string): Promise<this> {
    const existing = this.engine.header.getAllheadersFiles(this.engine.zip);
    for (const h of existing) await this.engine.header.removeHeader(h.fileName);
    const t = (text ?? "").trim();
    if (t) await this.engine.header.addHeader(t, "default", undefined, { registerInSectPr: true });
    return this;
  }

  /**
   * Set ONE document-wide page footer (text and/or page numbers). Empty text +
   * `pageNumbers:false` removes it. Replaces any existing footer and clears
   * per-section footers.
   */
  async setFooter(opts: FooterOptions = {}): Promise<this> {
    const existing = this.engine.footer.getAllFooterFiles(this.engine.zip);
    for (const f of existing) await this.engine.footer.removeFooter(f.fileName);
    const t = (opts.text ?? "").trim();
    if (!t && !opts.pageNumbers) return this; // removed
    const { footerPath } = await this.engine.footer.addFooter(t, "default", undefined, { registerInSectPr: true });
    if (opts.pageNumbers) {
      await this.engine.footer.insertPageNumber(footerPath, {
        alignment: opts.alignment ?? "center",
        prefix: opts.prefix || undefined,
        includeTotalPages: !!opts.includeTotalPages,
      });
    }
    return this;
  }

  /**
   * Make the block at `blockIndex` begin on a new page by inserting a section
   * break just before it (so e.g. each chapter starts a fresh page and can own
   * its header/footer). Returns `{ changed:false }` if the block is already the
   * first content. Block index ↔ paragraph index is handled internally.
   */
  async startOnNewPage(blockIndex: number, breakType: BreakType = "nextPage"): Promise<{ changed: boolean }> {
    const blocks = await this.engine.document.getBlocks();
    if (blockIndex < 0 || blockIndex >= blocks.length) {
      throw new Error(`startOnNewPage: block index ${blockIndex} out of range (0..${blocks.length - 1})`);
    }
    const paraIndex = blocks.slice(0, blockIndex).filter((b) => b.kind === "paragraph").length;
    if (paraIndex === 0) return { changed: false };
    await this.engine.sections.addSectionBreak(paraIndex - 1, breakType);
    return { changed: true };
  }

  /**
   * Set the page size / margins of the section CONTAINING the block at
   * `blockIndex` (byte-safe; only the named attributes change).
   */
  async setSectionGeometry(blockIndex: number, geometry: SectPrGeometry): Promise<SectionEditResult> {
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);
    await this.engine.sections.setSectionGeometry(sectionIndex, geometry);
    return { sectionIndex, totalSections: sections.length };
  }

  /**
   * Give the section CONTAINING the block at `blockIndex` its own running header
   * (call `startOnNewPage` on that heading first to make it its own section).
   * Empty text → a blank header for that section. Cleans up the section's
   * previous distinct header part.
   */
  async setSectionHeader(blockIndex: number, text: string): Promise<SectionEditResult> {
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);
    const oldRelId = sections[sectionIndex]?.headerRefs?.find((h) => h.type === "default")?.relId;
    // The old part's artwork must be READ before the new part replaces it — the
    // cleanup at the end of this method deletes the file it lives in.
    const carried = await this.readChromeDrawings(oldRelId);
    const { headerPath, relId } = await this.engine.header.addHeader((text ?? "").trim(), "default", undefined, { registerInSectPr: false });
    await this.carryChromeDrawings("header", headerPath, carried);
    await this.engine.sections.setSectionHeader(sectionIndex, relId, "default");
    if (oldRelId && oldRelId !== relId) await this.removeHeaderFooterByRel("header", oldRelId);
    return { sectionIndex, totalSections: sections.length };
  }

  /**
   * Give the section CONTAINING the block at `blockIndex` its own footer (text
   * and/or page numbers). Cleans up the section's previous distinct footer part.
   */
  async setSectionFooter(blockIndex: number, opts: FooterOptions = {}): Promise<SectionEditResult> {
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);
    const oldRelId = sections[sectionIndex]?.footerRefs?.find((f) => f.type === "default")?.relId;
    const carried = await this.readChromeDrawings(oldRelId);
    const { footerPath, relId } = await this.engine.footer.addFooter((opts.text ?? "").trim(), "default", undefined, { registerInSectPr: false });
    await this.carryChromeDrawings("footer", footerPath, carried);
    if (opts.pageNumbers) {
      await this.engine.footer.insertPageNumber(footerPath, {
        alignment: opts.alignment ?? "center",
        prefix: opts.prefix || undefined,
        includeTotalPages: !!opts.includeTotalPages,
      });
    }
    await this.engine.sections.setSectionFooter(sectionIndex, relId, "default");
    if (oldRelId && oldRelId !== relId) await this.removeHeaderFooterByRel("footer", oldRelId);
    return { sectionIndex, totalSections: sections.length };
  }

  /**
   * Vertically align the content of the section CONTAINING the block at
   * `blockIndex`. "center" places a divider page's title in the middle of the
   * page instead of at the top. Call `startOnNewPage` on that block first so it
   * is its own section, or this aligns whatever section it falls in.
   */
  async setSectionVerticalAlign(
    blockIndex: number,
    vAlign: "top" | "center" | "both" | "bottom",
  ): Promise<SectionEditResult> {
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);
    await this.engine.sections.setSectionVerticalAlign(sectionIndex, vAlign);
    return { sectionIndex, totalSections: sections.length };
  }

  /**
   * Draw a page border around the section CONTAINING the block at `blockIndex`
   * (the `frame` divider family). Overwrites any border that section had.
   */
  async setSectionPageBorders(
    blockIndex: number,
    opts: SectPrPageBorderOptions,
  ): Promise<SectionEditResult> {
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);
    await this.engine.sections.setSectionPageBorders(sectionIndex, opts);
    return { sectionIndex, totalSections: sections.length };
  }

  /**
   * Apply a COMPILED header and/or footer (raw OOXML region bodies) to the section
   * that contains the block at `blockIndex`, embedding any logo images into the
   * header/footer part's OWN relationships so they resolve in Word. This is how a
   * saved Header/Footer Studio template is applied onto a live thesis at full
   * fidelity (tab-stop segments, tables, live fields, logos).
   *
   * Each {@link ChromePart}.`xml` is the inner region body — this method wraps it in
   * `<w:hdr>`/`<w:ftr>` with the DrawingML namespaces. Every image `token` in that
   * xml is replaced with the part-local `r:embed` id once its bytes are embedded.
   * The section's previous distinct header/footer part (if any) is removed.
   */
  async applySectionChrome(
    blockIndex: number,
    parts: { header?: ChromePart; footer?: ChromePart },
  ): Promise<{ sectionIndex: number; warnings: string[] }> {
    const warnings: string[] = [];
    const { sections, sectionIndex } = await this.resolveSection(blockIndex);

    if (parts.header) {
      const oldRelId = sections[sectionIndex]?.headerRefs?.find((h) => h.type === "default")?.relId;
      let xml = `<w:hdr ${CHROME_PART_NS}>${parts.header.xml}</w:hdr>`;
      const { headerPath, relId } = await this.engine.header.addHeader("", "default", xml, { registerInSectPr: false });
      for (const img of parts.header.images) {
        const rid = await this.engine.media.addImageToPartRels(headerPath, img.bytes, img.ext);
        xml = xml.split(img.token).join(rid);
      }
      if (parts.header.images.length) this.engine.header.updateHeader(headerPath, xml);
      await this.engine.sections.setSectionHeader(sectionIndex, relId, "default");
      if (oldRelId && oldRelId !== relId) await this.removeHeaderFooterByRel("header", oldRelId);
    }

    if (parts.footer) {
      const oldRelId = sections[sectionIndex]?.footerRefs?.find((f) => f.type === "default")?.relId;
      let xml = `<w:ftr ${CHROME_PART_NS}>${parts.footer.xml}</w:ftr>`;
      const { footerPath, relId } = await this.engine.footer.addFooter("", "default", xml, { registerInSectPr: false });
      for (const img of parts.footer.images) {
        const rid = await this.engine.media.addImageToPartRels(footerPath, img.bytes, img.ext);
        xml = xml.split(img.token).join(rid);
      }
      if (parts.footer.images.length) this.engine.footer.updateFooter(footerPath, xml);
      await this.engine.sections.setSectionFooter(sectionIndex, relId, "default");
      if (oldRelId && oldRelId !== relId) await this.removeHeaderFooterByRel("footer", oldRelId);
    }

    return { sectionIndex, warnings };
  }

  /** Map a block index → its owning section (+ the full section list). */
  private async resolveSection(blockIndex: number): Promise<{ sections: SectionEntry[]; sectionIndex: number }> {
    const blocks = await this.engine.document.getBlocks();
    if (blockIndex < 0 || blockIndex >= blocks.length) {
      throw new Error(`block index ${blockIndex} out of range (0..${blocks.length - 1})`);
    }
    const paraIndex = blocks.slice(0, blockIndex).filter((b) => b.kind === "paragraph").length;
    const sections = await this.engine.sections.getSections();
    return { sections, sectionIndex: sectionIndexForParagraph(sections, paraIndex) };
  }

  /**
   * Per-section header/footer info, read-only companion to setSectionHeader /
   * setSectionFooter. Inheritance is resolved the way Word renders it
   * (ECMA-376): a section without its own reference uses the previous
   * section's part; a first section without one has none.
   *
   * @param preloadedBlocks Pass blocks you already fetched from
   *   `engine.document.getBlocks()` to avoid a second body parse; they must be
   *   CURRENT for this document state.
   */
  async sections(preloadedBlocks?: BodyBlock[]): Promise<SectionInfo[]> {
    const [blocks, entries] = await Promise.all([
      preloadedBlocks ? Promise.resolve(preloadedBlocks) : this.engine.document.getBlocks(),
      this.engine.sections.getSections(),
    ]);
    // Section boundaries live on paragraphs; map paragraph index → block index.
    const paraToBlock: number[] = [];
    blocks.forEach((b, i) => {
      if (b.kind === "paragraph") paraToBlock.push(i);
    });

    const out: SectionInfo[] = [];
    // Sections often share one part (that's what inheritance references) —
    // read each part at most once per call.
    const partCache = new Map<string, HeaderFooterContent | null>();
    let header: HeaderFooterContent | null = null;
    let footer: HeaderFooterContent | null = null;

    // The body sectPr — found by isFinal, not position, so this survives any
    // reordering of getSections(). Falling back to its geometry for a section
    // that omits w:pgSz/w:pgMar is not ECMA-376 behaviour — see
    // resolveSectionPageGeometry's doc comment — it works around the bare
    // sectPr addSectionBreak writes, not the spec.
    const bodyEntry = entries.find((e) => e.isFinal) ?? entries[entries.length - 1];

    for (let k = 0; k < entries.length; k++) {
      const prev = entries[k - 1];
      const prevBreakBlock =
        prev?.paragraphIndex !== undefined ? paraToBlock[prev.paragraphIndex] : undefined;
      const startBlockIndex =
        k === 0 ? 0 : Math.min((prevBreakBlock ?? -1) + 1, blocks.length);

      const own = entries[k];
      const [ownHeader, ownFooter] = await Promise.all([
        this.readHeaderFooterPart(own.headerRefs, partCache),
        this.readHeaderFooterPart(own.footerRefs, partCache),
      ]);
      if (ownHeader) header = ownHeader;
      if (ownFooter) footer = ownFooter;

      out.push({
        index: k,
        startBlockIndex,
        headerText: header ? header.text : null,
        headerSegments: header ? header.segments : null,
        headerBorder: header ? header.border : null,
        headerDrawings: header ? header.drawings : [],
        footerText: footer ? footer.text : null,
        footerDrawings: footer ? footer.drawings : [],
        footerHasPageNumbers: !!footer?.hasPage,
        // Explicit pgNumType wins; else the format the effective footer's PAGE
        // field renders with (it travels with the inherited part).
        pageNumberFormat: own.pageNumberType?.format ?? footer?.pageFormat ?? null,
        pageNumberStart: own.pageNumberType?.start ?? null,
        page: resolveSectionPageGeometry(own, bodyEntry),
      });
    }
    return out;
  }

  /**
   * Content of the DEFAULT-type header/footer part behind `refs` (first/even
   * page-only refs are not the running chrome and are ignored), memoized by
   * relId in `cache` so sections sharing a part read it once per
   * {@link sections} call. null when no part resolves — including on any
   * read/parse failure, so chrome extraction can never throw.
   */
  private async readHeaderFooterPart(
    refs: SectionHeaderFooterRef[],
    cache: Map<string, HeaderFooterContent | null>,
  ): Promise<HeaderFooterContent | null> {
    const ref = refs.find((r) => r.type === "default");
    if (!ref?.relId) return null;
    if (cache.has(ref.relId)) return cache.get(ref.relId) ?? null;
    let content: HeaderFooterContent | null = null;
    try {
      const target = await this.engine.rels.getTarget(ref.relId);
      if (target) {
        const path = target.startsWith("word/") ? target : `word/${target.replace(/^\/+/, "")}`;
        const xml = this.engine.zip.readAsText(path);
        if (xml) {
          content = extractHeaderFooterContent(xml);
          // `r:embed` resolves against the PART's own rels, not the document's,
          // and a scheme colour only means something against the theme — both
          // need the engine, so they are filled in here rather than in the
          // pure extractor.
          if (content.drawings.length) await this.resolveChromeDrawings(path, content.drawings);
        }
      }
    } catch {
      content = null;
    }
    cache.set(ref.relId, content);
    return content;
  }

  /**
   * Fill in what {@link extractChromeDrawings} could not know from the part XML
   * alone: which media file each `r:embed` points at (resolved against the
   * part's OWN `_rels`), and the hex behind any theme-slot recolour.
   *
   * Mutates in place — the drawings were just built for this content object and
   * have no other reader yet.
   */
  private async resolveChromeDrawings(partPath: string, drawings: ChromeDrawing[]): Promise<void> {
    const partName = partPath.replace(/^word\//, "");
    const rels = new RelManager(this.engine.zip, `word/_rels/${partName}.rels`);
    for (const d of drawings) {
      const target = await rels.getTarget(d.embedId).catch(() => null);
      if (target) d.image = target.replace(/^.*\//, "");
      if (d.duotone) {
        if (!d.duotone.dark && d.duotone.darkScheme) {
          d.duotone.dark = await this.resolveThemeColor(d.duotone.darkScheme);
        }
        if (!d.duotone.light && d.duotone.lightScheme) {
          d.duotone.light = await this.resolveThemeColor(d.duotone.lightScheme);
        }
      }
    }
  }

  /** Theme slot ("accent4", "dk1"…) → 6-hex from theme1.xml's `<a:clrScheme>`.
   *  `<a:sysClr>` carries the resolved value in `lastClr`. Slots the document
   *  does not define fall back to Word's built-in Office palette, which is what
   *  Word itself paints for a package with no theme part. null when the slot is
   *  not a colour-scheme name at all. */
  private async resolveThemeColor(slot: string): Promise<string | null> {
    if (!this.themeColors) {
      const map = new Map<string, string>(Object.entries(OFFICE_DEFAULT_THEME_COLORS));
      try {
        const xml = this.engine.zip.readAsText("word/theme/theme1.xml") ?? "";
        const scheme = xml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/)?.[0] ?? "";
        for (const m of scheme.matchAll(
          /<a:(\w+)>\s*<a:(?:srgbClr\s+val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")/g,
        )) {
          const hex = (m[2] ?? m[3])!.toUpperCase();
          map.set(m[1]!, hex);
        }
      } catch {
        // Unreadable theme — the Office defaults above still stand in.
      }
      // Word's document-level aliases for the first two pairs.
      if (map.has("dk1")) map.set("tx1", map.get("dk1")!);
      if (map.has("lt1")) map.set("bg1", map.get("lt1")!);
      if (map.has("dk2")) map.set("tx2", map.get("dk2")!);
      if (map.has("lt2")) map.set("bg2", map.get("lt2")!);
      this.themeColors = map;
    }
    return this.themeColors.get(slot) ?? null;
  }

  /** Best-effort delete of a header/footer part by its relationship id (cleanup). */
  private async removeHeaderFooterByRel(which: "header" | "footer", relId: string): Promise<void> {
    const target = await this.engine.rels.getTarget(relId);
    if (!target) return;
    const path = target.startsWith("word/") ? target : `word/${target.replace(/^\/+/, "")}`;
    if (which === "header") await this.engine.header.removeHeader(path).catch(() => {});
    else await this.engine.footer.removeFooter(path).catch(() => {});
  }

  /**
   * Read the artwork-bearing paragraphs out of the header/footer part behind
   * `relId`, together with the bytes each one's images resolve to.
   *
   * Setting a section's header/footer text builds a BRAND-NEW part and deletes
   * the old one — which silently destroyed full-page decorative frames, the
   * near-universal shape of an Algerian thesis cover (a `<wp:anchor behindDoc>`
   * picture living in the header, drawn behind the whole page). The student's
   * only clue was the border vanishing. So the artwork is lifted out first and
   * replanted by {@link carryChromeDrawings} into the replacement part.
   *
   * Images are carried as BYTES, not relationship ids: `r:embed` resolves against
   * the part's own `_rels`, so an id from the old part means nothing in the new
   * one and would leave Word showing a repair prompt.
   */
  private async readChromeDrawings(relId: string | undefined): Promise<CarriedChromeDrawings> {
    const empty: CarriedChromeDrawings = { paragraphs: [], images: [] };
    if (!relId) return empty;
    try {
      const target = await this.engine.rels.getTarget(relId);
      if (!target) return empty;
      const path = target.startsWith("word/") ? target : `word/${target.replace(/^\/+/, "")}`;
      const xml = this.engine.zip.readAsText(path);
      if (!xml) return empty;

      // Only whole paragraphs that carry a drawing — a run holding an anchor is
      // meaningless outside the paragraph Word anchors it to.
      const paragraphs = (xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []).filter((p) =>
        /<w:drawing\b|<w:pict\b/.test(p),
      );
      if (paragraphs.length === 0) return empty;

      // Resolve every image the carried paragraphs reference against the OLD
      // part's own rels, keeping the bytes so they can be re-embedded.
      const partName = path.replace(/^word\//, "");
      const rels = new RelManager(this.engine.zip, `word/_rels/${partName}.rels`);
      const ids = new Set(
        paragraphs.flatMap((p) => Array.from(p.matchAll(/r:embed="([^"]+)"/g)).map((m) => m[1]!)),
      );
      const images: CarriedChromeImage[] = [];
      for (const id of ids) {
        const mediaTarget = await rels.getTarget(id).catch(() => null);
        if (!mediaTarget) continue;
        const name = mediaTarget.replace(/^.*\//, "");
        const bytes = this.engine.media.extractImage(name);
        if (!bytes) continue;
        images.push({ oldRelId: id, bytes, ext: name.replace(/^.*\./, "") || "png" });
      }
      return { paragraphs, images };
    } catch {
      // Preservation is best-effort: a part we cannot read must not fail the edit.
      return empty;
    }
  }

  /**
   * Replant the artwork {@link readChromeDrawings} lifted out of the part being
   * replaced, re-embedding each image into the NEW part's own relationships and
   * rewriting its `r:embed` to the id that resolves there.
   *
   * The paragraphs are appended, so the artwork keeps its z-order behind the new
   * text: an anchored `behindDoc` picture paints behind regardless of document
   * order, and an inline logo reads as trailing content rather than displacing
   * the text the caller just set.
   */
  private async carryChromeDrawings(
    which: "header" | "footer",
    partPath: string,
    carried: CarriedChromeDrawings,
  ): Promise<void> {
    if (carried.paragraphs.length === 0) return;
    try {
      const xml = this.engine.zip.readAsText(partPath);
      if (!xml) return;
      let block = carried.paragraphs.join("");
      for (const img of carried.images) {
        const rid = await this.engine.media.addImageToPartRels(partPath, img.bytes, img.ext);
        block = block.split(`r:embed="${img.oldRelId}"`).join(`r:embed="${rid}"`);
      }
      const close = which === "header" ? "</w:hdr>" : "</w:ftr>";
      const at = xml.lastIndexOf(close);
      if (at === -1) return;
      const merged = xml.slice(0, at) + block + xml.slice(at);
      if (which === "header") this.engine.header.updateHeader(partPath, merged);
      else this.engine.footer.updateFooter(partPath, merged);
    } catch {
      // Same contract as the read half — never fail the caller's edit.
    }
  }

  // ─── Describe (generated map) ────────────────────────────────────────────────

  /** A generated, always-accurate structural map of the document. */
  async describe(): Promise<DocMap> {
    const blocks = await this.blocks();
    const counts = { paragraphs: 0, headings: 0, tables: 0, images: 0, sections: 0 };
    for (const b of blocks) {
      if (b.kind === "table") counts.tables++;
      else if (b.kind === "image") counts.images++;
      else if (b.kind === "paragraph") counts.paragraphs++;
      if (b.headingLevel > 0) counts.headings++;
    }

    const [meta, page, margins, sections, wordCount, outline] = await Promise.all([
      this.engine.metadata.getCoreProperties().catch(() => ({} as any)),
      this.engine.pageLayout.getPageSize().catch(() => ({ width: 0, height: 0, orientation: "portrait" } as any)),
      this.engine.pageLayout.getMargins().catch(() => ({ top: 0, right: 0, bottom: 0, left: 0 } as any)),
      this.engine.sections.getSections().then((s) => s.length).catch(() => 1),
      this.wordCount(),
      this.outline(),
    ]);
    counts.sections = sections;

    const entries = this.engine.zip.getEntries().map((e) => e.entryName);
    const hasHeader = entries.some((n) => /^word\/header\d+\.xml$/.test(n));
    const hasFooter = entries.some((n) => /^word\/footer\d+\.xml$/.test(n));
    const docXml = this.engine.zip.getFileAsString("word/document.xml") ?? "";
    const rtl = /<w:bidi\b/.test(docXml) || /<w:rtl\b/.test(docXml);

    return {
      title: meta.title ?? "",
      wordCount,
      counts,
      page: { width: page.width, height: page.height, orientation: page.orientation },
      margins: { top: margins.top, right: margins.right, bottom: margins.bottom, left: margins.left },
      hasHeader,
      hasFooter,
      rtl,
      outline,
    };
  }

  /** The structural map rendered as human-readable Markdown. */
  async toMarkdownMap(): Promise<string> {
    const m = await this.describe();
    const lines: string[] = [];
    lines.push(`# ${m.title || "Document"} — map`);
    lines.push("");
    lines.push(
      `**${m.counts.paragraphs}** paragraphs · **${m.counts.headings}** headings · ` +
        `**${m.counts.tables}** tables · **${m.counts.images}** images · ` +
        `**${m.counts.sections}** sections · **${m.wordCount}** words`,
    );
    lines.push(
      `Page: ${m.page.width}×${m.page.height} (${m.page.orientation}) · ` +
        `Header: ${m.hasHeader ? "yes" : "no"} · Footer: ${m.hasFooter ? "yes" : "no"} · ` +
        `RTL: ${m.rtl ? "yes" : "no"}`,
    );
    lines.push("");
    lines.push("## Outline");
    if (!m.outline.length) lines.push("_(no headings)_");
    const walk = (nodes: OutlineNode[], depth: number) => {
      for (const n of nodes) {
        lines.push(`${"  ".repeat(depth)}- H${n.level} ${n.title}  _(block ${n.index})_`);
        walk(n.children, depth + 1);
      }
    };
    walk(m.outline, 0);
    return lines.join("\n");
  }

  // ─── Save ────────────────────────────────────────────────────────────────────

  /** Write the document to a file. */
  async save(outputPath: string): Promise<void> {
    await this.engine.saveToFile(outputPath);
  }

  /** Get the document as an in-memory buffer. */
  toBuffer(): Buffer {
    return this.engine.zip.toBuffer();
  }
}
