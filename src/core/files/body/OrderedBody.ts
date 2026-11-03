/**
 * OrderedBody — order-faithful, string-level parse/serialize for the document
 * body.
 *
 * ## Why a string splitter (not an XML library)
 *
 * The previous implementation used `fast-xml-parser` in `preserveOrder` mode.
 * On a real-world template (El Bayadh STAPS) that parser MIS-NESTED two sibling
 * `<w:tbl>` elements *inside* paragraph 0's `<w:r>` run. `editParagraphText`
 * then rebuilt that paragraph's children and silently DISCARDED the runs — so
 * two of four tables vanished on an unrelated paragraph edit. Identity
 * round-trip looked fine; an actual edit lost data.
 *
 * This module instead treats `word/document.xml` as a string and splits the
 * `<w:body>` into its top-level children, keeping each child as its **exact
 * original XML substring**. Untouched blocks are never re-serialized, so:
 *   - identity round-trip is byte-perfect by construction, and
 *   - a paragraph edit can only ever rewrite that one paragraph's substring;
 *     it is physically impossible to corrupt a sibling table or image.
 *
 * The scanner is depth-aware and handles self-closing tags, attribute values
 * that contain `>`/`<` inside quotes, XML comments and processing
 * instructions. `pre + blocks.map(b => b.xml).join("") + post` always equals
 * the original `document.xml` byte-for-byte when nothing is edited.
 *
 * It is intentionally self-contained: it does NOT touch the existing
 * `Paragraph` / `Table` classes or `DocumentManager.saveChanges` (the
 * xml2js-based path used elsewhere).
 */

import { ownPPr } from "@/core/ooxml/ownPPr";
import * as XmlUtils from "@/utils/xmlUtils";
import { Table } from "@/core/files/table/index";
import { toHalfPoints } from "@/core/ooxml/hps";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// DrawingML / picture / wordprocessingDrawing namespace URIs. Declared LOCALLY
// on the drawing block so it is self-contained regardless of the host document.
const NS_A_DML = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WP_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";

export type BlockKind = "paragraph" | "table" | "sectPr" | "other";

/**
 * A single ordered body child, stored as its exact original XML substring.
 * `xml` is the verbatim slice of `document.xml` for this top-level child —
 * including any whitespace/comment text captured as an "other" block.
 */
export interface BodyBlock {
  kind: BlockKind;
  tag: string;
  xml: string;
}

/** The result of splitting a full `word/document.xml`. */
export interface SplitDocument {
  /** Everything up to and including the `<w:body ...>` open tag. */
  pre: string;
  /** The `<w:body ...>` open tag itself (subset of `pre`, for convenience). */
  bodyOpen: string;
  /** Top-level children of `<w:body>`, in document order. */
  blocks: BodyBlock[];
  /** `</w:body>` and everything after it (trailing whitespace, `</w:document>`, …). */
  post: string;
}

// ─── Tag classification ──────────────────────────────────────────────────────

/** Classify a (possibly namespaced) tag name into a block kind. */
export function kindOf(tag: string): BlockKind {
  if (tag === "w:p") return "paragraph";
  if (tag === "w:tbl") return "table";
  if (tag === "w:sectPr") return "sectPr";
  return "other";
}

// ─── XML entity (de)coding ───────────────────────────────────────────────────

/**
 * Remove characters that are illegal in XML 1.0 even when escaped — the C0
 * control chars except tab/newline/CR, plus the non-characters U+FFFE/U+FFFF.
 * Inserted text (e.g. from an AI) can carry a stray form-feed (U+000C) that makes
 * `document.xml` unparseable ("PCDATA invalid Char value 12") and breaks strict
 * readers like docx-preview / OnlyOffice. Stripping here keeps the part valid.
 */
// eslint-disable-next-line no-control-regex
const INVALID_XML_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g;
export function stripInvalidXmlChars(text: string): string {
  return text.replace(INVALID_XML_CHARS, "");
}

/** Escape the 5 predefined XML entities in element text (invalid chars stripped first). */
export function escapeXmlText(text: string): string {
  return stripInvalidXmlChars(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Decode the 5 predefined XML entities (and numeric refs) in element text. */
export function decodeXmlText(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_m, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

// ─── String scanner ──────────────────────────────────────────────────────────

/**
 * Read an element/PI/comment tag name starting at `xml[i]` where `xml[i] === '<'`.
 * Returns the bare tag name (e.g. `w:p`), or "" for comments/PIs/closing tags.
 */
function readTagName(xml: string, ltIndex: number): string {
  let j = ltIndex + 1;
  // closing tag, comment, PI, CDATA, doctype → not an opening element
  if (xml[j] === "/" || xml[j] === "!" || xml[j] === "?") return "";
  let name = "";
  while (j < xml.length) {
    const ch = xml[j];
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === ">" || ch === "/") {
      break;
    }
    name += ch;
    j++;
  }
  return name;
}

/**
 * Given `xml` and the index of a `<` that begins a tag, return the index just
 * PAST the closing `>` of that tag (handling quoted attribute values that may
 * contain `>` / `<`). For comments/PIs/CDATA, skips to the proper terminator.
 */
function endOfTag(xml: string, ltIndex: number): number {
  // XML comment
  if (xml.startsWith("<!--", ltIndex)) {
    const end = xml.indexOf("-->", ltIndex + 4);
    return end === -1 ? xml.length : end + 3;
  }
  // CDATA section
  if (xml.startsWith("<![CDATA[", ltIndex)) {
    const end = xml.indexOf("]]>", ltIndex + 9);
    return end === -1 ? xml.length : end + 3;
  }
  // Processing instruction or declaration that we treat generically: scan for
  // the matching `>`, respecting quoted strings.
  let i = ltIndex + 1;
  let quote: string | null = null;
  while (i < xml.length) {
    const ch = xml[i];
    if (quote) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      return i + 1;
    }
    i++;
  }
  return xml.length;
}

/** True if the tag whose `<` is at `ltIndex` is self-closing (`<w:x .../>`). */
function isSelfClosing(xml: string, ltIndex: number, tagEnd: number): boolean {
  // tagEnd points just past `>`. The char before `>` is at tagEnd-2.
  return xml[tagEnd - 2] === "/";
}

/**
 * Split a full `word/document.xml` into the body region's top-level children,
 * each preserved as its exact original substring.
 *
 * Guarantee: `pre + blocks.map(b => b.xml).join("") + post === documentXml`.
 */
export function splitDocument(documentXml: string): SplitDocument {
  // Locate the <w:body ...> open tag (with optional attributes) or a
  // self-closing <w:body/>.
  const bodyOpenMatch = /<w:body(?:\s[^>]*?)?(\/?)>/.exec(documentXml);
  if (!bodyOpenMatch) {
    throw new Error("OrderedBody.splitDocument: no <w:body> found");
  }
  const bodyOpenStart = bodyOpenMatch.index;
  const bodyOpenEnd = bodyOpenStart + bodyOpenMatch[0].length;
  const selfClosingBody = bodyOpenMatch[1] === "/";

  if (selfClosingBody) {
    // <w:body/> → no children.
    return {
      pre: documentXml.slice(0, bodyOpenEnd),
      bodyOpen: bodyOpenMatch[0],
      blocks: [],
      post: documentXml.slice(bodyOpenEnd),
    };
  }

  // Find the matching </w:body>. <w:body> cannot nest, so the first one is it.
  const bodyCloseIdx = documentXml.indexOf("</w:body>", bodyOpenEnd);
  if (bodyCloseIdx === -1) {
    throw new Error("OrderedBody.splitDocument: no closing </w:body> found");
  }

  const pre = documentXml.slice(0, bodyOpenEnd);
  const post = documentXml.slice(bodyCloseIdx);
  const inner = documentXml.slice(bodyOpenEnd, bodyCloseIdx);

  const blocks = scanTopLevelChildren(inner);
  return { pre, bodyOpen: bodyOpenMatch[0], blocks, post };
}

/**
 * Scan a body's inner XML into top-level children, depth-aware.
 * Any inter-element text/whitespace/comments are captured as "other" blocks so
 * the join is byte-lossless.
 *
 * Exported for reuse by other string-level scanners in this package (e.g.
 * `core/ooxml/canonicalOrder.ts`) that need the same comment/CDATA/PI-aware,
 * same-tag-name depth tracking this module already gets right — rather than
 * having them duplicate it, imperfectly, on their own.
 */
export function scanTopLevelChildren(inner: string): BodyBlock[] {
  const blocks: BodyBlock[] = [];
  let i = 0;
  const n = inner.length;

  while (i < n) {
    const lt = inner.indexOf("<", i);
    if (lt === -1) {
      // Trailing text after the last element — keep it lossless.
      const text = inner.slice(i);
      if (text.length > 0) blocks.push({ kind: "other", tag: "#text", xml: text });
      break;
    }
    if (lt > i) {
      // Text/whitespace between top-level children → "other" block.
      blocks.push({ kind: "other", tag: "#text", xml: inner.slice(i, lt) });
      i = lt;
      continue;
    }

    // We're at a `<`.
    if (inner.startsWith("<!--", lt)) {
      const end = endOfTag(inner, lt);
      blocks.push({ kind: "other", tag: "#comment", xml: inner.slice(lt, end) });
      i = end;
      continue;
    }
    if (inner.startsWith("<?", lt)) {
      const end = endOfTag(inner, lt);
      blocks.push({ kind: "other", tag: "#pi", xml: inner.slice(lt, end) });
      i = end;
      continue;
    }
    if (inner.startsWith("<![CDATA[", lt)) {
      const end = endOfTag(inner, lt);
      blocks.push({ kind: "other", tag: "#cdata", xml: inner.slice(lt, end) });
      i = end;
      continue;
    }

    const tag = readTagName(inner, lt);
    if (!tag) {
      // Unexpected stray closing tag or malformed — capture generically.
      const end = endOfTag(inner, lt);
      blocks.push({ kind: "other", tag: "#raw", xml: inner.slice(lt, end) });
      i = end;
      continue;
    }

    const openTagEnd = endOfTag(inner, lt);
    if (isSelfClosing(inner, lt, openTagEnd)) {
      blocks.push({ kind: kindOf(tag), tag, xml: inner.slice(lt, openTagEnd) });
      i = openTagEnd;
      continue;
    }

    // Normal open tag — scan forward, depth-aware, to its matching close.
    const elementEnd = findMatchingClose(inner, openTagEnd, tag);
    blocks.push({ kind: kindOf(tag), tag, xml: inner.slice(lt, elementEnd) });
    i = elementEnd;
  }

  return blocks;
}

/**
 * Whether a body child counts as an indexable block.
 *
 * Excludes the trailing w:sectPr AND whitespace-only #text runs (left behind by
 * pretty-printing round-trips) — counting the latter as blocks would corrupt
 * every consumer's block indices. The excluded blocks stay in the document on
 * byte-safe reassembly; they're just never indexable.
 *
 * Exported because block indices are a CONTRACT with consumers outside this
 * package (the server's page-map renderer indexes blocks the same way, and a
 * private copy of this rule would drift into wrong page numbers).
 */
export function isEditableBlock(b: BodyBlock): boolean {
  if (b.kind === "sectPr") return false;
  if (b.kind === "other" && b.tag === "#text" && b.xml.trim() === "") return false;
  return true;
}

/**
 * Starting just past the open tag of `<tag>` (at `from`), find the index just
 * past the matching `</tag>`, tracking depth across same-named nested elements
 * and skipping comments / PIs / CDATA / self-closing tags / quoted attrs.
 */
function findMatchingClose(inner: string, from: string | number, tag: string): number {
  // `from` is the numeric index just past the open tag's `>`.
  let i = from as number;
  let depth = 1;
  const openMarker = `<${tag}`;
  const closeMarker = `</${tag}>`;
  const n = inner.length;

  while (i < n) {
    const lt = inner.indexOf("<", i);
    if (lt === -1) break;

    if (inner.startsWith("<!--", lt) || inner.startsWith("<?", lt) || inner.startsWith("<![CDATA[", lt)) {
      i = endOfTag(inner, lt);
      continue;
    }

    if (inner.startsWith(closeMarker, lt)) {
      depth--;
      const end = lt + closeMarker.length;
      if (depth === 0) return end;
      i = end;
      continue;
    }

    // Is this an opening of the SAME tag (and not just a prefix like
    // `<w:tblPr` when tag is `<w:tbl`)? Require the char after the name to be a
    // delimiter.
    if (inner.startsWith(openMarker, lt)) {
      const after = inner[lt + openMarker.length];
      const isSameTag =
        after === " " || after === "\t" || after === "\n" || after === "\r" || after === ">" || after === "/";
      const tagEnd = endOfTag(inner, lt);
      if (isSameTag && !isSelfClosing(inner, lt, tagEnd)) {
        depth++;
      }
      i = tagEnd;
      continue;
    }

    // Any other tag — skip past it (we only track depth on the SAME tag name,
    // which is sufficient because XML is well-formed and elements are properly
    // nested; a `</tag>` can only close a same-named ancestor).
    i = endOfTag(inner, lt);
  }

  // Malformed — return end so we don't loop forever.
  return n;
}

/** Reassemble a split document into a full `document.xml` string. */
export function assembleDocument(split: {
  pre: string;
  blocks: BodyBlock[];
  post: string;
}): string {
  return split.pre + split.blocks.map((b) => b.xml).join("") + split.post;
}

// ─── Paragraph construction ──────────────────────────────────────────────────

/**
 * Build a `<w:p>` paragraph XML string from plain text + optional style/RTL:
 * `<w:p>(<w:pPr>(<w:pStyle w:val="ID"/>)(<w:bidi/>)</w:pPr>)?<w:r><w:t xml:space="preserve">ESCAPED</w:t></w:r></w:p>`
 */
export function makeParagraphXml(text: string, styleId?: string, rtl?: boolean): string {
  let pPr = "";
  if (styleId || rtl) {
    const style = styleId ? `<w:pStyle w:val="${escapeXmlText(styleId)}"/>` : "";
    const bidi = rtl ? "<w:bidi/>" : "";
    pPr = `<w:pPr>${style}${bidi}</w:pPr>`;
  }
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`;
}

/**
 * Build a paragraph `BodyBlock` (kept under the legacy export name
 * `makeParagraphNode` for API stability). Now string-based.
 */
export function makeParagraphNode(text: string, styleId?: string, rtl?: boolean): BodyBlock {
  return { kind: "paragraph", tag: "w:p", xml: makeParagraphXml(text, styleId, rtl) };
}

/** A `<w:sectPr>` living inside a `<w:pPr>` — a section break, never cloned. */
const PPR_SECTPR_RE = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g;

/**
 * Build a `<w:p>` carrying `text` but wearing `templateXml`'s clothes: its whole
 * `<w:pPr>` (style, alignment, indent, spacing, bidi, numbering) plus the run
 * properties of its first text-bearing run.
 *
 * Use this — NOT `makeParagraphXml`/`makeStyledParagraphXml` — whenever a new
 * paragraph is added beside existing content (split, range rewrite, AI insert).
 * Rebuilding from a styleId alone only carries what the style sheet defines, and
 * in an imported thesis that is usually nothing: the formatting lives on the
 * runs. The result is a new paragraph that renders identically to the one it was
 * derived from, instead of falling back to renderer defaults.
 *
 * The template's `<w:sectPr>` is dropped — cloning it would duplicate a section
 * break and repaginate the document.
 */
/**
 * A run's content for `text`: `<w:t>` pieces with a real `<w:tab/>` wherever
 * the text holds "\t". The mirror of paragraphText's tab reading — a TOC entry
 * edited as "label\tpage" must come back as `label<w:tab/>page`, or the tab
 * STOP it was written for has nothing to act on and Word shows a glued line.
 * One run may hold several `<w:t>` and `<w:tab/>` children (CT_R allows it).
 */
function runContentXml(text: string): string {
  return text
    .split("\t")
    .map((piece) => (piece ? `<w:t xml:space="preserve">${escapeXmlText(piece)}</w:t>` : ""))
    .join("<w:tab/>");
}

export function makeParagraphXmlLike(templateXml: string, text: string): string {
  const pPr = paragraphPropsXml(templateXml).replace(PPR_SECTPR_RE, "");
  const rPr = paragraphRunPropsXml(templateXml);
  return `<w:p>${pPr}<w:r>${rPr}${runContentXml(text)}</w:r></w:p>`;
}

/** {@link makeParagraphXmlLike} as a `BodyBlock`. */
export function makeParagraphNodeLike(templateXml: string, text: string): BodyBlock {
  return { kind: "paragraph", tag: "w:p", xml: makeParagraphXmlLike(templateXml, text) };
}

/** Options for {@link makeStyledParagraphXml} — a fully-formatted single-run paragraph. */
export interface StyledParagraphOptions {
  styleId?: string;
  alignment?: "left" | "center" | "right" | "both";
  outlineLevel?: number;
  bold?: boolean;
  italic?: boolean;
  /** Font size in POINTS (emitted as half-points). */
  fontSizePt?: number;
  fontFamily?: string;
  /** Hex colour, with or without leading '#'. */
  color?: string;
  /** Right-to-left: emits paragraph `<w:bidi/>` and run `<w:rtl/>`. */
  rtl?: boolean;
}

/**
 * Build a fully-formatted single-run `<w:p>` XML STRING (byte-safe for the
 * OrderedBody block path — unlike `Paragraph.toXml()`, which wraps in `<root>`).
 * `<w:pPr>`/`<w:rPr>` children are emitted in canonical OOXML order. Empty `text`
 * yields a property-only spacer paragraph.
 */
export function makeStyledParagraphXml(text: string, opts: StyledParagraphOptions = {}): string {
  // `w:pPr` is an ORDERED sequence (CT_PPrBase): pStyle(1) … bidi(19) … jc(28)
  // … outlineLvl(31). Emitting bidi last — after jc and outlineLvl — makes Word
  // refuse the file, and it only ever showed on RTL paragraphs, i.e. every
  // Arabic thesis this builder writes a heading or paragraph into.
  const pPrParts: string[] = [];
  if (opts.styleId) pPrParts.push(`<w:pStyle w:val="${escapeXmlText(opts.styleId)}"/>`);
  if (opts.rtl) pPrParts.push(`<w:bidi/>`);
  if (opts.alignment) pPrParts.push(`<w:jc w:val="${opts.alignment}"/>`);
  if (opts.outlineLevel !== undefined) pPrParts.push(`<w:outlineLvl w:val="${opts.outlineLevel}"/>`);
  const pPr = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";

  const rPrParts: string[] = [];
  if (opts.fontFamily) {
    const f = escapeXmlText(opts.fontFamily);
    rPrParts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`);
  }
  if (opts.bold) rPrParts.push(`<w:b/>`);
  if (opts.italic) rPrParts.push(`<w:i/>`);
  if (opts.fontSizePt) {
    const hp = toHalfPoints(opts.fontSizePt, "makeStyledParagraphXml fontSizePt");
    rPrParts.push(`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`);
  }
  if (opts.color) rPrParts.push(`<w:color w:val="${escapeXmlText(opts.color.replace("#", ""))}"/>`);
  if (opts.rtl) rPrParts.push(`<w:rtl/>`);
  const rPr = rPrParts.length ? `<w:rPr>${rPrParts.join("")}</w:rPr>` : "";

  const run = text ? `<w:r>${rPr}${runContentXml(text)}</w:r>` : "";
  return `<w:p>${pPr}${run}</w:p>`;
}

/** Build a fully-formatted paragraph `BodyBlock` (string-based, byte-safe). */
export function makeStyledParagraphNode(text: string, opts: StyledParagraphOptions = {}): BodyBlock {
  return { kind: "paragraph", tag: "w:p", xml: makeStyledParagraphXml(text, opts) };
}

// ─── Table construction ──────────────────────────────────────────────────────

/**
 * Build a `<w:tbl>` XML string from a grid of cell texts.
 *
 * Reuses the existing `Table` class (the same path the export pipeline uses, via
 * `docx-blocks.buildTable`) so the produced table is byte-identical to a
 * natively built one:
 *   - `w:tblPr` with single-line borders (top/bottom/left/right/insideH/insideV)
 *     + `w:tblW` 100% (`w:type="pct" w:w="5000"`) + `w:bidiVisual` when `rtl`;
 *   - `w:tblGrid` with N `w:gridCol` (N = the widest row);
 *   - the first row marked as a shaded + bold header when `headerRow`;
 *   - each cell `<w:tc><w:p><w:r>(<w:rPr><w:b/></w:rPr>)?<w:t xml:space="preserve">…</w:t></w:r></w:p></w:tc>`.
 *
 * Cell text escaping is delegated to the xml2js serializer (it escapes the
 * predefined entities for the run text + attribute values).
 */
export function makeTableXml(
  rows: string[][],
  opts?: { headerRow?: boolean; rtl?: boolean },
): string {
  const headerRow = opts?.headerRow ?? false;
  const rtl = opts?.rtl ?? false;

  // Strip XML-illegal control chars from cell text (the AI-insertion path can
  // carry them); then delegate to the shared Table.fromGrid builder so the
  // produced table is byte-identical to a natively built one.
  const sanitized = rows.map((r) => r.map((c) => stripInvalidXmlChars(c ?? "")));
  const t = Table.fromGrid(sanitized, {
    headerRow,
    boldHeader: headerRow,
    headerFill: "D9D9D9",
    rtl,
  });

  // Serialize compactly (no pretty whitespace) for parity with the string model.
  return XmlUtils.buildXml(t.toObject(), {
    rootName: "w:tbl",
    headless: true,
    pretty: false,
  });
}

/** Build a table `BodyBlock` from a grid of cell texts. */
export function makeTableNode(
  rows: string[][],
  opts?: { headerRow?: boolean; rtl?: boolean },
): BodyBlock {
  return { kind: "table", tag: "w:tbl", xml: makeTableXml(rows, opts) };
}

// ─── Drawing (inline image) construction ─────────────────────────────────────

/**
 * Build a `<w:p>` paragraph wrapping an inline `<w:drawing>` that references an
 * already-registered image relationship (`relId`, see MediaManager.insertImage).
 *
 * Namespaces are declared LOCALLY so the block is self-contained regardless of
 * the host document: `xmlns:wp` + `xmlns:a` on `<wp:inline>`, `xmlns:pic` on
 * `<pic:pic>`. `r:` is NOT redeclared — it is universally declared on
 * `<w:document>`. A plain `<w:drawing>` is used (no `mc:AlternateContent`,
 * which modern Word does not require for inline pictures).
 *
 * @param relId      The image relationship id (e.g. "rId7").
 * @param widthEmu   Inline width in EMU (1 px @96dpi = 9525 EMU).
 * @param heightEmu  Inline height in EMU.
 * @param shapeId    A document-unique id for `wp:docPr` / `pic:cNvPr`.
 * @param name       The picture name (escaped).
 */
export function makeDrawingParagraphXml(
  relId: string,
  widthEmu: number,
  heightEmu: number,
  shapeId: number,
  name: string,
): string {
  const cx = String(Math.round(widthEmu));
  const cy = String(Math.round(heightEmu));
  const id = String(shapeId);
  const safeName = escapeXmlText(name);
  const safeRelId = escapeXmlText(relId);

  return (
    `<w:p><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"` +
    ` xmlns:wp="${NS_WP_DRAWING}" xmlns:a="${NS_A_DML}">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${safeName}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic>` +
    `<a:graphicData uri="${PICTURE_URI}">` +
    `<pic:pic xmlns:pic="${NS_PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${safeName}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${safeRelId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</pic:spPr>` +
    `</pic:pic>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`
  );
}

/** Build a drawing (inline image) paragraph `BodyBlock`. */
export function makeDrawingParagraphNode(
  relId: string,
  widthEmu: number,
  heightEmu: number,
  shapeId: number,
  name: string,
): BodyBlock {
  return {
    kind: "paragraph",
    tag: "w:p",
    xml: makeDrawingParagraphXml(relId, widthEmu, heightEmu, shapeId, name),
  };
}

/**
 * Scan a full `document.xml` for the highest existing drawing id (`wp:docPr @id`
 * and `pic:cNvPr @id`) and return max+1 (minimum 1). Use it to assign a
 * document-unique id to a newly inserted drawing.
 */
export function nextDrawingId(documentXml: string): number {
  let max = 0;
  const re = /<(?:wp:docPr|pic:cNvPr)\b[^>]*\bid="(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(documentXml)) !== null) {
    const n = parseInt(m[1], 10);
    if (Number.isFinite(n) && n > max) max = n;
  }
  return max + 1;
}

// ─── Paragraph inspection ────────────────────────────────────────────────────

/** Concatenate the decoded text of every `<w:t ...>...</w:t>` in a block. */
export function paragraphText(xml: string): string {
  // Match `<w:t>`/`<w:t ...>...</w:t>` (NOT `<w:tab>`/`<w:tbl>`), `<w:t/>` — and
  // a run's `<w:tab/>` as "\t": a TOC entry is `text<w:tab/>page`, and dropping
  // the tab glued the two ("مقدمة5") everywhere this text surfaces.
  //
  // `<w:pPr>` spans are skipped WHOLESALE first, because tab-stop DEFINITIONS
  // (`<w:tabs><w:tab w:val="right" …/></w:tabs>`) live there and would read as
  // phantom tabs. Alternation order does the skipping: the pPr branch consumes
  // its whole span before the tab branch can see inside it.
  let out = "";
  const re = /<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>|<w:tab\b[^>]*\/>|<w:t(\/)?(?:\s[^>]*?)?(\/)?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[0].startsWith("<w:pPr")) continue; // tab STOPS, not tab characters
    if (m[0].startsWith("<w:tab")) {
      out += "\t";
      continue;
    }
    const selfClosed = m[1] === "/" || m[2] === "/";
    if (selfClosed) continue; // <w:t/> → empty
    const contentStart = re.lastIndex;
    const closeIdx = xml.indexOf("</w:t>", contentStart);
    if (closeIdx === -1) break;
    out += decodeXmlText(xml.slice(contentStart, closeIdx));
    re.lastIndex = closeIdx + "</w:t>".length;
  }
  return out;
}

/** Extract the `w:val` of the paragraph's `<w:pStyle>`, or null. */
export function paragraphStyleId(xml: string): string | null {
  const m = /<w:pStyle\b[^>]*\bw:val="([^"]*)"/.exec(xml);
  return m ? decodeXmlText(m[1]) : null;
}

/**
 * Map a paragraph style id to a heading level 1–6 (0 = not a heading). Tolerant
 * of the style-id variants real .docx files carry: English `Heading N` / `Title`,
 * French `Titre N`, and separators ("Heading 1", "heading-1"). Returns 0 for any
 * non-heading style or null.
 */
export function headingLevelFromStyleId(styleId: string | null): number {
  if (!styleId) return 0;
  const s = styleId.toLowerCase().replace(/[\s_-]/g, "");
  if (s === "title") return 1;
  const m = /^(?:heading|titre)([1-6])$/.exec(s);
  return m ? Number(m[1]) : 0;
}

/**
 * Read a paragraph's own outline level (`<w:outlineLvl w:val="N"/>`, 0-based) and
 * return it as a 1-based heading level (N+1), or 0 if absent / out of range.
 * Catches headings whose paragraph carries an outline level without a heading
 * STYLE — common in imported documents.
 */
export function paragraphOutlineLevel(xml: string): number {
  const m = /<w:outlineLvl\b[^>]*\bw:val="(\d+)"/.exec(xml);
  if (!m) return 0;
  const n = Number(m[1]);
  return n >= 0 && n <= 5 ? n + 1 : 0;
}

/**
 * The heading level (1–6) of a paragraph block, 0 if it is not a heading.
 * Prefers the paragraph style (`headingLevelFromStyleId`); falls back to the
 * paragraph's own outline level — so both Word-authored and imported headings are
 * recognised.
 */
export function paragraphHeadingLevel(xml: string): number {
  return headingLevelFromStyleId(paragraphStyleId(xml)) || paragraphOutlineLevel(xml);
}

/** The paragraph's alignment (`w:jc` value: left/center/right/both/…), or null. */
export function paragraphAlignment(xml: string): string | null {
  const m = /<w:jc\b[^>]*\bw:val="([^"]+)"/.exec(xml);
  return m ? m[1] : null;
}

/**
 * The paragraph's font size in POINTS, read from the first `<w:sz>`/`<w:szCs>`
 * (which store half-points), or null if none is set inline. Useful as a
 * "larger than body" signal when inferring untyped headings.
 */
export function paragraphFontSizePt(xml: string): number | null {
  const m = /<w:sz(?:Cs)?\b[^>]*\bw:val="(\d+)"/.exec(xml);
  return m ? Number(m[1]) / 2 : null;
}

const RUN_RE = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
/** True if a run's properties enable bold (`<w:b/>`/`<w:bCs/>`, not val="0/false/off"). */
function runIsBold(runXml: string): boolean {
  const m = /<w:b(?:Cs)?\b([^>]*?)\/?>/.exec(runXml);
  if (!m) return false;
  const val = /\bw:val="([^"]*)"/.exec(m[1] || "");
  return !(val && /^(0|false|off)$/i.test(val[1]));
}

/**
 * True when the paragraph's text is entirely bold — i.e. every text-bearing run
 * is bold. A common signal for an untyped (plain-style) title. Returns false for
 * paragraphs with no text.
 */
export function paragraphIsBold(xml: string): boolean {
  const runs = xml.match(RUN_RE) || [];
  let textRuns = 0;
  let boldRuns = 0;
  for (const r of runs) {
    if (!/<w:t[\s>/]/.test(r)) continue; // only runs that carry text
    textRuns++;
    if (runIsBold(r)) boldRuns++;
  }
  return textRuns > 0 && boldRuns === textRuns;
}

// ─── Paragraph text replacement (run-preserving where it matters) ────────────

// Content a paragraph carries that is NOT reachable through `<w:t>`, and that a
// naive "replace the runs" rewrite would therefore delete without trace.
// `m:oMath` belongs here for the same reason a picture does: an equation's
// characters live in `<m:t>`, so `paragraphText` reports the paragraph as empty
// (or as just its trailing "(II.1)" tab run) and rewriting it from that text
// dropped the whole equation — 108 of them in one real thesis.
const EMBED_RE = /<(?:w:(?:drawing|pict|object)|m:oMath(?:Para)?)\b/;

/** First `<w:rPr>…</w:rPr>` in `xml`, or "" (a self-closing `<w:rPr/>` = none). */
function firstRunProps(xml: string): string {
  const start = xml.indexOf("<w:rPr");
  if (start === -1) return "";
  const openEnd = endOfTag(xml, start);
  if (isSelfClosing(xml, start, openEnd)) return "";
  const close = xml.indexOf("</w:rPr>", openEnd);
  return close === -1 ? "" : xml.slice(start, close + "</w:rPr>".length);
}

/**
 * The `<w:pPr>…</w:pPr>` substring of a paragraph — style, alignment, indent,
 * spacing, bidi, numbering — or "" when it has none.
 */
export function paragraphPropsXml(paragraphXml: string): string {
  const own = ownPPr(paragraphXml);
  return own ? paragraphXml.slice(own.start, own.end) : "";
}

/**
 * The `<w:rPr>` that should style text (re)written into this paragraph: the
 * properties of its first TEXT-BEARING run, falling back to the paragraph-mark
 * properties inside `<w:pPr>` — the same source Word uses for newly typed text.
 *
 * This is what keeps an edited paragraph looking like its neighbours. Imported
 * theses carry their formatting almost entirely at run level — `<w:rFonts
 * w:cs="Times New Roman" w:hint="cs"/>` and `<w:rtl/>` for Arabic, plus
 * `w:sz`/`w:szCs`/`w:b` — and `word/styles.xml` frequently has an empty
 * `docDefaults`, so a run emitted WITHOUT these falls back to whatever the
 * renderer defaults to: wrong face, wrong size, and (because `w:lineRule="auto"`
 * scales with font size) visibly tighter lines than the paragraphs around it.
 */
export function paragraphRunPropsXml(paragraphXml: string): string {
  for (const run of paragraphXml.match(RUN_RE) || []) {
    if (!/<w:t[\s>/]/.test(run)) continue; // only runs that carry text
    const rPr = firstRunProps(run);
    if (rPr) return rPr;
  }
  return firstRunProps(paragraphPropsXml(paragraphXml));
}

/**
 * Replace the text of a paragraph's XML string IN PLACE (operates only on the
 * given paragraph substring — cannot affect any sibling block):
 *
 *  - If the paragraph contains a `<w:drawing>` / `<w:pict>` / `<w:object>`
 *    (an inline image or embedded object), DO NOT strip runs. Instead replace
 *    only the text inside the FIRST `<w:t>` (or append a text run after
 *    `<w:pPr>` if there is no `<w:t>`), leaving the drawing runs intact.
 *  - Otherwise (a plain text paragraph): preserve `<w:pPr>...</w:pPr>` if
 *    present and replace ALL `<w:r>...</w:r>` runs with a single run that
 *    CARRIES THE ORIGINAL RUN PROPERTIES — `<w:r>RPR<w:t
 *    xml:space="preserve">ESCAPED</w:t></w:r>`. Dropping the `<w:rPr>` here is
 *    what made an inline/AI edit render in the wrong font and line height while
 *    its untouched neighbours stayed correct; see {@link paragraphRunPropsXml}.
 */
/**
 * Split a paragraph into `<w:p …>` / inner / `</w:p>`, NORMALISING a
 * self-closing `<w:p …/>` into an open+close pair.
 *
 * Word writes an empty paragraph as `<w:p w:rsidR="…"/>` — extremely common, one
 * per blank line. Locating the end of the open tag with `indexOf(">")` finds the
 * `>` of `/>` on such a paragraph, so anything appended "inside" it actually
 * lands AFTER it, as a sibling. At body level that means a bare `<w:r>` next to
 * the paragraphs — invalid (a run is not block-level content), and Word refuses
 * to open the document at all. It also loses the edit: the text is no longer in
 * any paragraph.
 */
function splitParagraph(paragraphXml: string): { head: string; inner: string; tail: string } {
  const openEnd = endOfTag(paragraphXml, 0);
  if (isSelfClosing(paragraphXml, 0, openEnd)) {
    // `<w:p a="b"/>` → head `<w:p a="b">`, no inner, tail `</w:p>`.
    return { head: paragraphXml.slice(0, openEnd).replace(/\s*\/>$/, ">"), inner: "", tail: "</w:p>" };
  }
  const closeStart = paragraphXml.lastIndexOf("</w:p>");
  return {
    head: paragraphXml.slice(0, openEnd),
    inner: paragraphXml.slice(openEnd, closeStart === -1 ? paragraphXml.length : closeStart),
    tail: closeStart === -1 ? "</w:p>" : paragraphXml.slice(closeStart),
  };
}

export function setParagraphText(paragraphXml: string, text: string): string {
  const escaped = escapeXmlText(text);
  const rPr = paragraphRunPropsXml(paragraphXml);

  if (EMBED_RE.test(paragraphXml)) {
    // Image/object paragraph — never strip runs. Edit the first <w:t> only.
    const tMatch = /<w:t(?:\s[^>]*?)?>([\s\S]*?)<\/w:t>|<w:t(?:\s[^>]*?)?\/>/.exec(paragraphXml);
    if (tMatch) {
      // Replace the first <w:t> (whether empty self-closing or with content).
      const full = tMatch[0];
      const replacement = `<w:t xml:space="preserve">${escaped}</w:t>`;
      return paragraphXml.slice(0, tMatch.index) + replacement + paragraphXml.slice(tMatch.index + full.length);
    }
    // No <w:t> at all — append a text run after </w:pPr> (or after <w:p...>).
    const run = `<w:r>${rPr}<w:t xml:space="preserve">${escaped}</w:t></w:r>`;
    const pPrEnd = paragraphXml.indexOf("</w:pPr>");
    if (pPrEnd !== -1) {
      const at = pPrEnd + "</w:pPr>".length;
      return paragraphXml.slice(0, at) + run + paragraphXml.slice(at);
    }
    // Insert right after the <w:p ...> open tag — via splitParagraph, so a
    // self-closing paragraph becomes a real open/close pair instead of having
    // the run appended outside it.
    const { head, inner, tail } = splitParagraph(paragraphXml);
    return `${head}${run}${inner}${tail}`;
  }

  // Plain text paragraph: keep <w:pPr> (style/bidi/alignment) AND the original
  // run properties (font/size/bold/colour + the cs/rtl twins), replace runs.
  const newRun = `<w:r>${rPr}<w:t xml:space="preserve">${escaped}</w:t></w:r>`;

  // Find the paragraph's inner content (between <w:p ...> and </w:p>), turning a
  // self-closing empty paragraph into a real open/close pair first.
  const { head, inner, tail } = splitParagraph(paragraphXml);

  // Preserve a leading <w:pPr>...</w:pPr> if present.
  const pPrMatch = /^(\s*)<w:pPr\b/.exec(inner);
  let pPr = "";
  if (pPrMatch) {
    const pPrStart = inner.indexOf("<w:pPr", 0);
    // <w:pPr/> self-closing?
    const pPrOpenEnd = endOfTag(inner, pPrStart);
    if (isSelfClosing(inner, pPrStart, pPrOpenEnd)) {
      pPr = inner.slice(pPrStart, pPrOpenEnd);
    } else {
      const pPrClose = inner.indexOf("</w:pPr>", pPrOpenEnd);
      pPr = pPrClose === -1 ? "" : inner.slice(pPrStart, pPrClose + "</w:pPr>".length);
    }
  }

  return `${head}${pPr}${newRun}${tail}`;
}

// ─── Legacy-compatible helpers (kept for the existing public exports) ────────
//
// The package's `src/index.ts` re-exports `parseOrderedDoc`, `buildOrderedDoc`,
// `toBlocks`, and `nodeTag`. They are reimplemented here on top of the string
// model so external imports and the verify-roundtrip harness keep working.

/** Split a full document into its top-level body blocks. */
export function parseOrderedDoc(documentXml: string): {
  split: SplitDocument;
  blocks: BodyBlock[];
  bodyChildren: BodyBlock[];
} {
  const split = splitDocument(documentXml);
  return { split, blocks: split.blocks, bodyChildren: split.blocks };
}

/** Reassemble from a SplitDocument (or its parts). */
export function buildOrderedDoc(split: SplitDocument): string {
  return assembleDocument(split);
}

/**
 * Pass-through classifier kept for compatibility. Filters out pure-whitespace
 * "other" blocks would change semantics, so this returns every block as-is.
 */
export function toBlocks(bodyChildren: BodyBlock[]): BodyBlock[] {
  return bodyChildren;
}

/** The tag name of a block (compat shim — accepts a BodyBlock). */
export function nodeTag(block: BodyBlock | string): string {
  if (typeof block === "string") {
    // Treat as raw XML and read the first tag name.
    const lt = block.indexOf("<");
    return lt === -1 ? "" : readTagName(block, lt);
  }
  return block?.tag ?? "";
}

export { W_NS };
