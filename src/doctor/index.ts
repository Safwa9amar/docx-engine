/**
 * docx-doctor — inspect a .docx package for the corruption classes that actually
 * bite this product, and mechanically repair the ones that can be repaired.
 *
 * ## Why this exists
 *
 * An AI edit writes real OOXML into a real student's thesis. When it writes the
 * wrong SHAPE, Word does not warn and does not repair — it refuses the file
 * outright ("Word experienced an error trying to open the file"), and the damage
 * is already on disk. Every defect below is one we have shipped at least once:
 *
 *  • **Schema sequence.** `w:tblPr`, `w:tcPr`, `w:pPr`, `w:sectPr` and `w:style`
 *    are `xsd:sequence` — child ORDER is a hard constraint, and `bidiVisual`
 *    after `tblW` killed an Arabic thesis. Well-formed ≠ valid, so `xmllint` sees
 *    nothing. (`w:rPr` and `w:trPr` are deliberately NOT checked: their schema
 *    types are repeated CHOICES, so their child order is free.)
 *  • **Story shape.** A story must not end with a table, and two adjacent tables
 *    merge into one. Word writes an empty `<w:p/>` at both spots; hand-built
 *    OOXML forgets to.
 *  • **Body order.** Any manager that round-trips `word/document.xml` through
 *    xml2js regroups `<w:body>` children BY TAG, hoisting every table above every
 *    paragraph. Block indices here are positional (op queue, RAG chunks, edits),
 *    so that silently rewires the whole document. Unfixable in place — the order
 *    is genuinely gone — so we detect it loudly and point at history restore.
 *  • **Dropped spaces.** The same xml2js path runs `trim:true`, which turns
 *    Word's inter-word space runs (`<w:t xml:space="preserve"> </w:t>`) into
 *    `<w:t/>`. Words glue together. We can't recover an emptied run, but we CAN
 *    stop the next save from dropping the spaces still present.
 *  • **Package integrity.** Dangling relationship targets, missing
 *    `[Content_Types].xml` overrides, duplicate rIds — each on its own makes Word
 *    refuse the document.
 *  • **Lying zip flags.** adm-zip re-writes a streamed archive keeping the
 *    "a data descriptor follows" flag it did not honour. It then cannot read its
 *    own output back, and Word refuses the file. The byte-level wrappers report
 *    these as `zip.*` findings. See `inspectZipFlags`, `clearFalseDataDescriptors`.
 *
 * ## Shape of the module
 *
 * Pure functions over a duck-typed zip, exactly like `hf-part-repair.ts` (which
 * stays as the chrome-op fast path; both only ever move a part toward the same
 * canonical shape, so running either or both is idempotent). String surgery on
 * purpose: a parse/rebuild round-trip is the very thing that causes half the
 * defects listed above.
 *
 * `inspectDocx(zip)` reports. `inspectDocx(zip, { fix: true })` also repairs, in
 * place, and lists the parts it rewrote. `checkDocxBuffer` / `repairDocxBuffer`
 * are the byte-level wrappers; the repairing one re-opens its own output and
 * discards the whole repair unless every rewritten part still parses and no new
 * fatal appeared. Verified against real theses: four Word-authored student
 * documents come back completely clean and untouched, and a repair never changes
 * a single `<w:t>` payload — it only ever moves markup.
 */

import {
  findElements, splitChildren, reorderInner, rewriteElements, firstXmlError,
} from "./xml";
import {
  checkContainment, fillEmptyCells, fillEmptyStories, repairTables, unbalancedFields,
} from "./structure";
import { ZipManager } from "@/utils/ZipManager";
import { parseOrderedDoc, buildOrderedDoc } from "@/core/files/body/OrderedBody";
import type { BodyBlock } from "@/core/files/body/OrderedBody";
import { updateParagraphProps } from "@/core/ooxml/paragraphProps";
import { declareUsedPrefixes, undeclaredPrefixes } from "@/core/ooxml/namespaces";

// ─────────────────────────────────────────────────────────────────────────────
// Public types
// ─────────────────────────────────────────────────────────────────────────────

/** `fatal` — the document is broken (Word refuses it, or its content order is
 *  destroyed). `warning` — wrong or degraded, but it still opens. */
export type Severity = "fatal" | "warning";

export interface Finding {
  /** Stable id — the dashboard and the AI tools key their copy off this. */
  rule: string;
  severity: Severity;
  /** Zip entry the finding is in, or "package" for whole-file findings. */
  part: string;
  /** How many occurrences in that part. */
  count: number;
  /** One line, plain English, safe to show a human or hand to a model. */
  message: string;
  /** True when this doctor knows a mechanical repair for it. */
  fixable: boolean;
  /** True when `fix` was requested AND the repair was applied. */
  fixed: boolean;
  /** Optional specifics (offending ids, sample values) — already truncated. */
  detail?: string;
}

export interface DoctorReport {
  /** No fatal findings remain (after repairs, when repairing). */
  ok: boolean;
  /** Zip entries examined. */
  checkedParts: number;
  findings: Finding[];
  /** Parts actually rewritten. Empty unless `fix` was requested. */
  repairedParts: string[];
}

/** The three-and-a-bit methods we need off the engine's zip (or adm-zip). */
export interface DocxZip {
  getEntries(): { entryName: string }[];
  readAsText(entry: string): string;
  addFile(entry: string, content: Buffer): void;
}

export interface InspectOptions {
  /** Apply every repair this doctor considers safe. */
  fix?: boolean;
  /** Also apply repairs that DELETE something (an unreferenced dangling
   *  relationship). Never on the automatic save path. */
  aggressive?: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// Part classification
// ─────────────────────────────────────────────────────────────────────────────

/** Parts whose root holds wordprocessing content (paragraphs, tables, runs). */
const WORD_STORY = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;
/** Story parts PLUS the definition parts that use the same CT_* sequences. */
const SEQUENCED_PART = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments|styles|numbering|settings)\.xml$/;

const CONTENT_TYPES = "[Content_Types].xml";
const ROOT_RELS = "_rels/.rels";
const DOCUMENT = "word/document.xml";

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Content type each `word/*.xml` part must be declared as. */
function contentTypeFor(part: string): string | null {
  const m = /^word\/([a-zA-Z]+?)\d*\.xml$/.exec(part);
  if (!m) return part === "word/theme/theme1.xml"
    ? "application/vnd.openxmlformats-officedocument.theme+xml"
    : null;
  switch (m[1]) {
    case "document": return `${WML}.document.main+xml`;
    case "styles": return `${WML}.styles+xml`;
    case "numbering": return `${WML}.numbering+xml`;
    case "settings": return `${WML}.settings+xml`;
    case "webSettings": return `${WML}.webSettings+xml`;
    case "fontTable": return `${WML}.fontTable+xml`;
    case "header": return `${WML}.header+xml`;
    case "footer": return `${WML}.footer+xml`;
    case "footnotes": return `${WML}.footnotes+xml`;
    case "endnotes": return `${WML}.endnotes+xml`;
    case "comments": return `${WML}.comments+xml`;
    default: return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// CT_* child sequences (ECMA-376 Part 1)
//
// ONLY containers whose schema type is an `xsd:sequence` appear here. That
// distinction is load-bearing: `CT_TrPrBase` and `EG_RPrBase` are
// `xsd:choice maxOccurs="unbounded"`, so their children may legally appear in ANY
// order — "fixing" them would rewrite healthy documents for nothing. Verified
// against real files: four Microsoft-Word-authored theses produce zero violations
// of the tables below, while a LibreOffice-authored one reorders `w:rPr` freely
// (28 occurrences) and Word opens it without complaint.
//
// A nested array marks a CO-RANKED group — a repeated choice INSIDE an otherwise
// ordered sequence, where the members keep whatever order they were authored in.
// ─────────────────────────────────────────────────────────────────────────────

type SequenceOrder = (string | string[])[];

// CT_PPrBase, then CT_PPr's own trailing children (rPr, sectPr, pPrChange).
const CT_PPR: SequenceOrder = [
  "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
  "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
  "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
  "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
  "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
  "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
  "w:rPr", "w:sectPr", "w:pPrChange",
];

// CT_TblPrBase + CT_TblPr's tblPrChange. `w:bidiVisual` before `w:tblW` is the
// exact pair that made Word refuse an Arabic thesis outright.
const CT_TBLPR: SequenceOrder = [
  "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
  "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
  "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
  "w:tblPrChange",
];

const CT_TCPR: SequenceOrder = [
  "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd",
  "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
  "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
];

// EG_HdrFtrReferences (a repeated choice — headers and footers interleave, and
// Word really does emit `footerReference` before `headerReference`) precedes
// EG_SectPrContents.
const CT_SECTPR: SequenceOrder = [
  ["w:headerReference", "w:footerReference"],
  "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz",
  "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt",
  "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter",
  "w:docGrid", "w:printerSettings", "w:sectPrChange",
];

// CT_TblBorders / CT_TcBorders / CT_PBdr / CT_TblCellMar / CT_TcMar. All four
// sides in SCHEMA order — top, left, bottom, right — NOT the top/bottom/left/right
// order a human (or a CSS shorthand) reaches for. mdocxengine's
// `Table.setTableBorders` built the object in the human order and the XML builder
// emits keys as given, so every table it styled shipped `<w:bottom>` before
// `<w:left>` and Word refused the whole document. `w:start`/`w:end` are the strict
// spellings of left/right and co-rank with them.
const CT_BORDERS: SequenceOrder = [
  "w:top", ["w:start", "w:left"], "w:bottom", ["w:end", "w:right"],
  "w:insideH", "w:insideV", "w:tl2br", "w:tr2bl",
];
const CT_PBDR: SequenceOrder = ["w:top", "w:left", "w:bottom", "w:right", "w:between", "w:bar"];
const CT_MARGINS: SequenceOrder = ["w:top", ["w:start", "w:left"], "w:bottom", ["w:end", "w:right"]];
const CT_NUMPR: SequenceOrder = ["w:ilvl", "w:numId", "w:numberingChange", "w:ins"];

// CT_Lvl — one level of an abstract numbering definition. `w:lvlJc` sits LATE
// (just before the properties), which is the trap: writers that think of it as
// "alignment, like numFmt" emit it right after `w:start` and everything from
// `w:numFmt` on is then out of sequence.
//
// WARNING, not fatal, and the evidence is the seed asset: `thesis-base.docx`
// carries nine levels in exactly this shape and Word opens it without complaint
// (verified by driving Word itself). It was briefly marked fatal on the
// assumption that it was what made a combined thesis unopenable; the real cause
// turned out to be `rels.type-mismatch`. Normalise it, but never tell staff a
// thesis that opens fine is broken.
const CT_LVL: SequenceOrder = [
  "w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff", "w:lvlText",
  "w:lvlPicBulletId", "w:legacy", "w:lvlJc", "w:pPr", "w:rPr",
];

// CT_Settings. Long, and nearly all of it optional, so the only part that ever
// bites is the tail: `w:compat` comes near the END, after which almost nothing
// may follow. The seed `thesis-base.docx` emits `<w:displayBackgroundShape/>`
// AFTER `<w:compat>` — invalid in every thesis the product has ever made, and
// tolerated by Word, which is why it went unnoticed.
const CT_SETTINGS: SequenceOrder = [
  "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation", "w:removeDateAndTime",
  "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape", "w:printPostScriptOverText",
  "w:printFractionalCharacterWidth", "w:printFormsData", "w:embedTrueTypeFonts",
  "w:embedSystemFonts", "w:saveSubsetFonts", "w:saveFormsData", "w:mirrorMargins",
  "w:alignBordersAndEdges", "w:bordersDoNotSurroundHeader", "w:bordersDoNotSurroundFooter",
  "w:gutterAtTop", "w:hideSpellingErrors", "w:hideGrammaticalErrors", "w:activeWritingStyle",
  "w:proofState", "w:formsDesign", "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter",
  "w:stylePaneSortMethod", "w:documentType", "w:mailMerge", "w:revisionView", "w:trackChanges",
  "w:doNotTrackMoves", "w:doNotTrackFormatting", "w:documentProtection", "w:autoFormatOverride",
  "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop", "w:autoHyphenation",
  "w:consecutiveHyphenLimit", "w:hyphenationZone", "w:doNotHyphenateCaps", "w:showEnvelope",
  "w:summaryLength", "w:clickAndTypeStyle", "w:defaultTableStyle", "w:evenAndOddHeaders",
  "w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
  "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
  "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
  "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
  "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", "w:noPunctuationKerning",
  "w:characterSpacingControl", "w:printTwoOnOne", "w:strictFirstAndLastChars",
  "w:noLineBreaksAfter", "w:noLineBreaksBefore", "w:savePreviewPicture",
  "w:doNotValidateAgainstSchema", "w:saveInvalidXml", "w:ignoreMixedContent",
  "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly",
  "w:useXSLTWhenSaving", "w:saveThroughXslt", "w:showXMLTags", "w:alwaysMergeEmptyNamespace",
  "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
  "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang",
  "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
  "w:forceUpgrade", "w:captions", "w:readModeInkLockDown", "w:smartTagType",
  "sl:schemaLibrary", "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol",
  "w:listSeparator",
];

// CT_Style. The seed `thesis-base.docx` puts w:rPr first in every style — invalid.
const CT_STYLE: SequenceOrder = [
  "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden",
  "w:uiPriority", "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", "w:personal",
  "w:personalCompose", "w:personalReply", "w:rsid", "w:pPr", "w:rPr", "w:tblPr", "w:trPr",
  "w:tcPr", "w:tblStylePr",
];

/**
 * Applied innermost-first so an outer pass sees children already in order.
 *
 * `severity` is what the FINDING is worth, not what the repair is worth — the
 * repair always runs. Only `w:tblPr` is fatal, because that is the one we have
 * watched Word refuse a file over. The rest are schema violations Word tolerates:
 * worth normalising (our own writers should not emit them) but not worth telling
 * staff a healthy imported thesis is broken.
 */
const SEQUENCES: { tag: string; order: SequenceOrder; label: string; severity: Severity; part?: RegExp }[] = [
  // Innermost first: borders/margins sit inside tblPr and tcPr.
  { tag: "w:tblBorders", order: CT_BORDERS, label: "table borders", severity: "fatal" },
  { tag: "w:tcBorders", order: CT_BORDERS, label: "cell borders", severity: "fatal" },
  { tag: "w:tblCellMar", order: CT_MARGINS, label: "table cell margins", severity: "fatal" },
  { tag: "w:tcMar", order: CT_MARGINS, label: "cell margins", severity: "fatal" },
  { tag: "w:pBdr", order: CT_PBDR, label: "paragraph borders", severity: "warning" },
  { tag: "w:numPr", order: CT_NUMPR, label: "numbering properties", severity: "warning" },
  { tag: "w:tcPr", order: CT_TCPR, label: "table cell properties", severity: "warning" },
  { tag: "w:tblPr", order: CT_TBLPR, label: "table properties", severity: "fatal" },
  { tag: "w:sectPr", order: CT_SECTPR, label: "section properties", severity: "warning" },
  { tag: "w:pPr", order: CT_PPR, label: "paragraph properties", severity: "warning" },
  { tag: "w:style", order: CT_STYLE, label: "style definitions", severity: "warning", part: /^word\/styles\.xml$/ },
  { tag: "w:lvl", order: CT_LVL, label: "numbering levels", severity: "warning", part: /^word\/numbering\.xml$/ },
  { tag: "w:settings", order: CT_SETTINGS, label: "document settings", severity: "warning", part: /^word\/settings\.xml$/ },
];

/**
 * Properties elements that must LEAD their parent.
 *
 * A separate rule from SEQUENCES because the constraint is different: the parent
 * holds repeatable CONTENT (runs, cells, rows) that must not be reordered, so all
 * we can do is move the one properties element back to the front. A `w:trPr`
 * sitting after a `w:tc` is invalid and, like the border case, silently fatal.
 *
 * `after` names the elements the schema puts AHEAD of `child` — front means "the
 * first slot this element may occupy", not index 0. CT_Row is
 * `tblPrEx?, trPr?, EG_ContentCellContent*`, so a row that carries a table-level
 * exception has TWO leading properties elements and `w:trPr` belongs second.
 * Hoisting it to index 0 (which this did until 2026-08-15) took every valid row
 * Word had written with a `w:tblPrEx` and made it invalid — the doctor's own
 * repair was the corruption. `w:tblPrEx` is listed first so a row that arrives
 * already inverted is put right before the `w:trPr` rule looks at it.
 */
const FIRST_CHILD: { parent: string; child: string; label: string; after?: string[] }[] = [
  { parent: "w:tr", child: "w:tblPrEx", label: "row-level table exceptions" },
  { parent: "w:tr", child: "w:trPr", label: "row properties", after: ["w:tblPrEx"] },
  { parent: "w:tc", child: "w:tcPr", label: "cell properties" },
  { parent: "w:p", child: "w:pPr", label: "paragraph properties" },
  { parent: "w:r", child: "w:rPr", label: "run properties" },
];

/** Move `child` into its leading slot — past any `after` element already sitting
 *  there — or null when it is absent or in that slot already. */
function hoistFirstChild(inner: string, child: string, after: string[] = []): string | null {
  const split = splitChildren(inner);
  if (!split || split.items.length < 2) return null;
  const at = split.items.findIndex((i) => i.tag === child);
  if (at < 0) return null; // absent
  // Skip the elements the schema seats ahead of this one; the first index left is
  // where it belongs.
  let target = 0;
  while (target < at && after.includes(split.items[target].tag)) target++;
  if (at === target) return null; // already in its slot
  if (at < target) return null; // ahead of an element that may precede it — leave it
  const [moved] = split.items.splice(at, 1);
  // The moved element takes the leading whitespace of the element it displaces,
  // and hands its own to whatever now follows the gap it left.
  const head = split.items[target];
  const lead = head.lead;
  head.lead = moved.lead;
  moved.lead = lead;
  split.items.splice(target, 0, moved);
  return split.items.map((i) => i.lead + i.xml).join("") + split.tail;
}

/** Flatten a SequenceOrder into tag → rank, co-ranked groups sharing a rank. */
function rankMap(order: SequenceOrder): Map<string, number> {
  const out = new Map<string, number>();
  order.forEach((slot, i) => {
    if (Array.isArray(slot)) for (const tag of slot) out.set(tag, i);
    else out.set(slot, i);
  });
  return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// Finding helpers
// ─────────────────────────────────────────────────────────────────────────────

function finding(f: Omit<Finding, "fixed"> & { fixed?: boolean }): Finding {
  return { fixed: false, ...f };
}

/** Join a set of offending values into a short, bounded `detail` string. */
function sample(values: Iterable<string>, max = 8): string {
  const list = [...values];
  const head = list.slice(0, max).join(", ");
  return list.length > max ? `${head} (+${list.length - max} more)` : head;
}

// ─────────────────────────────────────────────────────────────────────────────
// Rules
// ─────────────────────────────────────────────────────────────────────────────

/** Missing parts that make the package not a Word document at all. */
function checkRequiredParts(names: Set<string>): Finding[] {
  const out: Finding[] = [];
  for (const required of [CONTENT_TYPES, ROOT_RELS, DOCUMENT]) {
    if (!names.has(required)) {
      out.push(finding({
        rule: "package.missing-part",
        severity: "fatal",
        part: required,
        count: 1,
        message: `The package is missing ${required}. Word cannot open it; restore the thesis from history.`,
        fixable: false,
      }));
    }
  }
  return out;
}

/** Every XML part must at least be well-formed. */
function checkWellFormed(zip: DocxZip, names: string[]): Finding[] {
  const out: Finding[] = [];
  for (const name of names) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    const xml = readPart(zip, name);
    if (xml == null) continue;
    const err = firstXmlError(xml);
    if (err) {
      out.push(finding({
        rule: "xml.malformed",
        severity: "fatal",
        part: name,
        count: 1,
        message: `${name} is not well-formed XML (${err}). Word will refuse the document; restore from history.`,
        fixable: false,
      }));
    }
  }
  return out;
}

/**
 * Every prefix a part uses must be declared. A tag matcher calls such a part well-formed, and Word
 * refuses the whole file: a root declaring only `xmlns:w` plus a section reference written as `r:id`
 * is enough. Word's own prefixes are repaired by declaring them on the root.
 */
function checkPrefixes(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const findings: Finding[] = [];
  const rewrote: string[] = [];
  for (const name of names) {
    if (!/^word\/[^/]+\.xml$/.test(name)) continue;
    const xml = readPart(zip, name);
    if (xml == null || firstXmlError(xml)) continue; // already reported as malformed
    const missing = undeclaredPrefixes(xml);
    if (!missing.length) continue;
    let left = missing;
    if (fix) {
      const next = declareUsedPrefixes(xml);
      if (next !== xml) {
        zip.addFile(name, Buffer.from(next, "utf8"));
        rewrote.push(name);
        left = undeclaredPrefixes(next);
      }
    }
    const fixed = fix && left.length === 0;
    const shown = fixed ? missing : left;
    findings.push(finding({
      rule: "xml.undeclared-prefix",
      severity: "fatal",
      part: name,
      count: shown.reduce((n, u) => n + u.count, 0),
      message: `${name} uses namespace prefixes it never declares. Word will refuse the document.`,
      fixable: left.length === 0 || missing.length !== left.length,
      fixed,
      detail: sample(shown.map((u) => u.prefix)),
    }));
  }
  return { findings, rewrote };
}

/** Every part that needs a content-type Override must have one. */
function checkContentTypes(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: boolean } {
  let xml = readPart(zip, CONTENT_TYPES);
  if (xml == null) return { findings: [], rewrote: false };

  // A second Default for an extension (their case does not count) or a second Override for a part: Word
  // refuses the file, and a fonts or media writer comparing case-sensitively writes exactly that.
  const seen = new Set<string>();
  const repeats: { tag: string; shown: string }[] = [];
  for (const m of xml.matchAll(/<(Default|Override)\b[^>]*?\b(Extension|PartName)="([^"]+)"[^>]*?(?:\/>|>\s*<\/\1>)/g)) {
    const key = `${m[1]}:${m[3].toLowerCase()}`;
    if (seen.has(key)) repeats.push({ tag: m[0], shown: m[3] });
    seen.add(key);
  }
  const duplicates: Finding[] = [];
  let dedup = false;
  if (repeats.length) {
    if (fix) {
      // Later ones go, so the file's first statement of each stands.
      const kept = new Set<string>();
      xml = xml.replace(/<(Default|Override)\b[^>]*?\b(Extension|PartName)="([^"]+)"[^>]*?(?:\/>|>\s*<\/\1>)/g, (tag, kind, _a, value) => {
        const key = `${kind}:${value.toLowerCase()}`;
        if (kept.has(key)) return "";
        kept.add(key);
        return tag;
      });
      zip.addFile(CONTENT_TYPES, Buffer.from(xml, "utf8"));
      dedup = true;
    }
    duplicates.push(finding({
      rule: "contenttypes.duplicate",
      severity: "fatal",
      part: CONTENT_TYPES,
      count: repeats.length,
      message: `[Content_Types].xml states ${repeats.length} entr${repeats.length === 1 ? "y" : "ies"} twice. Word refuses the document.`,
      detail: sample(repeats.map((r) => r.shown.toLowerCase())),
      fixable: true,
      fixed: dedup,
    }));
  }

  const declared = new Set<string>();
  for (const m of xml.matchAll(/<Override\b[^>]*PartName="\/?([^"]+)"/g)) declared.add(m[1]);

  const missing: { part: string; type: string }[] = [];
  for (const name of names) {
    const type = contentTypeFor(name);
    if (type && !declared.has(name)) missing.push({ part: name, type });
  }
  if (!missing.length) return { findings: duplicates, rewrote: dedup };

  let rewrote = dedup;
  if (fix) {
    const closeAt = xml.lastIndexOf("</Types>");
    if (closeAt !== -1) {
      const added = missing.map((m) => `<Override PartName="/${m.part}" ContentType="${m.type}"/>`).join("");
      zip.addFile(CONTENT_TYPES, Buffer.from(xml.slice(0, closeAt) + added + xml.slice(closeAt), "utf8"));
      rewrote = true;
    }
  }

  return {
    findings: [...duplicates, finding({
      rule: "contenttypes.missing-override",
      severity: "fatal",
      part: CONTENT_TYPES,
      count: missing.length,
      message: `${missing.length} part(s) are in the package but not declared in [Content_Types].xml. Word refuses a document with an undeclared part.`,
      detail: sample(missing.map((m) => m.part)),
      fixable: true,
      fixed: rewrote,
    })],
    rewrote,
  };
}

/** Resolve a relationship Target against the owning part's directory.
 *  Note the `[^/]*` — the package root's rels file is literally `_rels/.rels`,
 *  with an EMPTY stem, and a `+` there silently leaves the directory as `_rels/`
 *  so every root relationship resolves to a part that does not exist. */
function resolveTarget(relsPart: string, target: string): string {
  const dir = relsPart.replace(/_rels\/[^/]*\.rels$/, "");
  const joined = target.startsWith("/") ? target.slice(1) : dir + target;
  const segs: string[] = [];
  for (const seg of joined.split("/")) {
    if (seg === "." || seg === "") continue;
    if (seg === "..") segs.pop();
    else segs.push(seg);
  }
  return segs.join("/");
}

/** The part a `.rels` file describes ("word/_rels/document.xml.rels" →
 *  "word/document.xml"). The package root's `_rels/.rels` describes the package
 *  itself, not a part, so it has no owner. */
function ownerOf(relsPart: string): string | null {
  const m = /^(.*)_rels\/([^/]+)\.rels$/.exec(relsPart);
  return m ? m[1] + m[2] : null;
}

/** Elements that wrap a relationship reference and are meaningless without the
 *  part it points at — a picture whose bytes are not in the package renders as
 *  nothing at best, and makes Word reject the document at worst. Removing the
 *  whole wrapper is what Word's own "repair" does. */
const REF_WRAPPERS = ["w:drawing", "w:pict", "w:object"];

/**
 * Remove the content that references a relationship id, so a rel pointing at a
 * missing part can be dropped without leaving an unresolvable `r:id` behind.
 *
 * A `w:hyperlink` is UNWRAPPED rather than deleted — the link is dead but the
 * student's words are not, and silently deleting their text would be a far worse
 * bug than the one being fixed. Everything else in REF_WRAPPERS is removed whole.
 *
 * Returns `null` when some reference could not be accounted for; the caller then
 * leaves the part completely alone rather than half-fixing it.
 */
function stripReferences(xml: string, rId: string): { xml: string; removed: number } | null {
  const mentions = (s: string) =>
    [...s.matchAll(/\br:(?:id|embed|link|pict|href|dm|lo|qs|cs)="([^"]+)"/g)].filter((m) => m[1] === rId).length;
  if (mentions(xml) === 0) return { xml, removed: 0 };

  let out = xml;
  let removed = 0;

  // Hyperlinks first: unwrapping keeps the text, and a hyperlink can itself sit
  // inside nothing else we touch.
  for (;;) {
    const hit = findElements(out, "w:hyperlink").find(
      (r) => mentions(out.slice(r.start, r.innerStart)) > 0,
    );
    if (!hit) break;
    out = out.slice(0, hit.start) + out.slice(hit.innerStart, hit.innerEnd) + out.slice(hit.end);
    removed++;
  }

  for (const tag of REF_WRAPPERS) {
    for (;;) {
      const hit = findElements(out, tag).find((r) => mentions(out.slice(r.start, r.end)) > 0);
      if (!hit) break;
      out = out.slice(0, hit.start) + out.slice(hit.end);
      removed++;
    }
  }

  // Anything still pointing at the rel means we do not understand this document.
  if (mentions(out) > 0) return null;
  return { xml: out, removed };
}

/**
 * Remove only the wrappers holding a WRONGLY-TYPED reference.
 *
 * Deliberately not `stripReferences`: that keys on the rId and removes every use
 * of it, which is right for a dead rel (nothing can use it) and badly wrong here.
 * A mismatched id is usually a perfectly good relationship being pointed at by
 * the wrong kind of element — the thesis that prompted this had `<c:chart
 * r:id="rId7">` alongside a legitimate `<a:blip r:embed="rId7">`, and stripping
 * by id deleted the student's image too. Only the offending element's wrapper
 * goes; every other user of that rId is untouched.
 */
function stripMismatchedRefs(
  xml: string,
  bad: { id: string; re: RegExp }[],
): { xml: string; removed: number } | null {
  const offends = (s: string) =>
    bad.some(({ id, re }) => [...s.matchAll(new RegExp(re.source, "g"))].some((m) => m[1] === id));
  if (!offends(xml)) return { xml, removed: 0 };

  let out = xml;
  let removed = 0;
  for (const tag of REF_WRAPPERS) {
    for (;;) {
      const hit = findElements(out, tag).find((r) => offends(out.slice(r.start, r.end)));
      if (!hit) break;
      out = out.slice(0, hit.start) + out.slice(hit.end);
      removed++;
    }
  }
  // A mismatched reference outside any wrapper we know how to remove: leave the
  // whole part alone rather than half-fix it.
  if (offends(out)) return null;
  return { xml: out, removed };
}

/**
 * References that demand a specific relationship Type.
 *
 * Only elements whose Type is unambiguous are listed — `<w:hyperlink r:id>` is
 * deliberately absent, because Word also accepts an internal `document` target
 * there for a same-file link.
 */
const TYPED_REFS: { re: RegExp; want: RegExp; label: string }[] = [
  { re: /<c:chart\b[^>]*\br:id="([^"]+)"/g, want: /\/chart$/, label: "chart" },
  { re: /<a:blip\b[^>]*\br:embed="([^"]+)"/g, want: /\/image$/, label: "image" },
  { re: /<a:blip\b[^>]*\br:link="([^"]+)"/g, want: /\/image$/, label: "image" },
  { re: /<v:imagedata\b[^>]*\br:id="([^"]+)"/g, want: /\/image$/, label: "image" },
  { re: /<w:headerReference\b[^>]*\br:id="([^"]+)"/g, want: /\/header$/, label: "header" },
  { re: /<w:footerReference\b[^>]*\br:id="([^"]+)"/g, want: /\/footer$/, label: "footer" },
  { re: /<dgm:relIds\b[^>]*\br:dm="([^"]+)"/g, want: /\/diagramData$/, label: "diagram data" },
];

/** Dangling targets, duplicate ids, and r:id references with no relationship. */
function checkRelationships(
  zip: DocxZip,
  names: string[],
  present: Set<string>,
  opts: { fix: boolean; aggressive: boolean },
): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];

  for (const relsPart of names.filter((n) => n.endsWith(".rels"))) {
    const xml = readPart(zip, relsPart);
    if (xml == null) continue;

    const owner = ownerOf(relsPart);
    const ownerXml = owner && present.has(owner) && /\.xml$/.test(owner) ? readPart(zip, owner) ?? "" : "";
    // Every r:*="rIdN" attribute the owning part uses (r:id, r:embed, r:link, …).
    const used = new Set<string>();
    for (const m of ownerXml.matchAll(/\br:(?:id|embed|link|pict|href|dm|lo|qs|cs)="([^"]+)"/g)) used.add(m[1]);

    const seen = new Set<string>();
    const dupes = new Set<string>();
    const dangling: { id: string; target: string; xml: string }[] = [];

    for (const m of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
      const rel = m[0];
      const id = /\bId="([^"]*)"/.exec(rel)?.[1] ?? "";
      const target = /\bTarget="([^"]*)"/.exec(rel)?.[1] ?? "";
      const external = /\bTargetMode="External"/.test(rel);
      if (id) {
        if (seen.has(id)) dupes.add(id);
        seen.add(id);
      }
      if (external || !target) continue;
      if (!present.has(resolveTarget(relsPart, target))) dangling.push({ id, target, xml: rel });
    }

    if (dupes.size) {
      out.push(finding({
        rule: "rels.duplicate-id",
        severity: "fatal",
        part: relsPart,
        count: dupes.size,
        message: `${relsPart} declares the same relationship id more than once. Word resolves only one of them and refuses or mis-renders the rest.`,
        detail: sample(dupes),
        fixable: false,
      }));
    }

    if (dangling.length) {
      const orphans = dangling.filter((d) => !used.has(d.id));
      // A referenced dangling rel needs the CONTENT that points at it removed as
      // well, or dropping the rel just trades one fatal for another. Work out up
      // front whether we know how to do that — it decides `fixable`, and it is
      // the difference between "staff can fix this in one click" and "restore
      // from history".
      const referenced = dangling
        .filter((d) => used.has(d.id))
        .map((d) => ({ ...d, strip: owner && ownerXml ? stripReferences(ownerXml, d.id) : null }));
      const removable = referenced.filter((d) => d.strip);
      const stuck = referenced.filter((d) => !d.strip);

      const toDrop = opts.fix && opts.aggressive ? [...orphans, ...removable] : [];
      let droppedOrphans = 0;
      let droppedInUse = 0;

      if (toDrop.length) {
        // Rewrite the owning part first: if that fails we must not have already
        // removed the relationship it still points at.
        let ownerNext = ownerXml;
        let ownerOk = true;
        for (const d of removable) {
          const strip = stripReferences(ownerNext, d.id);
          if (!strip) { ownerOk = false; break; }
          ownerNext = strip.xml;
        }
        if (!ownerOk || (owner && ownerNext !== ownerXml && firstXmlError(ownerNext))) {
          // Leave everything alone — reported, not repaired.
        } else {
          if (owner && ownerNext !== ownerXml) {
            zip.addFile(owner, Buffer.from(ownerNext, "utf8"));
            rewrote.push(owner);
          }
          let next = xml;
          for (const d of toDrop) next = next.replace(d.xml, "");
          if (next !== xml) {
            zip.addFile(relsPart, Buffer.from(next, "utf8"));
            rewrote.push(relsPart);
            droppedOrphans = orphans.length;
            droppedInUse = removable.length;
          }
        }
      }

      if (orphans.length) {
        out.push(finding({
          rule: "rels.dangling-target",
          severity: "fatal",
          part: relsPart,
          count: orphans.length,
          message: `${orphans.length} relationship(s) point at a part that is not in the package, and nothing references them. Dropping them makes the document openable again.`,
          detail: sample(orphans.map((o) => `${o.id}→${o.target}`)),
          fixable: true,
          fixed: droppedOrphans > 0,
        }));
      }
      if (removable.length) {
        out.push(finding({
          rule: "rels.dangling-target-in-use",
          severity: "fatal",
          part: relsPart,
          count: removable.length,
          message:
            `${removable.length} relationship(s) point at a missing part and are still used by ${owner ?? relsPart} — ` +
            "typically a picture whose image bytes are not in the file. Repairing removes the broken picture (or unlinks the " +
            "dead hyperlink, keeping its text) and then drops the reference. Enable the dead-links option to apply it.",
          detail: sample(removable.map((o) => `${o.id}→${o.target}`)),
          fixable: true,
          fixed: droppedInUse > 0,
        }));
      }
      if (stuck.length) {
        out.push(finding({
          rule: "rels.dangling-target-in-use",
          severity: "fatal",
          part: relsPart,
          count: stuck.length,
          message: `${stuck.length} relationship(s) point at a missing part AND are referenced by ${owner ?? relsPart} from content this repair does not know how to remove safely. The missing part has to be restored — repairing cannot help here.`,
          detail: sample(stuck.map((o) => `${o.id}→${o.target}`)),
          fixable: false,
        }));
      }
    }

    // ── The reference resolves, but to the WRONG KIND of part ────────────────
    //
    // A `<c:chart r:id>` aimed at an `/image` relationship is invisible to every
    // other check here: the id exists, the target file exists, and the schema has
    // no idea what a relationship Type is. Word follows it, expects a chartSpace,
    // finds a PNG, and refuses the whole document. This is what a combined thesis
    // shipped with — unopenable from the day it was built.
    if (owner && ownerXml) {
      const typeOf = new Map<string, string>();
      for (const m of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
        const id = /\bId="([^"]*)"/.exec(m[0])?.[1];
        const type = /\bType="([^"]*)"/.exec(m[0])?.[1];
        if (id && type) typeOf.set(id, type);
      }
      const mismatched: { id: string; want: string; got: string; re: RegExp }[] = [];
      for (const ref of TYPED_REFS) {
        for (const m of ownerXml.matchAll(ref.re)) {
          const id = m[1];
          const got = typeOf.get(id);
          if (!got || ref.want.test(got)) continue;
          mismatched.push({ id, want: ref.label, got: got.replace(/^.*\//, ""), re: ref.re });
        }
      }
      if (mismatched.length) {
        // Keyed on element+id, not id alone: the same rId can be referenced
        // correctly elsewhere and must keep working.
        const uniq = [...new Map(mismatched.map((d) => [`${d.want} ${d.id}`, d])).values()];
        let fixed = 0;
        if (opts.fix && opts.aggressive) {
          const strip = stripMismatchedRefs(ownerXml, uniq.map((d) => ({ id: d.id, re: d.re })));
          if (strip && strip.xml !== ownerXml && !firstXmlError(strip.xml)) {
            zip.addFile(owner, Buffer.from(strip.xml, "utf8"));
            rewrote.push(owner);
            fixed = uniq.length;
          }
        }
        out.push(finding({
          rule: "rels.type-mismatch",
          severity: "fatal",
          part: owner,
          count: uniq.length,
          message:
            `${uniq.length} reference(s) in ${owner} point at a relationship of the wrong kind ` +
            `(${sample(uniq.map((d) => `${d.id}: wanted ${d.want}, got ${d.got}`))}). The id resolves and the target part ` +
            "exists, so neither a dangling-target check nor schema validation sees it — but Word follows the reference, " +
            "finds the wrong kind of part and refuses to open the document. The referenced content is not recoverable from " +
            "this package; repairing removes it. Enable the dead-links option to apply it.",
          detail: sample(uniq.map((d) => d.id)),
          fixable: true,
          fixed: fixed > 0,
        }));
      }
    }

    const unresolved = [...used].filter((id) => !seen.has(id));
    if (owner && unresolved.length) {
      out.push(finding({
        rule: "rels.missing-relationship",
        severity: "fatal",
        part: owner,
        count: unresolved.length,
        message: `${owner} references relationship id(s) that ${relsPart} does not declare. Word refuses a document with an unresolvable r:id.`,
        detail: sample(unresolved),
        fixable: false,
      }));
    }
  }

  return { findings: out, rewrote };
}

/** CT_* child order, across every part that carries those elements. */
function checkSequences(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];

  for (const name of names.filter((n) => SEQUENCED_PART.test(n))) {
    const original = readPart(zip, name);
    if (original == null) continue;

    let xml = original;
    const perTag: { label: string; tag: string; count: number; severity: Severity; firstChild?: boolean }[] = [];

    for (const seq of SEQUENCES) {
      if (seq.part && !seq.part.test(name)) continue;
      const rank = rankMap(seq.order);
      const res = rewriteElements(xml, seq.tag, (inner) => reorderInner(inner, rank));
      if (res.changed) {
        perTag.push({ label: seq.label, tag: seq.tag, count: res.changed, severity: seq.severity });
        // Keep the reordered XML even when not fixing: later passes must see the
        // same picture the fix would, so the counts we report are the real ones.
        xml = res.xml;
      }
    }

    for (const rule of FIRST_CHILD) {
      const res = rewriteElements(xml, rule.parent, (inner) => hoistFirstChild(inner, rule.child, rule.after));
      if (res.changed) {
        perTag.push({
          label: rule.after?.length
            ? `${rule.label} must sit directly after <${rule.after.join("> / <")}> in <${rule.parent}>`
            : `${rule.label} must be the first child of <${rule.parent}>`,
          tag: rule.child,
          count: res.changed,
          severity: "fatal",
          firstChild: true,
        });
        xml = res.xml;
      }
    }

    if (!perTag.length) continue;

    // Never write a part we just broke.
    const err = fix ? firstXmlError(xml) : null;
    const applied = fix && !err;
    if (applied) {
      zip.addFile(name, Buffer.from(xml, "utf8"));
      rewrote.push(name);
    }

    for (const t of perTag) {
      out.push(finding({
        rule: "sequence.out-of-order",
        severity: t.severity,
        part: name,
        count: t.count,
        message: t.firstChild
          ? `${t.count} <${t.tag}> element(s) are not where the schema puts them — ${t.label}. ` +
            "Word refuses to open the document rather than moving them."
          : `${t.count} <${t.tag}> element(s) have their children in the wrong order. ${t.label} are an ordered ` +
            "schema sequence" +
            (t.severity === "fatal"
              ? " — Word refuses to open the document rather than reordering them."
              : ", so the document is invalid even though Word tolerates it."),
        detail: t.tag,
        fixable: true,
        fixed: applied,
      }));
    }
    if (err) {
      out.push(finding({
        rule: "sequence.repair-refused",
        severity: "warning",
        part: name,
        count: 1,
        message: `The child-order repair for ${name} produced XML that does not check out (${err}), so it was discarded and the part left untouched.`,
        fixable: false,
      }));
    }
  }

  return { findings: out, rewrote };
}

/** The whitespace-only block kinds `parseOrderedDoc` emits between real children. */
function isFiller(b: BodyBlock): boolean {
  return b.kind === "other" && !b.xml.trim();
}

/**
 * What `<w:body>` is allowed to contain: EG_BlockLevelElts plus the range-markup
 * elements and the trailing sectPr. Anything else and Word refuses the document
 * — most often a bare `<w:r>`, which is a RUN and may only live inside a `<w:p>`.
 */
const BODY_CHILDREN = new Set([
  "w:customXml", "w:sdt", "w:p", "w:tbl", "w:proofErr", "w:permStart", "w:permEnd",
  "w:bookmarkStart", "w:bookmarkEnd", "w:moveFromRangeStart", "w:moveFromRangeEnd",
  "w:moveToRangeStart", "w:moveToRangeEnd", "w:commentRangeStart", "w:commentRangeEnd",
  "w:customXmlInsRangeStart", "w:customXmlInsRangeEnd", "w:customXmlDelRangeStart",
  "w:customXmlDelRangeEnd", "w:customXmlMoveFromRangeStart", "w:customXmlMoveFromRangeEnd",
  "w:customXmlMoveToRangeStart", "w:customXmlMoveToRangeEnd", "w:ins", "w:del",
  "w:moveFrom", "w:moveTo", "m:oMathPara", "m:oMath", "w:altChunk", "w:sectPr",
]);

/** Match a whole standalone `<w:pPr>` element — self-closing or paired. Capture
 *  group 1 is its inner XML (`undefined` for the self-closing form). */
const ORPHAN_PPR_RE = /^<w:pPr(?:\s*\/>|(?:\s[^>]*?)?>([\s\S]*?)<\/w:pPr>)$/;

/** The inner XML of a body-level `<w:pPr>`, or null when `xml` is not cleanly one
 *  (in which case it is left alone rather than merged into a paragraph). */
function orphanPPrInner(xml: string): string | null {
  const m = ORPHAN_PPR_RE.exec(xml.trim());
  return m ? (m[1] ?? "") : null;
}

/** Index of the paragraph immediately before `i`, skipping whitespace fillers.
 *  -1 when the previous real block is not a paragraph (nothing to merge into). */
function precedingParagraph(blocks: BodyBlock[], i: number): number {
  for (let j = i - 1; j >= 0; j--) {
    if (isFiller(blocks[j])) continue;
    return blocks[j].kind === "paragraph" ? j : -1;
  }
  return -1;
}

/** True when a run carries something a reader would actually see or miss. */
function runHasContent(xml: string): boolean {
  if (/<w:(?:drawing|pict|object|br|tab|sym|noBreakHyphen|softHyphen|fldChar|instrText|footnoteReference|endnoteReference|commentReference)\b/.test(xml)) return true;
  return [...xml.matchAll(/<w:(?:t|delText)(?:\s[^>]*)?>([^<]*)<\//g)].some((m) => m[1].length > 0);
}

/** Body-level structure: sectPr placement, table spacers, tag-grouped order. */
function checkBody(zip: DocxZip, fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const xml = readPart(zip, DOCUMENT);
  if (xml == null) return { findings: [], rewrote: [] };

  let parsed: ReturnType<typeof parseOrderedDoc>;
  try {
    parsed = parseOrderedDoc(xml);
  } catch {
    return { findings: [], rewrote: [] }; // xml.malformed already reported it
  }

  const out: Finding[] = [];
  const blocks = [...parsed.split.blocks];
  const real = blocks.filter((b) => !isFiller(b));
  let mutated = false;

  // ── Tag-grouped body: the xml2js round-trip signature. ────────────────────
  const paraIdx: number[] = [];
  const tableIdx: number[] = [];
  real.forEach((b, i) => {
    if (b.kind === "paragraph") paraIdx.push(i);
    else if (b.kind === "table") tableIdx.push(i);
  });
  if (tableIdx.length >= 3 && paraIdx.length >= 20) {
    const allTablesFirst = tableIdx[tableIdx.length - 1] < paraIdx[0];
    const allTablesLast = tableIdx[0] > paraIdx[paraIdx.length - 1];
    if (allTablesFirst || allTablesLast) {
      out.push(finding({
        rule: "body.tag-grouped",
        severity: "fatal",
        part: DOCUMENT,
        count: tableIdx.length,
        message:
          `All ${tableIdx.length} tables sit ${allTablesFirst ? "before" : "after"} all ${paraIdx.length} paragraphs — ` +
          "the signature of a body that was regrouped by tag on a save. The original order is not recoverable from the file; " +
          "restore the thesis from document history.",
        fixable: false,
      }));
    }
  }

  // ── Only block-level elements may sit in <w:body>. ────────────────────────
  // This is the defect that kept a real thesis unopenable after every other rule
  // here said it was healthy: a single `<w:r>` had ended up as a direct child of
  // <w:body>, next to an image paragraph. Word rejects the whole document.
  const strays = blocks
    .map((b, i) => ({ b, i }))
    .filter(({ b }) => !isFiller(b) && !BODY_CHILDREN.has(b.tag) && b.tag !== "#text");
  if (strays.length) {
    const runs = strays.filter(({ b }) => b.tag === "w:r");
    const others = strays.filter(({ b }) => b.tag !== "w:r");

    if (runs.length) {
      if (fix) {
        // An empty run is dropped (no content to save, and wrapping it would add
        // a stray blank line); a run carrying text or a picture is wrapped in the
        // paragraph it should always have been in. Either way nothing is lost.
        for (const { i } of [...runs].reverse()) {
          const xml = blocks[i].xml;
          if (runHasContent(xml)) blocks[i] = { kind: "paragraph", tag: "w:p", xml: `<w:p>${xml}</w:p>` };
          else blocks.splice(i, 1);
        }
        mutated = true;
      }
      out.push(finding({
        rule: "body.stray-run",
        severity: "fatal",
        part: DOCUMENT,
        count: runs.length,
        message:
          `${runs.length} text run(s) sit directly in <w:body> instead of inside a paragraph. ` +
          "A run is not a block-level element, so Word refuses to open the document — this is the classic " +
          "\"Word experienced an error trying to open the file\". Repairing wraps a run that carries content in a " +
          "paragraph and drops an empty one.",
        fixable: true,
        fixed: fix,
      }));
    }
    // An orphaned `<w:pPr>` is repairable, and worth repairing rather than
    // dropping: it is a paragraph's properties — very often a SECTION BREAK, the
    // page boundary of a front-matter page — that a writer spliced in after a
    // self-closing `<w:p/>`'s already-closed tag, so it landed BESIDE the
    // paragraph instead of inside it. Merging it back into that paragraph
    // restores exactly what the writer meant, page break included; deleting it
    // would silently repaginate the thesis.
    const reattachable = new Set(
      others.filter(({ b, i }) => b.tag === "w:pPr" && precedingParagraph(blocks, i) >= 0),
    );
    const unfixable = others.filter((o) => !reattachable.has(o));

    if (reattachable.size) {
      let merged = 0;
      if (fix) {
        // Reverse order: splicing an orphan out never shifts an earlier one.
        for (const { b, i } of [...reattachable].reverse()) {
          const p = precedingParagraph(blocks, i);
          const inner = orphanPPrInner(b.xml);
          if (inner === null) continue;
          try {
            // `updateParagraphProps` normalises the self-closing `<w:p/>` back to
            // paired form and re-canonicalises CT_PPr order, so a merged
            // `w:sectPr` lands where the schema wants it. An empty orphan leaves
            // the paragraph byte-identical and is simply dropped.
            if (inner) {
              blocks[p] = { ...blocks[p], xml: updateParagraphProps(blocks[p].xml, (cur) => cur + inner) };
            }
          } catch {
            continue; // malformed props: leave it in place, still reported below
          }
          blocks.splice(i, 1);
          merged++;
          mutated = true;
        }
      }
      out.push(finding({
        rule: "body.orphaned-ppr",
        severity: "fatal",
        part: DOCUMENT,
        count: fix ? merged : reattachable.size,
        message:
          `${reattachable.size} <w:pPr> element(s) sit directly in <w:body> instead of inside the paragraph they ` +
          "describe. Paragraph properties are not block-level content, so Word refuses to open the document. " +
          "This is what a writer produces when it splices properties in after a self-closing <w:p/> — the empty " +
          "paragraph Word writes for a blank line. Repairing merges each one back into the paragraph before it, " +
          "which preserves the section break it usually carries.",
        fixable: true,
        fixed: fix,
      }));
    }

    if (unfixable.length) {
      out.push(finding({
        rule: "body.illegal-child",
        severity: "fatal",
        part: DOCUMENT,
        count: unfixable.length,
        message: `${unfixable.length} element(s) in <w:body> are not block-level content. Word refuses to open the document.`,
        detail: sample(unfixable.map((o) => `<${o.b.tag}>`)),
        fixable: false,
      }));
    }
  }

  // ── The body-level sectPr must be the last child. ─────────────────────────
  const sectIdx = blocks.map((b, i) => (b.kind === "sectPr" ? i : -1)).filter((i) => i >= 0);
  if (sectIdx.length === 0) {
    out.push(finding({
      rule: "body.no-sectpr",
      severity: "warning",
      part: DOCUMENT,
      count: 1,
      message: "The body has no final <w:sectPr>, so page size, margins and headers fall back to Word's defaults.",
      fixable: false,
    }));
  } else {
    const lastReal = blocks.reduce((acc, b, i) => (isFiller(b) ? acc : i), -1);
    const misplaced = sectIdx.filter((i) => i !== lastReal);
    if (misplaced.length) {
      // Keep the LAST sectPr (the one Word would honour) and move it to the end.
      // The splice happens whether or not we are fixing, so the checks below see
      // the shape a repair would produce and their counts stay honest; nothing is
      // written unless `fix`.
      const keep = sectIdx[sectIdx.length - 1];
      const node = blocks[keep];
      for (let i = sectIdx.length - 1; i >= 0; i--) blocks.splice(sectIdx[i], 1);
      blocks.push(node);
      mutated = true;
      out.push(finding({
        rule: "body.sectpr-not-last",
        severity: "fatal",
        part: DOCUMENT,
        count: misplaced.length,
        message:
          `A <w:sectPr> sits in the middle of <w:body> (${sectIdx.length} at body level). ` +
          "Only the final child may be one; anything else makes Word report the document as unreadable.",
        fixable: true,
        fixed: fix,
      }));
    }
  }

  // ── Spacers: adjacent tables merge, and a story must not end on a table. ──
  const rtl = /<w:bidiVisual\b/.test(xml);
  const spacers = insertBodySpacers(blocks, rtl);
  if (spacers.adjacent) {
    mutated = true;
    out.push(finding({
      rule: "story.adjacent-tables",
      severity: "fatal",
      part: DOCUMENT,
      count: spacers.adjacent,
      message: `${spacers.adjacent} pair(s) of tables touch with no paragraph between them. Word merges touching tables into one, silently rewriting the document's block indices.`,
      fixable: true,
      fixed: fix,
    }));
  }
  if (spacers.trailing) {
    mutated = true;
    out.push(finding({
      rule: "story.ends-with-table",
      severity: "fatal",
      part: DOCUMENT,
      count: 1,
      message: "The body's last content block is a table. A story must not end with one — Word writes an empty paragraph there and rejects a file that has none.",
      fixable: true,
      fixed: fix,
    }));
  }

  const rewrote: string[] = [];
  if (fix && mutated) {
    const next = buildOrderedDoc({ ...parsed.split, blocks });
    if (!firstXmlError(next)) {
      zip.addFile(DOCUMENT, Buffer.from(next, "utf8"));
      rewrote.push(DOCUMENT);
    }
  }

  return { findings: out, rewrote };
}

/** Insert the empty paragraphs Word keeps between/after tables. Mutates `blocks`
 *  and returns how many spots needed one. */
function insertBodySpacers(blocks: BodyBlock[], rtl: boolean): { adjacent: number; trailing: number } {
  const spacer = (): BodyBlock => ({
    kind: "paragraph",
    tag: "w:p",
    xml: rtl ? "<w:p><w:pPr><w:bidi/></w:pPr></w:p>" : "<w:p/>",
  });

  let adjacent = 0;
  for (let i = blocks.length - 1; i > 0; i--) {
    if (blocks[i].kind !== "table") continue;
    let prev = i - 1;
    while (prev >= 0 && isFiller(blocks[prev])) prev--;
    if (prev >= 0 && blocks[prev].kind === "table") {
      blocks.splice(prev + 1, 0, spacer());
      adjacent++;
    }
  }

  // The last content block, ignoring the trailing sectPr and whitespace.
  let last = blocks.length - 1;
  while (last >= 0 && (isFiller(blocks[last]) || blocks[last].kind === "sectPr")) last--;
  const trailing = last >= 0 && blocks[last].kind === "table" ? 1 : 0;
  if (trailing) blocks.splice(last + 1, 0, spacer());

  return { adjacent, trailing };
}

/** Header/footer stories: same two table rules, applied with string surgery. */
function checkStories(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];

  for (const name of names) {
    const m = /^word\/(header|footer)\d*\.xml$/.exec(name);
    if (!m) continue;
    const xml = readPart(zip, name);
    if (xml == null) continue;

    const tag = m[1] === "header" ? "w:hdr" : "w:ftr";
    const rtl = /<w:bidiVisual\b/.test(xml);
    const spacer = rtl ? "<w:p><w:pPr><w:bidi/></w:pPr></w:p>" : "<w:p/>";

    let next = xml;
    let adjacent = 0;
    next = next.replace(/<\/w:tbl>\s*<w:tbl\b/g, () => {
      adjacent++;
      return `</w:tbl>${spacer}<w:tbl`;
    });

    let trailing = 0;
    const closeAt = next.lastIndexOf(`</${tag}>`);
    if (closeAt >= 0 && next.slice(0, closeAt).trimEnd().endsWith("</w:tbl>")) {
      next = next.slice(0, closeAt) + spacer + next.slice(closeAt);
      trailing = 1;
    }

    if (!adjacent && !trailing) continue;

    const applied = fix && !firstXmlError(next);
    if (applied) {
      zip.addFile(name, Buffer.from(next, "utf8"));
      rewrote.push(name);
    }
    if (adjacent) {
      out.push(finding({
        rule: "story.adjacent-tables",
        severity: "fatal",
        part: name,
        count: adjacent,
        message: `${adjacent} pair(s) of tables touch in ${name}. Word merges touching tables into one.`,
        fixable: true,
        fixed: applied,
      }));
    }
    if (trailing) {
      out.push(finding({
        rule: "story.ends-with-table",
        severity: "fatal",
        part: name,
        count: 1,
        message: `${name} ends with a table. A story must not — Word refuses to open the document.`,
        fixable: true,
        fixed: applied,
      }));
    }
  }

  return { findings: out, rewrote };
}

/** Word writes inter-word spaces as their own run. Without xml:space="preserve"
 *  the next reader is entitled to drop them, and words glue together. */
function checkTextSpacing(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];
  // Only ever spaces: a leading/trailing NEWLINE means a pretty-printer indented
  // the element, and pinning that as significant text would inject real
  // whitespace into the document.
  const EDGE_SPACE = /^ | $/;

  for (const name of names.filter((n) => WORD_STORY.test(n))) {
    const xml = readPart(zip, name);
    if (xml == null) continue;

    let unprotected = 0;
    const next = xml.replace(/<w:t(\s[^>]*)?>([^<]*)<\/w:t>/g, (whole, attrs: string | undefined, text: string) => {
      const a = attrs ?? "";
      if (/xml:space=/.test(a)) return whole;
      if (!EDGE_SPACE.test(text)) return whole;
      unprotected++;
      return `<w:t${a} xml:space="preserve">${text}</w:t>`;
    });
    const emptied = [...xml.matchAll(/<w:t\s*\/>|<w:t(?:\s[^>]*)?><\/w:t>/g)].length;

    if (unprotected) {
      const applied = fix && !firstXmlError(next);
      if (applied) {
        zip.addFile(name, Buffer.from(next, "utf8"));
        rewrote.push(name);
      }
      out.push(finding({
        rule: "text.unprotected-space",
        severity: "warning",
        part: name,
        count: unprotected,
        message: `${unprotected} text run(s) begin or end with a space but are not marked xml:space="preserve". Those spaces are dropped on the next round-trip and the words either side glue together.`,
        fixable: true,
        fixed: applied,
      }));
    }
    if (emptied >= 10) {
      out.push(finding({
        rule: "text.emptied-runs",
        severity: "warning",
        part: name,
        count: emptied,
        message: `${emptied} empty text run(s) — the fingerprint of a save that trimmed every text node. The lost characters cannot be recovered from the file; compare against document history if words look glued together.`,
        fixable: false,
      }));
    }
  }

  return { findings: out, rewrote };
}

/** A bookmarkStart with no End (or vice versa) — captions, cross-references and
 *  the TOC all hang off bookmarks, and Word reports the file as needing repair. */
function checkBookmarks(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];

  for (const name of names.filter((n) => WORD_STORY.test(n))) {
    const xml = readPart(zip, name);
    if (xml == null) continue;

    const starts = new Map<string, string>();
    for (const m of xml.matchAll(/<w:bookmarkStart\b[^>]*\/>/g)) {
      const id = /\bw:id="([^"]*)"/.exec(m[0])?.[1];
      if (id) starts.set(id, m[0]);
    }
    const ends = new Map<string, string>();
    for (const m of xml.matchAll(/<w:bookmarkEnd\b[^>]*\/>/g)) {
      const id = /\bw:id="([^"]*)"/.exec(m[0])?.[1];
      if (id) ends.set(id, m[0]);
    }

    const orphanStarts = [...starts.keys()].filter((id) => !ends.has(id));
    const orphanEnds = [...ends.keys()].filter((id) => !starts.has(id));
    if (!orphanStarts.length && !orphanEnds.length) continue;

    let applied = false;
    if (fix) {
      let next = xml;
      // An orphan start gets its End right after it (keeping the bookmark, which
      // a caption or cross-reference may point at); an orphan End is dropped.
      for (const id of orphanStarts) next = next.replace(starts.get(id)!, `${starts.get(id)!}<w:bookmarkEnd w:id="${id}"/>`);
      for (const id of orphanEnds) next = next.replace(ends.get(id)!, "");
      if (next !== xml && !firstXmlError(next)) {
        zip.addFile(name, Buffer.from(next, "utf8"));
        rewrote.push(name);
        applied = true;
      }
    }

    out.push(finding({
      rule: "bookmark.unmatched",
      severity: "warning",
      part: name,
      count: orphanStarts.length + orphanEnds.length,
      message: `${orphanStarts.length + orphanEnds.length} bookmark marker(s) have no matching partner. Captions, cross-references and the table of contents point at bookmarks, so Word flags the document for repair.`,
      detail: sample([...orphanStarts.map((i) => `start ${i}`), ...orphanEnds.map((i) => `end ${i}`)]),
      fixable: true,
      fixed: applied,
    }));
  }

  return { findings: out, rewrote };
}

/** A row wider than its own tblGrid — Word renders the overflow unpredictably. */
function checkTableGrids(zip: DocxZip, names: string[]): Finding[] {
  const out: Finding[] = [];

  for (const name of names.filter((n) => WORD_STORY.test(n))) {
    const xml = readPart(zip, name);
    if (xml == null) continue;

    let bad = 0;
    for (const tbl of findElements(xml, "w:tbl")) {
      const inner = xml.slice(tbl.innerStart, tbl.innerEnd);
      const grid = findElements(inner, "w:tblGrid")[0];
      if (!grid) continue;
      const cols = [...inner.slice(grid.innerStart, grid.innerEnd).matchAll(/<w:gridCol\b/g)].length;
      if (!cols) continue;

      for (const row of findElements(inner, "w:tr")) {
        const rowInner = inner.slice(row.innerStart, row.innerEnd);
        let width = 0;
        for (const cell of findElements(rowInner, "w:tc")) {
          const cellInner = rowInner.slice(cell.innerStart, cell.innerEnd);
          const pr = findElements(cellInner, "w:tcPr")[0];
          const span = pr
            ? /<w:gridSpan\b[^>]*w:val="(\d+)"/.exec(cellInner.slice(pr.innerStart, pr.innerEnd))?.[1]
            : undefined;
          width += span ? Number(span) : 1;
        }
        if (width > cols) bad++;
      }
    }

    if (bad) {
      out.push(finding({
        rule: "table.grid-mismatch",
        severity: "warning",
        part: name,
        count: bad,
        message: `${bad} table row(s) span more columns than their table's <w:tblGrid> declares. Word lays those rows out unpredictably and may report the document as repaired.`,
        fixable: false,
      }));
    }
  }

  return out;
}

/** Style and numbering ids referenced by content but never defined. */
function checkReferences(zip: DocxZip, names: string[], present: Set<string>): Finding[] {
  const out: Finding[] = [];

  if (present.has("word/styles.xml")) {
    const styles = readPart(zip, "word/styles.xml") ?? "";
    const defined = new Set<string>();
    for (const m of styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]*)"/g)) defined.add(m[1]);

    const missing = new Set<string>();
    for (const name of names.filter((n) => WORD_STORY.test(n))) {
      const xml = readPart(zip, name) ?? "";
      for (const m of xml.matchAll(/<w:(?:pStyle|rStyle|tblStyle)\b[^>]*w:val="([^"]*)"/g)) {
        if (!defined.has(m[1])) missing.add(m[1]);
      }
    }
    if (missing.size) {
      out.push(finding({
        rule: "ref.unknown-style",
        severity: "warning",
        part: "word/styles.xml",
        count: missing.size,
        message: `${missing.size} style id(s) are used by the content but never defined in styles.xml. Word silently falls back to Normal, so those paragraphs lose their formatting.`,
        detail: sample(missing),
        fixable: false,
      }));
    }
  }

  if (present.has("word/numbering.xml")) {
    const numbering = readPart(zip, "word/numbering.xml") ?? "";
    const defined = new Set<string>();
    for (const m of numbering.matchAll(/<w:num\b[^>]*w:numId="([^"]*)"/g)) defined.add(m[1]);

    const missing = new Set<string>();
    for (const name of names.filter((n) => WORD_STORY.test(n))) {
      const xml = readPart(zip, name) ?? "";
      for (const m of xml.matchAll(/<w:numId\b[^>]*w:val="([^"]*)"/g)) {
        if (m[1] !== "0" && !defined.has(m[1])) missing.add(m[1]);
      }
    }
    if (missing.size) {
      out.push(finding({
        rule: "ref.unknown-numid",
        severity: "warning",
        part: "word/numbering.xml",
        count: missing.size,
        message: `${missing.size} numbering id(s) are used by the content but never defined in numbering.xml. Word gives those paragraphs a default "1." instead of the list they meant.`,
        detail: sample(missing),
        fixable: false,
      }));
    }

    const bullets = new Set([...numbering.matchAll(/<w:numPicBullet\b[^>]*\bw:numPicBulletId="([^"]*)"/g)].map((m) => m[1]));
    const noBullet = new Set([...numbering.matchAll(/<w:lvlPicBulletId\b[^>]*\bw:val="([^"]*)"/g)].map((m) => m[1]).filter((id) => !bullets.has(id)));
    if (noBullet.size) {
      out.push(finding({
        rule: "ref.unknown-pic-bullet",
        severity: "fatal",
        part: "word/numbering.xml",
        count: noBullet.size,
        message: `${noBullet.size} list level(s) name a picture bullet the numbering part does not define. Word refuses the whole file.`,
        detail: sample(noBullet),
        fixable: false,
      }));
    }
  }

  return out;
}


/**
 * Structural content models: what an element must contain, and what it must not.
 * Runs over every story part — the sibling of checkBody, which owns `<w:body>`
 * because that one can splice whole ordered blocks.
 */
function checkStructure(zip: DocxZip, names: string[], fix: boolean): { findings: Finding[]; rewrote: string[] } {
  const out: Finding[] = [];
  const rewrote: string[] = [];

  for (const name of names.filter((n) => WORD_STORY.test(n))) {
    const original = readPart(zip, name);
    if (original == null) continue;
    let xml = original;

    // 1. Run-level content where block-level content belongs, and block-level
    //    content nested inside a paragraph.
    const contained = checkContainment(xml, fix);
    if (fix) xml = contained.xml;
    const { strayRuns, illegalInStory, blockInParagraph } = contained.report;
    if (strayRuns.count) {
      out.push(finding({
        rule: "story.stray-run",
        severity: "fatal",
        part: name,
        count: strayRuns.count,
        message:
          `${strayRuns.count} text run(s) sit directly in <${strayRuns.root}> instead of inside a paragraph. ` +
          "A run is not block-level content, so Word refuses to open the document. Repairing wraps a run that " +
          "carries content in a paragraph and drops an empty one.",
        fixable: true,
        fixed: fix,
      }));
    }
    for (const [tag, count] of illegalInStory) {
      out.push(finding({
        rule: "story.illegal-child",
        severity: "fatal",
        part: name,
        count,
        message: `${count} <${tag}> element(s) sit at story level, where only block-level content is allowed. Word refuses to open the document.`,
        detail: tag,
        fixable: false,
      }));
    }
    for (const [tag, count] of blockInParagraph) {
      out.push(finding({
        rule: "paragraph.block-child",
        severity: "fatal",
        part: name,
        count,
        message:
          `${count} <${tag}> element(s) are nested inside a paragraph, which may only hold run-level content. ` +
          "Moving them out would change where they appear, so this is reported rather than guessed at.",
        detail: tag,
        fixable: false,
      }));
    }

    // 2. Containers that may not be empty.
    const cells = fillEmptyCells(xml);
    if (cells.count) {
      if (fix) xml = cells.xml;
      out.push(finding({
        rule: "table.empty-cell",
        severity: "fatal",
        part: name,
        count: cells.count,
        message: `${cells.count} table cell(s) contain no paragraph. A cell must hold at least one block-level element; Word reports the document as needing repair. Repairing inserts the empty paragraph Word keeps in a blank cell.`,
        fixable: true,
        fixed: fix,
      }));
    }
    const stories = fillEmptyStories(xml);
    if (stories.count) {
      if (fix) xml = stories.xml;
      out.push(finding({
        rule: "story.empty",
        severity: "fatal",
        part: name,
        count: stories.count,
        message: `${stories.count} story/stories (body, header, footer or text box) contain no block-level content at all. Repairing inserts an empty paragraph.`,
        fixable: true,
        fixed: fix,
      }));
    }

    // 3. Tables that cannot render.
    const tables = repairTables(xml);
    if (tables.missingGrid || tables.emptyRows || tables.emptyTables) {
      if (fix) xml = tables.xml;
      if (tables.missingGrid) {
        out.push(finding({
          rule: "table.missing-grid",
          severity: "fatal",
          part: name,
          count: tables.missingGrid,
          message: `${tables.missingGrid} table(s) have no <w:tblGrid>, which the schema requires. Repairing rebuilds one from the widest row.`,
          fixable: true,
          fixed: fix,
        }));
      }
      if (tables.emptyRows) {
        out.push(finding({
          rule: "table.empty-row",
          severity: "fatal",
          part: name,
          count: tables.emptyRows,
          message: `${tables.emptyRows} table row(s) contain no cells. A row must have at least one; such a row renders nothing, so repairing removes it.`,
          fixable: true,
          fixed: fix,
        }));
      }
      if (tables.emptyTables) {
        out.push(finding({
          rule: "table.empty",
          severity: "fatal",
          part: name,
          count: tables.emptyTables,
          message: `${tables.emptyTables} table(s) contain no rows. Such a table renders nothing, so repairing removes it.`,
          fixable: true,
          fixed: fix,
        }));
      }
    }

    // 4. Half-deleted complex fields (captions, cross-references, the TOC).
    const broken = unbalancedFields(xml);
    if (broken) {
      out.push(finding({
        rule: "field.unbalanced",
        severity: "warning",
        part: name,
        count: broken,
        message: `${broken} complex field marker(s) have no partner. Captions, cross-references and the table of contents are built from begin/end field pairs, so Word may prompt to repair the document. Stitching a field back together is not something to guess at — check the affected captions by hand.`,
        fixable: false,
      }));
    }

    if (fix && xml !== original) {
      if (firstXmlError(xml)) {
        out.push(finding({
          rule: "structure.repair-refused",
          severity: "warning",
          part: name,
          count: 1,
          message: `The structural repair for ${name} produced XML that does not check out, so it was discarded and the part left untouched.`,
          fixable: false,
        }));
      } else {
        zip.addFile(name, Buffer.from(xml, "utf8"));
        rewrote.push(name);
      }
    }
  }

  return { findings: out, rewrote };
}

// ─────────────────────────────────────────────────────────────────────────────
// Entry points
// ─────────────────────────────────────────────────────────────────────────────

function readPart(zip: DocxZip, name: string): string | null {
  try {
    return zip.readAsText(name) || null;
  } catch {
    return null;
  }
}

const SEVERITY_ORDER: Record<Severity, number> = { fatal: 0, warning: 1 };

/**
 * Check a .docx package and, with `fix`, repair what can be repaired. The zip is
 * mutated in place; the caller decides whether to write the bytes back.
 */
export function inspectDocx(zip: DocxZip, opts: InspectOptions = {}): DoctorReport {
  const fix = opts.fix === true;
  const aggressive = opts.aggressive === true;

  const names = zip.getEntries().map((e) => e.entryName).filter((n) => !n.endsWith("/"));
  const present = new Set(names);

  const findings: Finding[] = [];
  const rewrote = new Set<string>();

  findings.push(...checkRequiredParts(present));
  findings.push(...checkWellFormed(zip, names));

  const prefixes = checkPrefixes(zip, names, fix);
  findings.push(...prefixes.findings);
  prefixes.rewrote.forEach((p) => rewrote.add(p));

  const ct = checkContentTypes(zip, names, fix);
  findings.push(...ct.findings);
  if (ct.rewrote) rewrote.add(CONTENT_TYPES);

  const rels = checkRelationships(zip, names, present, { fix, aggressive });
  findings.push(...rels.findings);
  rels.rewrote.forEach((p) => rewrote.add(p));

  // Body order first: it splices whole blocks, and the sequence pass below then
  // sees (and rewrites) the part those splices produced.
  const body = checkBody(zip, fix);
  findings.push(...body.findings);
  body.rewrote.forEach((p) => rewrote.add(p));

  const stories = checkStories(zip, names, fix);
  findings.push(...stories.findings);
  stories.rewrote.forEach((p) => rewrote.add(p));

  const structure = checkStructure(zip, names, fix);
  findings.push(...structure.findings);
  structure.rewrote.forEach((p) => rewrote.add(p));

  const seq = checkSequences(zip, names, fix);
  findings.push(...seq.findings);
  seq.rewrote.forEach((p) => rewrote.add(p));

  const spacing = checkTextSpacing(zip, names, fix);
  findings.push(...spacing.findings);
  spacing.rewrote.forEach((p) => rewrote.add(p));

  const bookmarks = checkBookmarks(zip, names, fix);
  findings.push(...bookmarks.findings);
  bookmarks.rewrote.forEach((p) => rewrote.add(p));

  findings.push(...checkTableGrids(zip, names));
  findings.push(...checkReferences(zip, names, present));

  findings.sort((a, b) => (SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity]) || a.part.localeCompare(b.part));

  return {
    ok: !findings.some((f) => f.severity === "fatal" && !f.fixed),
    checkedParts: names.length,
    findings,
    repairedParts: [...rewrote].sort(),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Zip-level inspection and repair
// ─────────────────────────────────────────────────────────────────────────────

const SIG_EOCD = 0x06054b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;
const SIG_DESCRIPTOR = 0x08074b50;
/** General-purpose bit 3: "sizes and CRC live in a trailing data descriptor". */
const FLAG_DATA_DESCRIPTOR = 0x0008;

/** One zip entry as the bytes state it. */
interface ZipEntryFacts {
  name: string;
  centralAt: number;
  localAt: number;
  centralFlags: number;
  localFlags: number;
  crc: number;
  size: number;
  /** Bit 3 is set in a header, and no descriptor stands behind the entry's data. */
  falseDescriptor: boolean;
}

/** Is a data descriptor really at `end`? The signed form (32 or 64 bit) is its own
 *  proof; the unsigned form has to repeat the central directory's values exactly. */
function descriptorAt(buf: Buffer, end: number, limit: number, crc: number, compressed: number, size: number): boolean {
  if (end + 4 > limit) return false;
  if (buf.readUInt32LE(end) === SIG_DESCRIPTOR) return true;
  return end + 12 <= limit
    && buf.readUInt32LE(end) === crc
    && buf.readUInt32LE(end + 4) === compressed
    && buf.readUInt32LE(end + 8) === size;
}

/**
 * Walk the central directory and each local header.
 *
 * Null on anything unexpected (Zip64, truncated directory, a local header that
 * doesn't line up, or a flagged entry with no descriptor whose local header
 * disagrees with the directory, since then the file's real values are unknown).
 * A half-understood zip is reported on by nobody and rewritten by nobody.
 */
function readZipEntries(buffer: Buffer): ZipEntryFacts[] | null {
  // The EOCD sits at the end, after an optional comment of up to 64 KB.
  const scanFrom = Math.max(0, buffer.length - (0xffff + 22));
  let eocd = -1;
  for (let i = buffer.length - 22; i >= scanFrom; i--) {
    if (buffer.readUInt32LE(i) === SIG_EOCD) { eocd = i; break; }
  }
  if (eocd === -1) return null;

  const count = buffer.readUInt16LE(eocd + 10);
  const dirOffset = buffer.readUInt32LE(eocd + 16);
  // 0xFFFF/0xFFFFFFFF are Zip64 escapes: out of scope.
  if (count === 0xffff || dirOffset === 0xffffffff || dirOffset >= buffer.length) return null;

  const entries: ZipEntryFacts[] = [];
  let p = dirOffset;
  for (let n = 0; n < count; n++) {
    if (p + 46 > buffer.length || buffer.readUInt32LE(p) !== SIG_CENTRAL) return null;
    const centralFlags = buffer.readUInt16LE(p + 8);
    const crc = buffer.readUInt32LE(p + 16);
    const compressed = buffer.readUInt32LE(p + 20);
    const size = buffer.readUInt32LE(p + 24);
    const nameLen = buffer.readUInt16LE(p + 28);
    const extraLen = buffer.readUInt16LE(p + 30);
    const commentLen = buffer.readUInt16LE(p + 32);
    const localAt = buffer.readUInt32LE(p + 42);
    if (compressed === 0xffffffff || size === 0xffffffff || localAt === 0xffffffff) return null;
    if (p + 46 + nameLen > buffer.length) return null;
    if (localAt + 30 > buffer.length || buffer.readUInt32LE(localAt) !== SIG_LOCAL) return null;

    const localFlags = buffer.readUInt16LE(localAt + 6);
    let falseDescriptor = false;
    if ((centralFlags | localFlags) & FLAG_DATA_DESCRIPTOR) {
      const dataEnd = localAt + 30 + buffer.readUInt16LE(localAt + 26) + buffer.readUInt16LE(localAt + 28) + compressed;
      if (!descriptorAt(buffer, dataEnd, dirOffset, crc, compressed, size)) {
        // No descriptor, so the local header has to carry the real values itself.
        if (buffer.readUInt32LE(localAt + 14) !== crc
          || buffer.readUInt32LE(localAt + 18) !== compressed
          || buffer.readUInt32LE(localAt + 22) !== size) return null;
        falseDescriptor = true;
      }
    }

    entries.push({
      name: buffer.toString("utf8", p + 46, p + 46 + nameLen),
      centralAt: p,
      localAt,
      centralFlags,
      localFlags,
      crc,
      size,
      falseDescriptor,
    });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/** What a package's own zip headers get wrong, by entry name. */
export interface ZipFlagReport {
  /** Non-empty entries whose local header promises a descriptor that is not there.
   *  Word refuses the whole file, even when it is a single entry. */
  falseDescriptors: string[];
  /** The same false promise where Word reads past it: an empty entry (a directory,
   *  a zero-length file), or a flag standing only in the central directory. */
  toleratedFalseDescriptors: string[];
  /** Non-empty entries that record CRC 0. Word opens them; adm-zip and `unzip -t`
   *  do not. Reported only: no flag explains it, so nothing here can repair it. */
  zeroCrc: string[];
}

/**
 * Read a package's zip flags without touching it. Null when the zip can't be
 * walked (see `readZipEntries`). `clearFalseDataDescriptors` clears exactly the
 * two descriptor lists, so what this names is what a repair clears.
 */
export function inspectZipFlags(buffer: Buffer): ZipFlagReport | null {
  const entries = readZipEntries(buffer);
  if (!entries) return null;
  const report: ZipFlagReport = { falseDescriptors: [], toleratedFalseDescriptors: [], zeroCrc: [] };
  for (const e of entries) {
    if (e.size > 0 && e.crc === 0) report.zeroCrc.push(e.name);
    if (!e.falseDescriptor) continue;
    if (e.size > 0 && (e.localFlags & FLAG_DATA_DESCRIPTOR)) report.falseDescriptors.push(e.name);
    else report.toleratedFalseDescriptors.push(e.name);
  }
  return report;
}

/**
 * Clear the "data descriptor present" flag on entries that do not have one.
 *
 * LibreOffice (and anything else that streams a zip) sets general-purpose bit 3
 * and appends a descriptor after each entry. adm-zip re-writes such an archive
 * with real sizes and CRCs in the headers but WITHOUT the descriptor, while
 * faithfully copying the flag that promises one. `unzip -t` still passes, but
 * real Word refuses the file (-10000, even with one such entry; measured on Word
 * for Mac), and adm-zip's own reader throws "No descriptor present".
 *
 * That matters here because the server re-opens its own output on the next edit:
 * one save of a LibreOffice-authored upload was enough to make every later AI
 * edit of that thesis fail to load. Clearing the lying flag is a two-byte fix per
 * entry, and it makes the bytes describe what is actually in the file.
 *
 * Each entry is judged on its own evidence: is there a descriptor behind its
 * data? So a truly streamed archive comes back as the same buffer, and so does
 * anything `readZipEntries` won't vouch for. A CRC of 0 proves nothing either
 * way: directories and empty files carry one legitimately, and they lose a false
 * flag like any other entry. Never mutates its input.
 */
export function clearFalseDataDescriptors(buffer: Buffer): Buffer {
  const lying = (readZipEntries(buffer) ?? []).filter((e) => e.falseDescriptor);
  if (!lying.length) return buffer;
  const out = Buffer.from(buffer);
  for (const e of lying) {
    out.writeUInt16LE(e.centralFlags & ~FLAG_DATA_DESCRIPTOR, e.centralAt + 8);
    out.writeUInt16LE(e.localFlags & ~FLAG_DATA_DESCRIPTOR, e.localAt + 6);
  }
  return out;
}

/** The report's findings for a zip's flags; `cleared` when a repair removed the false ones. */
function zipFindings(flags: ZipFlagReport | null, cleared: boolean): Finding[] {
  if (!flags) return [];
  const out: Finding[] = [];
  if (flags.falseDescriptors.length) {
    out.push(finding({
      rule: "zip.false-data-descriptor",
      severity: "fatal",
      part: "package",
      count: flags.falseDescriptors.length,
      message: `${flags.falseDescriptors.length} zip entr${flags.falseDescriptors.length === 1 ? "y promises" : "ies promise"} a data descriptor that is not there. Word refuses the whole file. Clearing the flag repairs it without touching any content.`,
      detail: sample(flags.falseDescriptors),
      fixable: true,
      fixed: cleared,
    }));
  }
  if (flags.toleratedFalseDescriptors.length) {
    out.push(finding({
      rule: "zip.false-data-descriptor-tolerated",
      severity: "warning",
      part: "package",
      count: flags.toleratedFalseDescriptors.length,
      message: `${flags.toleratedFalseDescriptors.length} zip entr${flags.toleratedFalseDescriptors.length === 1 ? "y promises" : "ies promise"} a data descriptor that is not there, on an empty entry or in the central directory only. Word reads past it, but stricter readers may not.`,
      detail: sample(flags.toleratedFalseDescriptors),
      fixable: true,
      fixed: cleared,
    }));
  }
  if (flags.zeroCrc.length) {
    out.push(finding({
      rule: "zip.crc-zero",
      severity: "warning",
      part: "package",
      count: flags.zeroCrc.length,
      message: `${flags.zeroCrc.length} zip entr${flags.zeroCrc.length === 1 ? "y records" : "ies record"} a checksum of 0 despite having content. Word opens the file, but unzip and this engine's own reader refuse those entries.`,
      detail: sample(flags.zeroCrc),
      fixable: false,
    }));
  }
  return out;
}

/**
 * A part the package names that cannot be read. `readPart` takes a throw for an absent part, so every
 * check passes over such a part and the report would say the document is fine, while Word refuses it
 * (a flag with no descriptor over a local header that disagrees, damaged data). A CRC-0 entry is left
 * to its own finding: Word opens those. Binary parts are not read: this runs on every save.
 */
function unreadableFindings(zip: DocxZip, flags: ZipFlagReport | null): Finding[] {
  const unreadable: string[] = [];
  for (const { entryName } of zip.getEntries()) {
    if (!/\.(?:xml|rels)$/i.test(entryName) || flags?.zeroCrc.includes(entryName)) continue;
    try {
      zip.readAsText(entryName);
    } catch {
      unreadable.push(entryName);
    }
  }
  if (!unreadable.length) return [];
  return [finding({
    rule: "zip.unreadable-part",
    severity: "fatal",
    part: "package",
    count: unreadable.length,
    message: `${unreadable.length} part${unreadable.length === 1 ? "" : "s"} of the package cannot be read, so none of them could be checked. Word is expected to refuse the file.`,
    detail: sample(unreadable),
    fixable: false,
    fixed: false,
  })];
}

/** A report with the zip's findings merged in, in the usual order, with `ok` restated. */
function withZipFindings<T extends DoctorReport>(report: T, extra: Finding[]): T {
  if (!extra.length) return report;
  const findings = [...report.findings, ...extra]
    .sort((a, b) => (SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity]) || a.part.localeCompare(b.part));
  return { ...report, findings, ok: !findings.some((f) => f.severity === "fatal" && !f.fixed) };
}

/** Open a .docx from bytes, healing the zip-level flag FIRST. adm-zip throws when a
 *  flagged entry is READ, not when the zip is opened, and `readPart` takes a throw
 *  for an absent part, so a healed copy is what makes every check below see the
 *  document at all. A plain ZipManager on purpose: the doctor only ever needs the
 *  parts, and constructing a full Mdocxengine runs every PartsManager, one of which
 *  throws on exactly the malformed packages this module exists to diagnose. */
async function openBuffer(buffer: Buffer): Promise<DocxZip> {
  return (await ZipManager.loadFromBuffer(clearFalseDataDescriptors(buffer))) as unknown as DocxZip;
}

/** Check a .docx held as bytes. Read-only: the buffer is never modified. */
export async function checkDocxBuffer(buffer: Buffer): Promise<DoctorReport> {
  const zip = await openBuffer(buffer);
  const flags = inspectZipFlags(buffer);
  return withZipFindings(inspectDocx(zip), [...zipFindings(flags, false), ...unreadableFindings(zip, flags)]);
}

export interface RepairResult extends DoctorReport {
  /** The repaired bytes, or the ORIGINAL buffer when nothing was rewritten. */
  buffer: Buffer;
  /** True when `buffer` differs from what was passed in. */
  changed: boolean;
}

/**
 * Repair a .docx held as bytes.
 *
 * The result is re-opened and re-checked before it is handed back: if the repair
 * somehow left a part unreadable, or turned a clean fatal into a new one, the
 * whole thing is discarded and the original bytes are returned. A doctor that
 * can corrupt a thesis is worse than no doctor.
 *
 * Zip flags are repaired too, and stay on the report as fixed `zip.*` findings.
 * `repairedParts` keeps its meaning, parts rewritten: clearing a flag rewrites
 * none, so go by `changed` for whether the bytes moved.
 */
export async function repairDocxBuffer(
  buffer: Buffer,
  opts: { aggressive?: boolean } = {},
): Promise<RepairResult> {
  const flags = inspectZipFlags(buffer);
  const healed = clearFalseDataDescriptors(buffer);
  const flagsCleared = healed !== buffer;

  const zip = await openBuffer(healed);
  const report = inspectDocx(zip, { fix: true, aggressive: opts.aggressive });
  const packageFindings = [...zipFindings(flags, flagsCleared), ...unreadableFindings(zip, flags)];
  if (!report.repairedParts.length) {
    return { ...withZipFindings(report, packageFindings), buffer: healed, changed: flagsCleared };
  }

  /** Give up on the repair and hand back exactly what we were given. */
  const abandon = async (why: string): Promise<RepairResult> => {
    const original = await checkDocxBuffer(buffer);
    original.findings.push(finding({
      rule: "repair.abandoned",
      severity: "warning",
      part: "package",
      count: 1,
      message: `The automatic repair was discarded and the document left exactly as it was (${why}).`,
      fixable: false,
    }));
    return { ...original, repairedParts: [], buffer, changed: false };
  };

  try {
    const repaired = clearFalseDataDescriptors(Buffer.from((zip as unknown as ZipManager).toBuffer()));

    // Verify against the bytes we would actually store: every rewritten part has
    // to read back and still parse, and the repair must not have invented a new
    // fatal. A doctor that can corrupt a thesis is worse than no doctor.
    const vzip = await openBuffer(repaired);
    for (const part of report.repairedParts) {
      const xml = readPart(vzip, part);
      if (xml == null || firstXmlError(xml)) return abandon(`${part} did not survive the rewrite`);
    }

    const after = inspectDocx(vzip);
    const fatalsBefore = report.findings.filter((f) => f.severity === "fatal").length;
    const fatalsAfter = after.findings.filter((f) => f.severity === "fatal").length;
    if (fatalsAfter > fatalsBefore) return abandon("it would have introduced a new problem");

    return {
      ...withZipFindings(after, packageFindings),
      repairedParts: report.repairedParts,
      buffer: repaired,
      changed: true,
    };
  } catch (e) {
    return abandon(e instanceof Error ? e.message : String(e));
  }
}

/** One-line summary for a log line or a tool reply. */
export function summarize(report: DoctorReport): string {
  const fatal = report.findings.filter((f) => f.severity === "fatal");
  const warn = report.findings.filter((f) => f.severity === "warning");
  if (!fatal.length && !warn.length) return "Document is healthy — no problems found.";
  const bits: string[] = [];
  if (fatal.length) bits.push(`${fatal.length} serious problem(s)`);
  if (warn.length) bits.push(`${warn.length} warning(s)`);
  const fixed = report.findings.filter((f) => f.fixed).length;
  return `${bits.join(", ")}${fixed ? `, ${fixed} repaired` : ""}.`;
}

// The scanner primitives are part of the doctor's public surface (the package
// index re-exports firstXmlError for callers that only need a well-formedness check).
export { firstXmlError } from "./xml";
