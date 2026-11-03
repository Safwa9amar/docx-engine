/**
 * MergeManager — copy another .docx's body into THIS document, fully remapped.
 *
 * mdocxengine stores each body child as its exact original XML substring
 * (see OrderedBody). Concatenating blocks from two documents is therefore
 * byte-safe — EXCEPT for anything a block references by id across the package
 * boundary: image relationships (`r:embed`/`r:link`), footnote references
 * (`w:footnoteReference/@w:id`), numbering (`w:numId`, Phase 2) and styles.
 *
 * `appendDocument` copies the source body and rewrites those references so they
 * resolve against THIS document:
 *   - images   → copied into word/media (new rId via MediaManager.insertImage)
 *   - footnotes→ copied into word/footnotes.xml (new id via FootnoteManager)
 *   - endnotes → copied into word/endnotes.xml (new id via EndnoteManager)
 *   - comments → NOT copied. Their anchors are stripped and counted: a thesis going
 *                to deposit never carries them, and a `w:commentReference` left with
 *                no comments part behind it makes Word refuse the file.
 *   - styles   → reconciled by NAME, Word's own identity for a style: see mergeStyles.ts.
 *                A caller's styleMap (the combine flow normalizes to one profile) outranks it.
 *   - charts   → the chart part and its whole closure (colours, style, embedded
 *                workbook) copied into word/charts + word/embeddings, with a new
 *                rId. A `<c:chart r:id>` addresses a PART, so copying only the
 *                paragraph aims it at whatever the TARGET has under that id.
 *   - equations (OMML) are inline in the paragraph XML and copied verbatim.
 *
 * The rewrite is attribute-aware and applied to EACH copied block string
 * individually — never a blind global string replace, and never a join/split
 * round-trip (block XML is full of spaces).
 */

import AdmZip from "adm-zip";
import DocumentManager from "./DocumentManager";
import { MediaManager } from "./MediaManager";
import { FootnoteManager, SEPARATOR_NOTE_TYPE } from "./FootnoteManager";
import { EndnoteManager } from "./EndnoteManager";
import { NumberingManager, type RawNumbering } from "./NumberingManager";
import { RelManager } from "./RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import { SettingsManager } from "./SettingsManager";
import { planStyles, retargetListRefs, retargetStyleRefs, stampNumbering, type StylePlan } from "./mergeStyles";
import { parseXml } from "@/utils/xmlUtils";
import { carryRootNamespaces } from "@/utils/rootNamespaces";
import { assembleDocument, splitDocument, type BodyBlock } from "@/core/files/body/OrderedBody";
import { clearFalseDataDescriptors } from "@/doctor";
import { canonicalMainPart, canonicalizeMainPart } from "@/core/package/canonicalMainPart";
// Imported lazily inside appendDocument() to avoid an index.ts ↔ MergeManager cycle.
import type { Mdocxengine as MdocxengineType } from "@/index";

const SRC_RELS_PATH = "word/_rels/document.xml.rels";
const NUMBERING_RELS_PATH = "word/_rels/numbering.xml.rels";
const PAGE_BREAK_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

function pageBreakBlock(): BodyBlock {
  return { kind: "paragraph", tag: "w:p", xml: PAGE_BREAK_XML };
}

/**
 * Content types for the parts a chart drags along, used only when the SOURCE
 * package declares no Default for the extension (Word always does; a generator
 * may not). The embedded workbook is the one that matters — without a Default
 * for `xlsx` Word cannot open the chart's data sheet.
 */
const DEFAULT_CONTENT_TYPES: Record<string, string> = {
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xls: "application/vnd.ms-excel",
  bin: "application/vnd.openxmlformats-officedocument.oleObject",
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
};

/** Resolve a `.rels` Target against the part that declared it, POSIX-style.
 *  `("word/charts/chart1.xml", "../embeddings/w.xlsx")` → `word/embeddings/w.xlsx`. */
function resolvePartPath(fromPart: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const segs = fromPart.split("/").slice(0, -1).concat(target.split("/"));
  const out: string[] = [];
  for (const s of segs) {
    if (s === "." || s === "") continue;
    if (s === "..") out.pop();
    else out.push(s);
  }
  return out.join("/");
}

/** Express `toPart` relative to the directory holding `fromPart`. */
function relativePartPath(fromPart: string, toPart: string): string {
  const from = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
}

/** `word/charts/chart1.xml` → `word/charts/_rels/chart1.xml.rels`. */
function partRelsPath(part: string): string {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
}

const asArray = <T>(x: T | T[] | undefined | null): T[] => (Array.isArray(x) ? x : x == null ? [] : [x]);

type NoteKind = "footnote" | "endnote";
const NOTE_KINDS: NoteKind[] = ["footnote", "endnote"];

/** The note ids a stretch of body XML refers to. */
function noteRefIds(xml: string, kind: NoteKind): Set<string> {
  return new Set([...xml.matchAll(new RegExp(`<w:${kind}Reference\\b[^>]*\\bw:id="([^"]+)"`, "g"))].map((m) => m[1]));
}

/** A note's own separators, by what its open tag says it is. */
const separatorNotes = (kind: NoteKind) =>
  new RegExp(`<w:${kind}\\b(?=[^>]*\\bw:type="(?:separator|continuationSeparator|continuationNotice)")[^>]*?(?:/>|>[\\s\\S]*?</w:${kind}>)`, "g");

const NOTE_PROPS = ["pos", "numFmt", "numStart", "numRestart"];
/** Word's values for what no setting names. */
const NOTE_PROP_DEFAULTS: Record<NoteKind, Record<string, string>> = {
  footnote: { pos: "pageBottom", numFmt: "decimal", numRestart: "continuous" },
  endnote: { pos: "docEnd", numFmt: "lowerRoman", numRestart: "continuous" },
};

/** The first `w:footnotePr` or `w:endnotePr` of a settings part or a section, as XML. */
function notePr(xml: string | null | undefined, kind: NoteKind): string | null {
  return new RegExp(`<w:${kind}Pr\\b[^>]*?(?:/>|>[\\s\\S]*?</w:${kind}Pr>)`).exec((xml ?? "").split("<w:sectPrChange")[0])?.[0] ?? null;
}

/** The numbering and placement a note property element states, as XML and by name. */
function noteProps(element: string | null): { xml: string; byName: Record<string, string> } {
  const byName: Record<string, string> = {};
  let xml = "";
  for (const m of (element ?? "").matchAll(new RegExp(`<w:(${NOTE_PROPS.join("|")})\\b[^>]*?(?:/>|>\\s*</w:\\1>)`, "g"))) {
    xml += m[0];
    const value = /\bw:val="([^"]*)"/.exec(m[0])?.[1] ?? "";
    const format = /\bw:format="([^"]*)"/.exec(m[0])?.[1];
    byName[m[1]] = format ? `${value} ${format}` : value;
  }
  return { xml, byName };
}

/** The separator list a `w:footnotePr` of a settings part holds. */
function noteSeparatorRefs(element: string | null, kind: NoteKind): string {
  return (element ?? "").match(new RegExp(`<w:${kind}\\b[^>]*?(?:/>|>[\\s\\S]*?</w:${kind}>)`, "g"))?.join("") ?? "";
}

/** What numbers and places a document's notes in its final section. A start value is left out:
 *  a part set to start at 31 is carrying on from the file before it, which a merge does anyway. */
function effectiveNoteProps(kind: NoteKind, settings: string | null | undefined, finalSect: string | null): Record<string, string> {
  const { numStart: _a, ...fromSettings } = noteProps(notePr(settings, kind)).byName;
  const { numStart: _b, ...fromSection } = noteProps(notePr(finalSect, kind)).byName;
  return { ...NOTE_PROP_DEFAULTS[kind], ...fromSettings, ...fromSection };
}

/** The final section of a document, as XML. */
function finalSectPr(documentXml: string): string | null {
  return documentXml ? (splitDocument(documentXml).blocks.filter((b) => b.kind === "sectPr").pop()?.xml ?? null) : null;
}

/** A section with this kind's note properties replaced, in `CT_SectPr` order. */
function withNotePr(sect: string, kind: NoteKind, element: string | null): string {
  const open = /^<w:sectPr\b[^>]*?(\/?)>/.exec(sect);
  if (!open) return sect;
  if (open[1]) return element ? `${open[0].slice(0, -2).trimEnd()}>${element}</w:sectPr>` : sect;
  const change = sect.indexOf("<w:sectPrChange");
  const head = change === -1 ? sect : sect.slice(0, change);
  const tail = change === -1 ? "" : sect.slice(change);
  const bare = head.replace(new RegExp(`\\s*<w:${kind}Pr\\b[^>]*?(?:/>|>[\\s\\S]*?</w:${kind}Pr>)`), "");
  if (!element) return bare + tail;
  const earlier = kind === "footnote" ? "headerReference|footerReference" : "headerReference|footerReference|footnotePr";
  const skip = new RegExp(`\\s*<w:(${earlier})\\b[^>]*?(?:/>|>[\\s\\S]*?</w:\\1>)`, "y");
  let at = open[0].length;
  skip.lastIndex = at;
  while (skip.exec(bare)) at = skip.lastIndex;
  return bare.slice(0, at) + element + bare.slice(at) + tail;
}

/**
 * Run `fix` over each field code as one string. Word splits a code across runs wherever it likes,
 * so a name can stand in two `w:instrText`. A changed code goes whole into the first piece; the
 * other pieces stay, empty. A nested field's code is its own.
 */
function rewriteFieldCodes(xml: string, fix: (code: string) => string): string {
  if (!xml.includes("<w:instrText")) return xml;
  type Piece = { start: number; end: number; open: string; text: string; close: string };
  const token = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"[^>]*?(?:\/>|>[\s\S]*?<\/w:fldChar>)|(<w:instrText\b[^>]*?(?<!\/)>)([^<]*)(<\/w:instrText>)/g;
  const edits: { start: number; end: number; xml: string }[] = [];
  const settle = (pieces: Piece[]) => {
    if (!pieces.length) return;
    const whole = pieces.map((p) => p.text).join("");
    const fixed = fix(whole);
    if (fixed === whole) return;
    pieces.forEach((p, i) => {
      const open = i === 0 && !/\bxml:space=/.test(p.open) ? p.open.replace(/>$/, ' xml:space="preserve">') : p.open;
      edits.push({ start: p.start, end: p.end, xml: open + (i === 0 ? fixed : "") + p.close });
    });
  };
  const open: Piece[][] = [];
  for (const m of xml.matchAll(token)) {
    if (m[1] === "begin") open.push([]);
    else if (m[1] === "separate" || m[1] === "end") {
      const top = open[open.length - 1];
      if (top) {
        settle(top);
        top.length = 0;
      }
      if (m[1] === "end") open.pop();
    } else {
      const piece = { start: m.index!, end: m.index! + m[0].length, open: m[2], text: m[3], close: m[4] };
      const top = open[open.length - 1];
      if (top) top.push(piece);
      else settle([piece]);
    }
  }
  open.forEach(settle);
  let out = xml;
  for (const e of edits.sort((a, b) => b.start - a.start)) out = out.slice(0, e.start) + e.xml + out.slice(e.end);
  return out;
}

// ─── Bookmark names ─────────────────────────────────────────────────────────

const bookmarkKey = (name: string) => name.toLowerCase();
const BOOKMARK_NAME_MAX = 40;
/** A name not yet in use, within what Word allows: it does not keep a longer one. */
function renamedBookmark(name: string, k: number): string {
  const suffix = `_m${k}`;
  return name.slice(0, BOOKMARK_NAME_MAX - suffix.length) + suffix;
}

/** A word at the head of a field is its type when Word knows it, and a bookmark only otherwise. */
const FIELD_TYPES = new Set(
  ("addin addressblock advance ask author autonum autonumlgl autonumout autotext autotextlist barcode bibliography bidioutline citation " +
    "comments compare createdate database date docproperty docvariable edittime eq filename filesize fillin formcheckbox formdropdown " +
    "formtext ftnref gotobutton greetingline hyperlink if includepicture includetext index info keywords lastsavedby link listnum " +
    "macrobutton mergefield mergerec mergeseq next nextif noteref numchars numpages numwords page pageref print printdate private quote " +
    "rd ref revnum savedate section sectionpages seq set skipif styleref subject symbol ta tc template time title toa toc useraddress " +
    "userinitials username xe").split(" "),
);
const FIELD_TOKEN = /&quot;[\s\S]*?&quot;|"[^"]*"|[^\s"]+/g;
const unquoted = (token: string) => token.replace(/^(?:&quot;|")|(?:&quot;|")$/g, "");
const isQuoted = (token: string) => /^(?:&quot;|")/.test(token);
const afterSwitch = (tokens: string[], name: string) => {
  const at = tokens.findIndex((t) => t.toLowerCase() === name);
  return at < 0 ? -1 : at + 1;
};

/**
 * Which word of a field code is a bookmark, and put the new name there. Field codes are not
 * case-sensitive, and a field can be nothing but the bookmark's name.
 */
function renameBookmarksInField(code: string, renames: Map<string, string>): string {
  const tokens = [...code.matchAll(FIELD_TOKEN)];
  if (!tokens.length) return code;
  const words = tokens.map((m) => m[0]);
  const type = unquoted(words[0]).toLowerCase();

  const rewrite = (edit: (token: string, i: number) => string) => {
    let out = "";
    let at = 0;
    tokens.forEach((m, i) => {
      out += code.slice(at, m.index!) + edit(m[0], i);
      at = m.index! + m[0].length;
    });
    return out + code.slice(at);
  };

  // A formula names bookmarks as its operands, anywhere that is not a quoted string.
  if (type.startsWith("=") || type === "if") {
    return rewrite((token) => {
      if (isQuoted(token)) return token;
      let out = token;
      for (const [key, to] of renames) {
        const name = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        out = out.replace(new RegExp(`(?<![\\p{L}\\p{N}_&])${name}(?![\\p{L}\\p{N}_])`, "giu"), to);
      }
      return out;
    });
  }

  let at = -1;
  if (type === "ref" || type === "pageref" || type === "noteref") at = 1;
  else if (type === "hyperlink") at = afterSwitch(words, "\\l");
  else if (type === "toc" || type === "index" || type === "toa") at = afterSwitch(words, "\\b");
  else if (type === "seq") at = words[2]?.startsWith("\\") ? -1 : 2;
  else if (!FIELD_TYPES.has(type)) at = 0;
  if (at < 0 || at >= words.length) return code;

  const renamed = renames.get(bookmarkKey(unquoted(words[at])));
  if (!renamed) return code;
  return rewrite((token, i) => (i === at ? token.replace(unquoted(token), renamed) : token));
}

/** Elements whose relationship attributes have a path of their own, or which a merge removes. */
const REL_OWN_PATH = new Set([
  "w:hyperlink", "c:chart", "v:imagedata", "v:fill", "o:OLEObject", "w:headerReference", "w:footerReference", "w:printerSettings",
]);
const ANY_OPEN_TAG = /<([\w.-]+:[\w.-]+)\b[^>]*>/g;
/** `r:embed` and `r:link` are the media path's, whatever element holds them. */
const OTHER_REL_ATTR = /\s(r:(?!embed\b|link\b)[a-zA-Z]+|o:relid)="([^"]+)"/g;

/**
 * What ties a content control to a store: the package's core properties, its custom XML. Word
 * replaces the control's text from that store when it opens the file, and a later part's store stays
 * behind, so its control would print part 1's title, or nothing. Without the binding it keeps the
 * words it holds. An attribute may hold a `>`: an XPath can.
 */
const DATA_BINDING = /<(w|w15):dataBinding\b(?:"[^"]*"|'[^']*'|[^>"'])*?(?:\/>|>\s*<\/\1:dataBinding>)/g;

const safeDecode = (s: string) => {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
};

// ─── Comment markup ─────────────────────────────────────────────────────────

const COMMENT_MARK_ID = /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*\bw:id="([^"]*)"/g;
/**
 * A run that holds the reference and nothing else (an rPr aside) goes with it. `(?<!\/)` keeps
 * a self-closed tag (`<w:rPr />`, the form .NET's XmlWriter gives an empty element) from being
 * read as an opening one, which would run the removal on through the student's words.
 */
const COMMENT_ONLY_RUN =
  /<w:r(?:\s[^>]*?)?(?<!\/)>\s*(?:<w:rPr(?:\s[^>]*?)?\/>|<w:rPr(?:\s[^>]*?)?(?<!\/)>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?\s*<w:commentReference\b[^>]*?(?:\/>|>\s*<\/w:commentReference>)\s*<\/w:r>/g;
const COMMENT_MARK = /<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*?(?:\/>|>\s*<\/w:comment\1>)/g;
/** What a comment anchor can never hold: words and the elements around them. */
const COMMENT_FREE_CONTENT = /<w:t(?:\s[^>]*)?>[^<]*<\/w:t>|<\/?w:(?:p|tc|tr|tbl|hyperlink|footnote|endnote)(?=[\s>\/])/g;
const commentFreeContent = (xml: string) => (xml.match(COMMENT_FREE_CONTENT) ?? []).join("");

/**
 * `xml` with every comment anchor gone, nothing else changed; their ids land in `ids`. A removal
 * that took anything else with it throws: it reaches every route, including those with no check
 * behind them, and losing a student's words is worse than refusing the file.
 */
function stripCommentMarks(xml: string, ids: Set<string>): string {
  if (!xml.includes("<w:comment")) return xml;
  for (const m of xml.matchAll(COMMENT_MARK_ID)) ids.add(m[1]);
  const out = xml.replace(COMMENT_ONLY_RUN, "").replace(COMMENT_MARK, "");
  if (commentFreeContent(out) !== commentFreeContent(xml)) {
    throw new Error("merge: removing the comment anchors would have changed the text around them");
  }
  return out;
}

/** Blocks with their comment anchors gone. A block that was nothing BUT an anchor (a
 *  range mark standing between paragraphs) goes too; every other block keeps its bytes. */
function stripCommentMarksFromBlocks(blocks: BodyBlock[], ids: Set<string>): BodyBlock[] {
  const out: BodyBlock[] = [];
  for (const b of blocks) {
    const xml = stripCommentMarks(b.xml, ids);
    if (xml === b.xml) out.push(b);
    else if (xml.trim()) out.push({ ...b, xml });
  }
  return out;
}

/** The ids of the comments a comments part defines. */
function definedCommentIds(commentsXml: string | null | undefined, ids: Set<string>): void {
  if (!commentsXml) return;
  for (const m of commentsXml.matchAll(/<w:comment\b[^>]*\bw:id="([^"]*)"/g)) ids.add(m[1]);
}

const COMMENT_PART = /^word\/comments(?:Extended|Ids|Extensible)?\.xml$/i;
const COMMENT_REL_TYPE = /\/comments(?:Extended|Ids|Extensible)?$/i;
/** Stories that could hold an anchor besides the body. Word's UI only puts them in
 *  notes, but the schema allows any run, and one left behind makes Word refuse the file. */
const COMMENTABLE_PART = /^word\/(?:footnotes|endnotes|header\d*|footer\d*)\.xml$/;
const RELATIONSHIP_ELEMENT = /<Relationship\b[^>]*?(?:\/>|>\s*<\/Relationship>)/g;
const attrOf = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

/** A part's identity across rels files: Targets may be percent-encoded or differ in case. */
function partKey(path: string): string {
  try {
    return decodeURIComponent(path).toLowerCase();
  } catch {
    return path.toLowerCase();
  }
}

interface IdPlan {
  nextDoc: () => number;
  bmIds: Map<string, string>;
  renames: Map<string, string>;
}

export interface AppendOptions {
  /** Prepend a page break before everything this call appends. */
  startOnNewPage?: boolean;
  /** Blocks inserted BEFORE the copied source body (e.g. a part-title heading). */
  leadingBlocks?: BodyBlock[];
  /** source styleId → target styleId (applied to `<w:pStyle w:val>`). */
  styleMap?: Record<string, string>;
}

/** A later part wanted its notes numbered or placed otherwise. One final section, one answer. */
export interface NoteConflict {
  kind: NoteKind;
  property: string;
  kept: string;
  dropped: string;
}

export interface AppendResult {
  /** Comments the source carried, none of which came across. */
  commentsStripped: number;
  /** Where the source's note properties gave way to this document's. */
  noteConflicts: NoteConflict[];
  /** Content controls of the source that lost their tie to a store left behind, keeping their words. */
  bindingsDropped: number;
}

/**
 * Common French/English Word heading + body style aliases → canonical target
 * styleIds. The combine flow passes this so mismatched source heading styles map
 * onto the template's styles by name.
 */
export const DEFAULT_STYLE_ALIASES: Record<string, string> = {
  Titre1: "Heading1",
  Titre2: "Heading2",
  Titre3: "Heading3",
  Titre4: "Heading4",
  Heading1: "Heading1",
  Heading2: "Heading2",
  Heading3: "Heading3",
  Heading4: "Heading4",
  Normal: "Normal",
  Standard: "Normal",
  Corpsdetexte: "Normal",
  BodyText: "Normal",
};

export class MergeManager {
  private zip: AdmZip;
  private document: DocumentManager;
  private media: MediaManager;
  private footnotes: FootnoteManager;
  private endnotes: EndnoteManager;
  private numbering: NumberingManager;
  private rels: RelManager;
  private contentTypes: ContentTypesManager;
  private settings: SettingsManager;

  constructor(zip: AdmZip) {
    this.zip = zip;
    this.document = new DocumentManager(zip);
    this.media = new MediaManager(zip);
    this.footnotes = new FootnoteManager(zip);
    this.endnotes = new EndnoteManager(zip);
    this.numbering = new NumberingManager(zip);
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
    this.settings = new SettingsManager(zip);
  }

  /**
   * Copy `sourceBuffer`'s body into this document, fully remapped, appended after
   * the existing body (before the trailing sectPr, handled by saveBlocks).
   * The source's comments do not come across; the result counts them. Comments
   * already in THIS document are left alone: see `stripComments`.
   */
  async appendDocument(sourceBuffer: Buffer, opts: AppendOptions = {}): Promise<AppendResult> {
    const { Mdocxengine } = await import("@/index");
    // Both packages in the shape every reader below expects: a part whose main document is
    // elsewhere would otherwise come across as nothing at all. One with none stops here.
    canonicalizeMainPart(this.zip);
    // A false zip flag is a two-byte fault that makes every read of that entry throw.
    const source = (await (Mdocxengine as typeof MdocxengineType).loadFromBuffer(
      canonicalMainPart(clearFalseDataDescriptors(sourceBuffer)),
    )) as MdocxengineType;

    const commentIds = new Set<string>();
    definedCommentIds(source.zip.readAsText("word/comments.xml"), commentIds);
    const srcBlocks = stripCommentMarksFromBlocks(await source.document.getBlocks(), commentIds);
    const noteConflicts: NoteConflict[] = [];
    let bindingsDropped = 0;
    const unbind = (xml: string) =>
      xml.replace(DATA_BINDING, () => {
        bindingsDropped++;
        return "";
      });
    const remapped = await this.remapBlocks(source, srcBlocks, opts.styleMap ?? {}, commentIds, noteConflicts, unbind);

    const existing = await this.document.getBlocks();
    const prefix: BodyBlock[] = [];
    if (opts.startOnNewPage) prefix.push(pageBreakBlock());
    if (opts.leadingBlocks?.length) prefix.push(...opts.leadingBlocks);

    await this.document.saveBlocks([...existing, ...prefix, ...remapped]);
    this.carryNamespaces(source);
    return { commentsStripped: commentIds.size, noteConflicts, bindingsDropped };
  }

  // ─── Comments ───────────────────────────────────────────────────────────────

  /**
   * Remove every comment from THIS document and return how many it had: the
   * anchors in the body and the notes, the comment parts with their relationships
   * and content types, and any picture only a comment used. Everything else keeps
   * its exact bytes; a document with no comments is not rewritten at all.
   */
  async stripComments(): Promise<number> {
    canonicalizeMainPart(this.zip);
    const ids = new Set<string>();
    definedCommentIds(this.zip.readAsText("word/comments.xml"), ids);

    const blocks = await this.document.getBlocks();
    const bare = stripCommentMarksFromBlocks(blocks, ids);
    if (bare.length !== blocks.length || bare.some((b, i) => b !== blocks[i])) {
      await this.document.saveBlocks(bare);
    }
    for (const name of this.zip.getEntries().map((e) => e.entryName).filter((n) => COMMENTABLE_PART.test(n))) {
      const xml = this.zip.readAsText(name);
      const stripped = stripCommentMarks(xml, ids);
      if (stripped !== xml) this.zip.addFile(name, Buffer.from(stripped, "utf8"));
    }
    this.removeCommentParts();
    return ids.size;
  }

  /** Delete the comment parts and everything that names them, by string surgery. */
  private removeCommentParts(): void {
    const names = this.zip.getEntries().map((e) => e.entryName);
    const parts = new Set(names.filter((n) => COMMENT_PART.test(n)));

    const relsXml = this.zip.readAsText(SRC_RELS_PATH);
    if (relsXml) {
      const kept = relsXml.replace(RELATIONSHIP_ELEMENT, (rel) => {
        const target = attrOf(rel, "Target");
        if (!target || !COMMENT_REL_TYPE.test(attrOf(rel, "Type") ?? "")) return rel;
        parts.add(resolvePartPath("word/document.xml", target));
        return "";
      });
      if (kept !== relsXml) this.zip.addFile(SRC_RELS_PATH, Buffer.from(kept, "utf8"));
    }
    if (parts.size === 0) return;

    const dependants = [...parts].flatMap((part) => this.internalTargets(partRelsPath(part)));
    const doomed = new Set(parts);
    for (const part of parts) doomed.add(partRelsPath(part));
    for (const part of doomed) if (this.zip.getEntry(part)) this.zip.deleteFile(part);

    // What only a comment pointed at (a pasted picture) must not stay behind in the
    // package; anything another part still reaches stays.
    const reached = new Set(
      this.zip
        .getEntries()
        .map((e) => e.entryName)
        .filter((n) => n.endsWith(".rels"))
        .flatMap((n) => this.internalTargets(n))
        .map(partKey),
    );
    for (const dep of dependants) {
      if (reached.has(partKey(dep)) || !this.zip.getEntry(dep)) continue;
      this.zip.deleteFile(dep);
      doomed.add(dep);
    }

    const contentTypes = this.zip.readAsText("[Content_Types].xml");
    if (contentTypes) {
      const next = contentTypes.replace(/<Override\b[^>]*?(?:\/>|>\s*<\/Override>)/g, (o) =>
        doomed.has((attrOf(o, "PartName") ?? "").replace(/^\//, "")) ? "" : o,
      );
      if (next !== contentTypes) this.zip.addFile("[Content_Types].xml", Buffer.from(next, "utf8"));
    }
  }

  /** The package paths a rels part points at, external targets excepted. */
  private internalTargets(relsPath: string): string[] {
    const xml = this.zip.readAsText(relsPath);
    if (!xml) return [];
    // "word/_rels/comments.xml.rels" describes "word/comments.xml"; "_rels/.rels" the root.
    const owner = relsPath.replace(/(^|\/)_rels\/([^/]*)\.rels$/, "$1$2");
    const out: string[] = [];
    for (const m of xml.matchAll(RELATIONSHIP_ELEMENT)) {
      const target = attrOf(m[0], "Target");
      if (target && attrOf(m[0], "TargetMode") !== "External") out.push(resolvePartPath(owner, target));
    }
    return out;
  }

  /**
   * The copied blocks/footnotes/numbering use prefixes (`w14`, `mc`, `wps`, `r`…)
   * the SOURCE root declared. Bind them on the matching TARGET root, or the part
   * has unbound prefixes and Word refuses to open it. A prefix bound to a
   * different URI is a real conflict: throw rather than silently mis-bind.
   */
  private carryNamespaces(source: MdocxengineType): void {
    for (const part of ["word/document.xml", "word/styles.xml", "word/footnotes.xml", "word/endnotes.xml", "word/numbering.xml"]) {
      const targetXml = this.zip.readAsText(part);
      const sourceXml = source.zip.readAsText(part);
      if (!targetXml || !sourceXml) continue;
      const { xml, conflicts } = carryRootNamespaces(targetXml, sourceXml);
      if (conflicts.length) {
        throw new Error(
          `merge: ${part} binds prefix(es) ${conflicts.join(", ")} to a different namespace in the source than in the target`,
        );
      }
      if (xml !== targetXml) this.zip.addFile(part, Buffer.from(xml, "utf8"));
    }
  }

  // ─── Remap orchestration ───────────────────────────────────────────────────

  /**
   * Build the id maps from a READ-ONLY concatenation of source blocks, then apply
   * them to EACH block's xml independently (kind/tag preserved verbatim).
   */
  private async remapBlocks(
    source: MdocxengineType,
    blocks: BodyBlock[],
    styleMap: Record<string, string>,
    commentIds: Set<string>,
    noteConflicts: NoteConflict[],
    unbind: (xml: string) => string,
  ): Promise<BodyBlock[]> {
    const scan = blocks.map((b) => b.xml).join("\n");
    // First, before anything is written to this document: it can refuse the source.
    const notesScan = NOTE_KINDS.map((kind) => this.neededNotesXml(source, kind, scan)).join("\n");
    const content = scan + "\n" + notesScan;
    const srcRels = await this.readSourceRels(source);
    const carry = this.relationshipCarrier(source, srcRels);
    const mediaMap = await this.buildMediaMap(carry, scan);
    const otherRelMap = await this.buildOtherRelMap(carry, scan);
    const hyperlinkMap = await this.buildHyperlinkMap(srcRels, scan);
    // A note brings its own styles and lists: `FootnoteText` often appears nowhere
    // but in the notes, and undefined it renders as Normal, at body size.
    const sourceLists = await source.numbering.getRawDefinitions();
    const styles = planStyles({
      sourceStyles: source.zip.readAsText("word/styles.xml"),
      targetStyles: this.zip.readAsText("word/styles.xml"),
      sourceLists,
      targetLists: await this.numbering.getRawDefinitions(),
      styleMap,
      contentXml: content,
    });
    // One numbering pass over blocks, notes AND imported styles, so a list used by
    // more than one maps to the same copied definition.
    const numberingMap = await this.buildNumberingMap(source, sourceLists, styles);
    await this.importStyles(styles, numberingMap);
    const ids = this.planIds(content);
    const prepareNote = (xml: string) =>
      this.applyIdPlan(
        retargetStyleRefs(this.applyNumIdMap(stampNumbering(stripCommentMarks(unbind(xml), commentIds), styles), numberingMap), styles),
        ids,
      );
    for (const kind of NOTE_KINDS) {
      if (noteRefIds(scan, kind).size) noteConflicts.push(...(await this.settleNoteFurniture(source, kind)));
    }
    const footnoteMap = await this.buildNoteMap(source, "footnote", scan, prepareNote);
    const endnoteMap = await this.buildNoteMap(source, "endnote", scan, prepareNote);
    const chartMap = await this.buildChartMap(source, srcRels, scan);
    const vmlOleMap = await this.buildVmlOleMap(source, srcRels, scan);

    return blocks.map((b) => {
      let xml = stampNumbering(unbind(b.xml), styles);
      xml = this.applyAttrMap(xml, ["r:embed", "r:link"], mediaMap);
      xml = this.applyOtherRelMap(xml, otherRelMap);
      xml = this.applyHyperlinkMap(xml, hyperlinkMap);
      xml = this.applyNoteRefMap(xml, "footnote", footnoteMap);
      xml = this.applyNoteRefMap(xml, "endnote", endnoteMap);
      xml = this.applyNumIdMap(xml, numberingMap);
      xml = this.applyChartMap(xml, chartMap);
      xml = this.applyVmlOleMap(xml, vmlOleMap);
      xml = this.applyIdPlan(xml, ids);
      xml = retargetStyleRefs(xml, styles);
      xml = this.stripSectionChrome(xml);
      return { kind: b.kind, tag: b.tag, xml };
    });
  }

  /**
   * Drop the source's `headerReference` / `footerReference` from any `sectPr`
   * that rides along with a copied paragraph.
   *
   * Those rIds point at header/footer parts of the SOURCE package, which this
   * merge deliberately does not copy — the target's template owns the running
   * chrome. Left in place they are dangling relationships: a 40-header source
   * produced ~75 references to parts that do not exist in the merged package,
   * which the docx doctor reports as an unrepairable fatal (it cannot invent the
   * missing parts). Removing them lets the section inherit the target's chrome,
   * which is the behaviour the combine flow wants anyway.
   *
   * A section's printer settings go with them: a relationship to a binary this merge
   * does not copy, and of no use to a deposit.
   *
   * Everything else in the copied `sectPr` — page size, margins, break type — is
   * preserved.
   */
  private stripSectionChrome(xml: string): string {
    if (!xml.includes("Reference") && !xml.includes("<w:printerSettings")) return xml;
    return xml
      .replace(/<w:(?:header|footer)Reference\b[^>]*\/>/g, "")
      .replace(/<w:(header|footer)Reference\b[^>]*>[\s\S]*?<\/w:\1Reference>/g, "")
      .replace(/<w:printerSettings\b[^>]*?(?:\/>|>\s*<\/w:printerSettings>)/g, "");
  }

  // ─── Media ─────────────────────────────────────────────────────────────────

  /** Read the source document's relationships as { rId: { type, target, targetMode } }. */
  private async readSourceRels(
    source: MdocxengineType,
  ): Promise<Record<string, { type: string; target: string; targetMode?: string }>> {
    const xml = source.zip.readAsText(SRC_RELS_PATH);
    if (!xml) return {};
    const obj: any = await parseXml(xml);
    const raw = obj?.Relationships?.Relationship;
    const arr = Array.isArray(raw) ? raw : raw ? [raw] : [];
    const map: Record<string, { type: string; target: string; targetMode?: string }> = {};
    for (const r of arr) {
      if (r?.$?.Id && r?.$?.Target != null) {
        map[r.$.Id] = {
          type: String(r.$.Type ?? ""),
          target: String(r.$.Target),
          targetMode: r.$.TargetMode ? String(r.$.TargetMode) : undefined,
        };
      }
    }
    return map;
  }

  /**
   * What copies one source relationship here, once each: an external target stays external,
   * a picture gets its own media part, and any other part comes with its closure, keeping its
   * relationship type. Undefined when the source itself has nothing behind the id.
   */
  private relationshipCarrier(
    source: MdocxengineType,
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
  ): (rId: string) => Promise<string | undefined> {
    const carried = new Map<string, string>();
    const copied = new Map<string, string>();
    const carry = async (rId: string): Promise<string | undefined> => {
      const known = carried.get(rId);
      if (known) return known;
      const rel = srcRels[rId];
      if (!rel) return undefined;
      let id: string;
      if (rel.targetMode === "External") {
        id = await this.rels.genId();
        await this.rels.addRelationship(id, rel.type, rel.target, "External");
      } else {
        const srcPath = resolvePartPath("word/document.xml", rel.target);
        const entry = source.zip.getEntry(srcPath) ?? source.zip.getEntry(safeDecode(srcPath));
        if (!entry) return undefined;
        if (/\/image$/.test(rel.type)) {
          const ext = (/\.([a-zA-Z0-9]+)$/.exec(srcPath)?.[1] ?? "png").toLowerCase();
          id = (await this.media.insertImage(entry.getData(), ext)).relId;
        } else {
          const newPath = await this.copyPartClosure(
            source,
            entry.entryName,
            await this.readContentTypeOverrides(source.zip),
            await this.readContentTypeDefaults(source.zip),
            copied,
          );
          id = await this.rels.genId();
          await this.rels.addRelationship(id, rel.type, relativePartPath("word/document.xml", newPath));
          // SmartArt's data part names its drawing by a relationship of the DOCUMENT.
          if (/\/diagramData$/.test(rel.type)) {
            const data = this.zip.readAsText(newPath) ?? "";
            const named = /(<dsp:dataModelExt\b[^>]*\brelId=")([^"]+)(")/.exec(data);
            const drawing = named && (await carry(named[2]));
            if (named && drawing) this.zip.addFile(newPath, Buffer.from(data.replace(named[0], named[1] + drawing + named[3]), "utf8"));
          }
        }
      }
      carried.set(rId, id);
      return id;
    };
    return carry;
  }

  /** Copy what each `r:embed`/`r:link` in the blocks reaches; return source-rId → new-rId. */
  private async buildMediaMap(carry: (rId: string) => Promise<string | undefined>, blocksXml: string): Promise<Record<string, string>> {
    const map: Record<string, string> = {};
    for (const m of new Set([...blocksXml.matchAll(/r:(?:embed|link)="([^"]+)"/g)].map((x) => x[1]))) {
      const id = await carry(m);
      if (id) map[m] = id;
    }
    return map;
  }

  /** Copy what every other relationship reference in the blocks reaches: a link on a picture,
   *  SmartArt's four parts, an embedded chunk, ink. The elements in REL_OWN_PATH are handled by
   *  their own paths or removed. */
  private async buildOtherRelMap(carry: (rId: string) => Promise<string | undefined>, blocksXml: string): Promise<Record<string, string>> {
    const map: Record<string, string> = {};
    for (const tag of blocksXml.matchAll(ANY_OPEN_TAG)) {
      if (REL_OWN_PATH.has(tag[1])) continue;
      for (const m of tag[0].matchAll(OTHER_REL_ATTR)) {
        if (map[m[2]]) continue;
        const id = await carry(m[2]);
        if (id) map[m[2]] = id;
      }
    }
    return map;
  }

  private applyOtherRelMap(xml: string, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(ANY_OPEN_TAG, (tag, name) =>
      REL_OWN_PATH.has(name) ? tag : tag.replace(OTHER_REL_ATTR, (full, attr, val) => (idMap[val] ? ` ${attr}="${idMap[val]}"` : full)),
    );
  }

  /**
   * Copy external hyperlink relationships referenced by `<w:hyperlink r:id>` into
   * the target (preserving TargetMode="External"); return source-rId → new-rId.
   */
  private async buildHyperlinkMap(
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refRIds = new Set(
      [...blocksXml.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]),
    );
    if (refRIds.size === 0) return {};

    const map: Record<string, string> = {};
    for (const rId of refRIds) {
      const rel = srcRels[rId];
      if (!rel || !rel.type.endsWith("/hyperlink")) continue;
      const newId = await this.rels.genId();
      await this.rels.addRelationship(newId, rel.type, rel.target, rel.targetMode ?? "External");
      map[rId] = newId;
    }
    return map;
  }

  // ─── VML pictures + OLE objects ─────────────────────────────────────────────
  //
  // `<v:imagedata r:id>` (a VML picture, or the preview of an OLE object) and
  // `<o:OLEObject r:id>` (the embedded object) address parts by relationship id,
  // just like `r:embed`. Unlike a chart, losing one silently is not acceptable —
  // a missing equation object is lost thesis content — so an unresolvable
  // reference THROWS rather than being dropped.

  /** Attributes on `<v:imagedata>` / `<v:fill>` (picture fill) that name a relationship. */
  private static readonly VML_REL_ATTRS = ["r:id", "r:href", "r:pict", "o:relid"];

  /** Copy each VML image and OLE object referenced by the blocks; return
   *  source-rId → new-rId. */
  private async buildVmlOleMap(
    source: MdocxengineType,
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refs: { tag: string; attr: string; rId: string; ole: boolean }[] = [];
    for (const m of blocksXml.matchAll(/<(v:imagedata|v:fill|o:OLEObject)\b[^>]*>/g)) {
      const ole = m[1] === "o:OLEObject";
      for (const attr of ole ? ["r:id"] : MergeManager.VML_REL_ATTRS) {
        const v = new RegExp(`\\b${attr}="([^"]+)"`).exec(m[0])?.[1];
        if (v) refs.push({ tag: m[1], attr, rId: v, ole });
      }
    }
    if (refs.length === 0) return {};

    const srcOverrides = await this.readContentTypeOverrides(source.zip);
    const srcDefaults = await this.readContentTypeDefaults(source.zip);
    const copied = new Map<string, string>();
    const map: Record<string, string> = {};

    for (const { tag, attr, rId, ole } of refs) {
      if (map[rId]) continue;
      const rel = srcRels[rId];
      if (!rel) {
        throw new Error(`merge: <${tag} ${attr}="${rId}"> has no relationship in the source document`);
      }
      if (rel.targetMode === "External") {
        const newId = await this.rels.genId();
        await this.rels.addRelationship(newId, rel.type, rel.target, "External");
        map[rId] = newId;
        continue;
      }
      const srcPath = resolvePartPath("word/document.xml", rel.target);
      const entry = source.zip.getEntry(srcPath);
      if (!entry) {
        throw new Error(
          `merge: <${tag} ${attr}="${rId}"> points at ${srcPath}, which is missing from the source package`,
        );
      }
      if (ole) {
        const newPath = await this.copyPartClosure(source, srcPath, srcOverrides, srcDefaults, copied);
        const newId = await this.rels.genId();
        await this.rels.addRelationship(newId, rel.type, newPath.replace(/^word\//, ""));
        map[rId] = newId;
      } else {
        const ext = (/\.([a-zA-Z0-9]+)$/.exec(srcPath)?.[1] ?? "png").toLowerCase();
        map[rId] = (await this.media.insertImage(entry.getData(), ext)).relId;
      }
    }
    return map;
  }

  /** Remap the relationship attributes inside `<v:imagedata>` / `<o:OLEObject>`
   *  open tags only (never another element's r:id). */
  private applyVmlOleMap(xml: string, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(/<(?:v:imagedata|v:fill|o:OLEObject)\b[^>]*>/g, (tag) =>
      tag.replace(/\b((?:r:id|r:href|r:pict|o:relid)=")([^"]+)(")/g, (full, a, val, c) =>
        idMap[val] ? `${a}${idMap[val]}${c}` : full,
      ),
    );
  }

  // ─── Styles ─────────────────────────────────────────────────────────────────
  //
  // A merge copies blocks, not style DEFINITIONS, so a `pStyle`/`rStyle`/`tblStyle`
  // the target does not define dangles (one real thesis: 1052 uses of 8 ids) and
  // Word silently renders it as Normal. mergeStyles.ts decides what each one becomes;
  // here is only the writing. Existing definitions are never touched.

  private async importStyles(plan: StylePlan, numberingMap: Record<string, string>): Promise<void> {
    if (plan.imports.length === 0) return;
    const tgt = await this.stylesPartForImport();
    const close = tgt.lastIndexOf("</w:styles>");
    const add = plan.imports.map((x) => this.applyNumIdMap(x.xml, numberingMap)).join("");
    this.zip.addFile("word/styles.xml", Buffer.from(tgt.slice(0, close) + add + tgt.slice(close), "utf8"));
  }

  /** This document's styles part. One with none gets an empty one, without defaults, so that
   *  its own text keeps rendering as it did. */
  private async stylesPartForImport(): Promise<string> {
    const existing = this.zip.readAsText("word/styles.xml");
    if (existing) return existing;
    const fresh = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`;
    if (this.zip.getEntry("word/styles.xml")) return fresh;
    await this.contentTypes.addOverride("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
    await this.rels.addRelationship(await this.rels.genId(), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
    return fresh;
  }

  // ─── Unique ids: drawing ids + bookmarks ─────────────────────────────────────
  //
  // `wp:docPr/@id` and `w:bookmarkStart/@w:id` must be unique across the merged
  // document, and a bookmark NAME must be unique too. Two split parts that both
  // start at id 1 / `_Toc123` otherwise give Word a repair prompt (or make a
  // cross-reference jump to the wrong part).

  private planIds(blocksXml: string): IdPlan {
    // The notes hold bookmarks too, and every story shares one set of names and ids.
    const targetXml = ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml"].map((p) => this.zip.readAsText(p) ?? "").join("\n");
    const maxOf = (re: RegExp) =>
      [...targetXml.matchAll(re)].reduce((m, x) => Math.max(m, parseInt(x[1], 10) || 0), 0);

    let nextDoc = maxOf(/<wp:docPr\b[^>]*\bid="(\d+)"/g);
    // docPr ids are unique per element, not per value: every source drawing gets
    // its own new id, handed out in document order by applyIdPlan.
    let nextBm = maxOf(/<w:bookmarkStart\b[^>]*\bw:id="(\d+)"/g);
    const bmIds = new Map<string, string>();
    // Word's bookmark names are not case-sensitive: "Resultat" and "RESULTAT" are one bookmark.
    const taken = new Set(
      [...targetXml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)].map((m) => bookmarkKey(m[1])),
    );
    const srcNames = new Map<string, string>();
    for (const m of blocksXml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)) {
      if (!srcNames.has(bookmarkKey(m[1]))) srcNames.set(bookmarkKey(m[1]), m[1]);
    }
    const renames = new Map<string, string>();
    for (const m of blocksXml.matchAll(/<w:bookmark(?:Start|End)\b[^>]*\bw:id="([^"]+)"/g)) {
      if (!bmIds.has(m[1])) bmIds.set(m[1], String(++nextBm));
    }
    for (const [key, name] of srcNames) {
      if (!taken.has(key)) continue;
      let k = 1;
      while (taken.has(bookmarkKey(renamedBookmark(name, k))) || srcNames.has(bookmarkKey(renamedBookmark(name, k)))) k++;
      renames.set(key, renamedBookmark(name, k));
      taken.add(bookmarkKey(renamedBookmark(name, k)));
    }
    return { nextDoc: () => ++nextDoc, bmIds, renames };
  }

  private applyIdPlan(xml: string, plan: IdPlan): string {
    let out = xml;
    if (out.includes("<wp:docPr")) {
      out = out.replace(/(<wp:docPr\b[^>]*\bid=")(\d+)(")/g, (_f, a, _v, c) => `${a}${plan.nextDoc()}${c}`);
    }
    if (out.includes("<w:bookmark")) {
      out = out.replace(/<w:bookmark(Start|End)\b[^>]*>/g, (tag) => {
        let t = tag.replace(/(\bw:id=")([^"]+)(")/, (f, a, v, c) => (plan.bmIds.has(v) ? `${a}${plan.bmIds.get(v)}${c}` : f));
        if (tag.startsWith("<w:bookmarkStart")) {
          t = t.replace(/(\bw:name=")([^"]*)(")/, (f, a, v, c) => {
            const renamed = plan.renames.get(bookmarkKey(v));
            return renamed ? `${a}${renamed}${c}` : f;
          });
        }
        return t;
      });
    }
    if (plan.renames.size) out = this.renameBookmarkRefs(out, plan.renames);
    return out;
  }

  /** Point the fields and every `w:anchor` naming a renamed bookmark at its new name. */
  private renameBookmarkRefs(xml: string, renames: Map<string, string>): string {
    const fix = (code: string) => renameBookmarksInField(code, renames);
    return rewriteFieldCodes(xml, fix)
      .replace(/(\bw:instr=")([^"]*)(")/g, (_m, a, t, c) => a + fix(t) + c)
      .replace(/(<w:hyperlink\b[^>]*\bw:anchor=")([^"]*)(")/g, (m, a, v, c) => {
        const renamed = renames.get(bookmarkKey(v));
        return renamed ? `${a}${renamed}${c}` : m;
      });
  }

  // ─── Charts (part-closure copy) ─────────────────────────────────────────────
  //
  // A chart is NOT inline in the paragraph: `<c:chart r:id="rIdN"/>` points at a
  // whole part (`word/charts/chart1.xml`), which in turn points at its own colour
  // and style parts and an embedded .xlsx workbook. Copying the paragraph without
  // copying that closure leaves the reference aimed at whatever rId happens to
  // occupy that slot in the TARGET — in practice an image — and Word refuses to
  // open the document ("Word experienced an error trying to open the file").
  // That is not hypothetical: it shipped a student's combined thesis that could
  // never be opened, and neither schema validation nor a dangling-target check
  // sees it, because the rId resolves fine — to the wrong KIND of part.

  /** Copy each chart referenced by the blocks, with its whole part closure;
   *  return source-rId → new-rId. */
  private async buildChartMap(
    source: MdocxengineType,
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refRIds = new Set(
      [...blocksXml.matchAll(/<c:chart\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]),
    );
    if (refRIds.size === 0) return {};

    const srcOverrides = await this.readContentTypeOverrides(source.zip);
    const srcDefaults = await this.readContentTypeDefaults(source.zip);
    const copied = new Map<string, string>();
    const map: Record<string, string> = {};

    for (const rId of refRIds) {
      const rel = srcRels[rId];
      // Only follow a rel that really is a chart; anything else stays unmapped and
      // is dropped below rather than left pointing at a stranger's part.
      if (!rel || !/\/chart$/.test(rel.type) || rel.targetMode === "External") continue;
      const srcPath = resolvePartPath("word/document.xml", rel.target);
      if (!source.zip.getEntry(srcPath)) continue;

      const newPath = await this.copyPartClosure(source, srcPath, srcOverrides, srcDefaults, copied);
      const newId = await this.rels.genId();
      // document.xml.rels targets are relative to word/.
      await this.rels.addRelationship(newId, rel.type, newPath.replace(/^word\//, ""));
      map[rId] = newId;
    }
    return map;
  }

  /**
   * Copy `srcPath` and every part it transitively references into this package
   * under a free name, carrying content types across. Returns the new part path.
   *
   * `.rels` targets are rewritten to the copied parts' new names, so two merged
   * documents that both ship `charts/chart1.xml` cannot collide.
   */
  private async copyPartClosure(
    source: MdocxengineType,
    srcPath: string,
    srcOverrides: Record<string, string>,
    srcDefaults: Record<string, string>,
    copied: Map<string, string>,
  ): Promise<string> {
    const already = copied.get(srcPath);
    if (already) return already;

    const entry = source.zip.getEntry(srcPath);
    if (!entry) return srcPath;
    const bytes = entry.getData();

    const newPath = this.freePartPath(srcPath);
    // Record BEFORE recursing so a cyclic reference terminates.
    copied.set(srcPath, newPath);

    // Content type: an explicit Override travels with the part; otherwise the
    // extension's Default must exist in this package.
    const ct = srcOverrides["/" + srcPath];
    if (ct) {
      await this.contentTypes.addOverride("/" + newPath, ct);
    } else {
      const ext = (/\.([a-zA-Z0-9]+)$/.exec(srcPath)?.[1] ?? "").toLowerCase();
      const def = srcDefaults[ext] ?? DEFAULT_CONTENT_TYPES[ext];
      if (ext && def) await this.contentTypes.addDefault(ext, def);
    }

    // Follow this part's own relationships.
    const srcRelsPath = partRelsPath(srcPath);
    const relsXml = source.zip.readAsText(srcRelsPath);
    if (relsXml) {
      const obj: any = await parseXml(relsXml);
      const raw = obj?.Relationships?.Relationship;
      const arr = Array.isArray(raw) ? raw : raw ? [raw] : [];
      const out: any[] = [];
      for (const r of arr) {
        const id = r?.$?.Id;
        const target = r?.$?.Target;
        if (!id || target == null) continue;
        if (r?.$?.TargetMode === "External") {
          out.push({ $: { ...r.$ } });
          continue;
        }
        const depSrc = resolvePartPath(srcPath, String(target));
        if (!source.zip.getEntry(depSrc)) continue; // dangling in the source — drop it
        const depNew = await this.copyPartClosure(source, depSrc, srcOverrides, srcDefaults, copied);
        out.push({ $: { ...r.$, Target: relativePartPath(newPath, depNew) } });
      }
      if (out.length) {
        const relsMgr = new RelManager(this.zip, partRelsPath(newPath));
        for (const r of out) {
          await relsMgr.addRelationship(r.$.Id, r.$.Type, r.$.Target, r.$.TargetMode);
        }
      }
    }

    this.zip.addFile(newPath, bytes);
    return newPath;
  }

  /** `word/charts/chart1.xml` → the same name, or `chart2.xml`, … until free. */
  private freePartPath(srcPath: string): string {
    const dir = srcPath.slice(0, srcPath.lastIndexOf("/") + 1);
    const file = srcPath.slice(dir.length);
    const m = /^(.*?)(\d*)(\.[^.]+)$/.exec(file);
    const stem = m?.[1] ?? file;
    const ext = m?.[3] ?? "";
    for (let i = 1; i < 10000; i++) {
      const candidate = `${dir}${stem}${i}${ext}`;
      if (!this.zip.getEntry(candidate)) return candidate;
    }
    return srcPath;
  }

  /** Remap r:id ONLY inside <c:chart .../> (never other r:id). Any chart whose
   *  part could not be carried has its whole drawing removed — a reference to a
   *  part that is not there is exactly what makes Word reject the file. */
  private applyChartMap(xml: string, idMap: Record<string, string>): string {
    if (!xml.includes("<c:chart")) return xml;
    let out = xml.replace(/<c:chart\b[^>]*>/g, (tag) =>
      tag.replace(/\br:id="([^"]+)"/, (full, val) => (idMap[val] ? `r:id="${idMap[val]}"` : full)),
    );
    // Drop drawings still pointing at an un-copied chart.
    out = out.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, (drawing) => {
      const refs = [...drawing.matchAll(/<c:chart\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]);
      const orphan = refs.some((r) => !Object.values(idMap).includes(r));
      return refs.length && orphan ? "" : drawing;
    });
    return out;
  }

  /** `/word/charts/chart1.xml` → content type, from a package's [Content_Types].xml. */
  private async readContentTypeOverrides(zip: AdmZip): Promise<Record<string, string>> {
    const xml = zip.readAsText("[Content_Types].xml");
    if (!xml) return {};
    const obj: any = await parseXml(xml);
    const raw = obj?.Types?.Override;
    const arr = Array.isArray(raw) ? raw : raw ? [raw] : [];
    const out: Record<string, string> = {};
    for (const o of arr) if (o?.$?.PartName) out[String(o.$.PartName)] = String(o.$.ContentType ?? "");
    return out;
  }

  /** extension (lower-case, no dot) → content type. */
  private async readContentTypeDefaults(zip: AdmZip): Promise<Record<string, string>> {
    const xml = zip.readAsText("[Content_Types].xml");
    if (!xml) return {};
    const obj: any = await parseXml(xml);
    const raw = obj?.Types?.Default;
    const arr = Array.isArray(raw) ? raw : raw ? [raw] : [];
    const out: Record<string, string> = {};
    for (const d of arr) {
      if (d?.$?.Extension) out[String(d.$.Extension).toLowerCase()] = String(d.$.ContentType ?? "");
    }
    return out;
  }

  // ─── Footnotes + endnotes (verbatim element copy, rich-content fidelity) ─────

  /**
   * The source notes `blocksXml` refers to, as XML. A reference whose note does not
   * exist THROWS: Word refuses such a file, and left in place the id could land on
   * a different note this document already has, which nothing downstream would see.
   */
  private neededNotesXml(source: MdocxengineType, kind: NoteKind, blocksXml: string): string {
    const wanted = noteRefIds(blocksXml, kind);
    if (wanted.size === 0) return "";
    const partXml = source.zip.readAsText(`word/${kind}s.xml`) ?? "";
    const found = new Map<string, string>();
    for (const m of partXml.matchAll(new RegExp(`<w:${kind}\\b[^>]*\\bw:id="([^"]+)"[^>]*?(?:/>|>[\\s\\S]*?</w:${kind}>)`, "g"))) {
      if (wanted.has(m[1]) && !SEPARATOR_NOTE_TYPE.test(m[0].slice(0, m[0].indexOf(">")))) found.set(m[1], m[0]);
    }
    const missing = [...wanted].filter((id) => !found.has(id));
    if (missing.length) {
      throw new Error(`merge: ${kind} reference(s) ${missing.join(", ")} have no ${kind} in the source document`);
    }
    return [...found.values()].join("\n");
  }

  /** Copy each referenced source note verbatim; return source-id → new-id. `prepare`
   *  runs on each note before its relationships are remapped. */
  private async buildNoteMap(
    source: MdocxengineType,
    kind: NoteKind,
    blocksXml: string,
    prepare: (noteXml: string) => string,
  ): Promise<Record<string, string>> {
    const refIds = noteRefIds(blocksXml, kind);
    if (refIds.size === 0) return {};
    const notesPart = `word/${kind}s.xml`;
    const relsPath = `word/_rels/${kind}s.xml.rels`;
    const sourceNotesXml = source.zip.readAsText(notesPart);
    // A note's `r:embed` resolves against its own part's rels, not the document
    // rels — so what it points at is copied into THAT part's rels.
    const srcRelsXml = source.zip.readAsText(relsPath);
    const noteRels: Record<string, { type: string; target: string; external: boolean }> = {};
    // Parsed, not regex-scanned: Targets are XML-escaped on disk (`&amp;`), and
    // addRelationship escapes again on write — raw text would double-escape a URL.
    if (srcRelsXml) {
      const obj: any = await parseXml(srcRelsXml);
      const raw = obj?.Relationships?.Relationship;
      for (const r of Array.isArray(raw) ? raw : raw ? [raw] : []) {
        if (r?.$?.Id && r?.$?.Target != null) {
          noteRels[r.$.Id] = {
            type: String(r.$.Type ?? ""),
            target: String(r.$.Target),
            external: r.$.TargetMode === "External",
          };
        }
      }
    }
    const relMgr = new RelManager(this.zip, relsPath);
    const copied = new Map<string, string>();
    const rewrite = async (raw: string): Promise<string> => {
      const xml = prepare(raw);
      const ids = new Map<string, string>();
      for (const m of xml.matchAll(/\b(?:r:[a-zA-Z]+|o:relid)="([^"]+)"/g)) {
        const rel = noteRels[m[1]];
        if (!rel || ids.has(m[1])) continue;
        if (rel.external) {
          // Hyperlinks (and linked images) live in the note part's OWN rels.
          const newId = await relMgr.genId();
          await relMgr.addRelationship(newId, rel.type, rel.target, "External");
          ids.set(m[1], newId);
          continue;
        }
        const srcPath = resolvePartPath(notesPart, rel.target);
        const entry = source.zip.getEntry(srcPath);
        if (!entry) continue;
        if (/\/image$/.test(rel.type)) {
          const ext = (/\.([a-zA-Z0-9]+)$/.exec(rel.target)?.[1] ?? "png").toLowerCase();
          ids.set(m[1], await this.media.addImageToPartRels(notesPart, entry.getData(), ext));
          continue;
        }
        // Not a picture (an OLE object, a chart): the part keeps its own
        // relationship type and brings its closure along.
        const newPath = await this.copyPartClosure(
          source,
          srcPath,
          await this.readContentTypeOverrides(source.zip),
          await this.readContentTypeDefaults(source.zip),
          copied,
        );
        const newId = await relMgr.genId();
        await relMgr.addRelationship(newId, rel.type, newPath.replace(/^word\//, ""));
        ids.set(m[1], newId);
      }
      return this.applyAttrMap(xml, ["r:embed", "r:link", "r:id", "r:pict", "r:href", "r:dm", "r:lo", "r:qs", "r:cs", "o:relid"], Object.fromEntries(ids));
    };
    const map = await (kind === "footnote"
      ? this.footnotes.copyFootnotesVerbatim(sourceNotesXml, refIds, rewrite)
      : this.endnotes.copyEndnotesVerbatim(sourceNotesXml, refIds, rewrite));
    // The source had them all (neededNotesXml), so a shortfall means this document's
    // notes part cannot take them, and the references would be left dangling.
    if ([...refIds].some((id) => !map[id])) {
      throw new Error(`merge: ${notesPart} in this document cannot take the copied ${kind}s`);
    }
    return map;
  }

  // ─── A note's furniture: its properties and separators ───────────────────────
  //
  // How notes are numbered and placed lives in word/settings.xml and in each section; the rule
  // above them lives in the notes part. A merge's text flows into one final section, so the first
  // part with notes of a kind decides for that kind. Whatever a later part wanted is reported.

  private async settleNoteFurniture(source: MdocxengineType, kind: NoteKind): Promise<NoteConflict[]> {
    const targetDoc = this.zip.readAsText("word/document.xml") ?? "";
    const sourceSettings = source.zip.readAsText("word/settings.xml");
    const sourceSect = finalSectPr(source.zip.readAsText("word/document.xml") ?? "");
    const targetSettings = this.zip.readAsText("word/settings.xml");
    if (noteRefIds(targetDoc, kind).size) {
      const mine = effectiveNoteProps(kind, targetSettings, finalSectPr(targetDoc));
      const theirs = effectiveNoteProps(kind, sourceSettings, sourceSect);
      return Object.keys(mine)
        .filter((property) => mine[property] !== theirs[property])
        .map((property) => ({ kind, property, kept: mine[property], dropped: theirs[property] }));
    }
    // Nothing here is numbered this way yet, so nothing can change under the target's feet.
    const adopted = await this.adoptSeparators(source, kind);
    const inner = noteProps(notePr(sourceSettings, kind)).xml + noteSeparatorRefs(notePr(adopted ? sourceSettings : targetSettings, kind), kind);
    await this.settings.setSettingElement(`${kind}Pr`, inner ? `<w:${kind}Pr>${inner}</w:${kind}Pr>` : null);
    this.setFinalSectionNotePr(kind, notePr(sourceSect, kind));
    return [];
  }

  /** The source's separators in place of this document's unused ones. False: it has none, or an
   *  unreferenced note here already holds one of their ids. */
  private async adoptSeparators(source: MdocxengineType, kind: NoteKind): Promise<boolean> {
    const theirs = source.zip.readAsText(`word/${kind}s.xml`)?.match(separatorNotes(kind));
    if (!theirs?.length) return false;
    await (kind === "footnote" ? this.footnotes : this.endnotes).ensureRegistered();
    const part = `word/${kind}s.xml`;
    const ours = this.zip.readAsText(part) ?? "";
    const open = new RegExp(`<w:${kind}s\\b[^>]*?(?<!/)>`).exec(ours);
    if (!open) return false;
    const rest = ours.replace(separatorNotes(kind), "");
    const idOf = (xml: string) => /\bw:id="([^"]*)"/.exec(xml)?.[1];
    const held = new Set([...rest.matchAll(new RegExp(`<w:${kind}\\b[^>]*`, "g"))].map((m) => idOf(m[0])));
    if (theirs.some((note) => held.has(idOf(note)))) return false;
    const at = rest.indexOf(open[0]) + open[0].length;
    this.zip.addFile(part, Buffer.from(rest.slice(0, at) + theirs.join("") + rest.slice(at), "utf8"));
    return true;
  }

  /** Put a note property element in this document's final section, at its schema position. */
  private setFinalSectionNotePr(kind: NoteKind, element: string | null): void {
    const xml = this.zip.readAsText("word/document.xml");
    if (!xml) return;
    const split = splitDocument(xml);
    const block = split.blocks.filter((b) => b.kind === "sectPr").pop();
    if (!block) {
      if (!element) return;
      split.blocks.push({ kind: "sectPr", tag: "w:sectPr", xml: `<w:sectPr>${element}</w:sectPr>` });
    } else {
      const next = withNotePr(block.xml, kind, element);
      if (next === block.xml) return;
      block.xml = next;
    }
    this.zip.addFile("word/document.xml", Buffer.from(assembleDocument(split), "utf8"));
  }

  // ─── Numbering (Phase 2) ─────────────────────────────────────────────────────

  /**
   * Copy the source's lists something copied uses (abstractNum + num) into the target
   * with fresh, collision-free ids; return source-numId → new-numId. Without this, two
   * parts that both use `numId="1"` would share one list and renumber wrongly.
   */
  private async buildNumberingMap(
    source: MdocxengineType,
    srcDefs: RawNumbering,
    styles: StylePlan,
  ): Promise<Record<string, string>> {
    if (styles.numIds.size === 0) return {};

    const { absMax, numMax, picMax } = await this.numbering.maxIds();
    let nextAbs = absMax + 1;
    // Above what this document POINTS at too: a numId with no definition prints a default "1."
    // of its own (measured), and a copied list taking that id would change what those paragraphs say.
    let nextNum = Math.max(numMax, this.highestNumIdInUse()) + 1;
    let nextPic = picMax + 1;

    const numIdMap: Record<string, string> = {};
    const absIdMap: Record<string, string> = {};
    const picIdMap: Record<string, string> = {};
    const addAbs: any[] = [];
    const addNum: any[] = [];
    const addPic: any[] = [];

    /** A level's picture bullet, copied with its image: a `w:lvlPicBulletId` naming nothing
     *  makes Word refuse the whole file, so one the source itself lacks is dropped. */
    const carryPicBullet = async (level: any) => {
      const old = level?.["w:lvlPicBulletId"]?.$?.["w:val"];
      if (old == null) return;
      if (picIdMap[old] == null) {
        const def = srcDefs.picBullets.find((p) => p?.$?.["w:numPicBulletId"] === old);
        if (!def) {
          delete level["w:lvlPicBulletId"];
          return;
        }
        const clone = JSON.parse(JSON.stringify(def));
        picIdMap[old] = String(nextPic++);
        clone.$ = { ...clone.$, "w:numPicBulletId": picIdMap[old] };
        await this.rehomeNumberingRels(source, clone);
        addPic.push(clone);
      }
      level["w:lvlPicBulletId"].$["w:val"] = picIdMap[old];
    };

    for (const num of srcDefs.nums) {
      const oldNumId = num?.$?.["w:numId"];
      if (oldNumId == null || !styles.numIds.has(String(oldNumId))) continue;

      const oldAbs = num?.["w:abstractNumId"]?.$?.["w:val"];
      if (oldAbs == null || !srcDefs.abstractNums.some((a) => a?.$?.["w:abstractNumId"] === oldAbs)) continue;
      // Copy the referenced abstractNum once, with a fresh id.
      if (oldAbs != null && absIdMap[oldAbs] == null) {
        const absDef = srcDefs.abstractNums.find((a) => a?.$?.["w:abstractNumId"] === oldAbs);
        if (absDef) {
          const newAbs = String(nextAbs++);
          const clone = JSON.parse(JSON.stringify(absDef));
          clone.$ = { ...clone.$, "w:abstractNumId": newAbs };
          // Drop nsid so Word doesn't fold this into the source list's identity.
          if (clone["w:nsid"]) delete clone["w:nsid"];
          retargetListRefs(clone, styles);
          for (const level of asArray(clone["w:lvl"])) await carryPicBullet(level);
          addAbs.push(clone);
          absIdMap[oldAbs] = newAbs;
        }
      }

      const newNumId = String(nextNum++);
      const numClone = JSON.parse(JSON.stringify(num));
      numClone.$ = { ...numClone.$, "w:numId": newNumId };
      if (numClone["w:abstractNumId"]?.$ && oldAbs != null && absIdMap[oldAbs] != null) {
        numClone["w:abstractNumId"].$ = { ...numClone["w:abstractNumId"].$, "w:val": absIdMap[oldAbs] };
      }
      retargetListRefs(numClone, styles);
      for (const override of asArray<any>(numClone["w:lvlOverride"])) {
        if (override?.["w:lvl"]) await carryPicBullet(override["w:lvl"]);
      }
      addNum.push(numClone);
      numIdMap[String(oldNumId)] = newNumId;
    }

    await this.numbering.appendRawDefinitions(addAbs, addNum, addPic);
    // A numId with no list behind it prints a default "1." of its own, and has to go on doing so:
    // left as it is, it would fall into whatever list this document has of that id. Each stays
    // missing under an id of its own, since Word counts each on its own.
    for (const id of styles.numIds) if (numIdMap[id] == null && id !== "0") numIdMap[id] = String(nextNum++);
    return numIdMap;
  }

  private highestNumIdInUse(): number {
    let max = 0;
    for (const { entryName } of this.zip.getEntries()) {
      if (!/^word\/[^/]+\.xml$/.test(entryName) || entryName === "word/numbering.xml") continue;
      for (const m of (this.zip.readAsText(entryName) ?? "").matchAll(/<w:numId\b[^>]*\bw:val="(\d+)"/g)) max = Math.max(max, Number(m[1]));
    }
    return max;
  }

  /** A picture bullet's image is reached through the numbering part's own relationships. */
  private async rehomeNumberingRels(source: MdocxengineType, node: any): Promise<void> {
    const relsXml = source.zip.readAsText(NUMBERING_RELS_PATH);
    if (!relsXml) return;
    const parsed: any = await parseXml(relsXml);
    const rels = new Map<string, any>(asArray<any>(parsed?.Relationships?.Relationship).map((r) => [r?.$?.Id, r?.$]));
    const relMgr = new RelManager(this.zip, NUMBERING_RELS_PATH);
    const moved = new Map<string, string>();
    const walk = async (n: any): Promise<void> => {
      if (Array.isArray(n)) {
        for (const x of n) await walk(x);
        return;
      }
      if (!n || typeof n !== "object") return;
      for (const [attr, value] of Object.entries<any>(n.$ ?? {})) {
        if (!/^(?:r:[a-zA-Z]+|o:relid)$/.test(attr)) continue;
        const rel = rels.get(value);
        if (!rel) continue;
        if (!moved.has(value)) {
          if (rel.TargetMode === "External") {
            const id = await relMgr.genId();
            await relMgr.addRelationship(id, rel.Type, rel.Target, "External");
            moved.set(value, id);
          } else {
            const bytes = source.zip.getEntry(resolvePartPath("word/numbering.xml", rel.Target))?.getData();
            if (!bytes) continue;
            const ext = (/\.([a-zA-Z0-9]+)$/.exec(rel.Target)?.[1] ?? "png").toLowerCase();
            moved.set(value, await this.media.addImageToPartRels("word/numbering.xml", bytes, ext));
          }
        }
        n.$[attr] = moved.get(value);
      }
      for (const [key, child] of Object.entries(n)) if (key !== "$") await walk(child);
    };
    await walk(node);
  }

  /** Remap w:val inside <w:numId .../> elements only. */
  private applyNumIdMap(xml: string, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(/<w:numId\b[^>]*\/?>/g, (tag) =>
      tag.replace(/\bw:val="([^"]+)"/, (full, val) => (idMap[val] ? `w:val="${idMap[val]}"` : full)),
    );
  }

  // ─── Attribute rewriters (scoped to the passed block xml) ───────────────────

  /** Replace attr="old" → attr="new" for each given attribute name. */
  private applyAttrMap(
    xml: string,
    attrs: string[],
    idMap: Record<string, string>,
  ): string {
    if (Object.keys(idMap).length === 0) return xml;
    let out = xml;
    for (const attr of attrs) {
      const re = new RegExp(`(${attr.replace(":", "\\:")}=")([^"]+)(")`, "g");
      out = out.replace(re, (full, p1, val, p3) => (idMap[val] ? `${p1}${idMap[val]}${p3}` : full));
    }
    return out;
  }

  /** Remap r:id ONLY inside <w:hyperlink ...> open tags (never other r:id). */
  private applyHyperlinkMap(xml: string, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(/<w:hyperlink\b[^>]*>/g, (tag) =>
      tag.replace(/\br:id="([^"]+)"/, (full, val) => (idMap[val] ? `r:id="${idMap[val]}"` : full)),
    );
  }

  /** Remap w:id ONLY inside <w:footnoteReference/> or <w:endnoteReference/> (never other w:id). */
  private applyNoteRefMap(xml: string, kind: NoteKind, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(new RegExp(`<w:${kind}Reference\\b[^>]*>`, "g"), (tag) =>
      tag.replace(/\bw:id="([^"]+)"/, (full, val) => (idMap[val] ? `w:id="${idMap[val]}"` : full)),
    );
  }
}
