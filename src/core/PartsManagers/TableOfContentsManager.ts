import { StylesManager } from "@/core/PartsManagers/StylesManager";
import DocumentManager from "@/core/PartsManagers/DocumentManager";
import {
  escapeXmlText,
  decodeXmlText,
  stripInvalidXmlChars,
  paragraphText,
  paragraphHeadingLevel,
  type BodyBlock,
} from "@/core/files/body/OrderedBody";
import AdmZip from "adm-zip";

const DOCUMENT_PATH = "word/document.xml";

/** Right tab stop for the dot leader before the page number (twips). */
const PAGE_TAB_POS = 8748;

export interface TocOptions {
  /** Deepest heading level collected (1–9). Default 3. */
  headingDepth?: number;
  /** Heading shown above the table (pass "" for none). Default "Table of Contents". */
  title?: string;
  /** Show page numbers with a dot leader. Default true. */
  includePageNumbers?: boolean;
  /** Make the entries clickable links to their headings. Default true. */
  useHyperlinks?: boolean;
  /** Write the table right-to-left (Arabic thesis). Default false. */
  rtl?: boolean;
  /** Delete any table of contents already in the document first. Default true. */
  replaceExisting?: boolean;
}

/** What {@link TableOfContentsManager.insertTOC} actually wrote. */
export interface TocResult {
  /** Block index the table now starts at. */
  atIndex: number;
  /** One per heading collected. */
  entries: number;
  /** Deepest heading level collected. */
  headingDepth: number;
  /** How many blocks a replaced table occupied (0 when none was replaced). */
  replaced: number;
}

/** What {@link TableOfContentsManager.removeTOC} deleted. */
export interface TocRemoval {
  /** Blocks deleted across every table of contents found. */
  removed: number;
  /** Block index the FIRST removed table started at, or -1 if none. */
  at: number;
}

/**
 * A table of contents the student TYPED BY HAND — ordinary paragraphs reading
 * "المقدمة .......... 5", not a `TOC` field. It is the usual state of an
 * imported thesis, and no field-based operation can see it.
 */
export interface TypedTocSpan {
  /** First block of the typed table (its title, when it has one). */
  startIndex: number;
  /** Last block, INCLUSIVE. */
  endIndex: number;
  /** The heading above the entries ("الفهرس", "Table des matières"), or null. */
  title: string | null;
  /** How many entry lines it holds. */
  entries: number;
  /** The first few entry lines, for showing the student what was found. */
  sample: string[];
}

/** Titles a hand-typed table of contents is given, across the app's languages. */
const TYPED_TOC_TITLE =
  /^(?:الفهرس|فهرس\s+المحتويات|قائمة\s+المحتويات|المحتويات|فهرس|table\s+des\s+mati[eè]res|sommaire|table\s+of\s+contents|contents|index)\s*:?\s*$/i;

/** Trailing page number: "… 12", "… ١٢", "… IV". */
const TRAILING_PAGE = /(?:\d{1,4}|[٠-٩]{1,4}|[ivxlcdm]{1,7})\s*$/i;

/** A dot/underscore leader, as typed ("....." or "……"). */
const TEXT_LEADER = /[.…_·]{3,}/;

/**
 * Escape for ELEMENT TEXT (`<w:t>`, `<w:instrText>`), where `"` is a legal
 * literal — Word writes `TOC \o "1-3"` with bare quotes, and `escapeXmlText`'s
 * attribute-grade escaping would turn them into `&quot;`.
 */
function escapeXmlContent(text: string): string {
  return stripInvalidXmlChars(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/** Word's "TOC Heading" style — a heading in LOOK that stays OUT of the table. */
const TOC_HEADING_STYLE = (basedOn: string) => ({
  $: { "w:type": "paragraph", "w:styleId": "TOCHeading" },
  "w:name":       { $: { "w:val": "TOC Heading" } },
  "w:basedOn":    { $: { "w:val": basedOn } },
  "w:next":       { $: { "w:val": "Normal" } },
  "w:uiPriority": { $: { "w:val": "39" } },
  "w:unhideWhenUsed": {},
  "w:qFormat":    {},
  // `w:pPr` is an ORDERED sequence — spacing (22) before outlineLvl (31).
  "w:pPr": {
    "w:spacing":    { $: { "w:before": "240", "w:after": "0" } },
    // outlineLvl 9 = body text. Without it the title itself would be collected
    // into the very table it introduces on the next update.
    "w:outlineLvl": { $: { "w:val": "9" } },
  },
});

/**
 * Word's "toc N" entry styles. Each level indents a further 220 twips and ends
 * in a right tab stop with a dot leader, which is what draws the "…… 12".
 */
const TOC_ENTRY_STYLE = (level: number) => ({
  $: { "w:type": "paragraph", "w:styleId": `TOC${level}` },
  "w:name":       { $: { "w:val": `toc ${level}` } },
  "w:basedOn":    { $: { "w:val": "Normal" } },
  "w:next":       { $: { "w:val": "Normal" } },
  "w:autoRedefine": {},
  "w:uiPriority": { $: { "w:val": "39" } },
  "w:unhideWhenUsed": {},
  // `w:pPr` is an ORDERED sequence — tabs (11), spacing (22), ind (23).
  "w:pPr": {
    "w:tabs": {
      "w:tab": { $: { "w:val": "right", "w:leader": "dot", "w:pos": String(PAGE_TAB_POS) } },
    },
    "w:spacing": { $: { "w:after": "100" } },
    ...(level > 1 ? { "w:ind": { $: { "w:left": String((level - 1) * 220) } } } : {}),
  },
});

/**
 * What the `TC` fields of a paragraph put in a table of contents. A title drawn in a shape (WordArt, a
 * text box) is no text of its paragraph, so Word's `\\o` never collects it: a TC field is how it gets a
 * line, on its true page, printing nothing. Word collects one only when its runs are NOT hidden, and the
 * table's instruction carries `\\f`: measured in real Word, 2026-11-04.
 */
export function tcFieldEntries(paragraphXml: string): Array<{ text: string; level: number }> {
  const out: Array<{ text: string; level: number }> = [];
  for (const m of paragraphXml.matchAll(/<w:instrText\b[^>]*>([^<]*)<\/w:instrText>/g)) {
    const code = decodeXmlText(m[1]);
    const tc = /^\s*TC\s+"((?:[^"\\]|\\.)*)"([\s\S]*)$/.exec(code);
    if (!tc) continue;
    const level = /\\l\s*"?(\d)"?/.exec(tc[2])?.[1];
    out.push({ text: tc[1].replace(/\\(.)/g, "$1").trim(), level: level ? Number(level) : 1 });
  }
  return out;
}

/**
 * The document's Table of Contents — Word's References → Table of Contents.
 *
 * Every read/write goes through DocumentManager's ORDER-PRESERVING block API
 * (OrderedBody string splicing). Rebuilding `document.xml` through an XML object
 * model regroups the body by tag — hoisting every table away from its paragraphs
 * — and trims the whitespace-only runs Word uses between words.
 *
 * Like Word, the table is a `TOC` FIELD, so it renumbers itself on repagination.
 * Unlike Word, it is written PRE-POPULATED with one entry per heading (each a
 * hyperlink to a `_Toc…` bookmark plus a `PAGEREF` page number), so it reads
 * correctly in the app and in viewers that never update fields.
 */
export class TableOfContentsManager {
  private zip: AdmZip;
  private doc: DocumentManager;
  private styles: StylesManager;
  /** Deepest level whose TOC style is known registered — NOT a boolean: the same
   *  manager is reused across calls (the engine is cached per thesis), and a
   *  second call at depth 5 after a first at depth 3 must still register TOC4/5
   *  or those entries silently render as body text. */
  private registeredDepth = 0;

  constructor(zip: AdmZip) {
    this.zip    = zip;
    this.doc    = new DocumentManager(zip);
    this.styles = new StylesManager(zip);
  }

  // ── Internal ────────────────────────────────────────────────────────────────

  /** Register TOCHeading + TOC1…TOC{depth} into styles.xml the first time. */
  private async ensureStyles(depth: number): Promise<void> {
    if (depth <= this.registeredDepth) return;
    const existing = await this.styles.listStyles();
    const ids      = new Set(existing.map((s) => s.id));
    if (!ids.has("TOCHeading")) {
      // Base it on Heading1 for the thesis's own heading look, but only if the
      // document actually defines Heading1 — a dangling w:basedOn makes Word
      // fall back to nothing and renders the title as body text.
      await this.styles.addStyle(TOC_HEADING_STYLE(ids.has("Heading1") ? "Heading1" : "Normal"));
    }
    for (let lvl = 1; lvl <= depth; lvl++) {
      if (!ids.has(`TOC${lvl}`)) await this.styles.addStyle(TOC_ENTRY_STYLE(lvl));
    }
    this.registeredDepth = depth;
  }

  /**
   * A bookmark allocator handing out a DISTINCT id/name per call.
   *
   * document.xml is only re-read once: a batch that mints many bookmarks before
   * saving would otherwise stamp the same `w:id`/`_Toc…` on all of them, leaving
   * every entry's PAGEREF pointing at the first heading.
   */
  private bookmarkSeries(): () => { id: number; name: string } {
    const xml = this.zip.readAsText(DOCUMENT_PATH) ?? "";
    let maxId = 0;
    for (const m of xml.matchAll(/<w:bookmark(?:Start|End)\b[^>]*\bw:id="(\d+)"/g)) {
      maxId = Math.max(maxId, Number(m[1]));
    }
    let maxRef = 100000000;
    for (const m of xml.matchAll(/w:name="_(?:Ref|Toc)(\d+)"/g)) {
      maxRef = Math.max(maxRef, Number(m[1]));
    }
    let n = 0;
    return () => {
      const k = n++;
      return { id: maxId + 1 + k, name: `_Toc${maxRef + 1 + k}` };
    };
  }

  /**
   * The name of a bookmark PAGEREF can already aim at, or null.
   *
   * Only `_Toc…`/`_Ref…` bookmarks count: Word's own `_GoBack` marks the last
   * edit position and moves, so an entry anchored to it would drift.
   */
  private existingAnchor(paragraphXml: string): string | null {
    for (const m of paragraphXml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)) {
      const name = decodeXmlText(m[1]);
      if (/^_(?:Toc|Ref)/.test(name)) return name;
    }
    return null;
  }

  /** Wrap a heading paragraph in a bookmark so PAGEREF/hyperlink can find it. */
  private addAnchor(paragraphXml: string, bm: { id: number; name: string }): string {
    const pPrEnd   = paragraphXml.indexOf("</w:pPr>");
    const insertAt = pPrEnd === -1 ? paragraphXml.indexOf(">") + 1 : pPrEnd + "</w:pPr>".length;
    const closeIdx = paragraphXml.lastIndexOf("</w:p>");
    return (
      paragraphXml.slice(0, insertAt) +
      `<w:bookmarkStart w:id="${bm.id}" w:name="${escapeXmlText(bm.name)}"/>` +
      paragraphXml.slice(insertAt, closeIdx) +
      `<w:bookmarkEnd w:id="${bm.id}"/>` +
      paragraphXml.slice(closeIdx)
    );
  }

  /** The `TOC` field instruction, mirroring the switches Word writes. */
  private buildInstrText(options: TocOptions, depth: number, tcFields = false): string {
    let instr = ` TOC \\o "1-${depth}"`;
    if (options.useHyperlinks !== false)   instr += " \\h";
    if (options.includePageNumbers === false) instr += " \\n";     // omit page numbers
    instr += " \\z \\u ";
    if (tcFields) instr += "\\f ";
    return instr;
  }

  /** One pre-populated entry: `<heading text><dot leader tab><page number>`. */
  private buildEntryParagraph(
    text: string,
    level: number,
    anchor: string,
    opts: { rtl: boolean; pageNumbers: boolean; hyperlinks: boolean },
  ): string {
    const rPr  = opts.rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    const bidi = opts.rtl ? "<w:bidi/>" : "";

    const textRun = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(text)}</w:t></w:r>`;
    const pageRun = opts.pageNumbers
      ? `<w:r>${rPr}<w:tab/></w:r>` +
        `<w:fldSimple w:instr="${escapeXmlText(` PAGEREF ${anchor} \\h `)}">` +
        `<w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>`
      : "";

    // An internal hyperlink (w:anchor) needs NO relationship — nothing to add to
    // document.xml.rels, so the entry is clickable the moment it is written.
    const body = opts.hyperlinks
      ? `<w:hyperlink w:anchor="${escapeXmlText(anchor)}" w:history="1">${textRun}${pageRun}</w:hyperlink>`
      : textRun + pageRun;

    // `w:pPr` is an ORDERED sequence: tabs (11) BEFORE bidi (19). Emitting bidi
    // first makes Word refuse the file — and it is exactly the Arabic theses this
    // serves that carry the bidi.
    return (
      `<w:p><w:pPr><w:pStyle w:val="TOC${level}"/>` +
      `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${PAGE_TAB_POS}"/></w:tabs>${bidi}</w:pPr>` +
      body +
      `</w:p>`
    );
  }

  /** Every field instruction in a block, concatenated (Word splits them across runs). */
  private instructionsIn(xml: string): string {
    return (
      [...xml.matchAll(/<w:instrText\b[^>]*?>([\s\S]*?)<\/w:instrText>/g)].map((m) => decodeXmlText(m[1])).join("") +
      [...xml.matchAll(/<w:fldSimple\b[^>]*?\bw:instr="([^"]*)"/g)].map((m) => decodeXmlText(m[1])).join("")
    );
  }

  /**
   * True when the block opens a table of CONTENTS field.
   *
   * `TOC \c "Figure"` is a caption list (List of Figures / List of Tables) and is
   * deliberately excluded — replacing the contents must never eat one. A
   * `PAGEREF _Toc…` entry does not match either: the needle is `TOC \`.
   */
  private isTocField(xml: string): boolean {
    const instr = this.instructionsIn(xml);
    return /(?:^|\s)TOC\s+\\/.test(instr) && !/\\c\s*"/.test(instr);
  }

  // ── Public API ──────────────────────────────────────────────────────────────

  /**
   * Insert a Table of Contents at the given BODY BLOCK index (default 0 = top).
   *
   * Collects every heading down to `headingDepth`, anchoring each one so the
   * entries link and paginate. Headings are recognised by style (`Heading1`,
   * `Titre 1`, `Title`) OR by their own outline level, so imported theses whose
   * headings carry no heading style are still collected — as long as something
   * marked them as headings (use `set_heading` / infer_structure first if not).
   */
  public async insertTOC(options: TocOptions = {}, index = 0): Promise<TocResult> {
    const depth       = Math.max(1, Math.min(9, Math.round(options.headingDepth ?? 3)));
    const rtl         = options.rtl === true;
    const pageNumbers = options.includePageNumbers !== false;
    const hyperlinks  = options.useHyperlinks !== false;

    // Replace first, so a re-run refreshes the table instead of stacking a second
    // one — and so the heading scan below never collects the OLD table's own
    // TOCHeading title.
    let replaced = 0;
    let insertAt = index;
    if (options.replaceExisting !== false) {
      const removal = await this.removeTOC();
      replaced = removal.removed;
      // The caller's index was measured BEFORE the removal; shift it back by
      // whatever was deleted ahead of it.
      if (removal.at !== -1 && removal.at < insertAt) insertAt = Math.max(removal.at, insertAt - removal.removed);
    }

    await this.ensureStyles(depth);

    const blocks    = await this.doc.getBlocks();
    const nextBm    = this.bookmarkSeries();
    const entries: Array<{ text: string; level: number; anchor: string }> = [];
    let tcFields = false;

    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind !== "paragraph") continue;
      const fields = tcFieldEntries(b.xml);
      if (fields.length) tcFields = true;
      // A TC field says what its paragraph is listed as: Word reads no heading text out of a shape.
      const wanted = fields.length
        ? fields.filter((f) => f.text && f.level >= 1 && f.level <= depth)
        : [{ text: paragraphText(b.xml).trim(), level: paragraphHeadingLevel(b.xml) }].filter((f) => f.text && f.level >= 1 && f.level <= depth);
      if (!wanted.length) continue;

      let anchor = this.existingAnchor(b.xml);
      if (!anchor) {
        const bm = nextBm();
        blocks[i] = { ...b, xml: this.addAnchor(b.xml, bm) };
        anchor = bm.name;
      }
      for (const w of wanted) entries.push({ text: w.text, level: w.level, anchor });
    }

    const heading   = options.title !== undefined ? options.title : "Table of Contents";
    const instrText = this.buildInstrText(options, depth, tcFields);
    const rPr       = rtl ? "<w:rPr><w:rtl/></w:rPr>" : "";
    const bidi      = rtl ? "<w:bidi/>" : "";
    const para = (xml: string): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml });
    const newParas: BodyBlock[] = [];

    if (heading) {
      newParas.push(
        para(
          `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/>${bidi}</w:pPr>` +
            `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent(heading)}</w:t></w:r></w:p>`,
        ),
      );
    }

    // Field begin + instruction + separate …
    //
    // NOTE on page numbers: nothing here can paginate, so every entry below
    // caches page 1 until a layout engine recomputes the field. Marking the field
    // `w:dirty="true"` would make Word recompute it on open, but this engine
    // deliberately emits NO dirty fields (see caption-list.spec / hanachi spec) —
    // do not add one without changing that decision.
    newParas.push(
      para(
        `<w:p><w:pPr>${bidi}</w:pPr>` +
          `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
          `<w:r>${rPr}<w:instrText xml:space="preserve">${escapeXmlContent(instrText)}</w:instrText></w:r>` +
          `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r></w:p>`,
      ),
    );

    if (entries.length) {
      for (const e of entries) {
        newParas.push(para(this.buildEntryParagraph(e.text, e.level, e.anchor, { rtl, pageNumbers, hyperlinks })));
      }
    } else {
      newParas.push(
        para(
          `<w:p><w:pPr><w:pStyle w:val="TOC1"/>${bidi}</w:pPr>` +
            `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlContent("No headings found")}</w:t></w:r></w:p>`,
        ),
      );
    }

    // … field end.
    newParas.push(
      para(`<w:p><w:pPr>${bidi}</w:pPr><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r></w:p>`),
    );

    insertAt = Math.max(0, Math.min(insertAt, blocks.length));
    blocks.splice(insertAt, 0, ...newParas);
    await this.doc.saveBlocks(blocks);

    return { atIndex: insertAt, entries: entries.length, headingDepth: depth, replaced };
  }

  /**
   * Remove every Table of Contents — its title, the field, and all pre-populated
   * entry paragraphs. Caption lists (`TOC \c "Figure"`) are left alone.
   *
   * A TOC field SPANS paragraphs (begin … separate … entries … end), so this
   * walks from the paragraph holding the instruction to the one holding the
   * matching `fldChar end` rather than deleting a single paragraph.
   */
  public async removeTOC(): Promise<TocRemoval> {
    const blocks = await this.doc.getBlocks();
    let removed = 0;
    let firstAt = -1;

    // A document can hold more than one (a stale table plus a fresh one); loop
    // until no table of contents is left. `start` only ever moves forward, so the
    // scan terminates even if a field is malformed.
    for (let guard = 0; guard < 32; guard++) {
      const start = blocks.findIndex((b) => this.isTocField(b.xml));
      if (start === -1) break;

      let end = start;
      if (/w:fldCharType="begin"/.test(blocks[start].xml)) {
        let depth = 0;
        for (let i = start; i < blocks.length; i++) {
          depth += (blocks[i].xml.match(/w:fldCharType="begin"/g) ?? []).length;
          depth -= (blocks[i].xml.match(/w:fldCharType="end"/g) ?? []).length;
          end = i;
          if (depth <= 0) break;
        }
      }

      // The heading directly above ("Table of Contents", "الفهرس") belongs to it.
      const from = start > 0 && /w:val="TOCHeading"/.test(blocks[start - 1].xml) ? start - 1 : start;
      if (firstAt === -1) firstAt = from;
      removed += end - from + 1;
      blocks.splice(from, end - from + 1);
    }

    if (removed) await this.doc.saveBlocks(blocks);
    return { removed, at: firstAt };
  }

  /** True when the document already carries a Table of Contents field. */
  public async hasTOC(): Promise<boolean> {
    const blocks = await this.doc.getBlocks();
    return blocks.some((b) => this.isTocField(b.xml));
  }

  // ── Hand-typed tables of contents ───────────────────────────────────────────

  /** Block indices that sit INSIDE any field (a real TOC, a caption list, …). */
  private fieldBlockIndices(blocks: BodyBlock[]): Set<number> {
    const inField = new Set<number>();
    let depth = 0;
    for (let i = 0; i < blocks.length; i++) {
      const opens  = (blocks[i].xml.match(/w:fldCharType="begin"/g) ?? []).length;
      const closes = (blocks[i].xml.match(/w:fldCharType="end"/g) ?? []).length;
      if (depth > 0 || opens > 0) inField.add(i);
      depth += opens - closes;
      if (depth < 0) depth = 0;
    }
    return inField;
  }

  /**
   * Does this paragraph read as a hand-typed contents line?
   *
   * Two signals, BOTH required: it ends in a page number, and it separates that
   * number from the title with a leader — typed dots, or a real tab. A tab
   * carries no text, so "المقدمة<tab>5" arrives as "المقدمة5"; the XML is what
   * says which it was.
   */
  private isTypedTocEntry(xml: string): boolean {
    const text = paragraphText(xml).trim();
    if (!text || text.length > 300) return false;
    if (!TRAILING_PAGE.test(text)) return false;
    return TEXT_LEADER.test(text) || /<w:(?:tab|ptab)\b/.test(xml);
  }

  /**
   * Find a table of contents the student TYPED as ordinary paragraphs.
   *
   * This is the one the model could not delete: it is not a field, so
   * {@link removeTOC} never saw it, and the model was left guessing a block
   * range — "صعوبة تقنية في تحديد نهاية الفهرس اليدوي" — and deleting the wrong
   * thing. Here the span is computed exactly.
   *
   * Deliberately conservative: at least `minEntries` consecutive entry lines
   * (single blank lines tolerated), and anything inside a real field is skipped
   * so a generated table or a list of figures can never be mistaken for one.
   */
  public async findTypedTOC(minEntries = 3): Promise<TypedTocSpan | null> {
    const blocks  = await this.doc.getBlocks();
    const inField = this.fieldBlockIndices(blocks);

    for (let i = 0; i < blocks.length; i++) {
      if (inField.has(i) || blocks[i].kind !== "paragraph") continue;
      if (!this.isTypedTocEntry(blocks[i].xml)) continue;

      // Consume the run of entry lines, allowing a single blank between them.
      let end = i;
      let entries = 0;
      const sample: string[] = [];
      for (let j = i; j < blocks.length; j++) {
        const b = blocks[j];
        if (inField.has(j)) break;
        if (b.kind !== "paragraph") break;
        if (this.isTypedTocEntry(b.xml)) {
          entries++;
          if (sample.length < 5) sample.push(paragraphText(b.xml).trim());
          end = j;
          continue;
        }
        if (!paragraphText(b.xml).trim()) continue;  // a blank line inside the list
        break;
      }
      if (entries < minEntries) { i = end; continue; }

      // The heading directly above ("الفهرس", "Table des matières") belongs to it.
      let start = i;
      for (let k = i - 1; k >= 0 && k >= i - 2; k--) {
        if (inField.has(k) || blocks[k].kind !== "paragraph") break;
        const t = paragraphText(blocks[k].xml).trim();
        if (!t) continue;                            // skip a blank line between
        if (TYPED_TOC_TITLE.test(t)) { start = k; break; }
        break;
      }
      const title = start < i ? paragraphText(blocks[start].xml).trim() : null;
      return { startIndex: start, endIndex: end, title, entries, sample };
    }
    return null;
  }

  /**
   * Delete a hand-typed table of contents (title + every entry line).
   *
   * DESTRUCTIVE in a way {@link removeTOC} is not: these are paragraphs the
   * student wrote, not a field this engine generated. Callers must have the
   * student's agreement.
   */
  public async removeTypedTOC(span?: TypedTocSpan): Promise<TocRemoval> {
    const target = span ?? (await this.findTypedTOC());
    if (!target) return { removed: 0, at: -1 };
    const blocks = await this.doc.getBlocks();
    const count  = target.endIndex - target.startIndex + 1;
    blocks.splice(target.startIndex, count);
    await this.doc.saveBlocks(blocks);
    return { removed: count, at: target.startIndex };
  }
}
