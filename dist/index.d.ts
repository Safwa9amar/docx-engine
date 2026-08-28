import { default as default_2 } from 'adm-zip';

export declare interface AppendOptions {
    /** Prepend a page break before everything this call appends. */
    startOnNewPage?: boolean;
    /** Blocks inserted BEFORE the copied source body (e.g. a part-title heading). */
    leadingBlocks?: BodyBlock[];
    /** source styleId → target styleId (applied to `<w:pStyle w:val>`). */
    styleMap?: Record<string, string>;
}

export declare function applyBodyPageLayout(documentXml: string, opts: BodyPageLayoutOpts): string;

/**
 * Place the first picture in a paragraph, converting its container as needed.
 *
 * - Asking for a `vertical` position (or `float: true`) FLOATS an inline
 *   picture: `wp:inline` → `wp:anchor` with both axes written as named
 *   alignments relative to `relativeTo` (default `page`). This is the only way
 *   "centre it on the page" can be expressed in OOXML.
 * - `float: false` returns a floating picture to the text flow — the positioning
 *   children are dropped and the container becomes `wp:inline` again.
 * - Passing ONLY `horizontal` on an inline picture changes nothing here: it is
 *   the carrier PARAGRAPH's `w:jc` that moves an in-flow picture sideways, and
 *   floating a picture just to nudge it left would change how text flows around
 *   it. The caller handles that case (`changed: false` with `floating: false`
 *   is the signal). On an already-floating picture, `horizontal` IS applied.
 *
 * Everything outside the container's geometry prefix — the whole `a:graphic`
 * subtree, so the image itself, its crop, its effects — is copied byte for byte.
 *
 * ON `mc:AlternateContent`. A great many real pictures ship as a pair: an
 * `mc:Choice` holding the modern DrawingML, and an `mc:Fallback` holding a VML
 * twin of the same picture for readers older than Word 2007. This function used
 * to refuse the whole paragraph on sight of one, on the theory that rewriting a
 * single branch leaves the file inconsistent. That was the wrong trade, and a
 * student's بسم الله page proved it within the hour: the tool refused, the
 * assistant reported the capability as missing, and the student got nothing —
 * over a compatibility branch that no Word since 2007 has read.
 *
 * So: the `mc:Choice` drawing IS the picture, and it is repositioned. The
 * `mc:Fallback` twin is left exactly as it was — Word never renders it, our own
 * reader already discards it (`stripAltContentFallback`), and rewriting VML
 * geometry is a second, unrelated language. `legacyTwin` reports that it is
 * there. Only a drawing found INSIDE the fallback is refused, because moving the
 * copy nobody renders would look, to the student, like nothing happened.
 *
 * @throws {Error} when `xml` holds no drawing this can move.
 */
export declare function applyDrawingLayout(xml: string, opts: DrawingLayout): DrawingLayoutResult;

/**
 * Apply pagination toggles to one paragraph's XML.
 *
 * `true` writes the bare element (`<w:keepNext/>`); `false` writes an EXPLICIT
 * `w:val="0"` rather than deleting the element. Deleting would only fall back to
 * whatever the paragraph's style says — and for a Heading style that is very
 * often `keepNext` ON, so "stop keeping this heading with the next paragraph"
 * would silently do nothing. `undefined` leaves the property alone.
 */
export declare function applyParagraphPagination(xml: string, opts: ParagraphPagination): string;

/**
 * Apply DIRECT run formatting to every run of ONE paragraph, plus its paragraph
 * mark.
 *
 * Direct, not style-level, on purpose. Two reasons, both learned the hard way:
 * an imported thesis carries its formatting on the RUNS, and direct run
 * formatting beats a style in the OOXML cascade — so a style patch alone is
 * invisible on exactly the documents students bring us; and the whole point of
 * this function is to hit ONE paragraph, which a shared style cannot do.
 *
 * The paragraph mark is included so the size sticks to the ¶ itself: Word draws
 * the empty line after a heading at the mark's size, and text typed at the end
 * of the paragraph inherits it.
 *
 * @throws {Error} if `xml` is not a `<w:p>` element, or if `props` carries an
 * invalid size or colour (see `runProps.mergeRunProps`).
 */
export declare function applyRunPropsToParagraph(xml: string, props: RunProps): ParagraphRunStyleResult;

/**
 * Pure transform: apply `props` to the `<w:rPr>` of the `styleId` style inside a
 * `word/styles.xml` string, creating the style first when `ensure` is supplied
 * and it does not exist. Returns the rewritten XML plus whether the style was
 * created and whether it ended up changed.
 *
 * The seed `thesis-base.docx` defines neither `Normal` nor `Caption`, so the
 * create path is the ordinary one, not an error path.
 *
 * @throws if the style fragment is malformed markup (propagated from
 * `canonicalizeRunProps`). This operates on ONE small styles fragment, not
 * per-run, so failing loudly is correct — a broken `styles.xml` must not be
 * half-applied. Callers doing PER-RUN work must catch per run instead.
 */
export declare function applyStyleRunPropsToXml(stylesXml: string, styleId: string, props: RunProps, ensure?: EnsureStyleSpec): {
    xml: string;
    created: boolean;
    updated: boolean;
};

export declare interface AppProperties {
    application?: string;
    pages?: number;
    words?: number;
    characters?: number;
}

/** Reassemble a split document into a full `document.xml` string. */
export declare function assembleDocument(split: {
    pre: string;
    blocks: BodyBlock[];
    post: string;
}): string;

/** One body block as plain data. */
export declare interface BlockInfo {
    index: number;
    kind: "paragraph" | "table" | "image" | "other";
    /** Visible text (paragraph/cell text); "" for images. */
    text: string;
    /** Paragraph style id, if any. */
    styleId: string | null;
    /** Heading level 1–6, or 0 if not a heading. */
    headingLevel: number;
}

export declare type BlockKind = "paragraph" | "table" | "sectPr" | "other";

/**
 * Caller hook to render a block the engine handles differently (or not at all).
 * Returns paragraphs to append plus optional table/image inserts whose
 * `afterParaCount` is LOCAL to the returned paragraphs (0 = before the first).
 * The engine rebases those offsets onto the running paragraph count. Return
 * `undefined` (or omit the hook) to fall back to the engine's built-in
 * rendering for that block.
 */
export declare type BlockRenderer = (block: MarkdownBlock, ctx: MarkdownRenderCtx) => RenderedChunk | undefined;

/**
 * A single ordered body child, stored as its exact original XML substring.
 * `xml` is the verbatim slice of `document.xml` for this top-level child —
 * including any whitespace/comment text captured as an "other" block.
 */
export declare interface BodyBlock {
    kind: BlockKind;
    tag: string;
    xml: string;
}

export declare interface BodyPageLayoutOpts {
    marginPreset?: MarginPreset;
    orientation?: Orientation;
    pageSizePreset?: PageSizePreset;
    columns?: number;
}

export declare interface BookmarkEntry {
    id: number;
    name: string;
    text: string;
}

export declare interface BorderSide {
    style?: string;
    size?: number;
    color?: string;
}

declare type BreakType = "nextPage" | "evenPage" | "oddPage";

/** Reassemble from a SplitDocument (or its parts). */
export declare function buildOrderedDoc(split: SplitDocument): string;

/** Build a minimal, schema-ordered paragraph `<w:style>`. */
export declare function buildParagraphStyleXml(styleId: string, name: string, opts?: {
    basedOn?: string;
    isDefault?: boolean;
}): string;

/**
 * Sort the INNER XML of a `<w:pPr>` into `CT_PPr` order.
 *
 * @throws {Error} on malformed markup - see {@link splitTopLevelElements}.
 * Callers doing PER-PARAGRAPH work over a whole document must catch per
 * paragraph rather than letting one bad paragraph abort the document.
 */
export declare const canonicalizeParagraphProps: (pPrInner: string) => string;

export declare interface CaptionEntry {
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

export declare type CaptionKind = "figure" | "table" | "equation";

export declare class CaptionManager {
    private zip;
    private styles;
    private doc;
    private stylesRegistered;
    constructor(zip: default_2);
    /** Register Caption + CaptionChar styles into styles.xml the first time. */
    private ensureStyles;
    /** Next free `w:id` for a bookmark, and a free `_Ref…` bookmark name. */
    private nextBookmark;
    /**
     * A bookmark allocator that hands out a DISTINCT id/name on every call.
     *
     * {@link nextBookmark} re-reads document.xml each time, so a batch that mints
     * many bookmarks before saving (the converter) would stamp the same
     * `w:id`/`_Ref…` on all of them — duplicate bookmark names, and every
     * cross-reference and list-of-figures entry pointing at the first one.
     */
    private bookmarkSeries;
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
    private buildCaptionXml;
    /**
     * Everything a caption paragraph holds AFTER its `<w:pPr>`: the bookmark, the
     * label runs, the chapter STYLEREF, the SEQ field and the caption text.
     *
     * Split out of {@link buildCaptionXml} because {@link convertTextCaptions}
     * needs the same content inside a paragraph that ALREADY exists — it keeps the
     * student's own `<w:pPr>` and run properties rather than writing fresh ones.
     * `rPr` is the run-properties string every generated run carries.
     */
    private captionContent;
    /**
     * Recompute every SEQ (and caption STYLEREF) field RESULT in the body, in
     * DOCUMENT order, honouring each field's own `\*` format and `\s` chapter
     * reset — exactly what Word does when you press F9.
     *
     * This matters because nothing outside Word updates fields: the app's preview
     * and the OnlyOffice PDF both render the cached result. Without this pass a
     * caption inserted before an existing one keeps the higher number on screen.
     */
    renumber(): Promise<void>;
    /** In-memory half of {@link renumber}. Returns true when something changed. */
    private renumberBlocks;
    /** Rewrite every SEQ / caption-STYLEREF field result inside one block's XML. */
    private renumberFieldsIn;
    /**
     * Read EVERY caption in one paragraph — Word numbers each SEQ field it finds,
     * and real theses do put two figures on one line ("Figure 19: … Figure 20: …").
     * Returns [] when the paragraph carries no caption field.
     */
    private readCaptions;
    /** The first caption in a paragraph, or null. */
    private readCaption;
    /**
     * Insert a caption paragraph next to the BLOCK at `nearIndex` — the same index
     * space as `document.getBlocks()`, so a table's index really does place the
     * caption against that table.
     *
     * `position: "below"` (default) inserts after `nearIndex`, `"above"` before it.
     *
     * @returns The BLOCK index at which the caption was inserted.
     */
    insertCaption(nearIndex: number, opts?: CaptionOptions): Promise<number>;
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
    convertTextCaptions(opts?: ConvertTextCaptionsOptions): Promise<ConvertTextCaptionsResult>;
    /**
     * The label each converted kind will carry. Existing REAL captions win — an
     * Arabic thesis whose figures already say "الشكل رقم" must not gain a second
     * "شكل" sequence — then the typed spelling used most often, then `override`.
     */
    private resolveConvertLabels;
    /**
     * Every caption in the document, in document order. `filterLabel` accepts
     * either the readable label ("الشكل رقم") or its SEQ identifier.
     */
    getCaptions(filterLabel?: string): Promise<CaptionEntry[]>;
    /** The distinct caption labels used in the document, most-used first. */
    listLabels(): Promise<Array<{
        label: string;
        displayLabel: string;
        count: number;
    }>>;
    /**
     * Replace a caption's TEXT while keeping its label, number field and bookmark
     * intact — the edit Word makes when you retype a caption's wording.
     * Returns false when the block isn't a caption.
     */
    setCaptionText(blockIndex: number, text: string): Promise<boolean>;
    /** String half of {@link setCaptionText}; null when `xml` isn't a caption. */
    private replaceCaptionText;
    /**
     * Remove every caption for a label (e.g. all "Figure" captions), then
     * renumber what remains.
     */
    removeCaptions(label: string): Promise<void>;
    /** Remove a single caption by BLOCK index, then renumber. */
    removeCaptionAt(blockIndex: number): Promise<void>;
    /** Convenience: insert a "Figure N — text" caption below a block. */
    insertFigureCaption(nearIndex: number, text: string, numbering?: CaptionNumberingOptions, opts?: Omit<CaptionOptions, "text" | "numbering">): Promise<number>;
    /** Convenience: insert a "Table N — text" caption above (Word default) a table. */
    insertTableCaption(nearIndex: number, text: string, position?: CaptionPosition, numbering?: CaptionNumberingOptions, opts?: Omit<CaptionOptions, "text" | "numbering" | "position">): Promise<number>;
    /** Convenience: insert an "Equation N" caption. */
    insertEquationCaption(nearIndex: number, text?: string, numbering?: CaptionNumberingOptions, opts?: Omit<CaptionOptions, "text" | "numbering">): Promise<number>;
    /** Insert a caption with a custom label (Word's "New Label"). */
    insertCustomCaption(nearIndex: number, label: string, text: string, opts?: Omit<CaptionOptions, "label" | "text">): Promise<number>;
    /** One pre-populated entry: `<entry text><dot leader tab><page number>`. */
    private buildListEntryParagraph;
    /** Give a caption paragraph a bookmark if it has none, so PAGEREF can find it. */
    private ensureCaptionBookmark;
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
    insertCaptionList(label: string, title?: string, index?: number, rtl?: boolean): Promise<void>;
    /** Shorthand: insert List of Figures at the given BLOCK index. */
    insertListOfFigures(title?: string, index?: number, rtl?: boolean): Promise<void>;
    /** Shorthand: insert List of Tables at the given BLOCK index. */
    insertListOfTables(title?: string, index?: number, rtl?: boolean): Promise<void>;
    /**
     * Remove the caption list (`TOC \c` field) for a given label — its heading,
     * the field, and every pre-populated entry paragraph.
     *
     * A TOC field SPANS paragraphs (begin … separate … entries … end), so this
     * walks from the paragraph holding the instruction to the one holding the
     * matching `fldChar end` rather than deleting a single paragraph.
     */
    removeCaptionList(label: string): Promise<void>;
}

/** Numbering format matching Word's Caption Numbering dialog */
export declare type CaptionNumberFormat = "arabic" | "ROMAN" | "roman" | "ALPHABETIC" | "alphabetic";

export declare interface CaptionNumberingOptions {
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

export declare interface CaptionOptions {
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

export declare type CaptionPosition = "below" | "above";

declare interface CellMargins {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
}

declare interface CellTextOptions {
    bold?: boolean;
    italic?: boolean;
    color?: string;
    fontSize?: number;
    fontFamily?: string;
    alignment?: "left" | "center" | "right" | "justify";
}

/** Chapter-number separator matching Word's "Use separator" dropdown */
export declare type ChapterSeparator = "-" | "." | ":" | "–" | "—";

declare type ChapterSeparator_2 = "hyphen" | "period" | "colon" | "emDash" | "enDash";

declare type ChapterSeparator_3 = "hyphen" | "period" | "colon" | "emDash" | "enDash";

/** Check a .docx held as bytes. Read-only — the buffer is never modified. */
export declare function checkDocxBuffer(buffer: Buffer): Promise<DoctorReport>;

/** A picture placed in a header/footer part, with the geometry needed to paint
 *  it where Word paints it. See {@link extractChromeDrawings}. */
export declare interface ChromeDrawing {
    /** The part-local `r:embed` id — resolution detail, kept for debugging. */
    embedId: string;
    /** Media file name inside `word/media` (e.g. "image1.png"), null if unresolved. */
    image: string | null;
    extent: {
        cxEmu: number;
        cyEmu: number;
    };
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

/** One placed axis of an anchored drawing. `offsetEmu` is signed — decorative
 *  full-page art is routinely offset NEGATIVELY so it overflows its anchor. */
export declare interface ChromeDrawingAxis {
    /** OOXML frame of reference: "page" | "margin" | "column" | "paragraph" | … */
    relativeTo: string;
    offsetEmu: number | null;
    /** Named alignment ("center", "right") when used instead of an offset. */
    align: string | null;
}

/** Word's "Recolor" on a picture: the stored bytes are painted between two
 *  colours rather than shown as-is. `dark`/`light` are 6-hex when the file names
 *  a literal colour; `*Scheme` carries a theme slot ("accent4") to resolve
 *  against theme1.xml. `shade`/`satMod` are fractions (0.45 = 45%). */
export declare interface ChromeDuotone {
    dark: string | null;
    darkScheme: string | null;
    light: string | null;
    lightScheme: string | null;
    shade: number | null;
    satMod: number | null;
}

/** One logo image embedded inside an applied header/footer part. `token` is a
 *  unique placeholder present in the region XML (e.g. inside `<a:blip r:embed>`),
 *  replaced with the part-local relationship id once the bytes are embedded. */
declare type ChromeImage = {
    token: string;
    bytes: Buffer;
    ext: string;
};

/** A compiled header or footer to apply to a section. `xml` is the INNER region
 *  body (paragraphs/tables), NOT wrapped in `<w:hdr>`/`<w:ftr>` — {@link Doc.applySectionChrome}
 *  adds the wrapper + namespaces. */
declare type ChromePart = {
    xml: string;
    images: ChromeImage[];
};

export declare class CitationManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private readSources;
    private writeSources;
    /**
     * Returns all citation sources stored in the document.
     */
    getSources(): Promise<CitationSource[]>;
    /**
     * Add or update a citation source.
     */
    addSource(source: CitationSource): Promise<void>;
    /**
     * Remove a citation source by tag.
     */
    removeSource(tag: string): Promise<void>;
    /**
     * Create a run array that inserts an inline citation (e.g. "(Smith, 2020)").
     * Insert all returned runs into a paragraph in order.
     */
    createCitationRuns(tag: string, locale?: string): Run_2[];
    /**
     * Insert a bibliography paragraph at the given body index (default: end of document).
     */
    insertBibliography(index?: number): Promise<void>;
}

export declare interface CitationSource {
    tag: string;
    sourceType?: string;
    author?: string;
    title?: string;
    year?: string;
    city?: string;
    publisher?: string;
}

/**
 * Clear the "data descriptor present" flag on entries that do not have one.
 *
 * LibreOffice (and anything else that streams a zip) sets general-purpose bit 3
 * and appends a descriptor after each entry. adm-zip re-writes such an archive
 * with real sizes and CRCs in the headers but WITHOUT the descriptor — while
 * faithfully copying the flag that promises one. The result is a perfectly valid
 * zip by every other reader (`unzip -t` passes, Word opens it) that adm-zip's own
 * reader then refuses with "No descriptor present".
 *
 * That matters here because the server re-opens its own output on the next edit:
 * one save of a LibreOffice-authored upload was enough to make every later AI
 * edit of that thesis fail to load. Clearing the lying flag is a two-byte fix per
 * entry, and it makes the bytes describe what is actually in the file.
 *
 * Bails out untouched on anything unexpected (Zip64, truncated directory, a
 * local header that doesn't line up) — a half-rewritten zip is far worse than an
 * unhelpful one.
 */
export declare function clearFalseDataDescriptors(buffer: Buffer): Buffer;

/** Convert centimetres to twips */
export declare const cmToTwips: (cm: number) => number;

export declare interface ColumnOptions {
    count: number;
    space?: number;
    equalWidth?: boolean;
}

export declare interface CommentEntry {
    id: number;
    author: string;
    initials: string;
    date: string;
    text: string;
}

export declare class CommentsManager {
    private zip;
    private rels;
    private contentTypes;
    constructor(zip: default_2);
    private readComments;
    private writeComments;
    private emptyCommentsDoc;
    private normalizeComments;
    private normalizeArr;
    private nextId;
    private extractText;
    private ensureRegistered;
    private readDocument;
    private writeDocument;
    private getBodyParagraphs;
    getComments(): Promise<CommentEntry[]>;
    addComment(paragraphIndex: number, author: string, text: string, date?: string): Promise<number>;
    deleteComment(id: number): Promise<void>;
    resolveComment(id: number): Promise<void>;
}

export declare class ContentTypesManager {
    private readonly zip;
    private readonly filePath;
    private readonly ns;
    constructor(zip: default_2);
    /** Generate a GUID */
    private generateGuid;
    /**
     * Reads current [Content_Types].xml or returns a default structure
     */
    private readTypes;
    /**
     * Writes back the [Content_Types].xml into the zip
     */
    private writeTypes;
    /**
     * Adds a Default content type if not exists
     * Example: Extension="xml" ContentType="application/xml"
     */
    addDefault(extension: string, contentType: string): Promise<void>;
    /**
     * Adds an Override element if it doesn't exist yet.
     * Example:
     * PartName="/word/header1.xml"
     * ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
     */
    addOverride(partName: string, contentType: string): Promise<void>;
    /**
     * Removes an Override entry by PartName
     */
    removeOverride(partName: string): Promise<void>;
    /**
     * Checks if an override exists
     */
    hasOverride(partName: string): Promise<boolean>;
    /**
     * Helper to create a new unique part name with GUID
     */
    generateUniquePartName(prefix: string, extension?: string): string;
}

export declare interface ConvertedCaption {
    blockIndex: number;
    kind: CaptionKind;
    /** The label written, after unification (e.g. "Figure"). */
    label: string;
    /** The paragraph's text before conversion. */
    before: string;
    /** Its text after — the same wording, now with a live number. */
    after: string;
}

export declare interface ConvertTextCaptionsOptions {
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
    /** Report what WOULD convert without writing anything. */
    dryRun?: boolean;
}

export declare interface ConvertTextCaptionsResult {
    converted: ConvertedCaption[];
    skipped: SkippedTextCaption[];
    /** The label chosen per kind. */
    labels: Partial<Record<CaptionKind, string>>;
    /** True when nothing was written. */
    dryRun: boolean;
}

export declare interface CoreProperties {
    title?: string;
    subject?: string;
    creator?: string;
    description?: string;
    lastModifiedBy?: string;
    created?: string;
    modified?: string;
}

export declare class CrossReferenceManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private getAllParagraphs;
    private nextBookmarkId;
    /**
     * Returns all bookmarks defined in the document.
     */
    getBookmarks(): Promise<BookmarkEntry[]>;
    /**
     * Add a bookmark wrapping a run in the paragraph at the given body index.
     * Returns the bookmark id.
     */
    addBookmark(paragraphIndex: number, name: string, text: string): Promise<number>;
    /**
     * Remove a bookmark by name (removes bookmarkStart/End from the paragraph).
     */
    removeBookmark(name: string): Promise<void>;
    /**
     * Returns three run objects that together form a cross-reference field.
     * All three must be added to the same paragraph in order.
     */
    createCrossRefRuns(bookmarkName: string, displayText?: string): Run_2[];
}

/**
 * `CT_PPr` child sequence — `CT_PPrBase` (ECMA-376 Part 1 17.3.1.26) followed by
 * `CT_PPr`'s own trailing children (`w:rPr`, `w:sectPr`, `w:pPrChange`).
 *
 * The order matters far more than it looks: `w:keepNext` / `w:keepLines` /
 * `w:pageBreakBefore` sit near the FRONT (right after `w:pStyle`) while
 * `w:widowControl` sits after `w:framePr`, and `w:bidi` / `w:jc` come much
 * later. Appending a pagination toggle to an existing `<w:pPr>` that already
 * carries `<w:bidi/>` and `<w:jc/>` — the normal state of an Arabic thesis
 * paragraph — produces a sequence violation, and Word refuses the file.
 *
 * Mirrors the `CT_PPR` table the document doctor validates against (src/doctor).
 */
export declare const CT_PPR_ORDER: readonly string[];

/**
 * Common French/English Word heading + body style aliases → canonical target
 * styleIds. The combine flow passes this so mismatched source heading styles map
 * onto the template's styles by name.
 */
export declare const DEFAULT_STYLE_ALIASES: Record<string, string>;

/** A body block enriched with formatting signals — for inferring structure or AI labelling. */
export declare interface DetailedBlockInfo extends BlockInfo {
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

export declare class Doc {
    /** Escape hatch: the underlying engine + all its managers. */
    readonly engine: Mdocxengine;
    /** theme1.xml's colour scheme, parsed once per Doc (slot → 6-hex). */
    private themeColors;
    private constructor();
    /** Open a document from a file path or an in-memory buffer. */
    static open(source: string | Buffer): Promise<Doc>;
    /** Wrap an already-loaded engine. */
    static from(engine: Mdocxengine): Doc;
    /** Whole-document plain text (blocks joined by `separator`, default newline). */
    text(separator?: string): Promise<string>;
    /** Total word count of the body. */
    wordCount(): Promise<number>;
    /** Every body block as plain data, in document order. */
    blocks(): Promise<BlockInfo[]>;
    private toBlockInfo;
    /** The heading outline as a nested tree (like a table of contents). */
    outline(): Promise<OutlineNode[]>;
    /** Every table as a plain text grid. */
    tables(): Promise<TableInfo[]>;
    /** Every embedded inline image as bytes + size + mime. */
    images(): Promise<ImageInfo[]>;
    /**
     * Every body block enriched with formatting signals (bold, font size,
     * alignment, word count, caption flag). The raw material for inferring
     * structure when a document has no heading markup — feed it to a heuristic
     * (`inferOutline`) or to an LLM to label headings.
     */
    blocksDetailed(captionPatterns?: RegExp[]): Promise<DetailedBlockInfo[]>;
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
    inferOutline(opts?: InferOutlineOptions): Promise<InferredHeading[]>;
    /** Append a paragraph (or insert at `at`). */
    addParagraph(text: string, opts?: ParagraphFormat, at?: number): Promise<this>;
    /** Append a heading at `level` (1–6) → maps to `Heading{level}`, bold, with outline level. */
    addHeading(text: string, level?: number, opts?: ParagraphFormat, at?: number): Promise<this>;
    /** Replace the text of the paragraph at `index` (preserves its formatting). */
    editParagraph(index: number, text: string): Promise<this>;
    /**
     * Promote (or demote) the paragraph at `index` to a real heading at `level`
     * (1–6): applies the `Heading{level}` style + matching outline level, keeping
     * the text and preserving RTL/alignment. This is how you turn a plain-text
     * title — e.g. one found by `inferOutline()` or labelled by an LLM — into a
     * structural heading the outline/TOC can see.
     */
    setHeadingLevel(index: number, level: number): Promise<this>;
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
    setTextStyle(targets: readonly TextStyleTargetInput[], props: RunProps): Promise<TargetReport[]>;
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
    getTextStyle(targets?: readonly TextStyleTargetInput[]): Promise<TextStyleInspection>;
    /** Delete the block at `index`. */
    deleteBlock(index: number): Promise<this>;
    /** Append a table from a row-major grid (or insert at `at`). */
    addTable(rows: string[][], opts?: {
        header?: boolean;
        rtl?: boolean;
    }, at?: number): Promise<this>;
    /** Set the text of one cell of the table at block `index`. */
    editTableCell(index: number, row: number, col: number, value: string): Promise<this>;
    private mutateTable;
    /**
     * Insert a row into the table at `index`: below row `at` (0-based), ABOVE it
     * when `before` is true (e.g. a new first row: at=0, before=true), or
     * appended when `at` is omitted.
     */
    addTableRow(index: number, at?: number, before?: boolean): Promise<this>;
    /** Remove row `row` (0-based) from the table at `index`. */
    removeTableRow(index: number, row: number): Promise<this>;
    /**
     * Insert a column into the table at `index`: to the right of column `at`
     * (0-based), to its LEFT when `before` is true (e.g. a new first column:
     * at=0, before=true), or appended when omitted.
     */
    insertTableColumn(index: number, at?: number, before?: boolean): Promise<this>;
    /** Sort the table's data rows by column `col` (numeric when both parse). Header row 0 stays put unless includeHeader. */
    sortTable(index: number, col: number, opts?: {
        desc?: boolean;
        includeHeader?: boolean;
    }): Promise<this>;
    /**
     * Column/table widths: `columnsTwips[i]` sets column i's preferred width
     * (1440 twips = 1 inch; ~600 per cm); `autofit` switches the layout mode.
     */
    setTableWidths(index: number, opts: {
        columnsTwips?: number[];
        autofit?: "contents" | "window";
    }): Promise<this>;
    /**
     * Format a cell (row+col) or a whole row (row only, col omitted) — existing
     * text preserved: bold/italic, font size (points), font family, vertical
     * alignment inside the cell, uniform cell padding (twips).
     */
    formatTableCellText(index: number, opts: {
        row?: number;
        col?: number;
        bold?: boolean;
        italic?: boolean;
        sizePt?: number;
        fontFamily?: string;
        vAlign?: "top" | "center" | "bottom";
        paddingTwips?: number;
    }): Promise<this>;
    /** Merge cells: horizontal (row + startCol..endCol) or vertical (col + startRow..endRow). */
    mergeTableCells(index: number, opts: {
        direction: "horizontal";
        row: number;
        start: number;
        end: number;
    } | {
        direction: "vertical";
        col: number;
        start: number;
        end: number;
    }): Promise<this>;
    /** Delete column `col` (0-based) from the table at `index`. */
    deleteTableColumn(index: number, col: number): Promise<this>;
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
    setTableLayout(index: number, opts: {
        alignment?: "left" | "center" | "right";
        direction?: "rtl" | "ltr";
        headerRow?: boolean;
        headerFill?: string;
        borders?: boolean | {
            style?: string;
            sizePt?: number;
            color?: string;
            sides?: ("top" | "bottom" | "left" | "right" | "insideH" | "insideV")[];
        };
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
    }): Promise<this>;
    /** Split a merged cell back apart: horizontal (gridSpan) or vertical (vMerge chain). */
    splitTableCells(index: number, opts: {
        direction: "horizontal" | "vertical";
        row: number;
        col: number;
    }): Promise<this>;
    /** Move a row or a column from one position to another (0-based). */
    moveTableLine(index: number, opts: {
        kind: "row" | "column";
        from: number;
        to: number;
    }): Promise<this>;
    /**
     * Style one cell (row+col), a whole row (row only), or the header row
     * (neither): `fill` = 6-hex background, `textColor` = 6-hex font colour of
     * the existing text. Pass either or both.
     */
    shadeTable(index: number, opts: {
        row?: number;
        col?: number;
        fill?: string;
        textColor?: string;
    }): Promise<this>;
    /**
     * Apply styling grids: fills[r][c] = 6-hex background, textColors[r][c] =
     * 6-hex font colour for that cell's existing text; null/undefined = leave
     * as-is. Either grid may be omitted.
     */
    shadeTableCells(index: number, fills?: (string | null | undefined)[][] | null, textColors?: (string | null | undefined)[][] | null): Promise<this>;
    /** Append an image from bytes (or insert at `at`). Size in pixels @96dpi. */
    addImage(bytes: Buffer, opts?: {
        format?: string;
        width?: number;
        height?: number;
        at?: number;
    }): Promise<this>;
    /** Render Markdown and append the result (headings, paragraphs, lists, tables). */
    addMarkdown(blocks: BodyBlock[], at?: number): Promise<this>;
    /** Find-and-replace across the document body (e.g. fill `{{tokens}}`). */
    replaceText(find: string | RegExp, replace: string): Promise<this>;
    setPageSize(preset: "A4" | "USLetter" | "Legal" | "A3" | "A5", orientation?: "portrait" | "landscape"): Promise<this>;
    setMargins(margins: {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
    }): Promise<this>;
    setOrientation(orientation: "portrait" | "landscape"): Promise<this>;
    /** Page numbers in the footer (centered decimal by default). */
    addPageNumbers(opts?: {
        alignment?: "left" | "center" | "right";
        format?: "decimal" | "lowerRoman" | "upperRoman";
    }): Promise<this>;
    /** Roman front-matter / arabic body footers, split at `bodyStartIndex`. */
    frontMatterNumbering(bodyStartIndex: number): Promise<this>;
    /**
     * Set ONE document-wide page header (top of every page). Empty text removes it.
     * Replaces any existing header (no duplicates) and clears per-section headers.
     */
    setHeader(text: string): Promise<this>;
    /**
     * Set ONE document-wide page footer (text and/or page numbers). Empty text +
     * `pageNumbers:false` removes it. Replaces any existing footer and clears
     * per-section footers.
     */
    setFooter(opts?: FooterOptions): Promise<this>;
    /**
     * Make the block at `blockIndex` begin on a new page by inserting a section
     * break just before it (so e.g. each chapter starts a fresh page and can own
     * its header/footer). Returns `{ changed:false }` if the block is already the
     * first content. Block index ↔ paragraph index is handled internally.
     */
    startOnNewPage(blockIndex: number, breakType?: BreakType): Promise<{
        changed: boolean;
    }>;
    /**
     * Give the section CONTAINING the block at `blockIndex` its own running header
     * (call `startOnNewPage` on that heading first to make it its own section).
     * Empty text → a blank header for that section. Cleans up the section's
     * previous distinct header part.
     */
    setSectionHeader(blockIndex: number, text: string): Promise<SectionEditResult>;
    /**
     * Give the section CONTAINING the block at `blockIndex` its own footer (text
     * and/or page numbers). Cleans up the section's previous distinct footer part.
     */
    setSectionFooter(blockIndex: number, opts?: FooterOptions): Promise<SectionEditResult>;
    /**
     * Vertically align the content of the section CONTAINING the block at
     * `blockIndex`. "center" places a divider page's title in the middle of the
     * page instead of at the top. Call `startOnNewPage` on that block first so it
     * is its own section, or this aligns whatever section it falls in.
     */
    setSectionVerticalAlign(blockIndex: number, vAlign: "top" | "center" | "both" | "bottom"): Promise<SectionEditResult>;
    /**
     * Draw a page border around the section CONTAINING the block at `blockIndex`
     * (the `frame` divider family). Overwrites any border that section had.
     */
    setSectionPageBorders(blockIndex: number, opts: SectPrPageBorderOptions): Promise<SectionEditResult>;
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
    applySectionChrome(blockIndex: number, parts: {
        header?: ChromePart;
        footer?: ChromePart;
    }): Promise<{
        sectionIndex: number;
        warnings: string[];
    }>;
    /** Map a block index → its owning section (+ the full section list). */
    private resolveSection;
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
    sections(preloadedBlocks?: BodyBlock[]): Promise<SectionInfo[]>;
    /**
     * Content of the DEFAULT-type header/footer part behind `refs` (first/even
     * page-only refs are not the running chrome and are ignored), memoized by
     * relId in `cache` so sections sharing a part read it once per
     * {@link sections} call. null when no part resolves — including on any
     * read/parse failure, so chrome extraction can never throw.
     */
    private readHeaderFooterPart;
    /**
     * Fill in what {@link extractChromeDrawings} could not know from the part XML
     * alone: which media file each `r:embed` points at (resolved against the
     * part's OWN `_rels`), and the hex behind any theme-slot recolour.
     *
     * Mutates in place — the drawings were just built for this content object and
     * have no other reader yet.
     */
    private resolveChromeDrawings;
    /** Theme slot ("accent4", "dk1"…) → 6-hex from theme1.xml's `<a:clrScheme>`.
     *  `<a:sysClr>` carries the resolved value in `lastClr`. Slots the document
     *  does not define fall back to Word's built-in Office palette, which is what
     *  Word itself paints for a package with no theme part. null when the slot is
     *  not a colour-scheme name at all. */
    private resolveThemeColor;
    /** Best-effort delete of a header/footer part by its relationship id (cleanup). */
    private removeHeaderFooterByRel;
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
    private readChromeDrawings;
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
    private carryChromeDrawings;
    /** A generated, always-accurate structural map of the document. */
    describe(): Promise<DocMap>;
    /** The structural map rendered as human-readable Markdown. */
    toMarkdownMap(): Promise<string>;
    /** Write the document to a file. */
    save(outputPath: string): Promise<void>;
    /** Get the document as an in-memory buffer. */
    toBuffer(): Buffer;
}

/** A generated, always-accurate structural map of the document. */
export declare interface DocMap {
    title: string;
    wordCount: number;
    counts: {
        paragraphs: number;
        headings: number;
        tables: number;
        images: number;
        sections: number;
    };
    page: {
        width: number;
        height: number;
        orientation: string;
    };
    margins: {
        top: number;
        right: number;
        bottom: number;
        left: number;
    };
    hasHeader: boolean;
    hasFooter: boolean;
    rtl: boolean;
    outline: OutlineNode[];
}

export declare interface DoctorReport {
    /** No fatal findings remain (after repairs, when repairing). */
    ok: boolean;
    /** Zip entries examined. */
    checkedParts: number;
    findings: Finding[];
    /** Parts actually rewritten. Empty unless `fix` was requested. */
    repairedParts: string[];
}

/**
 * A document-wide formatting profile. Every field is optional; only the
 * provided ones are applied. Mirrors the kind of "norm profile" a university
 * imposes (font, size, line spacing, margins) and is applied uniformly across
 * the body via direct OOXML rewriting.
 */
export declare interface DocumentFormatting {
    /** Font family applied to ascii/hAnsi/cs of every `<w:rFonts>`. */
    font?: string;
    /** Font size in POINTS (written as half-points to `<w:sz>`/`<w:szCs>`). */
    fontSizePt?: number;
    /** Line spacing as a multiplier (1, 1.5, 2 → `<w:spacing w:line>` 240ths). */
    lineSpacing?: number;
    /** Page margins (cm). */
    margins?: DocumentMargins;
    /** Which side the binding is on — decides whether `binding`/`opposite` map to left/right. Default "left". */
    bindingSide?: "left" | "right";
}

export declare class DocumentManager {
    zip: default_2;
    constructor(zip: default_2);
    addHeaderReferenceToDocument(relId: string, type?: "default" | "first" | "even"): Promise<void>;
    addFooterReferenceToDocument(relId: string, type?: "default" | "first" | "even"): Promise<void>;
    private _addSectPrReference;
    /**
     * Returns the top-level body paragraphs from word/document.xml.
     * Only direct children of <w:body> are returned — paragraphs nested
     * inside table cells are not included, preventing duplication on
     * round-trips through saveChanges().
     */
    getParagraphs(): Promise<Paragraph[]>;
    /**
     * Returns the paragraph whose w14:paraId matches, or null.
     */
    getParagraphById(paraId: string): Promise<Paragraph | null>;
    /**
     * Returns the paragraph at the given zero-based index, or null.
     */
    getParagraphByIndex(index: number): Promise<Paragraph | null>;
    /**
     * Writes a full Paragraph[] back to word/document.xml, preserving <w:sectPr>.
     */
    saveChanges(paragraphs: Paragraph[]): Promise<void>;
    /**
     * Inserts a paragraph at the given index (appends if no index given).
     */
    insertParagraph(paragraph: Paragraph, index?: number): Promise<void>;
    /**
     * Removes the paragraph whose w14:paraId matches.
     */
    deleteParagraph(paraId: string): Promise<void>;
    /**
     * Replaces the paragraph whose w14:paraId matches with newParagraph.
     */
    replaceParagraph(paraId: string, newParagraph: Paragraph): Promise<void>;
    /**
     * Finds and replaces text across every paragraph in the document.
     * @param search   String or RegExp to search for.
     * @param replace  Replacement string.
     */
    findAndReplaceAll(search: string | RegExp, replace: string): Promise<void>;
    /**
     * Returns all tables in word/document.xml as Table instances.
     */
    getTables(): Promise<Table[]>;
    /**
     * Inserts a Table into the document body at the given index (appends if omitted).
     */
    insertTable(table: Table, index?: number): Promise<void>;
    /**
     * Returns the body's ordered editable blocks (paragraphs / tables /
     * drawings) in document order, EXCLUDING the trailing w:sectPr (which stays
     * in the document). Each block carries its exact original XML substring.
     */
    getBlocks(): Promise<BodyBlock[]>;
    /**
     * Plain text of the whole document body: every block's visible text, joined by
     * the given separator (default newline). Blank/whitespace-only blocks are
     * dropped. Useful for word counts, search indexing, or AI classification.
     */
    getPlainText(separator?: string): Promise<string>;
    /** Total word count of the document body (whitespace-delimited). */
    getWordCount(): Promise<number>;
    /**
     * Replaces the body's editable children with `blocks` (in order), preserving
     * the existing trailing w:sectPr exactly, and writes document.xml. Callers
     * pass only the editable blocks; the sectPr is re-appended automatically.
     */
    saveBlocks(blocks: BodyBlock[]): Promise<void>;
    /**
     * Replaces the run text of the editable block at `index` (must be a
     * paragraph). Rewrites ONLY that paragraph's XML substring; every other
     * block — tables, drawings, sectPr — keeps its exact original bytes.
     */
    editParagraphText(index: number, text: string): Promise<void>;
    /**
     * Inserts `block` at editable position `index` (appends before sectPr if the
     * index is out of range), keeping sectPr last. Use
     * OrderedBody.makeParagraphNode to build a paragraph block.
     */
    insertBlockAt(block: BodyBlock, index: number): Promise<void>;
    /** Removes the editable block at `index`. */
    deleteBlockAt(index: number): Promise<void>;
    /**
     * Set the paragraph style at block `index`, preserving order + runs.
     * "Normal" demotes to body (drops w:pStyle + w:outlineLvl). "Heading{n}"
     * sets the style AND w:outlineLvl = n-1 so the outline/TOC detects it.
     */
    setBlockStyle(index: number, styleId: string): Promise<void>;
    /** Set paragraph alignment at block `index`, preserving order + runs. */
    setBlockAlignment(index: number, alignment: "left" | "center" | "right" | "both"): Promise<void>;
    /**
     * Set paragraph text direction at block `index` via <w:bidi>, preserving order
     * + runs. "rtl" → <w:bidi/> (right-to-left); "ltr" → <w:bidi w:val="0"/> (an
     * explicit left-to-right marker, so it overrides an RTL style default).
     */
    setBlockDirection(index: number, direction: "rtl" | "ltr"): Promise<void>;
    /** Strip run-level formatting (bold/italic/font) at block `index`; keeps text. */
    clearBlockFormatting(index: number): Promise<void>;
    /**
     * Move the block at `from` to position `to`, preserving every block's bytes.
     * Pure array reorder over the ordered block model — the moved block and all
     * others keep their exact XML; only their sequence changes.
     */
    moveBlock(from: number, to: number): Promise<void>;
    private _writeBody;
    private _getBody;
    private _writeDoc;
}

/** Page margins in centimetres, expressed relative to the binding. */
export declare interface DocumentMargins {
    /** Top margin (cm). */
    top: number;
    /** Bottom margin (cm). */
    bottom: number;
    /** Inner margin on the binding side (cm) — the gutter side. */
    binding: number;
    /** Outer margin on the side opposite the binding (cm). */
    opposite: number;
}

/** The three-and-a-bit methods we need off the engine's zip (or adm-zip). */
export declare interface DocxZip {
    getEntries(): {
        entryName: string;
    }[];
    readAsText(entry: string): string;
    addFile(entry: string, content: Buffer): void;
}

/**
 * Represents a drawing element.
 * A drawing object (e.g., a chart or picture) located in a run.
 * @example
 * <w:r>
 * <w:drawing>...</w:drawing>
 * </w:r>
 */
declare interface Drawing {
    "w:drawing": {
        $: Record<string, unknown>;
    };
}

/** Where a floating picture sits. `undefined` = leave that axis alone. */
export declare interface DrawingLayout {
    horizontal?: "left" | "center" | "right";
    vertical?: "top" | "center" | "bottom";
    /** Frame of reference for BOTH axes. Default `page`. */
    relativeTo?: DrawingRelativeTo;
    /**
     * Default `topAndBottom`, not Word's own `square`. A thesis figure is close to
     * the full column width, and square wrapping round one leaves a two-word
     * sliver of text down its side; top-and-bottom reserves the band the picture
     * occupies and never does. On the page this default was built for — a lone
     * picture centred on an otherwise empty page — the two are identical.
     */
    wrap?: DrawingWrap;
    /** `false` returns the picture to the text flow (anchor → inline). */
    float?: boolean;
}

export declare interface DrawingLayoutResult {
    xml: string;
    changed: boolean;
    placement: DrawingPlacement;
    /** The drawing ships with an `mc:Fallback` VML twin that was left untouched —
     *  see {@link applyDrawingLayout}. Informational; nothing needs to act on it. */
    legacyTwin: boolean;
}

/** What a drawing's placement currently is. */
export declare interface DrawingPlacement {
    /** `wp:anchor` rather than `wp:inline`. */
    floating: boolean;
    horizontal: {
        relativeTo: string;
        align: string | null;
        offsetEmu: number | null;
    } | null;
    vertical: {
        relativeTo: string;
        align: string | null;
        offsetEmu: number | null;
    } | null;
    wrap: DrawingWrap | "tight" | "through" | "inline";
    /** Painted behind the text (Word's "Send Behind Text"). */
    behindDoc: boolean;
    widthEmu: number;
    heightEmu: number;
}

/**
 * Picture PLACEMENT surgery: where a `<w:drawing>` sits on the page.
 *
 * Word stores a picture in one of two containers, and which one it is decides
 * everything about where it can go:
 *
 *  - `<wp:inline>` — the picture IS a character in the text flow. It can only be
 *    where the line is; "put it in the middle of the page" is not expressible.
 *  - `<wp:anchor>` — the picture FLOATS, with a `wp:positionH`/`wp:positionV`
 *    pair naming a frame of reference (`page`, `margin`, …) and either an offset
 *    or a named alignment. This is Word's Layout ▸ Position dialog, and
 *    "Position in Middle Center relative to Page" is exactly
 *    `<wp:align>center</wp:align>` on both axes relative to `page`.
 *
 * So vertical placement is not a property you can set on a picture — it is a
 * container conversion. That is what {@link applyDrawingLayout} does, in both
 * directions, copying every other byte of the drawing through untouched (the
 * `a:graphic` subtree, which holds the image reference, is never even read).
 *
 * TWO RULES the schema enforces, both of which would produce a file Word refuses
 * to open if broken:
 *
 *  1. `CT_Anchor` is an ORDERED sequence — simplePos, positionH, positionV,
 *     extent, effectExtent?, <wrap>, docPr, cNvGraphicFramePr?, graphic. The
 *     wrap element in particular goes immediately BEFORE `wp:docPr`, not at the
 *     end where appending would put it.
 *  2. Geometry is read and written on the PREFIX before `<a:graphic` only. A
 *     shape's text box can contain a whole nested drawing, so a free search for
 *     `wp:positionV` can find a child's and move the wrong picture.
 */
/** Frames of reference this module will write. Word accepts more; these are the
 *  two that mean anything to a student ("on the page" / "inside the margins"). */
export declare type DrawingRelativeTo = "page" | "margin";

/** Text wrapping around a floating picture, in Word's UI vocabulary. */
export declare type DrawingWrap = "none" | "square" | "topAndBottom";

/**
 * Every outermost `<w:p>…</w:p>` inside a block's XML, in document order. A
 * `tables` target block is a whole `<w:tbl>` — its cell paragraphs are nested
 * several levels deep, and this walks in regardless of depth. Mirrors
 * `CaptionManager`'s private `eachParagraph`; duplicated rather than imported
 * because that helper is not exported and this task's scope is limited to the
 * one-line `CAPTION_PARA_STYLE` export (Decision B).
 *
 * Exported for `TextStyleReader`, which must walk a target's paragraphs exactly
 * the way the write does or it would report on text the write wouldn't touch.
 */
export declare function eachParagraphIn(xml: string): Array<{
    start: number;
    end: number;
    xml: string;
}>;

export declare const EMU_PER_CM = 360000;

export declare const EMU_PER_INCH = 914400;

/** EMU per pixel at 96 DPI (1 px = 9525 EMU). */
export declare const EMU_PER_PIXEL = 9525;

/** Convert EMU to whole pixels at 96 DPI. */
export declare const emuToPixels: (emu: number) => number;

export declare interface EndnoteEntry {
    id: number;
    text: string;
}

export declare class EndnoteManager {
    private zip;
    private rels;
    private contentTypes;
    constructor(zip: default_2);
    private readEndnotes;
    private writeEndnotes;
    private emptyEndnotesDoc;
    private normalizeArray;
    private nextId;
    private ensureRegistered;
    /**
     * Returns all user-defined endnotes (excludes separator entries).
     */
    getEndnotes(): Promise<EndnoteEntry[]>;
    /**
     * Adds a new endnote to endnotes.xml.
     * @param text   The endnote body text.
     * @returns      The endnote id and a run object to insert inline.
     */
    addEndnote(text: string): Promise<{
        id: number;
        run: Run_2;
    }>;
    /**
     * Removes an endnote by id.
     */
    removeEndnote(id: number): Promise<void>;
    /**
     * Returns a run object with an endnote reference mark for inline insertion.
     */
    createEndnoteRun(endnoteId: number): Run_2;
    private extractText;
}

/** How to create a paragraph style that does not exist yet. */
export declare interface EnsureStyleSpec {
    /** `w:name` — Word's display name. Maps the style onto a built-in when it matches. */
    name: string;
    /** `w:basedOn` target, omitted when absent. */
    basedOn?: string;
    /** Emit `w:default="1"` — used for `Normal`. */
    isDefault?: boolean;
}

/**
 * Expand `"headings"` into the six `headingN` targets, de-duplicate, and
 * validate every entry. Throws — rather than silently dropping an unrecognised
 * target — because a caller asking to restyle "bodytext" almost certainly
 * mistyped "body", and a silent no-op there would ship broken formatting to a
 * live thesis with no signal anything went wrong.
 */
export declare function expandTargets(targets: readonly string[]): TextStyleTarget[];

/** The `ObservedRunProps` keys reported as facets, in output order. */
declare const FACET_KEYS: readonly ["font", "fontCs", "sizePt", "sizeCsPt", "bold", "italic", "color"];

export declare type FacetKey = (typeof FACET_KEYS)[number];

/**
 * Represents a field, such as a page number or table of contents.
 * @example
 * <w:r>
 * <w:fldChar w:fldCharType="begin" />
 * </w:r>
 * <w:r>
 * <w:instrText xml:space="preserve">PAGEREF _Toc123456789 \h</w:instrText>
 * </w:r>
 * <w:r>
 * <w:fldChar w:fldCharType="end" />
 * </w:r>
 */
declare interface Field {
    "w:fldChar": {
        $: {
            "w:fldCharType": "begin" | "end";
        };
    };
    "w:instrText"?: string;
}

export declare interface Finding {
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

/** A cheap well-formedness check: tag balance with quote/comment awareness.
 *  Not a validator — its job is to catch a part we (or a previous writer) tore. */
export declare function firstXmlError(xml: string): string | null;

export declare class FooterManager {
    zip: default_2;
    rels: RelManager;
    contentTypes: ContentTypesManager;
    document: DocumentManager;
    footers: xmlFile[];
    constructor(zip: default_2);
    getFooterByName(name: string): xmlFile | false;
    getAllFooterFiles(zip: default_2): xmlFile[];
    private nextFooterPath;
    private buildFooterXml;
    private readDocObj;
    private writeDocObj;
    private getSectPr;
    private buildPageNumberRuns;
    /**
     * Adds a footer to the document.
     * @param registerInSectPr  When false, skips adding the <w:footerReference> to the main
     *                          w:sectPr. Useful when building multi-section documents.
     */
    addFooter(text: string, type?: FooterType, xml?: string, options?: {
        registerInSectPr?: boolean;
    }): Promise<{
        footerPath: string;
        relId: string;
        footerXml: string;
    }>;
    /**
     * Overwrites an existing footer file's content.
     */
    updateFooter(name: string, newXml: string): void;
    /**
     * Removes a footer: deletes zip entry, content-type, relationship and sectPr reference.
     */
    removeFooter(name: string): Promise<void>;
    /**
     * Insert a page number paragraph into the specified footer file.
     * Appends a new paragraph containing a PAGE field (optionally "X / Y").
     */
    insertPageNumber(footerPath: string, options?: PageNumberOptions): Promise<void>;
    /**
     * Remove all PAGE / NUMPAGES fields from the specified footer.
     */
    removePageNumbers(footerPath: string): Promise<void>;
    /**
     * Set page number format via w:pgNumType in w:sectPr.
     * Covers the full "Format Page Numbers" dialog:
     *   - number format (decimal, roman, letter…)
     *   - include chapter number + chapter heading style + separator
     *   - continue from previous section OR start at a specific number
     */
    formatPageNumbers(options: PageNumberFormatOptions): Promise<void>;
    /**
     * Enable or disable a different header/footer for the first page (w:titlePg in sectPr).
     */
    setDifferentFirstPage(enable: boolean): Promise<void>;
    /**
     * Enable or disable different odd/even page footers (w:evenAndOddHeaders in settings.xml).
     */
    setDifferentOddEvenPages(enable: boolean): Promise<void>;
    /**
     * Set the distance from the bottom of the page to the footer (in twips; 1 inch = 1440).
     * Default in Word is ~709 twips (0.49").
     */
    setFooterDistance(twips: number): Promise<void>;
    /**
     * Link (or unlink) the given footer to the previous section's footer.
     * Passing `true` removes the footer part so Word inherits from the previous section.
     */
    linkToPrevious(footerPath: string, linked: boolean): Promise<void>;
    private removeFooterRelAndReference;
    private removeFooterReferenceFromDocument;
}

/** Footer options (text and/or page numbers) for {@link Doc.setFooter}/{@link Doc.setSectionFooter}. */
export declare interface FooterOptions {
    text?: string;
    pageNumbers?: boolean;
    alignment?: "left" | "center" | "right";
    /** Text before the page number, e.g. "Page ". */
    prefix?: string;
    /** Render "current / total" (e.g. "3 / 40"). */
    includeTotalPages?: boolean;
}

declare type FooterType = "default" | "first" | "even";

export declare interface FootnoteEntry {
    id: number;
    text: string;
}

export declare class FootnoteManager {
    private zip;
    private rels;
    private contentTypes;
    constructor(zip: default_2);
    private readFootnotes;
    private writeFootnotes;
    private emptyFootnotesDoc;
    private normalizeArray;
    private nextId;
    private ensureRegistered;
    /**
     * Returns all user-defined footnotes (excludes separator entries).
     */
    getFootnotes(): Promise<FootnoteEntry[]>;
    /**
     * Adds a new footnote to footnotes.xml and returns a run object
     * that should be inserted into the paragraph at the desired position.
     *
     * @param text   The footnote body text.
     * @returns      The footnote id and a run object containing the reference mark.
     */
    addFootnote(text: string): Promise<{
        id: number;
        run: Run_2;
    }>;
    /**
     * Copy footnote ELEMENTS verbatim from another document's footnotes.xml string,
     * preserving their rich content (runs, rPr, hyperlinks). Each needed footnote is
     * appended under a fresh, collision-free id; returns source-id → new-id.
     *
     * Byte-faithful string injection (no xml2js round-trip) so footnote content can
     * never be mis-nested. NOTE: footnote-internal media (r:embed) / numbering are
     * NOT remapped here — footnotes embedding images are a rare deferred edge.
     */
    copyFootnotesVerbatim(sourceFootnotesXml: string | null, neededIds: Set<string>): Promise<Record<string, string>>;
    /**
     * Removes a footnote by id from footnotes.xml.
     * You must also manually remove the <w:footnoteReference> run from the document.
     */
    removeFootnote(id: number): Promise<void>;
    /**
     * Returns a run object (w:r) with a footnote reference mark for inline insertion.
     * Use after addFootnote() to get the inline anchor run.
     */
    createFootnoteRun(footnoteId: number): Run_2;
    private extractText;
}

/** Render a sequence number the way Word's `\*` format switch would. */
export declare function formatSeqNumber(n: number, fmt: string): string;

/**
 * Applies a {@link DocumentFormatting} profile uniformly across the document body
 * by direct OOXML rewriting — font, font size, line spacing and page margins.
 * A coarse but deterministic "normalise the whole document" pass, distinct from
 * per-paragraph or style-based formatting.
 */
export declare class FormattingManager {
    private zip;
    constructor(zip: ZipManager);
    /**
     * Apply the profile to `word/document.xml` in place. Returns which transforms
     * ran. Only fields present on `formatting` are attempted.
     */
    apply(formatting: DocumentFormatting): FormattingResult;
    /**
     * Pure transform: apply the profile to a `document.xml` string and return the
     * rewritten XML plus the applied/skipped breakdown. Useful when you already
     * hold the part bytes (e.g. a freshly merged buffer) and want to avoid a
     * second load.
     */
    static applyToXml(documentXml: string, formatting: DocumentFormatting): {
        xml: string;
        result: FormattingResult;
    };
}

/** Which transforms ran (`applied`) vs. failed (`skipped`). */
export declare interface FormattingResult {
    applied: string[];
    skipped: string[];
}

export declare interface FromGridOptions {
    /** Mark row 0 as a repeating header row (`w:tblHeader`). */
    headerRow?: boolean;
    /** Bold the header row's text (default false; only relevant with `headerRow`). */
    boldHeader?: boolean;
    /** Shade the header row with this hex fill (e.g. "D9D9D9"); omit for none. */
    headerFill?: string;
    /** Right-to-left table direction (`w:bidiVisual`). */
    rtl?: boolean;
    /** Table width as a percentage of page width (default 100). */
    widthPct?: number;
    /** Border set; defaults to a light grey single-line grid. Pass to override. */
    borders?: TableBorderOptions;
}

/** Options for {@link Mdocxengine.applyFrontMatterNumbering}. */
export declare interface FrontMatterNumberingOptions {
    /**
     * Paragraph index where the body begins. A `nextPage` section break is
     * inserted here, splitting the document into front-matter and body sections.
     */
    bodyStartParaIndex: number;
    /** Front-matter page-number format (default "lowerRoman" → i, ii, iii…). */
    frontMatterFormat?: PageNumberFormat;
    /** Body page-number format (default "decimal" → 1, 2, 3…). */
    bodyFormat?: PageNumberFormat;
    /** Footer alignment for both sections (default "center"). */
    alignment?: "left" | "center" | "right";
    /** Page number the body restarts at (default 1). */
    bodyStartAt?: number;
}

export declare class HeaderManager {
    zip: default_2;
    rels: RelManager;
    contentTypes: ContentTypesManager;
    document: DocumentManager;
    headers: xmlFile_2[];
    constructor(zip: default_2);
    getHeaderByName(name: string): xmlFile_2 | false;
    getAllheadersFiles(zip: default_2): xmlFile_2[];
    private nextHeaderPath;
    private buildHeaderXml;
    private readDocObj;
    private writeDocObj;
    private getSectPr;
    private buildPageNumberRuns;
    /**
     * Adds a header to the document.
     * @param registerInSectPr  When false, skips adding the <w:headerReference> to the main
     *                          w:sectPr. Useful when building multi-section documents where the
     *                          reference will be placed in an intermediate section-break paragraph.
     */
    addHeader(text: string, type?: HeaderType, xml?: string, options?: {
        registerInSectPr?: boolean;
    }): Promise<{
        headerPath: string;
        relId: string;
        headerXml: string;
    }>;
    /**
     * Overwrites an existing header file's content.
     */
    updateHeader(name: string, newXml: string): void;
    /**
     * Removes a header: deletes zip entry, content-type, relationship and sectPr reference.
     */
    removeHeader(name: string): Promise<void>;
    /**
     * Insert a page number paragraph into the specified header file.
     * Appends a new paragraph containing a PAGE field (optionally "X / Y").
     */
    insertPageNumber(headerPath: string, options?: PageNumberOptions_2): Promise<void>;
    /**
     * Remove all PAGE / NUMPAGES fields from the specified header.
     */
    removePageNumbers(headerPath: string): Promise<void>;
    /**
     * Set page number format via w:pgNumType in w:sectPr.
     * Covers the full "Format Page Numbers" dialog:
     *   - number format (decimal, roman, letter…)
     *   - include chapter number + chapter heading style + separator
     *   - continue from previous section OR start at a specific number
     */
    formatPageNumbers(options: PageNumberFormatOptions_2): Promise<void>;
    /**
     * Enable or disable a different header/footer for the first page (w:titlePg in sectPr).
     */
    setDifferentFirstPage(enable: boolean): Promise<void>;
    /**
     * Enable or disable different odd/even page headers (w:evenAndOddHeaders in settings.xml).
     */
    setDifferentOddEvenPages(enable: boolean): Promise<void>;
    /**
     * Set the distance from the top of the page to the header (in twips; 1 inch = 1440).
     * Default in Word is ~709 twips (0.49").
     */
    setHeaderDistance(twips: number): Promise<void>;
    /**
     * Link (or unlink) the given header to the previous section's header.
     * In OOXML this is controlled by the absence/presence of a headerReference for this section.
     * Passing `false` removes the reference so Word inherits from the previous section.
     */
    linkToPrevious(headerPath: string, linked: boolean): Promise<void>;
    private removeHeaderRelAndReference;
    private removeHeaderReferenceFromDocument;
}

declare type HeaderType = "default" | "first" | "even";

/**
 * Map a paragraph style id to a heading level 1–6 (0 = not a heading). Tolerant
 * of the style-id variants real .docx files carry: English `Heading N` / `Title`,
 * French `Titre N`, and separators ("Heading 1", "heading-1"). Returns 0 for any
 * non-heading style or null.
 */
export declare function headingLevelFromStyleId(styleId: string | null): number;

/** A regex → heading-level rule used by {@link Doc.inferOutline}. */
export declare interface HeadingPattern {
    re: RegExp;
    level: number;
}

/** Run-level formatting to force onto a heading style's `<w:rPr>`. Only the
 *  provided fields are touched. */
declare interface HeadingRunFormatting {
    /** Font size in POINTS (written as half-points to `<w:sz>`/`<w:szCs>`). */
    fontSizePt?: number;
    /** true adds `<w:b/><w:bCs/>`; false removes them. */
    bold?: boolean;
    /** Hex colour, with or without leading '#'. */
    color?: string;
}

/** Which `HeadingN` styles were found (and rewritten) vs. missing from styles.xml. */
declare interface HeadingStyleResult {
    updatedLevels: number[];
    missingLevels: number[];
}

/**
 * Interface for a hyperlink element that contains one or more runs.
 * It's a container for text that serves as a link to another part of the document or an external URL.
 * @example
 * <w:hyperlink w:anchor="_Toc106885920" w:history="1">
 * <w:r>
 * <w:t>Liste des organigrammes</w:t>
 * </w:r>
 * </w:hyperlink>
 */
declare interface Hyperlink {
    $: {
        "w:anchor"?: string;
        "w:history"?: string;
        "r:id"?: string;
    };
    "w:r"?: Run_2[];
    "w:fldChar"?: Field[];
}

export declare interface ImageEntry {
    name: string;
    path: string;
    buffer: Buffer;
}

/** An embedded image as plain data. */
export declare interface ImageInfo extends InlineImage {
    index: number;
}

/** Convert inches to twips */
export declare const inchesToTwips: (inches: number) => number;

/** Tuning for {@link Doc.inferOutline}; all fields have sensible Arabic-academic defaults. */
export declare interface InferOutlineOptions {
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
export declare interface InferredHeading {
    index: number;
    level: number;
    title: string;
    /** `styled` = real heading markup; `high`/`medium` = inferred from text/format. */
    confidence: "styled" | "high" | "medium";
    reason: string;
}

/** An inline image resolved from a paragraph block's `<w:drawing>`. */
export declare interface InlineImage {
    /** The `r:embed` relationship id. */
    relId: string;
    /** The resolved relationship target (e.g. "media/image1.png"). */
    target: string;
    /** The raw image bytes. */
    bytes: Buffer;
    /** File extension without dot (e.g. "png"). */
    extension: string;
    /** MIME type for the extension (e.g. "image/png"). */
    mime: string;
    /** Inline display width in EMU (0 if no `<wp:extent>`). */
    widthEmu: number;
    /** Inline display height in EMU (0 if no `<wp:extent>`). */
    heightEmu: number;
    /** Inline display width in pixels @96dpi. */
    widthPx: number;
    /** Inline display height in pixels @96dpi. */
    heightPx: number;
}

/**
 * Splice `fragment` in immediately after a paragraph's OPENING TAG.
 *
 * Use this — never a hand-rolled `xml.replace(/<w:p\b[^>]*>/, …)` — whenever
 * anything is inserted at the head of a paragraph. `[^>]*` swallows the slash of
 * a SELF-CLOSING `<w:p w:rsidR="00A1"/>`, which is exactly what Word writes for
 * an empty paragraph, so the naive replace appends after a tag that has ALREADY
 * CLOSED: the fragment lands as the paragraph's SIBLING — a direct child of
 * `<w:body>`. A body holding a bare `<w:pPr>` is not openable in Word
 * (`body.illegal-child`, fatal), and the app draws it as an unknown-block chip.
 * {@link splitParagraph} normalises `<w:p/>` to paired form first, so the
 * fragment always lands INSIDE the paragraph.
 *
 * Anchored to the opening tag, so a paragraph nested in a text box
 * (`w:txbxContent`) can never be the one that gets written to.
 *
 * Returns `xml` unchanged when it does not begin with a `<w:p>` element, or when
 * `fragment` is empty.
 */
export declare function insertAfterParagraphOpen(xml: string, fragment: string): string;

export declare interface InsertLineOptions {
    color?: string;
    arrowEnd?: boolean;
    arrowStart?: boolean;
    lineWidthPt?: number;
    paragraphIndex?: number;
    label?: string;
    labelOffsetX?: number;
    labelOffsetY?: number;
}

export declare interface InsertShapeOptions {
    position?: ShapePosition;
    size?: ShapeSize;
    paragraphIndex?: number;
    fillColor?: string;
    borderColor?: string;
    floating?: boolean;
    name?: string;
    text?: string;
    textColor?: string;
    fontSize?: number;
    bold?: boolean;
    textAlign?: "l" | "ctr" | "r";
}

/**
 * Check a .docx package and, with `fix`, repair what can be repaired. The zip is
 * mutated in place; the caller decides whether to write the bytes back.
 */
export declare function inspectDocx(zip: DocxZip, opts?: InspectOptions): DoctorReport;

export declare interface InspectOptions {
    /** Apply every repair this doctor considers safe. */
    fix?: boolean;
    /** Also apply repairs that DELETE something (an unreferenced dangling
     *  relationship). Never on the automatic save path. */
    aggressive?: boolean;
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
export declare function isEditableBlock(b: BodyBlock): boolean;

/** True when `xml` is a `<w:p>` element this module can rewrite. */
export declare function isParagraphXml(xml: string): boolean;

export declare interface LineNumberingOptions {
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: "newPage" | "newSection" | "continuous";
}

/** Build a drawing (inline image) paragraph `BodyBlock`. */
export declare function makeDrawingParagraphNode(relId: string, widthEmu: number, heightEmu: number, shapeId: number, name: string): BodyBlock;

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
export declare function makeDrawingParagraphXml(relId: string, widthEmu: number, heightEmu: number, shapeId: number, name: string): string;

/**
 * Build a paragraph `BodyBlock` (kept under the legacy export name
 * `makeParagraphNode` for API stability). Now string-based.
 */
export declare function makeParagraphNode(text: string, styleId?: string, rtl?: boolean): BodyBlock;

/** {@link makeParagraphXmlLike} as a `BodyBlock`. */
export declare function makeParagraphNodeLike(templateXml: string, text: string): BodyBlock;

/**
 * Build a `<w:p>` paragraph XML string from plain text + optional style/RTL:
 * `<w:p>(<w:pPr>(<w:pStyle w:val="ID"/>)(<w:bidi/>)</w:pPr>)?<w:r><w:t xml:space="preserve">ESCAPED</w:t></w:r></w:p>`
 */
export declare function makeParagraphXml(text: string, styleId?: string, rtl?: boolean): string;

export declare function makeParagraphXmlLike(templateXml: string, text: string): string;

/** Build a fully-formatted paragraph `BodyBlock` (string-based, byte-safe). */
export declare function makeStyledParagraphNode(text: string, opts?: StyledParagraphOptions): BodyBlock;

/**
 * Build a fully-formatted single-run `<w:p>` XML STRING (byte-safe for the
 * OrderedBody block path — unlike `Paragraph.toXml()`, which wraps in `<root>`).
 * `<w:pPr>`/`<w:rPr>` children are emitted in canonical OOXML order. Empty `text`
 * yields a property-only spacer paragraph.
 */
export declare function makeStyledParagraphXml(text: string, opts?: StyledParagraphOptions): string;

/** Build a table `BodyBlock` from a grid of cell texts. */
export declare function makeTableNode(rows: string[][], opts?: {
    headerRow?: boolean;
    rtl?: boolean;
}): BodyBlock;

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
export declare function makeTableXml(rows: string[][], opts?: {
    headerRow?: boolean;
    rtl?: boolean;
}): string;

export declare const MARGIN_PRESETS: Record<MarginPreset, PageMargins>;

export declare type MarginPreset = "normal" | "narrow" | "moderate" | "wide" | "mirrored";

export declare type MarkdownAlign = "left" | "center" | "right" | "both";

export declare type MarkdownBlock = {
    kind: "heading";
    level: 1 | 2 | 3 | 4 | 5 | 6;
    text: string;
} | {
    kind: "paragraph";
    text: string;
} | {
    kind: "list";
    ordered: boolean;
    items: string[];
} | {
    kind: "table";
    header: string[];
    rows: string[][];
} | {
    kind: "quote";
    text: string;
} | {
    kind: "code";
    lang: string;
    text: string;
};

export declare interface MarkdownRenderCtx {
    /** Body alignment for paragraphs/lists/headings. */
    align: MarkdownAlign;
    /** Right-to-left tables (and a hint for caller hooks). */
    rtl: boolean;
    /**
     * Word heading level a top-level markdown `#` maps to (default 2). Use 3 to
     * reserve Heading1/Heading2 for an outer structure (e.g. Partie/Chapitre).
     */
    headingBase?: 2 | 3;
}

/** Word's own ceiling (its font-size box refuses more). */
export declare const MAX_FONT_SIZE_PT = 1638;

export declare class Mdocxengine {
    zip: ZipManager;
    rels: RelManager;
    contentTypes: ContentTypesManager;
    document: DocumentManager;
    footer: FooterManager;
    header: HeaderManager;
    rootRels: RootRelManager;
    styles: StylesManager;
    numbering: NumberingManager;
    metadata: MetadataManager;
    media: MediaManager;
    footnotes: FootnoteManager;
    endnotes: EndnoteManager;
    toc: TableOfContentsManager;
    crossRef: CrossReferenceManager;
    citations: CitationManager;
    pageLayout: PageLayoutManager;
    captions: CaptionManager;
    comments: CommentsManager;
    trackedChanges: TrackedChangesManager;
    sections: SectionManager;
    shapes: ShapeManager;
    merge: MergeManager;
    formatting: FormattingManager;
    private constructor();
    static loadFromFile(path: string): Promise<Mdocxengine>;
    static loadFromBuffer(buffer: Buffer): Promise<Mdocxengine>;
    saveToFile(outputPath: string): Promise<void>;
    /**
     * Apply the canonical thesis/report footer scheme: front-matter pages numbered
     * in one format (default lowercase roman — i, ii, iii…) and the body restarted
     * in another (default decimal — 1, 2, 3…), each centered in the footer.
     *
     * Splits the document into two sections with a `nextPage` break at
     * `bodyStartParaIndex`, attaches a distinct footer (with a page-number field)
     * to each section, and restarts the body numbering. A composition over the
     * PageLayout / Footer / Section managers — no new OOXML logic.
     */
    applyFrontMatterNumbering(options: FrontMatterNumberingOptions): Promise<void>;
    /**
     * Replace the image embedded in the figure block at editable `index` with new
     * bytes — the engine-level primitive behind the app's crop / rotate / replace /
     * remove-background tools. Composes MediaManager (bytes, rels, content-types)
     * and DocumentManager (block XML) so callers never touch OOXML.
     *
     * Byte swap:
     *  • Same extension as the current image → overwrite the media part in place
     *    (relationship + inline reference untouched).
     *  • Different extension (e.g. jpeg → transparent PNG for background removal) →
     *    register a NEW media part + content-type + relationship and repoint the
     *    block's `r:embed` to it. The superseded part is left orphaned (Word ignores
     *    unreferenced media); we don't delete it because another block may share it.
     *
     * Optional `opts.widthPx`/`heightPx` are the NEW intrinsic pixel size. When the
     * aspect ratio changed (crop, rotate ±90°, replace) the drawing's display box is
     * rescaled to keep the current display WIDTH and apply the new ratio, so the
     * image isn't stretched. Omit them (or pass the same ratio) to leave the box as-is.
     *
     * @param index   Editable block index (matches DocumentManager.getBlocks order).
     * @param buffer  New image bytes.
     * @param ext     New image extension without dot ("png", "jpeg", "jpg", "gif").
     */
    replaceImageAtBlock(index: number, buffer: Buffer, ext: string, opts?: {
        widthPx?: number;
        heightPx?: number;
    }): Promise<void>;
}

export declare class MediaManager {
    private zip;
    private rels;
    private contentTypes;
    constructor(zip: default_2);
    /**
     * Returns all image files found in word/media/.
     */
    listImages(): ImageEntry[];
    /**
     * Returns the buffer for a specific image by filename (e.g. "image1.png"), or null.
     */
    extractImage(name: string): Buffer | null;
    /**
     * Resolve a relationship id (e.g. an image's `r:embed`) to its bytes via the
     * document relationships. Returns null if the rel or the target part is missing.
     */
    getImageByRelId(relId: string): Promise<Buffer | null>;
    /**
     * Read a media part by its relationship `Target` (e.g. "media/image1.png").
     * Targets are relative to `word/` unless already absolute/prefixed.
     */
    private readByTarget;
    /**
     * Extract the inline image embedded in a paragraph/run XML string. Handles both
     * forms Word writes:
     *   • DrawingML — `<w:drawing>` with `<a:blip r:embed="…">`, sized by `<wp:extent>`
     *   • VML       — `<w:pict>`/`<w:object>` with `<v:imagedata r:id="…">`, sized by
     *     the `<v:shape style="width:…pt;height:…pt">` CSS
     * The VML form is what an embedded OLE object (a legacy Equation.3 / MathType
     * equation, an Origin graph) uses for its on-page preview: the OLE binary itself
     * is unreadable outside Word, but Word always stores a rendered bitmap beside it,
     * and that bitmap is the only way to show the object anywhere else.
     *
     * Resolves the relationship to the media bytes. Returns null when the block has
     * no inline image or it can't be resolved.
     */
    extractInlineImage(blockXml: string): Promise<InlineImage | null>;
    /**
     * Inserts a new image into the document.
     * @param imageBuffer  Raw image bytes.
     * @param extension    File extension without dot (e.g. "png", "jpg").
     * @returns The image path and the generated relationship ID.
     */
    insertImage(imageBuffer: Buffer, extension: string): Promise<{
        imagePath: string;
        relId: string;
    }>;
    /**
     * Add image bytes to word/media (+ register the content-type Default) WITHOUT
     * creating any relationship. Use when the relationship must live in a specific
     * part's own _rels (e.g. a header/footer part) rather than the document rels —
     * see {@link addImageToPartRels}. Returns the media part path (e.g.
     * "word/media/image2.png").
     */
    addImagePart(imageBuffer: Buffer, extension: string): Promise<{
        imagePath: string;
    }>;
    /**
     * Embed an image into a SPECIFIC part's own relationships and return that
     * part-local `r:embed` id. A header/footer part resolves `r:embed` against its
     * own `word/_rels/<part>.rels`, NOT the document rels — so an id from
     * {@link insertImage} would not resolve inside a header. This adds the bytes +
     * content-type, then a part-local image relationship (creating the part's .rels
     * if absent), and hands back the id to drop into the part's `<a:blip r:embed>`.
     * @param partPath  Full part path, e.g. "word/header1.xml".
     */
    addImageToPartRels(partPath: string, imageBuffer: Buffer, extension: string): Promise<string>;
    /**
     * Replaces an existing image's bytes in-place (same filename, same relId).
     * @param name         Filename like "image1.png" (without path prefix).
     * @param newBuffer    New image bytes.
     */
    replaceImage(name: string, newBuffer: Buffer): void;
    /**
     * Deletes an image from the zip (does not remove the relationship or inline reference).
     */
    deleteImage(name: string): void;
}

export declare class MergeManager {
    private zip;
    private document;
    private media;
    private footnotes;
    private numbering;
    private rels;
    private contentTypes;
    constructor(zip: default_2);
    /**
     * Copy `sourceBuffer`'s body into this document, fully remapped, appended after
     * the existing body (before the trailing sectPr, handled by saveBlocks).
     */
    appendDocument(sourceBuffer: Buffer, opts?: AppendOptions): Promise<void>;
    /**
     * Build the id maps from a READ-ONLY concatenation of source blocks, then apply
     * them to EACH block's xml independently (kind/tag preserved verbatim).
     */
    private remapBlocks;
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
     * Everything else in the copied `sectPr` — page size, margins, break type — is
     * preserved.
     */
    private stripSectionChrome;
    /** Read the source document's relationships as { rId: { type, target, targetMode } }. */
    private readSourceRels;
    /** Copy each image referenced by the blocks; return source-rId → new-rId. */
    private buildMediaMap;
    /**
     * Copy external hyperlink relationships referenced by `<w:hyperlink r:id>` into
     * the target (preserving TargetMode="External"); return source-rId → new-rId.
     */
    private buildHyperlinkMap;
    /** Copy each chart referenced by the blocks, with its whole part closure;
     *  return source-rId → new-rId. */
    private buildChartMap;
    /**
     * Copy `srcPath` and every part it transitively references into this package
     * under a free name, carrying content types across. Returns the new part path.
     *
     * `.rels` targets are rewritten to the copied parts' new names, so two merged
     * documents that both ship `charts/chart1.xml` cannot collide.
     */
    private copyPartClosure;
    /** `word/charts/chart1.xml` → the same name, or `chart2.xml`, … until free. */
    private freePartPath;
    /** Remap r:id ONLY inside <c:chart .../> (never other r:id). Any chart whose
     *  part could not be carried has its whole drawing removed — a reference to a
     *  part that is not there is exactly what makes Word reject the file. */
    private applyChartMap;
    /** `/word/charts/chart1.xml` → content type, from a package's [Content_Types].xml. */
    private readContentTypeOverrides;
    /** extension (lower-case, no dot) → content type. */
    private readContentTypeDefaults;
    /** Copy each referenced source footnote verbatim; return source-id → new-id. */
    private buildFootnoteMap;
    /**
     * Copy the source's numbering definitions referenced by the blocks (abstractNum
     * + num) into the target with fresh, collision-free ids; return source-numId →
     * new-numId. Without this, two parts that both use `numId="1"` would share one
     * list and renumber wrongly.
     */
    private buildNumberingMap;
    /** Remap w:val inside <w:numId .../> elements only. */
    private applyNumIdMap;
    /** Replace attr="old" → attr="new" for each given attribute name. */
    private applyAttrMap;
    /** Remap r:id ONLY inside <w:hyperlink ...> open tags (never other r:id). */
    private applyHyperlinkMap;
    /** Remap w:id ONLY inside <w:footnoteReference .../> (never other w:id). */
    private applyFootnoteRefMap;
    /** Retarget <w:pStyle w:val="X"/> by name. */
    private applyStyleMap;
}

/**
 * Merge `props` into `w:rPr` inner XML: strip the tags each named property owns,
 * append the replacements, then canonicalise into `CT_RPr` order. Properties not
 * named on `props` survive untouched — including `w:rtl`, `w:cs`, `w:highlight`
 * and `w:u`.
 */
export declare function mergeRunProps(rPrInner: string, props: RunProps): string;

export declare class MetadataManager {
    private zip;
    constructor(zip: default_2);
    /**
     * Reads core document properties.
     */
    getCoreProperties(): Promise<CoreProperties>;
    /**
     * Updates core document properties. Merges with existing values.
     */
    setCoreProperties(props: CoreProperties): Promise<void>;
    /**
     * Reads application properties.
     */
    getAppProperties(): Promise<AppProperties>;
    /**
     * Updates application properties. Merges with existing values.
     */
    setAppProperties(props: AppProperties): Promise<void>;
}

/** Word's own floor: one half-point. */
export declare const MIN_FONT_SIZE_PT = 0.5;

/**
 * Scan a full `document.xml` for the highest existing drawing id (`wp:docPr @id`
 * and `pic:cNvPr @id`) and return max+1 (minimum 1). Use it to assign a
 * document-unique id to a newly inserted drawing.
 */
export declare function nextDrawingId(documentXml: string): number;

/** The tag name of a block (compat shim — accepts a BodyBlock). */
export declare function nodeTag(block: BodyBlock | string): string;

export declare interface NumberingDefinition {
    abstractNumId: string;
    numId: string;
}

export declare class NumberingManager {
    private zip;
    constructor(zip: default_2);
    private readNumbering;
    private writeNumbering;
    /**
     * Returns all concrete numbering definitions (numId → abstractNumId).
     */
    getNumberingDefinitions(): Promise<NumberingDefinition[]>;
    /**
     * Applies numbering to a paragraph by setting <w:numPr> on its <w:pPr>.
     * @param paragraph  The Paragraph instance to modify.
     * @param numId      The concrete numId from numbering.xml.
     * @param ilvl       Indent level (0-based). Defaults to 0.
     */
    applyNumbering(paragraph: Paragraph, numId: string, ilvl?: number): void;
    /**
     * Removes numbering from a paragraph.
     */
    removeNumbering(paragraph: Paragraph): void;
    /**
     * Adds a new concrete <w:num> definition referencing an abstract numbering ID.
     * @param numId         The new numId to use.
     * @param abstractNumId The abstractNumId to reference.
     */
    addNumberingDefinition(numId: string, abstractNumId: string): Promise<void>;
    /** Raw `<w:abstractNum>` + `<w:num>` arrays from this document's numbering.xml. */
    getRawDefinitions(): Promise<RawNumbering>;
    /** Highest existing abstractNumId / numId (for collision-free allocation). */
    maxIds(): Promise<{
        absMax: number;
        numMax: number;
    }>;
    /**
     * Append already-id-rewritten `<w:abstractNum>` + `<w:num>` nodes. Keeps the
     * schema order (all abstractNum before all num) because they are separate keys.
     */
    appendRawDefinitions(abstractNums: any[], nums: any[]): Promise<void>;
}

/**
 * Run properties as OBSERVED in a document — the read counterpart of
 * {@link RunProps}, which describes what to WRITE.
 *
 * Every field is a tri-state: a value, `false` for a toggle explicitly turned
 * off (`<w:b w:val="0"/>`, which must beat an inherited bold rather than be
 * mistaken for "unspecified"), or `undefined` for "this rung of the cascade
 * says nothing", which is what lets a lower rung show through.
 */
export declare interface ObservedRunProps {
    /** Latin font — `w:rFonts/@w:ascii`, falling back to `@w:hAnsi`. */
    font?: string;
    /** Complex-script font — `w:rFonts/@w:cs`. What ARABIC text renders in. */
    fontCs?: string;
    /** Latin size in points (`w:sz`, stored as half-points). */
    sizePt?: number;
    /** Complex-script size in points (`w:szCs`). The size ARABIC text renders at. */
    sizeCsPt?: number;
    bold?: boolean;
    italic?: boolean;
    /** 6-digit upper-case hex, or the literal "auto". */
    color?: string;
}

export declare type Orientation = "portrait" | "landscape";

/** A node in the heading outline tree. */
export declare interface OutlineNode {
    index: number;
    level: number;
    title: string;
    children: OutlineNode[];
}

export declare const PAGE_SIZES: Record<PageSizePreset, {
    w: number;
    h: number;
}>;

export declare class PageLayoutManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private getSectPr;
    private twip;
    /**
     * Get the current page size in twips.
     */
    getPageSize(): Promise<PageSize>;
    /**
     * Set page size from a preset name.
     */
    setPageSizePreset(preset: PageSizePreset, orientation?: Orientation): Promise<void>;
    /**
     * Set a custom page size (values in twips).
     */
    setPageSize(size: PageSize): Promise<void>;
    /**
     * Toggle orientation while preserving current page size dimensions.
     */
    setOrientation(orientation: Orientation): Promise<void>;
    /**
     * Get current orientation.
     */
    getOrientation(): Promise<Orientation>;
    /**
     * Get the current page margins in twips.
     */
    getMargins(): Promise<PageMargins>;
    /**
     * Apply a named margin preset.
     */
    setMarginPreset(preset: MarginPreset): Promise<void>;
    /**
     * Set custom margins (values in twips). Pass partial to merge with existing.
     */
    setMargins(margins: Partial<PageMargins>): Promise<void>;
    /**
     * Set margins using inches (convenience wrapper).
     */
    setMarginsInInches(margins: {
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
        gutter?: number;
    }): Promise<void>;
    /**
     * Get current column configuration.
     */
    getColumns(): Promise<ColumnOptions>;
    /**
     * Set column layout.
     */
    setColumns(options: ColumnOptions): Promise<void>;
    /**
     * Insert a page or section break before the paragraph at `paragraphIndex`.
     * Defaults to appending at end of document.
     *
     * - "nextPage" / "evenPage" / "oddPage" / "continuous" = section breaks
     * - "nextColumn" = column break within a multi-column section
     */
    insertBreak(type: SectionBreakType, paragraphIndex?: number): Promise<void>;
    /**
     * Get current line numbering settings (null if not enabled).
     */
    getLineNumbering(): Promise<LineNumberingOptions | null>;
    /**
     * Enable line numbering.
     */
    setLineNumbering(options?: LineNumberingOptions): Promise<void>;
    /**
     * Remove line numbering.
     */
    removeLineNumbering(): Promise<void>;
}

export declare interface PageMargins {
    top: number;
    bottom: number;
    left: number;
    right: number;
    gutter?: number;
    header?: number;
    footer?: number;
}

export declare type PageNumberFormat = "decimal" | "upperRoman" | "lowerRoman" | "upperLetter" | "lowerLetter" | "arabicAlpha" | "arabicAbjad";

declare type PageNumberFormat_2 = "decimal" | "upperRoman" | "lowerRoman" | "upperLetter" | "lowerLetter" | "arabicAlpha" | "arabicAbjad";

declare interface PageNumberFormatOptions {
    /** Numbering style written to w:pgNumType. Default: "decimal". */
    format?: PageNumberFormat;
    /**
     * Include the chapter number before the page number (e.g. "1-1", "1-A").
     * Requires a heading style to be set via chapterStyle.
     */
    includeChapterNumber?: boolean;
    /**
     * Heading level whose numbering is used as the chapter prefix.
     * 1 = Heading 1, 2 = Heading 2, … Default: 1.
     */
    chapterStyle?: number;
    /**
     * Separator between chapter and page number.
     * hyphen → "-", period → ".", colon → ":", emDash → "—", enDash → "–".
     * Default: "hyphen".
     */
    chapterSeparator?: ChapterSeparator_2;
    /**
     * When true, page numbering continues from the previous section (no w:start attribute).
     * When false/omitted, startAt is used if provided.
     */
    continueFromPreviousSection?: boolean;
    /** Restart page numbering at this value. Ignored when continueFromPreviousSection is true. */
    startAt?: number;
}

declare interface PageNumberFormatOptions_2 {
    /** Numbering style written to w:pgNumType. Default: "decimal". */
    format?: PageNumberFormat_2;
    /**
     * Include the chapter number before the page number (e.g. "1-1", "1-A").
     * Requires a heading style to be set via chapterStyle.
     */
    includeChapterNumber?: boolean;
    /**
     * Heading level whose numbering is used as the chapter prefix.
     * 1 = Heading 1, 2 = Heading 2, … Default: 1.
     */
    chapterStyle?: number;
    /**
     * Separator between chapter and page number.
     * hyphen → "-", period → ".", colon → ":", emDash → "—", enDash → "–".
     * Default: "hyphen".
     */
    chapterSeparator?: ChapterSeparator_3;
    /**
     * When true, page numbering continues from the previous section (no w:start attribute).
     * When false/omitted, startAt is used if provided.
     */
    continueFromPreviousSection?: boolean;
    /** Restart page numbering at this value. Ignored when continueFromPreviousSection is true. */
    startAt?: number;
}

declare interface PageNumberOptions {
    /** Paragraph alignment for the page-number line. Default: "center". */
    alignment?: "left" | "center" | "right";
    /** Numbering style. Default: "decimal". */
    format?: PageNumberFormat;
    /** When true, renders "X / Y" (current page / total pages). */
    includeTotalPages?: boolean;
    /** Prefix text before the page number, e.g. "Page ". */
    prefix?: string;
}

declare interface PageNumberOptions_2 {
    /** Paragraph alignment for the page-number line. Default: "center". */
    alignment?: "left" | "center" | "right";
    /** Numbering style. Default: "decimal". */
    format?: PageNumberFormat_2;
    /** When true, renders "X / Y" (current page / total pages). */
    includeTotalPages?: boolean;
    /** Prefix text before the page number, e.g. "Page ". */
    prefix?: string;
}

export declare interface PageSize {
    width: number;
    height: number;
    orientation?: Orientation;
}

export declare type PageSizePreset = "USLetter" | "USLegal" | "A3" | "A4" | "A5" | "B5" | "JISB5" | "Tabloid" | "TabloidOversize" | "EnvelopeDL" | "Envelope10";

/**
 * A class representing a single paragraph from a WordprocessingML document.
 * It provides methods to easily get and modify the paragraph's text content.
 */
export declare class Paragraph {
    paragraph: Paragraph_2;
    /**
     * Constructs a new Paragraph instance.
     * @param paragraph The parsed JSON representation of the paragraph's XML.
     */
    constructor(paragraph: Paragraph_2);
    /**
     * Extracts plain text from a Word paragraph XML string and checks if it contains any text.
     * - If no `<w:t>` tags exist or they're empty, it returns `hasText = false` and `text = ""`.
     * - Otherwise, it returns the combined text and `hasText = true`.
     *
     * @param paragraphXml - Optional raw XML string of the <w:p> element. If not provided, it uses this.toXml().
     * @returns An object containing the extracted text and a boolean flag.
     */
    getPlainText(paragraphXml?: string): Promise<{
        hasText: boolean;
        text: string;
    }>;
    /**
     * Safely extracts visible text from a Word paragraph (<w:p>),
     * handling nested structures like hyperlinks, bookmarks, and tabs.
     *
     * @param paragraphXml Optional XML string of the paragraph.
     * @returns Object with a boolean (hasText) and the combined text string.
     */
    getPlainTextSafe(): Promise<{
        hasText: boolean;
        text: string;
    }>;
    /**
     * Recursively extracts text from any node including nested structures.
     */
    private extractTextFromNode;
    /**
     * Appends new text to the paragraph without removing existing runs.
     * @param text - The text to append.
     */
    appendText(text: string): void;
    /**
     * Recursively replaces text inside a paragraph without removing hyperlinks or nested structures.
     *
     * @param searchText - Text to search for. If null, replace all text.
     * @param replaceText - Text to replace with.
     */
    replaceText(searchText: string | null, replaceText: string): void;
    /**
     * Sets the paragraph alignment.
     * @param alignment - One of "left" | "center" | "right" | "both"
     */
    setAlignment(alignment: "left" | "center" | "right" | "both"): void;
    /**
     * Ensures <w:pPr> exists AND is the first child of <w:p>. OOXML (CT_P) requires
     * paragraph properties to precede all run content; xml2js's Builder serializes
     * keys in insertion order, so simply assigning this.paragraph["w:pPr"] = {} on a
     * paragraph that already has runs would emit <w:pPr> after <w:r> and produce a
     * file Word flags as corrupt. Rebuild the object with w:pPr first when creating it.
     */
    private ensurePPr;
    /**
     * Gets the current alignment of the paragraph.
     */
    getAlignment(): string | null;
    /**
     * Returns the total number of words in the paragraph.
     */
    getWordCount(): Promise<number>;
    /**
     * Applies a style to the entire paragraph.
     * @param styleId - The Word style ID (e.g., "Heading1", "Normal").
     */
    applyStyle(styleId: string): void;
    /**
     * Removes all formatting (bold, italic, etc.) from runs but keeps text.
     */
    removeFormatting(): void;
    /**
     * Creates a deep clone of the paragraph object.
     */
    clone(): Paragraph;
    /**
     * Merges another paragraph's runs into this one.
     * @param otherParagraph - The paragraph to merge into this one.
     */
    mergeWith(otherParagraph: Paragraph): void;
    /**
     * Splits the paragraph into two at the specified character index.
     * @param index - Character position to split at.
     * @returns A tuple: [firstPart, secondPart]
     */
    splitAt(index: number): Promise<[Paragraph, Paragraph]>;
    /**
     * Adds a new hyperlink to the paragraph.
     * @param url - The URL of the hyperlink.
     * @param displayText - The visible text for the hyperlink.
     */
    addHyperlink(url: string, displayText: string, rsidRPr: string): void;
    /**
     * Extracts all hyperlinks inside the paragraph.
     * Returns array of { displayText, url }.
     */
    getHyperlinks(): Promise<{
        displayText: string;
        url: string;
    }[]>;
    /**
     * Removes all hyperlinks but keeps the visible text.
     */
    removeHyperlinks(): void;
    /**
     * Build a single-run, formatted paragraph from plain text + options.
     *
     * The `<w:pPr>` children are emitted in canonical CT_PPr order
     * (pStyle → jc → outlineLvl) so the result is schema-valid regardless of which
     * options are passed. Empty `text` yields a property-only (spacer) paragraph
     * with no run.
     *
     * Replaces the repeated `new Paragraph({ $: {}, "w:pPr": {}, "w:r": [] })`
     * boilerplate used across document builders.
     */
    static make(text: string, opts?: ParagraphOptions): Paragraph;
    /**
     * Creates a Paragraph instance from an XML string.
     * @param xmlString The XML string of the paragraph.
     * @returns A Promise that resolves with the new Paragraph instance.
     */
    static createFromXml(xmlString: string): Promise<Paragraph>;
    /**
     * Returns all highlighted runs in the current paragraph, optionally filtered by fill and value.
     *
     * @param {string} [fill] - Optional. Filter by highlight fill color (e.g., "FFFF00").
     * @param {string} [value] - Optional. Filter by shading value (e.g., "clear").
     * @returns {Run[] | false} - Array of highlighted runs or false if none found.
     */
    getHighlightedRuns(fill?: string, value?: string): Run_2[] | false;
    /**
     * Returns true if any run in the paragraph is highlighted.
     */
    hasHighlight(): boolean;
    /**

     * Modifies the text content of the paragraph.
     * Clears all existing runs and child elements, then replaces them
     * with a single new run containing the provided text.
     * @param newText - The new string to set as the paragraph's content.

     */
    modifyText(newText: string): void;
    generateUniqueParaId(zip: default_2): string;
    /**
     * Detects the primary language of the paragraph (if available).
     */
    detectLanguage(): string | null;
    /**
     * Returns all runs in the paragraph as Run class instances.
     */
    getRuns(): Run[];
    /**
     * Appends a Run instance to the paragraph.
     */
    addRun(run: Run): void;
    /**
     * Removes the run at the given zero-based index.
     */
    removeRun(index: number): void;
    /**
     * Converts the internal paragraph object back into an XML string.
     * @returns A Promise that resolves with the XML string.
     */
    toXml(): Promise<string>;
}

/**
 * Interface for a paragraph element, which contains a collection of content elements.
 * A paragraph is a block-level element that can contain multiple runs, hyperlinks, and other elements.
 * @example
 * <w:p>
 * <w:pPr>
 * <w:pStyle w:val="Normal" />
 * </w:pPr>
 * <w:r>
 * <w:t>Hello</w:t>
 * </w:r>
 * <w:r>
 * <w:t xml:space="preserve"> World</w:t>
 * </w:r>
 * </w:p>
 */
declare interface Paragraph_2 {
    $: {
        "w14:paraId"?: string;
        "w14:textId"?: string;
        "w:rsidR"?: string;
        [key: string]: string | undefined;
    };
    "w:pPr"?: ParagraphProperties;
    "w:r"?: Run_2[];
    "w:hyperlink"?: Hyperlink[];
    "w:fldChar"?: Field[];
    "w:drawing"?: Drawing[];
}

/** The paragraph's alignment (`w:jc` value: left/center/right/both/…), or null. */
export declare function paragraphAlignment(xml: string): string | null;

/**
 * The paragraph's font size in POINTS, read from the first `<w:sz>`/`<w:szCs>`
 * (which store half-points), or null if none is set inline. Useful as a
 * "larger than body" signal when inferring untyped headings.
 */
export declare function paragraphFontSizePt(xml: string): number | null;

/** Options accepted by paragraph/heading verbs. */
export declare type ParagraphFormat = Omit<StyledParagraphOptions, "outlineLevel" | "styleId"> & {
    styleId?: string;
};

/**
 * The heading level (1–6) of a paragraph block, 0 if it is not a heading.
 * Prefers the paragraph style (`headingLevelFromStyleId`); falls back to the
 * paragraph's own outline level — so both Word-authored and imported headings are
 * recognised.
 */
export declare function paragraphHeadingLevel(xml: string): number;

/**
 * True when the paragraph's text is entirely bold — i.e. every text-bearing run
 * is bold. A common signal for an untyped (plain-style) title. Returns false for
 * paragraphs with no text.
 */
export declare function paragraphIsBold(xml: string): boolean;

/**
 * Options for {@link Paragraph.make} — a single-run, formatted paragraph factory.
 * Mirrors the hand-rolled `new Paragraph({...})` + applyStyle + setAlignment +
 * addRun(Run.fromText().setBold()...) pattern that document builders repeat.
 */
export declare interface ParagraphOptions {
    /** Paragraph style id, e.g. "Heading1", "ListParagraph". */
    styleId?: string;
    /** Horizontal alignment (`w:jc`). */
    alignment?: "left" | "center" | "right" | "both";
    /** Outline level (`w:outlineLvl`, 0-based) — drives TOC depth for headings. */
    outlineLevel?: number;
    /** Bold the run. */
    bold?: boolean;
    /** Italicise the run. */
    italic?: boolean;
    /** Font size in POINTS (converted to half-points internally). */
    fontSizePt?: number;
    /** Font family (applied to both ascii/hAnsi and complex-script). */
    fontFamily?: string;
    /** Text colour (hex, with or without leading '#'). */
    color?: string;
    /** Mark the run right-to-left (`<w:rtl/>`) for Arabic / Hebrew shaping. */
    rtl?: boolean;
}

/**
 * Read a paragraph's own outline level (`<w:outlineLvl w:val="N"/>`, 0-based) and
 * return it as a 1-based heading level (N+1), or 0 if absent / out of range.
 * Catches headings whose paragraph carries an outline level without a heading
 * STYLE — common in imported documents.
 */
export declare function paragraphOutlineLevel(xml: string): number;

/** Word's paragraph pagination toggles. `undefined` = leave as it is. */
export declare interface ParagraphPagination {
    /** Keep with next — the paragraph never ends a page alone, ahead of its body text. */
    keepWithNext?: boolean;
    /** Keep lines together — the paragraph is never split across two pages. */
    keepLines?: boolean;
    /** Widow/orphan control — never leave a single line of it on a page by itself. */
    widowControl?: boolean;
    /** Always start this paragraph on a new page. */
    pageBreakBefore?: boolean;
}

/**
 * Interface for the properties of a paragraph.
 * These properties apply to the entire paragraph, such as style, alignment, and spacing.
 */
declare interface ParagraphProperties {
    "w:pStyle"?: {
        $: {
            "w:val": string;
        };
    };
    "w:tabs"?: {};
    "w:spacing"?: {};
    "w:jc"?: {
        $: {
            "w:val": string;
        };
    };
    "w:outlineLvl"?: {
        $: {
            "w:val": string;
        };
    };
    "w:numPr"?: {
        "w:ilvl"?: {
            $: {
                "w:val": string;
            };
        };
        "w:numId"?: {
            $: {
                "w:val": string;
            };
        };
    };
    /**
     * Represents the default run properties for the paragraph mark.
     * Any run properties defined here are inherited by all runs in the paragraph,
     * unless overridden by a specific run's own properties.
     */
    "w:rPr"?: RunProperties;
}

/**
 * The `<w:pPr>…</w:pPr>` substring of a paragraph — style, alignment, indent,
 * spacing, bidi, numbering — or "" when it has none.
 */
export declare function paragraphPropsXml(paragraphXml: string): string;

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
export declare function paragraphRunPropsXml(paragraphXml: string): string;

export declare interface ParagraphRunStyleResult {
    xml: string;
    /** How many runs were restyled. Zero means the paragraph holds no text runs. */
    runs: number;
}

/** Extract the `w:val` of the paragraph's `<w:pStyle>`, or null. */
export declare function paragraphStyleId(xml: string): string | null;

/** Concatenate the decoded text of every `<w:t ...>...</w:t>` in a block. */
export declare function paragraphText(xml: string): string;

/**
 * Read EVERY caption in one paragraph's XML — Word numbers each SEQ field it
 * finds, and real theses do put two figures on one line ("Figure 19: … Figure
 * 20: …"). Returns [] when the paragraph carries no caption field.
 *
 * Exported because reading a caption needs no document: the app's figure-caption
 * endpoint uses it to strip the label and number off a caption before handing the
 * wording to the model.
 */
export declare function parseCaptionParagraph(xml: string, blockIndex?: number): CaptionEntry[];

/** Split a full document into its top-level body blocks. */
export declare function parseOrderedDoc(documentXml: string): {
    split: SplitDocument;
    blocks: BodyBlock[];
    bodyChildren: BodyBlock[];
};

/**
 * Parse the INSIDE of a `<w:rPr>` into observed properties. The read inverse of
 * {@link mergeRunProps}: it reports only what this fragment states, never a
 * default, so callers can merge rungs of the cascade without a lower one
 * silently masking a higher one.
 */
export declare function parseRunProps(rPrInner: string): ObservedRunProps;

/**
 * Index `styles.xml`: every `<w:style>`'s run properties and `w:basedOn`, plus
 * `w:docDefaults`. String surgery rather than an xml2js round-trip, matching
 * every other read on this file — and cheap enough to do once per inspection.
 */
export declare function parseStylesIndex(stylesXml: string | null | undefined): StylesIndex;

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
export declare function parseTextCaption(text: string): TextCaptionMatch | null;

/** Convert pixels (96 DPI) to EMU. */
export declare const pixelsToEmu: (px: number) => number;

/** Raw parsed numbering parts (xml2js objects), for cross-document copy. */
declare interface RawNumbering {
    abstractNums: any[];
    nums: any[];
}

/** Read where the first drawing in `xml` is placed. `null` when there is none. */
export declare function readDrawingLayout(xml: string): DrawingPlacement | null;

/** Read back the effective (directly-set) pagination flags of a paragraph. */
export declare function readParagraphPagination(xml: string): ParagraphPagination;

/**
 * Read a paragraph's `<w:pPr>` inner XML.
 *
 * Anchored to the paragraph's OPENING TAG rather than searching the whole
 * string, because `w:pPr` — when present — is by schema the first child of
 * `w:p`. A free search would happily find the `<w:pPr>` of a paragraph nested
 * inside a text box (`w:txbxContent`) and rewrite the wrong one.
 *
 * Returns `null` when `xml` is not a paragraph element, and `""` when it is a
 * paragraph with no properties of its own.
 */
export declare function readParagraphProps(xml: string): string | null;

/** One entry from a `.rels` part. */
export declare interface RelationshipEntry {
    id: string;
    type: string;
    /** Target path, usually relative to the part's folder (e.g. "media/image1.png"). */
    target: string;
    /** "External" for hyperlinks/external targets; undefined for internal parts. */
    targetMode?: string;
}

export declare class RelManager {
    zip: default_2;
    relsPath: string;
    ns: string;
    constructor(zip: default_2, relsPath?: string);
    private readRels;
    private writeRels;
    /**
     * Adds a relationship entry: Id must be unique (caller responsible).
     * target should be relative to 'word/' (e.g. 'header1.xml' or 'media/image1.png')
     */
    addRelationship(id: string, type: string, target: string, targetMode?: string): Promise<void>;
    /**
     * Quick helper to generate a new rId (checks existing ones)
     */
    genId(prefix?: string): Promise<string>;
    /**
     * Read all relationships from the .rels part as a typed list. Empty when the
     * part is missing.
     */
    getRelationships(): Promise<RelationshipEntry[]>;
    /** Resolve a relationship id to its `Target` (e.g. "media/image1.png"), or null. */
    getTarget(relId: string): Promise<string | null>;
}

declare enum RelsType {
    Root = "_rels/.rels",
    Document = "word/_rels/document.xml.rels"
}

/** A chunk a {@link BlockRenderer} hook may return; insert offsets are local. */
export declare interface RenderedChunk {
    paragraphs: Paragraph[];
    tables?: RenderedTable[];
    images?: RenderedImage[];
}

/** An inline image (PNG bytes + EMU dimensions) to embed after the Nth paragraph. */
export declare interface RenderedImage {
    afterParaCount: number;
    png: Buffer;
    widthEmu: number;
    heightEmu: number;
}

/** Output of {@link renderMarkdownBlocks}. */
export declare interface RenderedMarkdown {
    paragraphs: Paragraph[];
    tables: RenderedTable[];
    images: RenderedImage[];
}

/** A `Table` to be inserted after the Nth rendered paragraph. */
export declare interface RenderedTable {
    afterParaCount: number;
    table: Table;
}

/**
 * Render markdown blocks into engine paragraphs + table/image inserts.
 *
 * Heading levels map to `Heading{N}` styles starting at `ctx.headingBase`
 * (default 2): a markdown `#` → `Heading{base}`, `##` → `Heading{base+1}`, …
 * capped at Heading6. Headings are bold with a matching `w:outlineLvl` so a TOC
 * can be derived. Tables are recorded as inserts positioned after the paragraph
 * count at the point they appear (the same contract the export pipeline uses).
 *
 * Pass `renderBlock` to intercept blocks (e.g. `code` blocks of a custom
 * language); anything the hook does not claim is rendered by the engine.
 */
export declare function renderMarkdownBlocks(blocks: MarkdownBlock[], ctx: MarkdownRenderCtx, renderBlock?: BlockRenderer): RenderedMarkdown;

/**
 * Repair a .docx held as bytes.
 *
 * The result is re-opened and re-checked before it is handed back: if the repair
 * somehow left a part unreadable, or turned a clean fatal into a new one, the
 * whole thing is discarded and the original bytes are returned. A doctor that
 * can corrupt a thesis is worse than no doctor.
 */
export declare function repairDocxBuffer(buffer: Buffer, opts?: {
    aggressive?: boolean;
}): Promise<RepairResult>;

export declare interface RepairResult extends DoctorReport {
    /** The repaired bytes, or the ORIGINAL buffer when nothing was rewritten. */
    buffer: Buffer;
    /** True when `buffer` differs from what was passed in. */
    changed: boolean;
}

/**
 * Resolve a style through its `w:basedOn` ancestry, base-first so the nearest
 * definition wins. Returns the chain too: an AI reporting "Heading1, which
 * inherits from Normal" is telling the student where to make the change.
 *
 * A dangling reference is the NORMAL state in this corpus (the seed thesis has
 * styles based on a `Normal` it never defines), so a missing link truncates the
 * walk instead of throwing.
 */
export declare function resolveStyleChain(styleId: string, index: StylesIndex): {
    chain: string[];
    props: ObservedRunProps;
    defined: boolean;
};

export declare interface RevisionEntry {
    id: number;
    type: RevisionType;
    author: string;
    date: string;
    text: string;
    paragraphIndex: number;
}

export declare type RevisionType = "ins" | "del" | "rPrChange" | "pPrChange";

export declare class RootRelManager extends RelManager {
    constructor(zip: ZipManager, relsPath?: RelsType);
}

/**
 * A class to represent and manipulate a single WordprocessingML run (<w:r>).
 */
export declare class Run {
    private run;
    constructor(run: Run_2);
    /**
     * Get the raw run object.
     */
    getRaw(): Run_2;
    /**
     * Get the text content of the run.
     */
    getText(): string;
    /**
     * Set or replace the text content of the run.
     * @param newText - The new text to set.
     */
    setText(newText: string): void;
    /**
     * Append text to the existing text in the run.
     * @param textToAppend - Text to append.
     */
    appendText(textToAppend: string): void;
    /**
     * Get the run's formatting properties (<w:rPr>).
     */
    getProperties(): RunProperties | undefined;
    /**
     * Ensure the run has a <w:rPr> object to work with.
     */
    private ensureProperties;
    /**
     * Set bold formatting.
     */
    setBold(enable?: boolean): void;
    /**
     * Set italic formatting.
     */
    setItalic(enable?: boolean): void;
    /**
     * Set underline formatting with optional style.
     * @param style - Underline style (e.g., "single", "double")
     */
    setUnderline(enable?: boolean, style?: string): void;
    /**
     * Set shading (background color) for the run.
     */
    setShading(color?: string, value?: string): void;
    /**
     * Remove shading.
     */
    removeShading(): void;
    /**
     * Mark the run as right-to-left (`<w:rtl/>`) so Word shapes Arabic / Hebrew
     * text correctly. Pass `false` to remove the marker.
     */
    setRtl(enable?: boolean): void;
    /** Determine if the run is marked right-to-left. */
    isRtl(): boolean;
    /**
     * Determine if the run is bold.
     */
    isBold(): boolean;
    /**
     * Determine if the run is italic.
     */
    isItalic(): boolean;
    /**
     * Check if the run has underline formatting.
     */
    hasUnderline(): boolean;
    /**
     * Determine if the run is empty (no text and no special fields).
     */
    isEmpty(): boolean;
    /**
     * Add a field to the run, like a page number or TOC reference.
     */
    setField(fieldType: "begin" | "end", instrText?: string): void;
    /**
     * Add a drawing (image or graphic) to the run.
     */
    setDrawing(drawing: Drawing): void;
    /**
     * Set font size.
     * @param halfPoints Size in half-points (e.g. 24 = 12pt, 28 = 14pt).
     */
    setFontSize(halfPoints: number): void;
    /**
     * Set font family.
     * @param ascii  Font name for ASCII/Latin characters.
     * @param cs     Font name for complex scripts (defaults to ascii).
     */
    setFontFamily(ascii: string, cs?: string): void;
    /**
     * Set text color.
     * @param hex  6-character hex color without '#' (e.g. "FF0000" for red).
     */
    setColor(hex: string): void;
    /**
     * Clear all run formatting (empties w:rPr).
     */
    clearFormatting(): void;
    /**
     * Returns the underlying raw run object.
     */
    toObject(): Run_2;
    /**
     * Creates a Run from a plain text string with no formatting.
     */
    static fromText(text: string): Run;
    /**
     * Export the run to XML string.
     */
    toXml(): string;
}

/**
 * Represents a text run in the document.
 * A run may contain text, tabs, breaks, fields, and drawings.
 */
declare interface Run_2 {
    $?: {
        "w:rsidRPr"?: string;
        "w:rsidR"?: string;
        [key: string]: string | undefined;
    };
    /**
     * Run-level properties defining style and formatting.
     */
    "w:rPr"?: RunProperties;
    /**
     * Text nodes inside the run.
     * Some runs have multiple <w:t> tags if Word splits the text.
     */
    "w:t"?: TextNode | TextNode[] | string;
    /**
     * Line breaks inside a run.
     * <w:br/>
     */
    "w:br"?: Record<string, unknown> | Record<string, unknown>[];
    /**
     * Tab characters inside a run.
     * <w:tab/>
     */
    "w:tab"?: Record<string, unknown> | Record<string, unknown>[];
    /**
     * Dynamic fields or special instructions.
     * <w:instrText xml:space="preserve">PAGE</w:instrText>
     */
    "w:instrText"?: TextNode | TextNode[];
    /**
     * Drawings or images inside a run.
     * <w:drawing>...</w:drawing>
     */
    "w:drawing"?: Drawing | Drawing[];
    /**
     * A single field character, not a whole Field object
     * Example: <w:fldChar w:fldCharType="begin"/>
     */
    "w:fldChar"?: {
        $: {
            "w:fldCharType": "begin" | "end";
        };
    };
}

/**
 * Interface for the properties of a text run.
 * These properties define the formatting for the run's content.
 */
declare interface RunProperties {
    "w:rStyle"?: {
        $: {
            "w:val": string;
        };
    };
    "w:rFonts"?: {
        $: {
            "w:ascii"?: string;
            "w:hAnsi"?: string;
            "w:eastAsia"?: string;
            "w:cs"?: string;
        };
    };
    "w:noProof"?: {};
    "w:lang"?: {
        $: {
            "w:val": string;
            "w:eastAsia"?: string;
        };
    };
    "w:b"?: {};
    "w:i"?: {};
    "w:u"?: {
        $: {
            "w:val": string;
        };
        "w:sh": {
            $: {
                "w:fill": string;
                "w:val": string;
            };
        };
    };
    "w:shd"?: {
        $: {
            "w:fill": string;
            "w:val": string;
        };
    };
    "w:rtl"?: {};
    "w:sz"?: {
        $: {
            "w:val": string;
        };
    };
    "w:szCs"?: {
        $: {
            "w:val": string;
        };
    };
    "w:color"?: {
        $: {
            "w:val": string;
        };
    };
}

/**
 * The run-level properties `set_text_style` can apply. Every field is optional;
 * only the ones actually supplied are ever written, stripped, or considered.
 * `undefined` means "the student did not name this" — distinct from `false`,
 * which means "remove it".
 */
export declare interface RunProps {
    /** Font family. Written to `w:ascii`, `w:hAnsi` AND `w:cs`. */
    font?: string;
    /** Size in POINTS. Written as half-points to BOTH `w:sz` and `w:szCs`. */
    sizePt?: number;
    /** true writes `<w:b/><w:bCs/>`; false removes both. */
    bold?: boolean;
    /** true writes `<w:i/><w:iCs/>`; false removes both. */
    italic?: boolean;
    /** Hex colour, with or without a leading '#'. */
    color?: string;
}

export declare type SectionBreakType = "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";

/** Result of a section-scoped header/footer change. */
export declare interface SectionEditResult {
    sectionIndex: number;
    totalSections: number;
}

export declare interface SectionEntry {
    index: number;
    isFinal: boolean;
    type?: string;
    pageSize?: SectionPageSize;
    margins?: SectionMargins;
    headerRefs: SectionHeaderFooterRef[];
    footerRefs: SectionHeaderFooterRef[];
    /** Parsed w:pgNumType, when the section sets one (page-number format/restart). */
    pageNumberType?: {
        format: string;
        start?: number;
    };
    paragraphIndex?: number;
}

export declare interface SectionHeaderFooterRef {
    relId: string;
    type: "default" | "first" | "even";
}

export declare interface SectionInfo {
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
    headerBorder: {
        bottom: boolean;
        color: string | null;
    } | null;
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

export declare interface SectionLayout {
    type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
    pageSize?: SectionPageSize;
    margins?: SectionMargins;
}

export declare class SectionManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private getBodyParagraphs;
    private norm;
    private parseSectPr;
    private applyLayout;
    private findSectPr;
    getSections(): Promise<SectionEntry[]>;
    addSectionBreak(paragraphIndex: number, type: "nextPage" | "continuous" | "evenPage" | "oddPage"): Promise<void>;
    removeSectionBreak(paragraphIndex: number): Promise<void>;
    setSectionLayout(sectionIndex: number, layout: SectionLayout): Promise<void>;
    setSectionHeader(sectionIndex: number, relId: string, type?: "default" | "first" | "even"): Promise<void>;
    setSectionFooter(sectionIndex: number, relId: string, type?: "default" | "first" | "even"): Promise<void>;
    /**
     * Set ONE section's page-number format and/or restart value (`<w:pgNumType>`)
     * — e.g. roman for the front matter, arabic restarting at 1 for the body.
     *
     * `FooterManager.formatPageNumbers` only ever reaches the BODY sectPr, so it
     * cannot express per-section numbering; this can. Omit `start` to continue
     * the previous section's sequence.
     */
    setSectionPageNumbering(sectionIndex: number, opts: {
        format?: string;
        start?: number;
    }): Promise<void>;
    /**
     * Vertically align a section's page content (`<w:vAlign>`). "center" is what
     * puts a divider page's title in the true middle of the page.
     */
    setSectionVerticalAlign(sectionIndex: number, vAlign: "top" | "center" | "both" | "bottom"): Promise<void>;
    /**
     * Draw a page border on all four edges of ONE section (`<w:pgBorders>`) — the
     * `frame` divider family. Colour must be 6-hex (leading '#' tolerated):
     * ST_HexColor admits nothing else, and an invalid value here is the class of
     * defect that trips Word's "unreadable content" repair dialog.
     */
    setSectionPageBorders(sectionIndex: number, opts: SectPrPageBorderOptions): Promise<void>;
    /**
     * Apply `transform` to the `<w:sectPr>` of section `sectionIndex` — byte-safe
     * string surgery. Intermediate sections live inside a paragraph's `<w:pPr>`;
     * the final section is the body's `<w:sectPr>` (created if absent). Sections
     * are counted in document order: paragraph-level sectPrs first, final last —
     * the same order as getSections().
     */
    private editSectionSectPr;
}

export declare interface SectionMargins {
    top: number;
    bottom: number;
    left: number;
    right: number;
    header?: number;
    footer?: number;
    gutter?: number;
}

/** A section's page geometry in twips (1440 = 1 inch), inheritance resolved. */
export declare interface SectionPageGeometry {
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

export declare interface SectionPageSize {
    width: number;
    height: number;
    orientation?: "portrait" | "landscape";
}

export declare interface SectPrPageBorderOptions {
    style: "single" | "double" | "thick" | "dashed" | "dotted";
    /** 6-hex, with or without a leading '#'. */
    color: string;
    /** Border width in POINTS; Word stores eighths of a point (w:sz). */
    widthPt: number;
    /**
     * Distance from the page edge in points (w:space). Default 24. Word's own
     * dialog only accepts 24–31 when measuring from the page edge; other values
     * open fine but Word clamps them.
     */
    offsetPt?: number;
}

/**
 * The SEQ identifier for a caption label. A SEQ name is bookmark-like, so it
 * cannot contain whitespace — "الشكل رقم" numbers under "الشكل_رقم" while the
 * document still READS "الشكل رقم". The Table-of-Figures `\c` switch uses the
 * same identifier, so the two always agree.
 */
export declare function seqIdentifier(label: string): string;

export declare function setParagraphText(paragraphXml: string, text: string): string;

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
 *    "a data descriptor follows" flag it did not honour, and then cannot read its
 *    own output back. See `clearFalseDataDescriptors`.
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
/** `fatal` — the document is broken (Word refuses it, or its content order is
 *  destroyed). `warning` — wrong or degraded, but it still opens. */
export declare type Severity = "fatal" | "warning";

export declare interface ShapeEntry {
    id: number;
    name: string;
    type: "textbox" | ShapeType | "unknown";
    position?: ShapePosition;
    size: ShapeSize;
}

export declare class ShapeManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private getBodyParagraphs;
    private norm;
    private nextShapeId;
    private buildSpPr;
    private buildWsp;
    private buildGraphicData;
    private buildAnchor;
    private buildInline;
    private insertDrawingParagraph;
    /** Collect all w:drawing nodes from a run, handling both direct and
     *  mc:AlternateContent-wrapped drawings. */
    private drawingsFromRun;
    private extractShapesFromParagraph;
    getShapes(): Promise<ShapeEntry[]>;
    insertTextBox(opts: TextBoxOptions): Promise<number>;
    insertShape(type: ShapeType, opts?: InsertShapeOptions): Promise<number>;
    insertLine(x1: number, y1: number, x2: number, y2: number, opts?: InsertLineOptions): Promise<number>;
    /** Insert an inline picture referencing an already-registered image relId
     *  (see MediaManager.insertImage). width/height are in EMU (1 px @96dpi = 9525 EMU). */
    insertImage(relId: string, opts: {
        width: number;
        height: number;
        name?: string;
        paragraphIndex?: number;
    }): Promise<number>;
    deleteShape(id: number): Promise<void>;
}

export declare interface ShapePosition {
    x: number;
    y: number;
}

export declare interface ShapeSize {
    width: number;
    height: number;
}

export declare type ShapeType = "rect" | "roundRect" | "ellipse" | "triangle" | "diamond" | "line" | "rightArrow" | "leftArrow" | "star5" | "cloud" | "heart";

export declare interface SkippedTextCaption {
    blockIndex: number;
    reason: SkipReason;
    text: string;
}

/** Why a paragraph that reads like a caption was left alone. */
export declare type SkipReason = "already-a-caption" | "is-a-heading" | "contains-image" | "contains-field";

/** The result of splitting a full `word/document.xml`. */
export declare interface SplitDocument {
    /** Everything up to and including the `<w:body ...>` open tag. */
    pre: string;
    /** The `<w:body ...>` open tag itself (subset of `pre`, for convenience). */
    bodyOpen: string;
    /** Top-level children of `<w:body>`, in document order. */
    blocks: BodyBlock[];
    /** `</w:body>` and everything after it (trailing whitespace, `</w:document>`, …). */
    post: string;
}

/**
 * Split a full `word/document.xml` into the body region's top-level children,
 * each preserved as its exact original substring.
 *
 * Guarantee: `pre + blocks.map(b => b.xml).join("") + post === documentXml`.
 */
export declare function splitDocument(documentXml: string): SplitDocument;

/** Strip the common inline markdown markers, leaving plain text. */
export declare function stripInlineMarkdown(s: string): string;

declare interface StyleDef {
    id: string;
    type: string;
    name?: string;
    basedOn?: string;
    props: ObservedRunProps;
}

/** Options for {@link makeStyledParagraphXml} — a fully-formatted single-run paragraph. */
export declare interface StyledParagraphOptions {
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

export declare interface StyleEntry {
    id: string;
    name: string;
    type: string;
}

/** What one property resolves to across a whole target. */
export declare interface StyleFacet extends ValueShare {
    /** Runner-up values, share-descending — present only when the target is not uniform. */
    others?: ValueShare[];
}

declare interface StylesIndex {
    byId: Map<string, StyleDef>;
    /** The `w:type="paragraph" w:default="1"` style a paragraph with no `w:pStyle` resolves to. */
    defaultParagraphStyleId: string;
    docDefaults: ObservedRunProps;
}

export declare class StylesManager {
    private zip;
    constructor(zip: default_2);
    private readStyles;
    private writeStyles;
    private normalizeStylesArray;
    /**
     * Lists all styles in the document.
     */
    listStyles(): Promise<StyleEntry[]>;
    /**
     * Returns the raw style object for a given style ID, or null.
     */
    getStyle(styleId: string): Promise<any | null>;
    /**
     * Adds a new style. Does nothing if a style with the same ID already exists.
     */
    addStyle(styleObj: any): Promise<void>;
    /**
     * Removes a style by ID.
     */
    removeStyle(styleId: string): Promise<void>;
    /**
     * Forces `formatting` (font size / bold / color) onto the `Heading{level}`
     * styles for the given `levels` (default 1–6) — a STYLE-level change, so it
     * applies uniformly to every heading using that level, present and future,
     * rather than one paragraph's run. See {@link applyHeadingStyleToXml}.
     */
    setHeadingStyle(levels: number[] | undefined, formatting: HeadingRunFormatting): Promise<HeadingStyleResult>;
    /**
     * Ensure `styleId` exists (creating it from `ensure` when missing), then apply
     * `props` to its `<w:rPr>`. A STYLE-level change: it reaches every paragraph
     * using the style, present and future.
     */
    setStyleRunProps(styleId: string, props: RunProps, ensure?: EnsureStyleSpec): Promise<{
        created: boolean;
        updated: boolean;
    }>;
}

/** One-line summary for a log line or a tool reply. */
export declare function summarize(report: DoctorReport): string;

/**
 * Wraps a parsed <w:tbl> and provides a full Table Design + Layout API.
 */
export declare class Table {
    private table;
    constructor(table: TableObject);
    private getRows;
    private setRows;
    private getCells;
    private setCells;
    private extractCellText;
    private ensureCellProps;
    private ensureRowProps;
    private ensureTableProps;
    private buildBorderEl;
    private emptyCell;
    /** Number of rows. */
    getRowCount(): number;
    /** Number of columns in the given row (default row 0). */
    getColumnCount(rowIndex?: number): number;
    /** Raw cell object at [row, col], or null. */
    getCell(row: number, col: number): TableCell | null;
    /** Plain text of cell [row, col]. */
    getCellText(row: number, col: number): string;
    /**
     * Plain text of every cell as a row-major grid (`string[][]`). Ragged rows are
     * kept as-is (each row reflects its own cell count). Cell text is the joined
     * run text; intra-cell paragraph breaks become "\n".
     */
    getAllCellText(): string[][];
    /** Apply a named Word table style, e.g. "TableGrid", "LightShading-Accent1". */
    setTableStyle(styleId: string): this;
    /**
     * Table alignment on the page.
     * Equivalent to Table Properties → Table → Alignment.
     */
    setTableAlignment(alignment: "left" | "center" | "right"): this;
    /**
     * Table indent from the left margin in twips.
     * Equivalent to Table Properties → Table → Indent from Left.
     */
    setTableIndent(twips: number): this;
    /**
     * Text wrapping around the table.
     * Equivalent to Table Properties → Table → Text Wrapping.
     */
    setTextWrapping(type: "none" | "around"): this;
    /**
     * Right-to-left table direction.
     * Equivalent to Table Properties → Table → Table direction.
     */
    setTableDirection(rtl: boolean): this;
    /**
     * Set accessible alt text (title + description).
     * Equivalent to Table Properties → Alt Text.
     */
    setAltText(title: string, description?: string): this;
    /**
     * Set table width.
     * type "pct"  → percentage of page width (pass 0–100)
     * type "dxa"  → fixed twips
     * type "auto" → auto
     */
    setTableWidth(value: number, type?: "pct" | "dxa" | "auto"): this;
    /**
     * Layout mode.
     * "autofit" → AutoFit Contents / AutoFit Window
     * "fixed"   → Fixed Column Width
     */
    setLayoutMode(mode: "autofit" | "fixed"): this;
    /** Shorthand: AutoFit to contents. */
    autoFitContents(): this;
    /** Shorthand: AutoFit to window (page width). */
    autoFitWindow(): this;
    /** Shorthand: Fixed column widths. */
    fixedColumnWidth(): this;
    /**
     * Default cell margins for the entire table (twips).
     * Equivalent to Table Properties → Options → Default cell margins.
     */
    setDefaultCellMargins(margins: CellMargins): this;
    /**
     * Table-level borders.
     */
    setTableBorders(borders: TableBorderOptions): this;
    /**
     * Set row height.
     * rule: "atLeast" (At least) | "exact" (Exact)
     * Equivalent to Table Properties → Row → Specify height.
     */
    setRowHeight(rowIndex: number, heightTwips: number, rule?: "atLeast" | "exact"): this;
    /**
     * Allow/prevent a row from breaking across pages.
     * Equivalent to Table Properties → Row → Allow row to break across pages.
     */
    setRowAllowBreak(rowIndex: number, allow: boolean): this;
    /**
     * Mark a row as a repeating header at the top of each page.
     * Equivalent to Table Properties → Row → Repeat as header row.
     * Optionally set a background fill colour.
     */
    setHeaderRow(rowIndex: number, fillColor?: string): this;
    /** Shade all cells in a row with one background colour. */
    setRowShading(rowIndex: number, fillColor: string): this;
    /**
     * Set preferred width (twips) for every cell in a column.
     * Equivalent to Table Properties → Column → Preferred width.
     */
    setColumnWidth(colIndex: number, widthTwips: number): this;
    /** Set cell preferred width in twips. */
    setCellWidth(row: number, col: number, widthTwips: number): this;
    /**
     * Set cell vertical alignment.
     * Equivalent to Table Properties → Cell → Vertical Alignment.
     */
    setCellVerticalAlignment(row: number, col: number, alignment: "top" | "center" | "bottom"): this;
    /**
     * Set individual cell margins (overrides table default).
     * Equivalent to Table Properties → Cell → Options.
     */
    setCellMargins(row: number, col: number, margins: CellMargins): this;
    /** Set cell background shading colour. */
    setCellShading(row: number, col: number, fillColor: string): this;
    /**
     * Format a cell's EXISTING text (every run in every paragraph), preserving
     * the content — unlike setCellContent, which replaces it. Pass only the
     * properties you want to change. sizeHalfPoints: Word half-points (24 = 12pt).
     */
    setCellTextFormat(row: number, col: number, opts: {
        bold?: boolean;
        italic?: boolean;
        sizeHalfPoints?: number;
        fontFamily?: string;
    }): this;
    /** Grid column of the tc at [row, col] (cumulative gridSpans of preceding cells). */
    private gridColOf;
    /**
     * SPLIT a horizontally merged cell back into single cells: removes its
     * gridSpan and inserts the missing empty cells after it. Reverse of
     * mergeCellsHorizontal (the merged text stays in the first cell).
     */
    splitCellHorizontal(row: number, col: number): this;
    /**
     * SPLIT a vertically merged cell: removes vMerge from the restart cell at
     * [row, col] and from its continuation cells below (matched by GRID column).
     * Reverse of mergeCellsVertical (the merged text stays in the top cell).
     */
    splitCellVertical(row: number, col: number): this;
    /** Move row `from` to position `to` (both 0-based). */
    moveRow(from: number, to: number): this;
    /** Move column `from` to position `to` (both 0-based, tc-indexed). */
    moveColumn(from: number, to: number): this;
    /**
     * Set the font colour of a cell's EXISTING text (every run in every
     * paragraph), preserving the content — unlike setCellContent, which replaces
     * it. 6-hex colour, with or without '#'.
     */
    setCellTextColor(row: number, col: number, color: string): this;
    /** Set text with rich formatting in a cell. */
    setCellContent(row: number, col: number, text: string, opts?: CellTextOptions): this;
    /** Set plain text in a cell. */
    setCellText(row: number, col: number, text: string): this;
    private makeEmptyRow;
    /**
     * Insert a blank row above the given index.
     * Equivalent to Layout → Insert Above.
     */
    insertRowAbove(rowIndex: number): this;
    /**
     * Insert a blank row below the given index.
     * Equivalent to Layout → Insert Below.
     */
    insertRowBelow(rowIndex: number): this;
    /**
     * Append a row, optionally pre-filled with text.
     */
    addRow(cellTexts?: string[]): this;
    /**
     * Remove the row at the given index.
     * Equivalent to Layout → Delete → Delete Rows.
     */
    removeRow(index: number): this;
    /**
     * Insert a blank column to the left of colIndex.
     * Equivalent to Layout → Insert Left.
     */
    insertColumnLeft(colIndex: number): this;
    /**
     * Insert a blank column to the right of colIndex.
     * Equivalent to Layout → Insert Right.
     */
    insertColumnRight(colIndex: number): this;
    /**
     * Delete an entire column.
     * Equivalent to Layout → Delete → Delete Columns.
     */
    deleteColumn(colIndex: number): this;
    /**
     * Horizontally merge cells in a single row from startCol to endCol (inclusive).
     * Equivalent to Layout → Merge Cells.
     */
    mergeCellsHorizontal(rowIndex: number, startCol: number, endCol: number): this;
    /**
     * Vertically merge cells in a column from startRow to endRow (inclusive).
     * Equivalent to Layout → Merge Cells (vertical selection).
     */
    mergeCellsVertical(colIndex: number, startRow: number, endRow: number): this;
    /**
     * Set all rows to the same height (based on tallest row or a fixed value).
     * Equivalent to Layout → Distribute Rows.
     */
    distributeRows(heightTwips?: number): this;
    /**
     * Set all columns to equal width.
     * Equivalent to Layout → Distribute Columns.
     */
    distributeColumns(): this;
    /**
     * Sort data rows (skipping the header row) by the value in colIndex.
     * Equivalent to Layout → Sort.
     */
    sortByColumn(colIndex: number, direction?: "asc" | "desc", hasHeaderRow?: boolean): this;
    /** Raw table object for serialisation. */
    toObject(): TableObject;
    /**
     * Build a styled `Table` from a grid of cell texts. The widest row determines
     * the column count; short rows are padded with empty cells. By default the
     * table gets a light-grey single-line border grid and 100% width.
     *
     * Replaces the hand-rolled `buildTable(header, rows, rtl)` helpers used by
     * document builders, and backs the string-level `makeTableXml` in OrderedBody.
     */
    /**
     * Parse a `<w:tbl>…</w:tbl>` XML string into a `Table`. Useful for reading a
     * table out of an OrderedBody block (`block.xml`) without the xml2js
     * full-document path — e.g. `(await Table.fromXml(block.xml)).getAllCellText()`.
     */
    static fromXml(tableXml: string): Promise<Table>;
    static fromGrid(rows: string[][], opts?: FromGridOptions): Table;
}

export declare interface TableBorderOptions {
    top?: BorderSide;
    bottom?: BorderSide;
    left?: BorderSide;
    right?: BorderSide;
    insideH?: BorderSide;
    insideV?: BorderSide;
}

export declare interface TableCell {
    $?: Record<string, any>;
    "w:tcPr"?: TableCellProperties;
    "w:p"?: any | any[];
}

declare interface TableCellProperties {
    "w:tcW"?: {
        $: {
            "w:w": string;
            "w:type": string;
        };
    };
    "w:shd"?: {
        $: {
            "w:fill": string;
            "w:val": string;
            "w:color"?: string;
        };
    };
    "w:vAlign"?: {
        $: {
            "w:val": string;
        };
    };
    "w:tcMar"?: any;
    "w:gridSpan"?: {
        $: {
            "w:val": string;
        };
    };
    "w:vMerge"?: {
        $?: {
            "w:val"?: string;
        };
    } | string;
    [key: string]: any;
}

/** A table as a plain text grid. */
export declare interface TableInfo {
    index: number;
    rows: string[][];
}

export declare interface TableObject {
    $?: Record<string, any>;
    "w:tblPr"?: TableProperties;
    "w:tblGrid"?: any;
    "w:tr"?: TableRow | TableRow[];
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
export declare class TableOfContentsManager {
    private zip;
    private doc;
    private styles;
    /** Deepest level whose TOC style is known registered — NOT a boolean: the same
     *  manager is reused across calls (the engine is cached per thesis), and a
     *  second call at depth 5 after a first at depth 3 must still register TOC4/5
     *  or those entries silently render as body text. */
    private registeredDepth;
    constructor(zip: default_2);
    /** Register TOCHeading + TOC1…TOC{depth} into styles.xml the first time. */
    private ensureStyles;
    /**
     * A bookmark allocator handing out a DISTINCT id/name per call.
     *
     * document.xml is only re-read once: a batch that mints many bookmarks before
     * saving would otherwise stamp the same `w:id`/`_Toc…` on all of them, leaving
     * every entry's PAGEREF pointing at the first heading.
     */
    private bookmarkSeries;
    /**
     * The name of a bookmark PAGEREF can already aim at, or null.
     *
     * Only `_Toc…`/`_Ref…` bookmarks count: Word's own `_GoBack` marks the last
     * edit position and moves, so an entry anchored to it would drift.
     */
    private existingAnchor;
    /** Wrap a heading paragraph in a bookmark so PAGEREF/hyperlink can find it. */
    private addAnchor;
    /** The `TOC` field instruction, mirroring the switches Word writes. */
    private buildInstrText;
    /** One pre-populated entry: `<heading text><dot leader tab><page number>`. */
    private buildEntryParagraph;
    /** Every field instruction in a block, concatenated (Word splits them across runs). */
    private instructionsIn;
    /**
     * True when the block opens a table of CONTENTS field.
     *
     * `TOC \c "Figure"` is a caption list (List of Figures / List of Tables) and is
     * deliberately excluded — replacing the contents must never eat one. A
     * `PAGEREF _Toc…` entry does not match either: the needle is `TOC \`.
     */
    private isTocField;
    /**
     * Insert a Table of Contents at the given BODY BLOCK index (default 0 = top).
     *
     * Collects every heading down to `headingDepth`, anchoring each one so the
     * entries link and paginate. Headings are recognised by style (`Heading1`,
     * `Titre 1`, `Title`) OR by their own outline level, so imported theses whose
     * headings carry no heading style are still collected — as long as something
     * marked them as headings (use `set_heading` / infer_structure first if not).
     */
    insertTOC(options?: TocOptions, index?: number): Promise<TocResult>;
    /**
     * Remove every Table of Contents — its title, the field, and all pre-populated
     * entry paragraphs. Caption lists (`TOC \c "Figure"`) are left alone.
     *
     * A TOC field SPANS paragraphs (begin … separate … entries … end), so this
     * walks from the paragraph holding the instruction to the one holding the
     * matching `fldChar end` rather than deleting a single paragraph.
     */
    removeTOC(): Promise<TocRemoval>;
    /** True when the document already carries a Table of Contents field. */
    hasTOC(): Promise<boolean>;
    /** Block indices that sit INSIDE any field (a real TOC, a caption list, …). */
    private fieldBlockIndices;
    /**
     * Does this paragraph read as a hand-typed contents line?
     *
     * Two signals, BOTH required: it ends in a page number, and it separates that
     * number from the title with a leader — typed dots, or a real tab. A tab
     * carries no text, so "المقدمة<tab>5" arrives as "المقدمة5"; the XML is what
     * says which it was.
     */
    private isTypedTocEntry;
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
    findTypedTOC(minEntries?: number): Promise<TypedTocSpan | null>;
    /**
     * Delete a hand-typed table of contents (title + every entry line).
     *
     * DESTRUCTIVE in a way {@link removeTOC} is not: these are paragraphs the
     * student wrote, not a field this engine generated. Callers must have the
     * student's agreement.
     */
    removeTypedTOC(span?: TypedTocSpan): Promise<TocRemoval>;
}

declare interface TableProperties {
    "w:tblStyle"?: {
        $: {
            "w:val": string;
        };
    };
    "w:tblW"?: {
        $: {
            "w:w": string;
            "w:type": string;
        };
    };
    "w:jc"?: {
        $: {
            "w:val": string;
        };
    };
    "w:tblInd"?: {
        $: {
            "w:w": string;
            "w:type": string;
        };
    };
    "w:tblBorders"?: any;
    "w:tblLayout"?: {
        $: {
            "w:type": string;
        };
    };
    "w:tblCellMar"?: any;
    "w:bidiVisual"?: Record<string, any>;
    "w:tblCaption"?: {
        $: {
            "w:val": string;
        };
    };
    "w:tblDescription"?: {
        $: {
            "w:val": string;
        };
    };
    [key: string]: any;
}

export declare interface TableRow {
    $?: Record<string, any>;
    "w:trPr"?: TableRowProperties;
    "w:tc"?: TableCell | TableCell[];
}

declare interface TableRowProperties {
    "w:trHeight"?: {
        $: {
            "w:val": string;
            "w:hRule"?: string;
        };
    };
    "w:tblHeader"?: Record<string, any>;
    "w:cantSplit"?: Record<string, any>;
    [key: string]: any;
}

export declare const TARGET_SPECS: Record<TextStyleTarget, TargetSpec>;

/** Per-target outcome of {@link TextStyleManager.apply}. */
export declare interface TargetReport {
    target: TextStyleTarget;
    /** The style id patched in `word/styles.xml` for this target. */
    styleId: string;
    /** True if the style did not exist and was created (Phase 1 or Phase 2). */
    styleCreated: boolean;
    /** True if the style's `w:rPr` was rewritten (already existed or just created). */
    styleTouched: boolean;
    /** Runs whose named property was stripped so the patched style shows through. */
    runsStripped: number;
    /** Runs the per-run pass could not parse and left untouched (Decision A). */
    runsSkipped: number;
    /** Runs written directly because their paragraph would not resolve to the target's style. */
    directWrites: number;
    /** Paragraphs (or table-cell paragraphs) considered for this target. */
    paragraphsAffected: number;
}

declare interface TargetSpec {
    /** Style id(s) this target's runs are expected to resolve to. */
    styleIds: string[];
    /** Minimal style to create when absent, via string surgery. */
    ensure?: EnsureStyleSpec;
    /** Defer creation to a manager that owns a RICH definition. See Decision B. */
    richEnsure?: "caption";
    /** Which OOXML part the target's RUN edits happen in. The style patch itself
     *  always lands in `word/styles.xml` regardless of this field. */
    part: "body" | "footnotes";
}

/** What a target's text is formatted with, and where that formatting comes from. */
export declare interface TargetStyleReport {
    target: TextStyleTarget;
    /** Paragraphs matched (table-cell paragraphs for `tables`). 0 ⇒ this part is absent. */
    paragraphs: number;
    /** Characters weighed. 0 ⇒ matched paragraphs exist but are empty. */
    characters: number;
    /** The style id `set_text_style` would patch for this target. */
    styleId: string;
    /** Is that style actually DEFINED in `styles.xml`? Dangling refs are normal here. */
    styleDefined: boolean;
    /** `w:basedOn` chain walked, nearest first — `["Heading1", "Normal"]`. */
    styleChain: string[];
    /** What the style chain alone resolves to, before any direct run formatting. */
    styleProps: ObservedRunProps;
    /** What the text ACTUALLY renders with, per property. The answer to the question. */
    effective: Record<FacetKey, StyleFacet>;
    /** True when any property has more than one value across this target. */
    mixed: boolean;
}

export declare interface TextBoxOptions {
    text: string;
    position?: ShapePosition;
    size?: ShapeSize;
    paragraphIndex?: number;
    fillColor?: string;
    borderColor?: string;
    floating?: boolean;
}

export declare interface TextCaptionMatch {
    kind: CaptionKind;
    /** The label exactly as typed, e.g. "Fig." or "الشكل رقم". */
    rawLabel: string;
    /** The number exactly as typed, e.g. "I-1" or "٣"; "" when unnumbered. */
    rawNumber: string;
    /** What sat between the number and the text, e.g. " : ". */
    separator: string;
    /** The caption wording after the label, number and separator. */
    text: string;
}

/**
 * Represents a text run in the document.
 * A run is a region of text with a common set of properties, such as formatting,
 * inline elements like breaks, tabs, fields, and drawings.
 *
 * @example
 * <w:r>
 *   <w:rPr>
 *     <w:b/>
 *     <w:i/>
 *   </w:rPr>
 *   <w:t xml:space="preserve">quick</w:t>
 *   <w:br/>
 *   <w:tab/>
 * </w:r>
 */
/**
 * Represents a single text node inside a run.
 */
declare interface TextNode {
    $?: {
        "xml:space"?: "preserve";
    };
    _: string;
}

/** {@link TextStyleReader.inspect} output. */
export declare interface TextStyleInspection {
    /** `styles.xml`'s `w:docDefaults` — the bottom rung, inherited by everything. */
    documentDefaults: ObservedRunProps;
    targets: TargetStyleReport[];
}

/**
 * Apply a font/size/bold/italic/colour change to named PARTS of a thesis
 * ("body", "heading3", "captions", "footnotes", …).
 *
 * Three strictly ordered phases per {@link apply}:
 *  1. Object-model style ENSURES — today, only `captions`, via CaptionManager's
 *     own rich `Caption` definition (Decision B).
 *  2. String-surgery style PATCHES — `StylesManager.setStyleRunProps` for every
 *     selected target's style id.
 *  3. Body/footnote RUN edits — strip the named property from runs that would
 *     resolve to the now-patched style; direct-write everything else.
 *
 * Phase 2 must never run after Phase 3 starts, and Phase 1 must never run
 * after Phase 2: `StylesManager.addStyle` round-trips the ENTIRE styles.xml
 * through an xml2js object model (parse → pretty-print), while
 * `setStyleRunProps` is string surgery on the same file. An object-model
 * write after string surgery can reformat or disturb it — so object-model
 * ensures come first, string surgery always comes last within styles.xml.
 */
export declare class TextStyleManager {
    private zip;
    private styles;
    private doc;
    constructor(zip: default_2);
    /**
     * @param targets    Target names (accepts the `"headings"` shorthand).
     * @param props      Which run properties to force, and to what value.
     * @param blockInfos `Doc.blocks()` output, index-aligned 1:1 with
     *                    `DocumentManager.getBlocks()`. Required because
     *                    `BodyBlock` alone doesn't carry `headingLevel`/`styleId`.
     */
    apply(targets: readonly string[], props: RunProps, blockInfos: readonly BlockInfo[]): Promise<TargetReport[]>;
    /** Phase 3 for the `footnotes` target — a separate part, `word/footnotes.xml`. */
    private applyFootnotes;
}

/**
 * Report what each named PART of a document is actually formatted with.
 *
 * @see TextStyleManager for the write side. Any change to `matchesTarget` or
 * `TARGET_SPECS` reaches both, which is the point: a read that scoped its
 * targets differently from the write would answer a question about text the
 * write then wouldn't touch.
 */
export declare class TextStyleReader {
    private zip;
    private doc;
    constructor(zip: default_2);
    /**
     * @param targets    Already-expanded targets (see `expandTargets`).
     * @param blockInfos `Doc.blocks()` output, index-aligned 1:1 with
     *                   `DocumentManager.getBlocks()` — `matchesTarget` needs the
     *                   `headingLevel`/`styleId` a raw `BodyBlock` doesn't carry.
     */
    inspect(targets: readonly TextStyleTarget[], blockInfos: readonly BlockInfo[]): Promise<TextStyleInspection>;
    /**
     * Walk every paragraph in a block (cell paragraphs included, at any depth for
     * a `tables` target) and tally each run's effective properties.
     */
    private collect;
    private report;
}

export declare type TextStyleTarget = "body" | "heading1" | "heading2" | "heading3" | "heading4" | "heading5" | "heading6" | "title" | "captions" | "lists" | "tables" | "footnotes";

export declare type TextStyleTargetInput = TextStyleTarget | "headings";

/**
 * Pass-through classifier kept for compatibility. Filters out pure-whitespace
 * "other" blocks would change semantics, so this returns every block as-is.
 */
export declare function toBlocks(bodyChildren: BodyBlock[]): BodyBlock[];

export declare interface TocOptions {
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

/** What {@link TableOfContentsManager.removeTOC} deleted. */
export declare interface TocRemoval {
    /** Blocks deleted across every table of contents found. */
    removed: number;
    /** Block index the FIRST removed table started at, or -1 if none. */
    at: number;
}

/** What {@link TableOfContentsManager.insertTOC} actually wrote. */
export declare interface TocResult {
    /** Block index the table now starts at. */
    atIndex: number;
    /** One per heading collected. */
    entries: number;
    /** Deepest heading level collected. */
    headingDepth: number;
    /** How many blocks a replaced table occupied (0 when none was replaced). */
    replaced: number;
}

/**
 * `sizePt` in POINTS → whole half-points, or throw.
 *
 * @param sizePt  A real font size in points, e.g. 14 or 11.5.
 * @param what    What the caller was setting, for the error message.
 */
export declare function toHalfPoints(sizePt: number, what?: string): number;

export declare class TrackedChangesManager {
    private zip;
    constructor(zip: default_2);
    private readDocument;
    private writeDocument;
    private getBodyParagraphs;
    private norm;
    private textFromRuns;
    private extractRevisionsFromParagraph;
    private acceptInsInParagraph;
    private rejectInsInParagraph;
    private acceptDelInParagraph;
    private rejectDelInParagraph;
    private acceptRPrChanges;
    private rejectRPrChanges;
    private acceptPPrChange;
    private rejectPPrChange;
    getRevisions(): Promise<RevisionEntry[]>;
    acceptAll(): Promise<void>;
    rejectAll(): Promise<void>;
    acceptRevision(id: number): Promise<void>;
    rejectRevision(id: number): Promise<void>;
}

/** Convert twips to centimetres */
export declare const twipsToCm: (twips: number) => number;

/** Convert twips to inches */
export declare const twipsToInches: (twips: number) => number;

/**
 * A table of contents the student TYPED BY HAND — ordinary paragraphs reading
 * "المقدمة .......... 5", not a `TOC` field. It is the usual state of an
 * imported thesis, and no field-based operation can see it.
 */
export declare interface TypedTocSpan {
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

/**
 * Rewrite a paragraph's `<w:pPr>` through `mutate`, which receives the current
 * inner XML ("" when the paragraph has no `w:pPr` yet) and returns the new one.
 * The result is canonicalised into `CT_PPr` order and written back; a `w:pPr`
 * is created (or dropped, if `mutate` empties it) as needed.
 *
 * The rest of the paragraph — every run, bookmark, hyperlink, field and
 * drawing — is copied through byte for byte.
 *
 * @throws {Error} if `xml` is not a `<w:p>` element, or if the resulting
 * `w:pPr` cannot be cleanly parsed (see `canonicalOrder.splitTopLevelElements`).
 */
export declare function updateParagraphProps(xml: string, mutate: (pPrInner: string) => string): string;

/** One value of one property, with the share of characters carrying it. */
export declare interface ValueShare {
    /** `null` means "the cascade specifies nothing here" — Word falls back to its own default. */
    value: string | number | boolean | null;
    /** Share of this target's characters, 0–1, rounded to 3 decimals. */
    share: number;
}

declare type xmlFile = {
    fileName: string;
    xml: string;
};

declare type xmlFile_2 = {
    fileName: string;
    xml: string;
};

export declare class ZipManager extends default_2 {
    constructor(filePathOrBuffer?: string | Buffer);
    static loadFromFile(filePath: string): Promise<ZipManager>;
    static loadFromBuffer(buffer: Buffer): Promise<ZipManager>;
    /**
     * Guarantee the convenience readers exist on a (possibly adm-zip-returned)
     * instance. Idempotent: if they're already present (a real ZipManager), it's a
     * no-op. See the constructor note on why this is necessary.
     */
    private static ensureReaders;
    getFileAsBuffer(entryName: string): Buffer | null;
    getFileAsString(entryName: string): string | null;
    fileExists(entryName: string): boolean;
    saveToFile(filePath: string): Promise<void>;
    toBuffer(): Buffer;
}

export { }
