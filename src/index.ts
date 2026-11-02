import { ZipManager } from "./utils/ZipManager";
import { ContentTypesManager } from "./core/PartsManagers/ContentTypesManager";
import DocumentManager from "./core/PartsManagers/DocumentManager";
import FooterManager from "@/core/PartsManagers/FooterManager";
import HeaderManager from "./core/PartsManagers/HeaderManager";
import RootRelManager from "./core/PartsManagers/RootRelManager";
import { RelManager } from "./core/PartsManagers/RelManager";
import { StylesManager } from "./core/PartsManagers/StylesManager";
import { NumberingManager } from "./core/PartsManagers/NumberingManager";
import { MetadataManager } from "./core/PartsManagers/MetadataManager";
import { MediaManager } from "./core/PartsManagers/MediaManager";
import { FootnoteManager } from "./core/PartsManagers/FootnoteManager";
import { EndnoteManager } from "./core/PartsManagers/EndnoteManager";
import { TableOfContentsManager } from "./core/PartsManagers/TableOfContentsManager";
import { CrossReferenceManager } from "./core/PartsManagers/CrossReferenceManager";
import { CitationManager } from "./core/PartsManagers/CitationManager";
import { PageLayoutManager } from "./core/PartsManagers/PageLayoutManager";
import { CaptionManager } from "./core/PartsManagers/CaptionManager";
import { CommentsManager } from "./core/PartsManagers/CommentsManager";
import { TrackedChangesManager } from "./core/PartsManagers/TrackedChangesManager";
import { SectionManager } from "./core/PartsManagers/SectionManager";
import { ShapeManager } from "./core/PartsManagers/ShapeManager";
import { MergeManager } from "./core/PartsManagers/MergeManager";
import { FormattingManager } from "./core/PartsManagers/FormattingManager";
import { FontTableManager } from "./core/PartsManagers/FontTableManager";
import type { PageNumberFormat } from "./core/PartsManagers/FooterManager";
import fs from "fs/promises";
import path from "path";

/** Options for {@link Mdocxengine.applyFrontMatterNumbering}. */
export interface FrontMatterNumberingOptions {
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
class Mdocxengine {
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
  fontTable: FontTableManager;

  private constructor(zip: ZipManager) {
    this.zip = zip;
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
    this.document = new DocumentManager(zip);
    this.footer = new FooterManager(zip);
    this.header = new HeaderManager(zip);
    this.rootRels = new RootRelManager(zip);
    this.styles = new StylesManager(zip);
    this.numbering = new NumberingManager(zip);
    this.metadata = new MetadataManager(zip);
    this.media = new MediaManager(zip);
    this.footnotes = new FootnoteManager(zip);
    this.endnotes = new EndnoteManager(zip);
    this.toc = new TableOfContentsManager(zip);
    this.crossRef = new CrossReferenceManager(zip);
    this.citations = new CitationManager(zip);
    this.pageLayout = new PageLayoutManager(zip);
    this.captions       = new CaptionManager(zip);
    this.comments       = new CommentsManager(zip);
    this.trackedChanges = new TrackedChangesManager(zip);
    this.sections       = new SectionManager(zip);
    this.shapes         = new ShapeManager(zip);
    this.merge          = new MergeManager(zip);
    this.formatting     = new FormattingManager(zip);
    this.fontTable      = new FontTableManager(zip, this.contentTypes);
  }

  static async loadFromFile(path: string) {
    const zm = await ZipManager.loadFromFile(path);
    return new Mdocxengine(zm);
  }

  // Load a document from an in-memory Buffer (e.g. bytes fetched from object
  // storage) without writing a temp file. Mirrors loadFromFile.
  static async loadFromBuffer(buffer: Buffer) {
    const zm = await ZipManager.loadFromBuffer(buffer);
    return new Mdocxengine(zm);
  }

  async saveToFile(outputPath: string) {
    const buf: Buffer = this.zip.toBuffer();
    await fs.mkdir(path.dirname(outputPath), { recursive: true });
    await fs.writeFile(outputPath, buf);
  }

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
  async applyFrontMatterNumbering(options: FrontMatterNumberingOptions): Promise<void> {
    const {
      bodyStartParaIndex,
      frontMatterFormat = "lowerRoman",
      bodyFormat = "decimal",
      alignment = "center",
      bodyStartAt = 1,
    } = options;

    // 1. Section break before the body → two sections (front matter | body).
    await this.pageLayout.insertBreak("nextPage", bodyStartParaIndex);

    // 2. Front-matter footer (roman by default), 3. body footer (decimal).
    const front = await this.footer.addFooter("", "default", undefined, { registerInSectPr: false });
    await this.footer.insertPageNumber(front.footerPath, { alignment, format: frontMatterFormat });
    const body = await this.footer.addFooter("", "default", undefined, { registerInSectPr: false });
    await this.footer.insertPageNumber(body.footerPath, { alignment, format: bodyFormat });

    // 4. Attach a footer per section.
    const sections = await this.sections.getSections();
    if (sections.length >= 2) {
      await this.sections.setSectionFooter(0, front.relId, "default");
      await this.sections.setSectionFooter(1, body.relId, "default");
    }

    // 5. Restart body numbering (formatPageNumbers targets the body/final section).
    await this.footer.formatPageNumbers({ format: bodyFormat, startAt: bodyStartAt });
  }

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
  async replaceImageAtBlock(
    index: number,
    buffer: Buffer,
    ext: string,
    opts?: { widthPx?: number; heightPx?: number },
  ): Promise<void> {
    const blocks = await this.document.getBlocks();
    const block = blocks[index];
    if (!block || block.kind !== "paragraph" || !block.xml.includes("<w:drawing>")) {
      throw new Error(`replaceImageAtBlock: block ${index} is not an image`);
    }
    const embed = /r:embed="([^"]+)"/.exec(block.xml);
    if (!embed) throw new Error(`replaceImageAtBlock: block ${index} has no embedded image`);
    const oldRelId = embed[1];

    const target = await this.rels.getTarget(oldRelId); // e.g. "media/image1.jpeg"
    const oldExt = (target?.split(".").pop() ?? "").toLowerCase();
    const newExt = ext.toLowerCase().replace(/^\./, "");

    let xml = block.xml;
    if (target && oldExt === newExt) {
      // Same format → overwrite bytes in place; block XML unchanged.
      const name = target.split("/").pop()!;
      this.media.replaceImage(name, buffer);
    } else {
      // Format change (or unresolvable target) → new part + repoint the relationship.
      const { relId } = await this.media.insertImage(buffer, newExt);
      xml = xml.replace(/r:embed="[^"]+"/, `r:embed="${relId}"`);
    }

    if (opts?.widthPx && opts?.heightPx && opts.widthPx > 0 && opts.heightPx > 0) {
      xml = rescaleDrawingExtent(xml, opts.widthPx, opts.heightPx);
    }

    if (xml !== block.xml) {
      blocks[index].xml = xml;
      await this.document.saveBlocks(blocks);
    }
  }
}

// Rescale a drawing's display box to a new aspect ratio, preserving each extent's
// current display WIDTH (cx) and recomputing its height (cy) from wPx/hPx. Rewrites
// both the inline `<wp:extent>` and the picture shape's `<a:ext>` (they carry the
// same EMU size for an inline image). Leaves `<a:off>` (x/y offsets) untouched.
function rescaleDrawingExtent(xml: string, wPx: number, hPx: number): string {
  const ratio = hPx / wPx;
  return xml.replace(
    /(<(?:wp:extent|a:ext)\b[^>]*\bcx=")(\d+)("[^>]*\bcy=")(\d+)(")/g,
    (_m, pre, cx, mid, _cy, post) => `${pre}${cx}${mid}${Math.round(Number(cx) * ratio)}${post}`,
  );
}

export {
  Mdocxengine,
  ZipManager,
  RelManager,
  ContentTypesManager,
  DocumentManager,
  FooterManager,
  HeaderManager,
  RootRelManager,
  StylesManager,
  NumberingManager,
  MetadataManager,
  MediaManager,
  FootnoteManager,
  EndnoteManager,
  TableOfContentsManager,
  CrossReferenceManager,
  CitationManager,
  PageLayoutManager,
  CaptionManager,
  CommentsManager,
  TrackedChangesManager,
  SectionManager,
  ShapeManager,
  MergeManager,
  FormattingManager,
};
export { FontTableManager, appendRel } from "./core/PartsManagers/FontTableManager";
export type { EmbedFontInput, EmbedFontResult, FontVariant } from "./core/PartsManagers/FontTableManager";
export { readFsType, embeddingPermission, canEmbed } from "./core/fonts/fsType";
export type { EmbeddingPermission } from "./core/fonts/fsType";
export { obfuscateFont, fontKeyBytes, newFontKey } from "./core/fonts/obfuscate";
export type {
  DocumentFormatting,
  DocumentMargins,
  FormattingResult,
} from "./core/PartsManagers/FormattingManager";
export type { ParagraphOptions } from "./core/files/paragraph/index";
export type { FromGridOptions, TableBorderOptions, BorderSide } from "./core/files/table/index";
export {
  renderMarkdownBlocks,
  stripInlineMarkdown,
} from "./core/markdown/index";
export type {
  MarkdownBlock,
  MarkdownAlign,
  MarkdownRenderCtx,
  RenderedMarkdown,
  RenderedTable,
  RenderedImage,
  RenderedChunk,
  BlockRenderer,
} from "./core/markdown/index";
export type { PageNumberFormat } from "./core/PartsManagers/FooterManager";
export { DEFAULT_STYLE_ALIASES } from "./core/PartsManagers/MergeManager";
export type { AppendOptions } from "./core/PartsManagers/MergeManager";
export { default as Paragraph } from "./core/files/paragraph/index";
export { Run } from "./core/files/paragraph/Run";
export { Table } from "./core/files/table/index";
export {
  parseOrderedDoc,
  buildOrderedDoc,
  toBlocks,
  makeParagraphXml,
  makeParagraphNode,
  makeParagraphXmlLike,
  makeParagraphNodeLike,
  paragraphPropsXml,
  paragraphRunPropsXml,
  makeTableXml,
  makeTableNode,
  makeDrawingParagraphXml,
  makeDrawingParagraphNode,
  nextDrawingId,
  paragraphText,
  paragraphStyleId,
  headingLevelFromStyleId,
  paragraphOutlineLevel,
  paragraphHeadingLevel,
  paragraphAlignment,
  paragraphFontSizePt,
  paragraphIsBold,
  makeStyledParagraphXml,
  makeStyledParagraphNode,
  setParagraphText,
  nodeTag,
} from "./core/files/body/OrderedBody";
export type { StyledParagraphOptions } from "./core/files/body/OrderedBody";
export { Doc } from "./Doc";
export type {
  DocMap,
  OutlineNode,
  BlockInfo,
  TableInfo,
  ImageInfo,
  ParagraphFormat,
  DetailedBlockInfo,
  HeadingPattern,
  InferOutlineOptions,
  InferredHeading,
  FooterOptions,
  SectionEditResult,
  SectionInfo,
  SectionPageGeometry,
  ChromeDrawing,
  ChromeDrawingAxis,
  ChromeDuotone,
} from "./Doc";
export type { BodyBlock, BlockKind } from "./core/files/body/OrderedBody";
export { splitDocument, assembleDocument, isEditableBlock } from "./core/files/body/OrderedBody";
export type { SplitDocument } from "./core/files/body/OrderedBody";
export type { StyleEntry } from "./core/PartsManagers/StylesManager";
export type { NumberingDefinition } from "./core/PartsManagers/NumberingManager";
export type { CoreProperties, AppProperties } from "./core/PartsManagers/MetadataManager";
export { EMU_PER_PIXEL, emuToPixels, pixelsToEmu } from "./core/PartsManagers/MediaManager";
export type { ImageEntry, InlineImage } from "./core/PartsManagers/MediaManager";
export type { RelationshipEntry } from "./core/PartsManagers/RelManager";
export type { TableObject, TableRow, TableCell } from "./core/files/table/types";
export type { FootnoteEntry } from "./core/PartsManagers/FootnoteManager";
export type { EndnoteEntry } from "./core/PartsManagers/EndnoteManager";
export type { TocOptions, TocResult, TocRemoval, TypedTocSpan } from "./core/PartsManagers/TableOfContentsManager";
export type { BookmarkEntry } from "./core/PartsManagers/CrossReferenceManager";
export type { CitationSource } from "./core/PartsManagers/CitationManager";
export type {
  PageSizePreset,
  MarginPreset,
  Orientation,
  SectionBreakType,
  PageSize,
  PageMargins,
  ColumnOptions,
  LineNumberingOptions,
} from "./core/PartsManagers/PageLayoutManager";
export { PAGE_SIZES, MARGIN_PRESETS, inchesToTwips, cmToTwips, twipsToInches, twipsToCm } from "./core/PartsManagers/PageLayoutManager";
export { applyBodyPageLayout } from "./core/files/body/pageLayout";
export type { BodyPageLayoutOpts } from "./core/files/body/pageLayout";
export type { SectPrPageBorderOptions } from "./core/files/body/sectPr";
export type {
  CaptionOptions,
  CaptionNumberingOptions,
  CaptionNumberFormat,
  CaptionPosition,
  ChapterSeparator,
  CaptionEntry,
  CaptionKind,
  ConvertTextCaptionsOptions,
  ConvertTextCaptionsResult,
  ConvertedCaption,
  SkippedTextCaption,
  SkipReason,
  TextCaptionMatch,
} from "./core/PartsManagers/CaptionManager";
export { seqIdentifier, formatSeqNumber, parseCaptionParagraph, parseTextCaption } from "./core/PartsManagers/CaptionManager";
export type { CommentEntry } from "./core/PartsManagers/CommentsManager";
export type { RevisionEntry, RevisionType } from "./core/PartsManagers/TrackedChangesManager";
export type {
  SectionEntry,
  SectionLayout,
  SectionPageSize,
  SectionMargins,
  SectionHeaderFooterRef,
} from "./core/PartsManagers/SectionManager";
export type {
  ShapeEntry,
  TextBoxOptions,
  InsertShapeOptions,
  InsertLineOptions,
  ShapeType,
  ShapePosition,
  ShapeSize,
} from "./core/PartsManagers/ShapeManager";
export { EMU_PER_INCH, EMU_PER_CM } from "./core/PartsManagers/ShapeManager";
export {
  TextStyleManager,
  TARGET_SPECS,
  expandTargets,
  eachParagraphIn,
  type TextStyleTarget,
  type TextStyleTargetInput,
  type TargetReport,
} from "./core/PartsManagers/TextStyleManager";
export {
  TextStyleReader,
  parseRunProps,
  parseStylesIndex,
  resolveStyleChain,
  type ObservedRunProps,
  type FacetKey,
  type ValueShare,
  type StyleFacet,
  type TargetStyleReport,
  type TextStyleInspection,
} from "./core/PartsManagers/TextStyleReader";
export { type RunProps, mergeRunProps } from "./core/ooxml/runProps";
export { toHalfPoints, MIN_FONT_SIZE_PT, MAX_FONT_SIZE_PT } from "./core/ooxml/hps";
// Single-paragraph OOXML surgery: Word's pagination properties (keep with next,
// keep lines together, widow/orphan control, page break before) and DIRECT run
// formatting for ONE paragraph — what "make THIS heading 48pt" needs, where a
// style change would resize every heading in the thesis.
export {
  applyParagraphPagination,
  applyRunPropsToParagraph,
  insertAfterParagraphOpen,
  isParagraphXml,
  readParagraphPagination,
  readParagraphProps,
  updateParagraphProps,
  type ParagraphPagination,
  type ParagraphRunStyleResult,
} from "./core/ooxml/paragraphProps";
export { CT_PPR_ORDER, canonicalizeParagraphProps } from "./core/ooxml/canonicalOrder";
// Picture PLACEMENT: where a <w:drawing> sits on the page. Vertical position is
// not a property but a CONTAINER — an in-flow `wp:inline` picture has to become
// a floating `wp:anchor` before "centre it on the page" can even be expressed.
export {
  applyDrawingLayout,
  readDrawingLayout,
  type DrawingLayout,
  type DrawingLayoutResult,
  type DrawingPlacement,
  type DrawingRelativeTo,
  type DrawingWrap,
} from "./core/ooxml/drawingLayout";
export {
  applyStyleRunPropsToXml,
  buildParagraphStyleXml,
  type EnsureStyleSpec,
} from "./core/PartsManagers/StylesManager";

// ── Document doctor ──────────────────────────────────────────────────────────
// Package-level validation and repair for the corruption classes that make Word
// refuse a .docx. Lives here rather than in a consumer because it is pure OOXML
// knowledge — no thesis, database or storage concepts. See src/doctor/.
export {
  inspectDocx,
  checkDocxBuffer,
  repairDocxBuffer,
  clearFalseDataDescriptors,
  firstXmlError,
  summarize,
} from "./doctor/index";
export type {
  DoctorReport,
  Finding,
  Severity,
  DocxZip,
  InspectOptions,
  RepairResult,
} from "./doctor/index";
