import * as XmlUtils from "@/utils/xmlUtils";
import AdmZip from "adm-zip";
import { splitDocument, assembleDocument } from "@/core/files/body/OrderedBody";
import { declareUsedPrefixes } from "@/core/ooxml/namespaces";
import {
  upsertSectPrReference,
  upsertSectPrPageNumbering,
  upsertParagraphSectPr,
  removeParagraphSectPr,
  editSectPrWithinParagraph,
  editBodySectPr,
  upsertSectPrVAlign,
  upsertSectPrPageBorders,
  upsertSectPrGeometry,
  removeSectPrTitlePg,
  removeSectPrHeaderFooterRefs,
  extractSectPrGeometry,
  type SectPrGeometry,
  type SectPrPageBorderOptions,
} from "@/core/files/body/sectPr";

const DOCUMENT_PATH = "word/document.xml";

export interface SectionPageSize {
  width: number;
  height: number;
  orientation?: "portrait" | "landscape";
}

export interface SectionMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header?: number;
  footer?: number;
  gutter?: number;
}

export interface SectionLayout {
  type?: "nextPage" | "continuous" | "evenPage" | "oddPage";
  pageSize?: SectionPageSize;
  margins?: SectionMargins;
}

export interface SectionHeaderFooterRef {
  relId: string;
  type: "default" | "first" | "even";
}

export interface SectionEntry {
  index: number;
  isFinal: boolean;
  type?: string;
  pageSize?: SectionPageSize;
  margins?: SectionMargins;
  headerRefs: SectionHeaderFooterRef[];
  footerRefs: SectionHeaderFooterRef[];
  /** Parsed w:pgNumType, when the section sets one (page-number format/restart). */
  pageNumberType?: { format: string; start?: number };
  paragraphIndex?: number;
}

export class SectionManager {
  private zip: AdmZip;

  constructor(zip: AdmZip) {
    this.zip = zip;
  }

  // ─── Internal read/write ──────────────────────────────────────────────────

  private async readDocument(): Promise<any> {
    const xml = this.zip.readAsText(DOCUMENT_PATH);
    if (!xml) throw new Error("word/document.xml not found");
    return XmlUtils.parseXml(xml);
  }

  private async writeDocument(obj: any): Promise<void> {
    const xml = XmlUtils.buildXml(obj["w:document"], {
      rootName: "w:document",
      headless: false,
      pretty: true,
    });
    this.zip.addFile(DOCUMENT_PATH, Buffer.from(xml, "utf-8"));
  }

  private getBodyParagraphs(obj: any): any[] {
    const raw = obj?.["w:document"]?.["w:body"]?.["w:p"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  private norm(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
  }

  // ─── SectPr parsing ───────────────────────────────────────────────────────

  private parseSectPr(sectPr: any): Omit<SectionEntry, "index" | "isFinal" | "paragraphIndex"> {
    const type = sectPr?.["w:type"]?.$?.["w:val"];

    let pageSize: SectionPageSize | undefined;
    const pgSz = sectPr?.["w:pgSz"];
    if (pgSz) {
      pageSize = {
        width:       parseInt(pgSz.$?.["w:w"] ?? "12240", 10),
        height:      parseInt(pgSz.$?.["w:h"] ?? "15840", 10),
        orientation: pgSz.$?.["w:orient"] === "landscape" ? "landscape" : "portrait",
      };
    }

    let margins: SectionMargins | undefined;
    const pgMar = sectPr?.["w:pgMar"];
    if (pgMar) {
      margins = {
        top:    parseInt(pgMar.$?.["w:top"]    ?? "1440", 10),
        bottom: parseInt(pgMar.$?.["w:bottom"] ?? "1440", 10),
        left:   parseInt(pgMar.$?.["w:left"]   ?? "1440", 10),
        right:  parseInt(pgMar.$?.["w:right"]  ?? "1440", 10),
        header: pgMar.$?.["w:header"] !== undefined ? parseInt(pgMar.$["w:header"], 10) : undefined,
        footer: pgMar.$?.["w:footer"] !== undefined ? parseInt(pgMar.$["w:footer"], 10) : undefined,
        gutter: pgMar.$?.["w:gutter"] !== undefined ? parseInt(pgMar.$["w:gutter"], 10) : undefined,
      };
    }

    const headerRefs: SectionHeaderFooterRef[] = this.norm(sectPr?.["w:headerReference"]).map(
      (h: any) => ({ relId: h.$?.["r:id"] ?? "", type: (h.$?.["w:type"] ?? "default") as any }),
    );

    const footerRefs: SectionHeaderFooterRef[] = this.norm(sectPr?.["w:footerReference"]).map(
      (f: any) => ({ relId: f.$?.["r:id"] ?? "", type: (f.$?.["w:type"] ?? "default") as any }),
    );

    const pgNumType = sectPr?.["w:pgNumType"];
    const pageNumberType = pgNumType
      ? {
          format: (pgNumType.$?.["w:fmt"] as string) ?? "decimal",
          ...(pgNumType.$?.["w:start"] !== undefined
            ? { start: parseInt(pgNumType.$["w:start"], 10) }
            : {}),
        }
      : undefined;

    return { type, pageSize, margins, headerRefs, footerRefs, pageNumberType };
  }

  private applyLayout(sectPr: any, layout: SectionLayout): void {
    if (layout.type !== undefined) {
      sectPr["w:type"] = { $: { "w:val": layout.type } };
    }
    if (layout.pageSize) {
      const s = layout.pageSize;
      sectPr["w:pgSz"] = {
        $: {
          "w:w":      String(s.width),
          "w:h":      String(s.height),
          ...(s.orientation ? { "w:orient": s.orientation } : {}),
        },
      };
    }
    if (layout.margins) {
      const m = layout.margins;
      sectPr["w:pgMar"] = {
        $: {
          "w:top":    String(m.top),
          "w:bottom": String(m.bottom),
          "w:left":   String(m.left),
          "w:right":  String(m.right),
          ...(m.header !== undefined ? { "w:header": String(m.header) } : {}),
          ...(m.footer !== undefined ? { "w:footer": String(m.footer) } : {}),
          ...(m.gutter !== undefined ? { "w:gutter": String(m.gutter) } : {}),
        },
      };
    }
  }

  // ─── Section lookup ───────────────────────────────────────────────────────

  private async findSectPr(
    sectionIndex: number,
  ): Promise<{ obj: any; sectPr: any; isBody: boolean }> {
    const obj        = await this.readDocument();
    const paragraphs = this.getBodyParagraphs(obj);
    const body       = obj["w:document"]["w:body"];

    // Collect intermediate section sectPrs (in paragraph order)
    const intermediateParagraphIndices: number[] = [];
    for (let i = 0; i < paragraphs.length; i++) {
      if (paragraphs[i]?.["w:pPr"]?.["w:sectPr"]) {
        intermediateParagraphIndices.push(i);
      }
    }

    const totalSections = intermediateParagraphIndices.length + 1; // +1 for final

    if (sectionIndex < 0 || sectionIndex >= totalSections) {
      throw new Error(`Section index ${sectionIndex} out of range (0–${totalSections - 1})`);
    }

    if (sectionIndex < intermediateParagraphIndices.length) {
      const pIdx   = intermediateParagraphIndices[sectionIndex];
      const sectPr = paragraphs[pIdx]["w:pPr"]["w:sectPr"];
      return { obj, sectPr, isBody: false };
    }

    // Final section
    return { obj, sectPr: body["w:sectPr"], isBody: true };
  }

  // ─── Public API ───────────────────────────────────────────────────────────

  public async getSections(): Promise<SectionEntry[]> {
    const obj        = await this.readDocument();
    const paragraphs = this.getBodyParagraphs(obj);
    const body       = obj["w:document"]["w:body"];
    const sections:  SectionEntry[] = [];
    let   sectionIdx = 0;

    for (let i = 0; i < paragraphs.length; i++) {
      const sectPr = paragraphs[i]?.["w:pPr"]?.["w:sectPr"];
      if (sectPr) {
        sections.push({
          index:          sectionIdx++,
          isFinal:        false,
          paragraphIndex: i,
          ...this.parseSectPr(sectPr),
        });
      }
    }

    const bodySectPr = body["w:sectPr"];
    sections.push({
      index:   sectionIdx,
      isFinal: true,
      ...this.parseSectPr(bodySectPr ?? {}),
    });

    return sections;
  }

  // Byte-safe string edit of the body: locate the Nth `<w:p>` and inject/remove a
  // `<w:sectPr>` in its `<w:pPr>`, leaving every other block (tables, drawings,
  // the final sectPr) byte-identical. The xml2js full-body rebuild used elsewhere
  // would reorder tables; this never touches them.
  public async addSectionBreak(
    paragraphIndex: number,
    type: "nextPage" | "continuous" | "evenPage" | "oddPage",
    opts: { copyBodyGeometry?: boolean } = {},
  ): Promise<void> {
    const xml = this.zip.readAsText(DOCUMENT_PATH);
    if (!xml) throw new Error("word/document.xml not found");
    const split = splitDocument(xml);
    const paraBlocks = split.blocks.filter((b) => b.kind === "paragraph");
    if (!paraBlocks.length) return;
    const idx = Math.max(0, Math.min(paragraphIndex, paraBlocks.length - 1));
    // A bare `<w:type>` makes the section fall back to Word's default page
    // (Letter), so the break carries the body section's page size and margins.
    const bodySect = split.blocks.find((b) => b.kind === "sectPr");
    const geometry = opts.copyBodyGeometry === false || !bodySect ? "" : extractSectPrGeometry(bodySect.xml);
    const sectPr = `<w:sectPr><w:type w:val="${type}"/>${geometry}</w:sectPr>`;
    paraBlocks[idx].xml = upsertParagraphSectPr(paraBlocks[idx].xml, sectPr);
    this.zip.addFile(DOCUMENT_PATH, Buffer.from(assembleDocument(split), "utf-8"));
  }

  public async removeSectionBreak(paragraphIndex: number): Promise<void> {
    const xml = this.zip.readAsText(DOCUMENT_PATH);
    if (!xml) throw new Error("word/document.xml not found");
    const split = splitDocument(xml);
    const paraBlocks = split.blocks.filter((b) => b.kind === "paragraph");
    const blk = paraBlocks[paragraphIndex];
    if (!blk) return;
    blk.xml = removeParagraphSectPr(blk.xml);
    this.zip.addFile(DOCUMENT_PATH, Buffer.from(assembleDocument(split), "utf-8"));
  }

  public async setSectionLayout(sectionIndex: number, layout: SectionLayout): Promise<void> {
    const { obj, sectPr, isBody } = await this.findSectPr(sectionIndex);
    this.applyLayout(sectPr, layout);
    if (isBody) obj["w:document"]["w:body"]["w:sectPr"] = sectPr;
    await this.writeDocument(obj);
  }

  /**
   * Set ONE section's page size and margins (`pgSz` / `pgMar`), byte-safe: only
   * the named attributes change. Width/height are NOT swapped for `orient`.
   */
  public async setSectionGeometry(sectionIndex: number, geometry: SectPrGeometry): Promise<void> {
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrGeometry(s, geometry));
  }

  /** Remove `w:titlePg` ("different first page") from ONE section. */
  public async removeTitlePg(sectionIndex: number): Promise<void> {
    await this.editSectionSectPr(sectionIndex, removeSectPrTitlePg);
  }

  /** Drop every header/footer reference from ONE section's sectPr (the parts are left alone). */
  public async removeSectionHeaderFooterRefs(sectionIndex: number): Promise<void> {
    await this.editSectionSectPr(sectionIndex, removeSectPrHeaderFooterRefs);
  }

  public async setSectionHeader(
    sectionIndex: number,
    relId: string,
    type: "default" | "first" | "even" = "default",
  ): Promise<void> {
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrReference(s, "header", type, relId));
    this.declareRelationshipPrefix();
  }

  public async setSectionFooter(
    sectionIndex: number,
    relId: string,
    type: "default" | "first" | "even" = "default",
  ): Promise<void> {
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrReference(s, "footer", type, relId));
    this.declareRelationshipPrefix();
  }

  /**
   * Set ONE section's page-number format and/or restart value (`<w:pgNumType>`)
   * — e.g. roman for the front matter, arabic restarting at 1 for the body.
   *
   * `FooterManager.formatPageNumbers` only ever reaches the BODY sectPr, so it
   * cannot express per-section numbering; this can. Omit `start` to continue
   * the previous section's sequence.
   */
  public async setSectionPageNumbering(
    sectionIndex: number,
    opts: { format?: string; start?: number },
  ): Promise<void> {
    if (opts.format === undefined && opts.start === undefined) return;
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrPageNumbering(s, opts));
  }

  /**
   * Vertically align a section's page content (`<w:vAlign>`). "center" is what
   * puts a divider page's title in the true middle of the page.
   */
  public async setSectionVerticalAlign(
    sectionIndex: number,
    vAlign: "top" | "center" | "both" | "bottom",
  ): Promise<void> {
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrVAlign(s, vAlign));
  }

  /**
   * Draw a page border on all four edges of ONE section (`<w:pgBorders>`) — the
   * `frame` divider family. Colour must be 6-hex (leading '#' tolerated):
   * ST_HexColor admits nothing else, and an invalid value here is the class of
   * defect that trips Word's "unreadable content" repair dialog.
   */
  public async setSectionPageBorders(
    sectionIndex: number,
    opts: SectPrPageBorderOptions,
  ): Promise<void> {
    const hex = opts.color.replace(/^#/, "").toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(hex)) {
      throw new Error(`setSectionPageBorders: color must be a 6-digit hex value, got '${opts.color}'`);
    }
    if (!Number.isFinite(opts.widthPt) || opts.widthPt <= 0) {
      throw new Error(`setSectionPageBorders: widthPt must be a positive number, got '${opts.widthPt}'`);
    }
    await this.editSectionSectPr(sectionIndex, (s) => upsertSectPrPageBorders(s, opts));
  }

  /**
   * Apply `transform` to the `<w:sectPr>` of section `sectionIndex` — byte-safe
   * string surgery. Intermediate sections live inside a paragraph's `<w:pPr>`;
   * the final section is the body's `<w:sectPr>` (created if absent). Sections
   * are counted in document order: paragraph-level sectPrs first, final last —
   * the same order as getSections().
   */
  /** A reference is written as `r:id`; a root that declares only `xmlns:w` would then be refused by Word. */
  private declareRelationshipPrefix(): void {
    const xml = this.zip.readAsText(DOCUMENT_PATH);
    if (!xml) return;
    const next = declareUsedPrefixes(xml);
    if (next !== xml) this.zip.addFile(DOCUMENT_PATH, Buffer.from(next, "utf-8"));
  }

  private async editSectionSectPr(sectionIndex: number, transform: (sectPrXml: string) => string): Promise<void> {
    const xml = this.zip.readAsText(DOCUMENT_PATH);
    if (!xml) throw new Error("word/document.xml not found");
    const split = splitDocument(xml);
    const interParas = split.blocks.filter((b) => b.kind === "paragraph" && /<w:sectPr\b/.test(b.xml));
    const total = interParas.length + 1; // + final body section

    if (sectionIndex < 0 || sectionIndex >= total) {
      throw new Error(`Section index ${sectionIndex} out of range (0–${total - 1})`);
    }

    if (sectionIndex < interParas.length) {
      const blk = interParas[sectionIndex];
      blk.xml = editSectPrWithinParagraph(blk.xml, transform);
      this.zip.addFile(DOCUMENT_PATH, Buffer.from(assembleDocument(split), "utf-8"));
    } else {
      // Final/body section.
      this.zip.addFile(DOCUMENT_PATH, Buffer.from(editBodySectPr(xml, transform), "utf-8"));
    }
  }
}
