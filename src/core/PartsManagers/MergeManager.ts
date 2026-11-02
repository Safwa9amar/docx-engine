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
 *   - styles   → retargeted to this document's styles by name (caller-supplied
 *                map); the combine flow normalizes everything to one profile, so
 *                source style DEFINITIONS are intentionally not imported.
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
import { FootnoteManager } from "./FootnoteManager";
import { NumberingManager } from "./NumberingManager";
import { RelManager } from "./RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import { parseXml } from "@/utils/xmlUtils";
import { carryRootNamespaces } from "@/utils/rootNamespaces";
import type { BodyBlock } from "@/core/files/body/OrderedBody";
// Imported lazily inside appendDocument() to avoid an index.ts ↔ MergeManager cycle.
import type { Mdocxengine as MdocxengineType } from "@/index";

const SRC_RELS_PATH = "word/_rels/document.xml.rels";
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
  private numbering: NumberingManager;
  private rels: RelManager;
  private contentTypes: ContentTypesManager;

  constructor(zip: AdmZip) {
    this.zip = zip;
    this.document = new DocumentManager(zip);
    this.media = new MediaManager(zip);
    this.footnotes = new FootnoteManager(zip);
    this.numbering = new NumberingManager(zip);
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
  }

  /**
   * Copy `sourceBuffer`'s body into this document, fully remapped, appended after
   * the existing body (before the trailing sectPr, handled by saveBlocks).
   */
  async appendDocument(sourceBuffer: Buffer, opts: AppendOptions = {}): Promise<void> {
    const { Mdocxengine } = await import("@/index");
    const source = (await (Mdocxengine as typeof MdocxengineType).loadFromBuffer(
      sourceBuffer,
    )) as MdocxengineType;

    const srcBlocks = await source.document.getBlocks();
    const remapped = await this.remapBlocks(source, srcBlocks, opts.styleMap ?? {});

    const existing = await this.document.getBlocks();
    const prefix: BodyBlock[] = [];
    if (opts.startOnNewPage) prefix.push(pageBreakBlock());
    if (opts.leadingBlocks?.length) prefix.push(...opts.leadingBlocks);

    await this.document.saveBlocks([...existing, ...prefix, ...remapped]);
    this.carryNamespaces(source);
  }

  /**
   * The copied blocks/footnotes/numbering use prefixes (`w14`, `mc`, `wps`, `r`…)
   * the SOURCE root declared. Bind them on the matching TARGET root, or the part
   * has unbound prefixes and Word refuses to open it. A prefix bound to a
   * different URI is a real conflict: throw rather than silently mis-bind.
   */
  private carryNamespaces(source: MdocxengineType): void {
    for (const part of ["word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/numbering.xml"]) {
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
  ): Promise<BodyBlock[]> {
    const scan = blocks.map((b) => b.xml).join("\n");
    const srcRels = await this.readSourceRels(source);
    const mediaMap = await this.buildMediaMap(source, srcRels, scan);
    const hyperlinkMap = await this.buildHyperlinkMap(srcRels, scan);
    const footnoteMap = await this.buildFootnoteMap(source, scan);
    const numberingMap = await this.buildNumberingMap(source, scan);
    const chartMap = await this.buildChartMap(source, srcRels, scan);
    const vmlOleMap = await this.buildVmlOleMap(source, srcRels, scan);
    const ids = this.planIds(scan);

    return blocks.map((b) => {
      let xml = b.xml;
      xml = this.applyAttrMap(xml, ["r:embed", "r:link"], mediaMap);
      xml = this.applyHyperlinkMap(xml, hyperlinkMap);
      xml = this.applyFootnoteRefMap(xml, footnoteMap);
      xml = this.applyNumIdMap(xml, numberingMap);
      xml = this.applyChartMap(xml, chartMap);
      xml = this.applyVmlOleMap(xml, vmlOleMap);
      xml = this.applyIdPlan(xml, ids);
      xml = this.applyStyleMap(xml, styleMap);
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
   * Everything else in the copied `sectPr` — page size, margins, break type — is
   * preserved.
   */
  private stripSectionChrome(xml: string): string {
    if (!xml.includes("Reference")) return xml;
    return xml
      .replace(/<w:(?:header|footer)Reference\b[^>]*\/>/g, "")
      .replace(/<w:(header|footer)Reference\b[^>]*>[\s\S]*?<\/w:\1Reference>/g, "");
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

  /** Copy each image referenced by the blocks; return source-rId → new-rId. */
  private async buildMediaMap(
    source: MdocxengineType,
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refRIds = new Set(
      [...blocksXml.matchAll(/r:(?:embed|link)="([^"]+)"/g)].map((m) => m[1]),
    );
    if (refRIds.size === 0) return {};

    const map: Record<string, string> = {};
    for (const rId of refRIds) {
      const target = srcRels[rId]?.target;
      if (!target) continue; // not a media rel (or dangling)
      const name = target.replace(/^.*\//, ""); // "media/image1.png" → "image1.png"
      const buf = source.media.extractImage(name);
      if (!buf) continue;
      const ext = (/\.([a-zA-Z0-9]+)$/.exec(target)?.[1] ?? "png").toLowerCase();
      const { relId } = await this.media.insertImage(buf, ext);
      map[rId] = relId;
    }
    return map;
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

  /** Attributes on `<v:imagedata>` that name a relationship. */
  private static readonly VML_REL_ATTRS = ["r:id", "r:href", "o:relid"];

  /** Copy each VML image and OLE object referenced by the blocks; return
   *  source-rId → new-rId. */
  private async buildVmlOleMap(
    source: MdocxengineType,
    srcRels: Record<string, { type: string; target: string; targetMode?: string }>,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refs: { tag: string; attr: string; rId: string; ole: boolean }[] = [];
    for (const m of blocksXml.matchAll(/<(v:imagedata|o:OLEObject)\b[^>]*>/g)) {
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
    return xml.replace(/<(?:v:imagedata|o:OLEObject)\b[^>]*>/g, (tag) =>
      tag.replace(/\b((?:r:id|r:href|o:relid)=")([^"]+)(")/g, (full, a, val, c) =>
        idMap[val] ? `${a}${idMap[val]}${c}` : full,
      ),
    );
  }

  // ─── Unique ids: drawing ids + bookmarks ─────────────────────────────────────
  //
  // `wp:docPr/@id` and `w:bookmarkStart/@w:id` must be unique across the merged
  // document, and a bookmark NAME must be unique too. Two split parts that both
  // start at id 1 / `_Toc123` otherwise give Word a repair prompt (or make a
  // cross-reference jump to the wrong part).

  private planIds(blocksXml: string): IdPlan {
    const targetXml = this.zip.readAsText("word/document.xml") ?? "";
    const maxOf = (re: RegExp) =>
      [...targetXml.matchAll(re)].reduce((m, x) => Math.max(m, parseInt(x[1], 10) || 0), 0);

    let nextDoc = maxOf(/<wp:docPr\b[^>]*\bid="(\d+)"/g);
    // docPr ids are unique per element, not per value: every source drawing gets
    // its own new id, handed out in document order by applyIdPlan.
    let nextBm = maxOf(/<w:bookmarkStart\b[^>]*\bw:id="(\d+)"/g);
    const bmIds = new Map<string, string>();
    const taken = new Set(
      [...targetXml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)].map((m) => m[1]),
    );
    const srcNames = new Set(
      [...blocksXml.matchAll(/<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g)].map((m) => m[1]),
    );
    const renames = new Map<string, string>();
    for (const m of blocksXml.matchAll(/<w:bookmark(?:Start|End)\b[^>]*\bw:id="([^"]+)"/g)) {
      if (!bmIds.has(m[1])) bmIds.set(m[1], String(++nextBm));
    }
    for (const name of srcNames) {
      if (!taken.has(name)) continue;
      let k = 1;
      while (taken.has(`${name}_m${k}`) || srcNames.has(`${name}_m${k}`)) k++;
      renames.set(name, `${name}_m${k}`);
      taken.add(`${name}_m${k}`);
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
          t = t.replace(/(\bw:name=")([^"]*)(")/, (f, a, v, c) => (plan.renames.has(v) ? `${a}${plan.renames.get(v)}${c}` : f));
        }
        return t;
      });
    }
    if (plan.renames.size) out = this.renameBookmarkRefs(out, plan.renames);
    return out;
  }

  /** Point REF/PAGEREF/NOTEREF/HYPERLINK \l fields and `w:anchor` at renamed bookmarks. */
  private renameBookmarkRefs(xml: string, renames: Map<string, string>): string {
    const fix = (instr: string) => {
      let r = instr;
      for (const [from, to] of renames) {
        const f = from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        r = r
          .replace(new RegExp(`(\\b(?:REF|PAGEREF|NOTEREF)\\s+)(")?${f}(")?(?=\\s|$)`, "g"), (_m, a, q1, q2) => `${a}${q1 ?? ""}${to}${q2 ?? ""}`)
          .replace(new RegExp(`(\\\\l\\s+")${f}(")`, "g"), `$1${to}$2`);
      }
      return r;
    };
    return xml
      .replace(/(<w:instrText\b[^>]*>)([^<]*)(<\/w:instrText>)/g, (_m, a, t, c) => a + fix(t) + c)
      .replace(/(\bw:instr=")([^"]*)(")/g, (_m, a, t, c) => a + fix(t) + c)
      .replace(/(<w:hyperlink\b[^>]*\bw:anchor=")([^"]*)(")/g, (m, a, v, c) =>
        renames.has(v) ? `${a}${renames.get(v)}${c}` : m,
      );
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

  // ─── Footnotes (verbatim element copy, rich-content fidelity) ────────────────

  /** Copy each referenced source footnote verbatim; return source-id → new-id. */
  private async buildFootnoteMap(
    source: MdocxengineType,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const refIds = new Set(
      [...blocksXml.matchAll(/<w:footnoteReference\b[^>]*\bw:id="([^"]+)"/g)].map((m) => m[1]),
    );
    if (refIds.size === 0) return {};
    const sourceFootnotesXml = source.zip.readAsText("word/footnotes.xml");
    // A footnote's `r:embed` resolves against word/_rels/footnotes.xml.rels, not
    // the document rels — so its images are copied into THAT part's rels.
    const srcRelsXml = source.zip.readAsText("word/_rels/footnotes.xml.rels");
    const fnRels: Record<string, { type: string; target: string; external: boolean }> = {};
    // Parsed, not regex-scanned: Targets are XML-escaped on disk (`&amp;`), and
    // addRelationship escapes again on write — raw text would double-escape a URL.
    if (srcRelsXml) {
      const obj: any = await parseXml(srcRelsXml);
      const raw = obj?.Relationships?.Relationship;
      for (const r of Array.isArray(raw) ? raw : raw ? [raw] : []) {
        if (r?.$?.Id && r?.$?.Target != null) {
          fnRels[r.$.Id] = {
            type: String(r.$.Type ?? ""),
            target: String(r.$.Target),
            external: r.$.TargetMode === "External",
          };
        }
      }
    }
    const fnRelMgr = new RelManager(this.zip, "word/_rels/footnotes.xml.rels");
    const rewrite = async (xml: string): Promise<string> => {
      const ids = new Map<string, string>();
      for (const m of xml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)) {
        const rel = fnRels[m[1]];
        if (!rel || ids.has(m[1])) continue;
        if (rel.external) {
          // Hyperlinks (and linked images) live in the footnotes part's OWN rels.
          const newId = await fnRelMgr.genId();
          await fnRelMgr.addRelationship(newId, rel.type, rel.target, "External");
          ids.set(m[1], newId);
          continue;
        }
        const entry = source.zip.getEntry(resolvePartPath("word/footnotes.xml", rel.target));
        if (!entry) continue;
        const ext = (/\.([a-zA-Z0-9]+)$/.exec(rel.target)?.[1] ?? "png").toLowerCase();
        ids.set(m[1], await this.media.addImageToPartRels("word/footnotes.xml", entry.getData(), ext));
      }
      return this.applyAttrMap(xml, ["r:embed", "r:link", "r:id"], Object.fromEntries(ids));
    };
    return this.footnotes.copyFootnotesVerbatim(sourceFootnotesXml, refIds, rewrite);
  }

  // ─── Numbering (Phase 2) ─────────────────────────────────────────────────────

  /**
   * Copy the source's numbering definitions referenced by the blocks (abstractNum
   * + num) into the target with fresh, collision-free ids; return source-numId →
   * new-numId. Without this, two parts that both use `numId="1"` would share one
   * list and renumber wrongly.
   */
  private async buildNumberingMap(
    source: MdocxengineType,
    blocksXml: string,
  ): Promise<Record<string, string>> {
    const usedNumIds = new Set(
      [...blocksXml.matchAll(/<w:numId\b[^>]*\bw:val="([^"]+)"/g)].map((m) => m[1]),
    );
    if (usedNumIds.size === 0) return {};

    const srcDefs = await source.numbering.getRawDefinitions();
    if (srcDefs.nums.length === 0) return {};

    const { absMax, numMax } = await this.numbering.maxIds();
    let nextAbs = absMax + 1;
    let nextNum = numMax + 1;

    const numIdMap: Record<string, string> = {};
    const absIdMap: Record<string, string> = {};
    const addAbs: any[] = [];
    const addNum: any[] = [];

    for (const num of srcDefs.nums) {
      const oldNumId = num?.$?.["w:numId"];
      if (oldNumId == null || !usedNumIds.has(String(oldNumId))) continue;

      const oldAbs = num?.["w:abstractNumId"]?.$?.["w:val"];
      // Copy the referenced abstractNum once, with a fresh id.
      if (oldAbs != null && absIdMap[oldAbs] == null) {
        const absDef = srcDefs.abstractNums.find((a) => a?.$?.["w:abstractNumId"] === oldAbs);
        if (absDef) {
          const newAbs = String(nextAbs++);
          const clone = JSON.parse(JSON.stringify(absDef));
          clone.$ = { ...clone.$, "w:abstractNumId": newAbs };
          // Drop nsid so Word doesn't fold this into the source list's identity.
          if (clone["w:nsid"]) delete clone["w:nsid"];
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
      addNum.push(numClone);
      numIdMap[String(oldNumId)] = newNumId;
    }

    await this.numbering.appendRawDefinitions(addAbs, addNum);
    return numIdMap;
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

  /** Remap w:id ONLY inside <w:footnoteReference .../> (never other w:id). */
  private applyFootnoteRefMap(xml: string, idMap: Record<string, string>): string {
    if (Object.keys(idMap).length === 0) return xml;
    return xml.replace(/<w:footnoteReference\b[^>]*>/g, (tag) =>
      tag.replace(/\bw:id="([^"]+)"/, (full, val) => (idMap[val] ? `w:id="${idMap[val]}"` : full)),
    );
  }

  /** Retarget <w:pStyle w:val="X"/> by name. */
  private applyStyleMap(xml: string, styleMap: Record<string, string>): string {
    if (Object.keys(styleMap).length === 0) return xml;
    return xml.replace(/<w:pStyle\b[^>]*\bw:val="([^"]+)"[^>]*\/?>/g, (tag, val) =>
      styleMap[val] ? tag.replace(`w:val="${val}"`, `w:val="${styleMap[val]}"`) : tag,
    );
  }
}
