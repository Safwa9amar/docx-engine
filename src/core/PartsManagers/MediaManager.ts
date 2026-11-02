import { RelManager } from "@/core/PartsManagers/RelManager";
import { ContentTypesManager } from "./ContentTypesManager";
import AdmZip from "adm-zip";

const MEDIA_DIR = "word/media";
const IMAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const CONTENT_TYPE_MAP: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  tif: "image/tiff",
  wmf: "image/x-wmf",
  emf: "image/x-emf",
};

export interface ImageEntry {
  name: string;
  path: string;
  buffer: Buffer;
}

/** EMU per pixel at 96 DPI (1 px = 9525 EMU). */
export const EMU_PER_PIXEL = 9525;
/** EMU per point (914400 EMU per inch ÷ 72 pt per inch) — VML sizes in points. */
export const EMU_PER_POINT = 12700;
/** Convert EMU to whole pixels at 96 DPI. */
export const emuToPixels = (emu: number): number => Math.round(emu / EMU_PER_PIXEL);
/** Convert pixels (96 DPI) to EMU. */
export const pixelsToEmu = (px: number): number => Math.round(px * EMU_PER_PIXEL);

/** An inline image resolved from a paragraph block's `<w:drawing>`. */
export interface InlineImage {
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

export class MediaManager {
  private zip: AdmZip;
  private rels: RelManager;
  private contentTypes: ContentTypesManager;

  constructor(zip: AdmZip) {
    this.zip = zip;
    this.rels = new RelManager(zip);
    this.contentTypes = new ContentTypesManager(zip);
  }

  /**
   * Returns all image files found in word/media/.
   */
  public listImages(): ImageEntry[] {
    const images: ImageEntry[] = [];

    this.zip.getEntries().forEach((entry) => {
      if (
        entry.entryName.startsWith(`${MEDIA_DIR}/`) &&
        /\.(png|jpe?g|gif|bmp|tiff|svg)$/i.test(entry.entryName)
      ) {
        const buf = entry.getData();
        if (buf) {
          images.push({
            name: entry.entryName.replace(`${MEDIA_DIR}/`, ""),
            path: entry.entryName,
            buffer: buf,
          });
        }
      }
    });

    return images;
  }

  /**
   * Returns the buffer for a specific image by filename (e.g. "image1.png"), or null.
   */
  public extractImage(name: string): Buffer | null {
    const entry = this.zip.getEntry(`${MEDIA_DIR}/${name}`);
    return entry ? entry.getData() : null;
  }

  /**
   * Resolve a relationship id (e.g. an image's `r:embed`) to its bytes via the
   * document relationships. Returns null if the rel or the target part is missing.
   */
  public async getImageByRelId(relId: string): Promise<Buffer | null> {
    const target = await this.rels.getTarget(relId);
    if (!target) return null;
    return this.readByTarget(target);
  }

  /**
   * Read a media part by its relationship `Target` (e.g. "media/image1.png").
   * Targets are relative to `word/` unless already absolute/prefixed.
   */
  private readByTarget(target: string): Buffer | null {
    const path = target.startsWith("/")
      ? target.slice(1)
      : target.startsWith("word/")
        ? target
        : `word/${target}`;
    const entry = this.zip.getEntry(path);
    return entry ? entry.getData() : null;
  }

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
  public async extractInlineImage(blockXml: string): Promise<InlineImage | null> {
    // `r:id` is NOT specific enough on its own — the same `<w:object>` carries one
    // on `<o:OLEObject>` (the OLE binary, not an image), and hyperlinks carry one
    // too. Only `<v:imagedata>`'s points at media.
    const embed =
      /r:embed="([^"]+)"/.exec(blockXml) ?? /<v:imagedata\b[^>]*\br:id="([^"]+)"/.exec(blockXml);
    if (!embed) return null;
    const relId = embed[1];
    const target = await this.rels.getTarget(relId);
    if (!target) return null;
    const bytes = this.readByTarget(target);
    if (!bytes) return null;

    const extension = (target.split(".").pop() || "png").toLowerCase();
    const mime = CONTENT_TYPE_MAP[extension] ?? "application/octet-stream";

    let widthEmu = 0;
    let heightEmu = 0;
    const ext = /<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(blockXml);
    if (ext) {
      widthEmu = Number(ext[1]);
      heightEmu = Number(ext[2]);
    } else {
      // VML sizes in CSS on the shape: style="width:11.25pt;height:19.5pt".
      const style = /<v:shape\b[^>]*\bstyle="([^"]*)"/.exec(blockXml);
      if (style) {
        const w = /(?:^|;)\s*width\s*:\s*([\d.]+)pt/.exec(style[1]);
        const h = /(?:^|;)\s*height\s*:\s*([\d.]+)pt/.exec(style[1]);
        if (w) widthEmu = Math.round(Number(w[1]) * EMU_PER_POINT);
        if (h) heightEmu = Math.round(Number(h[1]) * EMU_PER_POINT);
      }
    }

    return {
      relId,
      target,
      bytes,
      extension,
      mime,
      widthEmu,
      heightEmu,
      widthPx: emuToPixels(widthEmu),
      heightPx: emuToPixels(heightEmu),
    };
  }

  /**
   * Next free `imageN` number, counted over EVERY file in word/media. `listImages`
   * only sees raster/svg extensions, so counting through it ignored `.wmf`/`.emf`
   * (and `.tif`): with only those present — or a mix — the counter did not advance
   * and each new image overwrote the last one.
   */
  private nextImageNumber(): number {
    let max = 0;
    for (const entry of this.zip.getEntries()) {
      if (!entry.entryName.startsWith(`${MEDIA_DIR}/`)) continue;
      const m = /(?:^|\/)image(\d+)\./i.exec(entry.entryName);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
    return max + 1;
  }

  /**
   * Inserts a new image into the document.
   * @param imageBuffer  Raw image bytes.
   * @param extension    File extension without dot (e.g. "png", "jpg").
   * @returns The image path and the generated relationship ID.
   */
  public async insertImage(
    imageBuffer: Buffer,
    extension: string,
  ): Promise<{ imagePath: string; relId: string }> {
    const ext = extension.toLowerCase();
    const next = this.nextImageNumber();
    const imageName = `image${next}.${ext}`;
    const imagePath = `${MEDIA_DIR}/${imageName}`;

    // 1. Add to zip
    this.zip.addFile(imagePath, imageBuffer);

    // 2. Register content type default for extension
    const contentType = CONTENT_TYPE_MAP[ext] ?? `image/${ext}`;
    await this.contentTypes.addDefault(ext, contentType);

    // 3. Register relationship
    const relId = await this.rels.genId();
    await this.rels.addRelationship(relId, IMAGE_REL_TYPE, `media/${imageName}`);

    return { imagePath, relId };
  }

  /**
   * Add image bytes to word/media (+ register the content-type Default) WITHOUT
   * creating any relationship. Use when the relationship must live in a specific
   * part's own _rels (e.g. a header/footer part) rather than the document rels —
   * see {@link addImageToPartRels}. Returns the media part path (e.g.
   * "word/media/image2.png").
   */
  public async addImagePart(
    imageBuffer: Buffer,
    extension: string,
  ): Promise<{ imagePath: string }> {
    const ext = extension.toLowerCase();
    const next = this.nextImageNumber();
    const imagePath = `${MEDIA_DIR}/image${next}.${ext}`;
    this.zip.addFile(imagePath, imageBuffer);
    const contentType = CONTENT_TYPE_MAP[ext] ?? `image/${ext}`;
    await this.contentTypes.addDefault(ext, contentType);
    return { imagePath };
  }

  /**
   * Embed an image into a SPECIFIC part's own relationships and return that
   * part-local `r:embed` id. A header/footer part resolves `r:embed` against its
   * own `word/_rels/<part>.rels`, NOT the document rels — so an id from
   * {@link insertImage} would not resolve inside a header. This adds the bytes +
   * content-type, then a part-local image relationship (creating the part's .rels
   * if absent), and hands back the id to drop into the part's `<a:blip r:embed>`.
   * @param partPath  Full part path, e.g. "word/header1.xml".
   */
  public async addImageToPartRels(
    partPath: string,
    imageBuffer: Buffer,
    extension: string,
  ): Promise<string> {
    const { imagePath } = await this.addImagePart(imageBuffer, extension);
    const partName = partPath.replace(/^word\//, "");
    const rels = new RelManager(this.zip, `word/_rels/${partName}.rels`);
    const relId = await rels.genId();
    // Target is relative to the part's folder (word/), so drop the "word/" prefix.
    await rels.addRelationship(relId, IMAGE_REL_TYPE, imagePath.replace(/^word\//, ""));
    return relId;
  }

  /**
   * Replaces an existing image's bytes in-place (same filename, same relId).
   * @param name         Filename like "image1.png" (without path prefix).
   * @param newBuffer    New image bytes.
   */
  public replaceImage(name: string, newBuffer: Buffer): void {
    const path = `${MEDIA_DIR}/${name}`;
    if (!this.zip.getEntry(path)) throw new Error(`Image not found: ${name}`);
    this.zip.addFile(path, newBuffer);
  }

  /**
   * Deletes an image from the zip (does not remove the relationship or inline reference).
   */
  public deleteImage(name: string): void {
    const path = `${MEDIA_DIR}/${name}`;
    if (!this.zip.getEntry(path)) throw new Error(`Image not found: ${name}`);
    this.zip.deleteFile(path);
  }
}
