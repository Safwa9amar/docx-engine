import AdmZip from "adm-zip";

/**
 * A small .docx written part by part, for the merge specs: what each case is about
 * stays visible in the case itself, with no sample file's styles or lists around it.
 */

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = (t: string) => `${R_NS}/${t}`;
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
export const XML_HEAD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

export const NS =
  `xmlns:w="${W_NS}" xmlns:r="${R_NS}" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
  `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ` +
  `xmlns:v="urn:schemas-microsoft-com:vml" ` +
  `xmlns:o="urn:schemas-microsoft-com:office:office" ` +
  `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
  `xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"`;

export const PAGE =
  `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>`;
export const SECT = `<w:sectPr>${PAGE}</w:sectPr>`;
const DOC_DEFAULTS =
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>`;

export const stylesDoc = (inner: string, defaults = DOC_DEFAULTS) => `${XML_HEAD}<w:styles xmlns:w="${W_NS}">${defaults}${inner}</w:styles>`;
export const NORMAL = `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`;
export const withNormal = (inner = "") => stylesDoc(NORMAL + inner);

/** One style. `attrs` go on the open tag, `inner` after the name. */
export const style = (type: string, id: string, name: string, inner = "", attrs = "") =>
  `<w:style w:type="${type}"${attrs ? " " + attrs : ""} w:styleId="${id}"><w:name w:val="${name}"/>${inner}</w:style>`;

export const P = (t: string, pPr = "") => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
export const noteRef = (kind: "footnote" | "endnote", id: number | string) => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
export const refP = (kind: "footnote" | "endnote", id: number | string, t = "body") =>
  `<w:p><w:r><w:t xml:space="preserve">${t} </w:t></w:r>${noteRef(kind, id)}</w:p>`;
export const note = (kind: "footnote" | "endnote", id: number | string, inner: string, attrs = "") =>
  `<w:${kind}${attrs ? " " + attrs : ""} w:id="${id}">${inner}</w:${kind}>`;
export const noteText = (kind: "footnote" | "endnote", t: string) => `<w:p><w:r><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> ${t}</w:t></w:r></w:p>`;
export const separators = (kind: "footnote" | "endnote", pPr = "", ids: [number, number] = [-1, 0]) =>
  `<w:${kind} w:type="separator" w:id="${ids[0]}"><w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="${ids[1]}"><w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}<w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
export const notesPart = (kind: "footnote" | "endnote", notes: string, seps = separators(kind)) =>
  `${XML_HEAD}<w:${kind}s ${NS}>${seps}${notes}</w:${kind}s>`;
export const settingsPart = (inner: string) => `${XML_HEAD}<w:settings xmlns:w="${W_NS}">${inner}</w:settings>`;
export const numberingPart = (inner: string) => `${XML_HEAD}<w:numbering ${NS}>${inner}</w:numbering>`;
export const relsPart = (inner: string) => `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

export interface FixtureOptions {
  body: string;
  /** The final section. `null` writes none. */
  sect?: string | null;
  /** The styles part. `null` writes none. */
  styles?: string | null;
  settings?: string;
  footnotes?: string;
  endnotes?: string;
  comments?: string;
  numbering?: string;
  /** More relationships of word/document.xml, as XML. */
  docRels?: string;
  /** More entries of [Content_Types].xml, as XML. */
  contentTypes?: string;
  /** More parts, by path. */
  parts?: Record<string, string | Buffer>;
}

export function fixture(o: FixtureOptions): Buffer {
  const zip = new AdmZip();
  const overrides: string[] = [];
  const rels: string[] = [];
  let n = 100;
  const part = (path: string, xml: string, contentType: string, rel: string) => {
    zip.addFile(path, Buffer.from(xml, "utf8"));
    overrides.push(`<Override PartName="/${path}" ContentType="${contentType}"/>`);
    rels.push(`<Relationship Id="rId${n++}" Type="${REL(rel)}" Target="${path.replace(/^word\//, "")}"/>`);
  };
  const sect = o.sect === null ? "" : (o.sect ?? SECT);
  zip.addFile("word/document.xml", Buffer.from(`${XML_HEAD}<w:document ${NS}><w:body>${o.body}${sect}</w:body></w:document>`, "utf8"));
  overrides.push(`<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>`);
  if (o.styles !== null) part("word/styles.xml", o.styles ?? withNormal(), `${WML}.styles+xml`, "styles");
  if (o.settings) part("word/settings.xml", o.settings, `${WML}.settings+xml`, "settings");
  if (o.footnotes) part("word/footnotes.xml", o.footnotes, `${WML}.footnotes+xml`, "footnotes");
  if (o.endnotes) part("word/endnotes.xml", o.endnotes, `${WML}.endnotes+xml`, "endnotes");
  if (o.comments) part("word/comments.xml", o.comments, `${WML}.comments+xml`, "comments");
  if (o.numbering) part("word/numbering.xml", o.numbering, `${WML}.numbering+xml`, "numbering");
  for (const [path, bytes] of Object.entries(o.parts ?? {})) zip.addFile(path, Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes, "utf8"));
  zip.addFile("word/_rels/document.xml.rels", Buffer.from(relsPart(rels.join("") + (o.docRels ?? "")), "utf8"));
  zip.addFile("_rels/.rels", Buffer.from(relsPart(`<Relationship Id="rId1" Type="${REL("officeDocument")}" Target="word/document.xml"/>`), "utf8"));
  zip.addFile(
    "[Content_Types].xml",
    Buffer.from(
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.join("")}${o.contentTypes ?? ""}</Types>`,
      "utf8",
    ),
  );
  return zip.toBuffer();
}

/** A one-pixel PNG, for the cases that need real image bytes. `seed` varies them. */
export function png(seed = 0): Buffer {
  const b = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
    "base64",
  );
  return seed ? Buffer.concat([b, Buffer.from([seed])]) : b;
}

export const texts = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]);
export const styleIds = (stylesXml: string) => [...stylesXml.matchAll(/<w:style\b[^>]*\bw:styleId="([^"]+)"/g)].map((m) => m[1]);
export const styleOf = (stylesXml: string, id: string) => new RegExp(`<w:style\\b[^>]*\\bw:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(stylesXml)?.[0] ?? "";
export const styleNames = (stylesXml: string) => [...stylesXml.matchAll(/<w:name w:val="([^"]+)"/g)].map((m) => m[1]);
