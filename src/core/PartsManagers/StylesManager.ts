import * as XmlUtils from "@/utils/xmlUtils";
import AdmZip from "adm-zip";
import { canonicalizeRunProps, canonicalizeStyleChildren } from "@/core/ooxml/canonicalOrder";
import { mergeRunProps, type RunProps } from "@/core/ooxml/runProps";
import { toHalfPoints } from "@/core/ooxml/hps";
import { applyParaStyleProps, type ParaStyleProps } from "@/core/ooxml/styleProps";
import { splitTopLevelElements, elementName } from "@/core/ooxml/canonicalOrder";

export type { ParaStyleProps } from "@/core/ooxml/styleProps";

const STYLES_PATH = "word/styles.xml";

export interface StyleEntry {
  id: string;
  name: string;
  type: string;
}

/** Run-level formatting to force onto a heading style's `<w:rPr>`. Only the
 *  provided fields are touched. */
export interface HeadingRunFormatting {
  /** Font size in POINTS (written as half-points to `<w:sz>`/`<w:szCs>`). */
  fontSizePt?: number;
  /** true adds `<w:b/><w:bCs/>`; false removes them. */
  bold?: boolean;
  /** Hex colour, with or without leading '#'. */
  color?: string;
}

/** Which `HeadingN` styles were found (and rewritten) vs. missing from styles.xml. */
export interface HeadingStyleResult {
  updatedLevels: number[];
  missingLevels: number[];
}

const selfClosingTag = (tag: string): RegExp => new RegExp(`<w:${tag}\\b[^>]*/>`, "g");

/** Strip every self-closing `<w:TAG .../>` occurrence from a fragment. */
function stripTag(xml: string, tag: string): string {
  return xml.replace(selfClosingTag(tag), "");
}

/**
 * Rewrite the INSIDE of a `<w:rPr>…</w:rPr>` fragment, replacing (or removing)
 * only the bold/color/size elements `formatting` specifies; every other child
 * (e.g. `<w:i/>`, `<w:rFonts/>`) is left untouched in its original position.
 * New elements are (re)inserted in CT_RPr schema order: b/bCs, …, color, …,
 * sz/szCs.
 */
function rewriteHeadingRunProps(rPrInner: string, formatting: HeadingRunFormatting): string {
  let body = rPrInner;
  let boldTags = "";
  if (formatting.bold !== undefined) {
    body = stripTag(stripTag(body, "b"), "bCs");
    if (formatting.bold) boldTags = `<w:b/><w:bCs/>`;
  }

  let colorTag = "";
  if (formatting.color !== undefined) {
    body = stripTag(body, "color");
    colorTag = `<w:color w:val="${formatting.color.replace(/^#/, "").toUpperCase()}"/>`;
  }

  let sizeTags = "";
  if (formatting.fontSizePt !== undefined) {
    body = stripTag(stripTag(body, "sz"), "szCs");
    const halfPts = toHalfPoints(formatting.fontSizePt, "heading style fontSizePt");
    sizeTags = `<w:sz w:val="${halfPts}"/><w:szCs w:val="${halfPts}"/>`;
  }

  return canonicalizeRunProps(`${boldTags}${body}${colorTag}${sizeTags}`);
}

/** Rewrite one `<w:style … w:styleId="HeadingN">…</w:style>` block's `<w:rPr>`. */
function rewriteHeadingStyleBlock(styleBlock: string, formatting: HeadingRunFormatting): string {
  const rPrMatch = styleBlock.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (rPrMatch) {
    const newInner = rewriteHeadingRunProps(rPrMatch[1], formatting);
    return styleBlock.replace(rPrMatch[0], `<w:rPr>${newInner}</w:rPr>`);
  }

  const newInner = rewriteHeadingRunProps("", formatting);
  if (!newInner) return styleBlock;
  const insertion = `<w:rPr>${newInner}</w:rPr>`;
  const nameMatch = styleBlock.match(/<w:name\b[^>]*\/>/);
  if (nameMatch) return styleBlock.replace(nameMatch[0], `${nameMatch[0]}${insertion}`);
  return styleBlock.replace(/(<w:style\b[^>]*>)/, `$1${insertion}`);
}

/**
 * Pure transform: force `formatting` onto the `Heading{level}` styles (1–6)
 * listed in `levels`, inside a `word/styles.xml` string. A style-level change —
 * not a per-run override — so it applies uniformly to every existing heading
 * that uses that style AND any new one added afterwards (matching how
 * `Doc.setHeadingLevel`/`addHeading` only ever set `w:pStyle`, never bold/size/
 * color directly). Returns the rewritten XML plus which levels existed.
 */
export function applyHeadingStyleToXml(
  stylesXml: string,
  levels: number[],
  formatting: HeadingRunFormatting,
): { xml: string; result: HeadingStyleResult } {
  let xml = stylesXml;
  const updatedLevels: number[] = [];
  const missingLevels: number[] = [];

  for (const level of levels) {
    const styleId = `Heading${level}`;
    const blockRe = new RegExp(`<w:style\\b[^>]*w:styleId="${styleId}"[^>]*>[\\s\\S]*?<\\/w:style>`);
    const match = blockRe.exec(xml);
    if (!match) {
      missingLevels.push(level);
      continue;
    }
    const rewritten = rewriteHeadingStyleBlock(match[0], formatting);
    xml = xml.slice(0, match.index) + rewritten + xml.slice(match.index + match[0].length);
    updatedLevels.push(level);
  }

  return { xml, result: { updatedLevels, missingLevels } };
}

/** How to create a paragraph style that does not exist yet. */
export interface EnsureStyleSpec {
  /** `w:name` — Word's display name. Maps the style onto a built-in when it matches. */
  name: string;
  /** `w:basedOn` target, omitted when absent. */
  basedOn?: string;
  /** Emit `w:default="1"` — used for `Normal`. */
  isDefault?: boolean;
}

/** Build a minimal, schema-ordered paragraph `<w:style>`. */
export function buildParagraphStyleXml(
  styleId: string,
  name: string,
  opts: { basedOn?: string; isDefault?: boolean } = {},
): string {
  const defaultAttr = opts.isDefault ? ` w:default="1"` : "";
  const basedOn = opts.basedOn ? `<w:basedOn w:val="${opts.basedOn}"/>` : "";
  return (
    `<w:style w:type="paragraph"${defaultAttr} w:styleId="${styleId}">` +
    `<w:name w:val="${name}"/>${basedOn}<w:qFormat/>` +
    `</w:style>`
  );
}

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
export function applyStyleRunPropsToXml(
  stylesXml: string,
  styleId: string,
  props: RunProps,
  ensure?: EnsureStyleSpec,
): { xml: string; created: boolean; updated: boolean } {
  let xml = stylesXml;
  let created = false;

  const blockRe = new RegExp(`<w:style\\b[^>]*w:styleId="${styleId}"[^>]*>[\\s\\S]*?<\\/w:style>`);
  if (!blockRe.test(xml)) {
    if (!ensure) return { xml, created: false, updated: false };
    const fresh = buildParagraphStyleXml(styleId, ensure.name, {
      basedOn: ensure.basedOn,
      isDefault: ensure.isDefault,
    });
    xml = xml.replace(/<\/w:styles>\s*$/, `${fresh}</w:styles>`);
    created = true;
  }

  const match = blockRe.exec(xml);
  if (!match) return { xml, created, updated: false };

  const block = match[0];
  const rPrMatch = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(block);
  const nextInner = mergeRunProps(rPrMatch ? rPrMatch[1] : "", props);

  let rewritten: string;
  if (rPrMatch) {
    rewritten = block.replace(rPrMatch[0], `<w:rPr>${nextInner}</w:rPr>`);
  } else if (nextInner) {
    const inner = /^(<w:style\b[^>]*>)([\s\S]*)(<\/w:style>)$/.exec(block);
    if (!inner) return { xml, created, updated: false };
    const children = canonicalizeStyleChildren(`${inner[2]}<w:rPr>${nextInner}</w:rPr>`);
    rewritten = `${inner[1]}${children}${inner[3]}`;
  } else {
    return { xml, created, updated: false };
  }

  xml = xml.slice(0, match.index) + rewritten + xml.slice(match.index + block.length);
  return { xml, created, updated: true };
}


const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const styleBlockRe = (styleId: string) =>
  new RegExp(`<w:style\\b[^>]*w:styleId="${escapeRe(styleId)}"[^>]*>[\\s\\S]*?<\\/w:style>`);

/**
 * Pure transform: apply paragraph `props` to the `<w:pPr>` of `styleId`, creating
 * the style first when `ensure` is supplied and it is missing. String surgery
 * only — every byte outside the one style's `pPr` is untouched.
 */
export function applyStyleParagraphPropsToXml(
  stylesXml: string,
  styleId: string,
  props: ParaStyleProps,
  ensure?: EnsureStyleSpec,
): { xml: string; created: boolean; updated: boolean } {
  let xml = stylesXml;
  let created = false;
  const re = styleBlockRe(styleId);
  if (!re.test(xml)) {
    if (!ensure) return { xml, created: false, updated: false };
    const fresh = buildParagraphStyleXml(styleId, ensure.name, { basedOn: ensure.basedOn, isDefault: ensure.isDefault });
    xml = xml.replace(/<\/w:styles>\s*$/, `${fresh}</w:styles>`);
    created = true;
  }
  const match = re.exec(xml);
  if (!match) return { xml, created, updated: false };

  const block = match[0];
  const parts = /^(<w:style\b[^>]*>)([\s\S]*)(<\/w:style>)$/.exec(block);
  if (!parts) return { xml, created, updated: false };
  const children = splitTopLevelElements(parts[2]);
  const at = children.findIndex((c) => elementName(c) === "w:pPr");
  const current = at === -1 ? "" : (/^<w:pPr(?:\s[^>]*)?>([\s\S]*)<\/w:pPr>$/.exec(children[at])?.[1] ?? "");
  const next = applyParaStyleProps(current, props);
  const pPr = next ? `<w:pPr>${next}</w:pPr>` : "";

  let rebuilt: string;
  if (at !== -1) {
    if (children[at] === pPr) return { xml, created, updated: false };
    children[at] = pPr;
    rebuilt = children.join("");
  } else if (pPr) {
    rebuilt = canonicalizeStyleChildren(`${parts[2]}${pPr}`);
  } else {
    return { xml, created, updated: false };
  }
  const out = parts[1] + rebuilt + parts[3];
  xml = xml.slice(0, match.index) + out + xml.slice(match.index + block.length);
  return { xml, created, updated: true };
}

/** How `ensureStyle` creates a style that is missing. */
export interface EnsureStyleOptions {
  name: string;
  type: "paragraph" | "character";
  basedOn?: string;
  next?: string;
  uiPriority?: number;
  qFormat?: boolean;
}

/**
 * Pure transform: add a minimal style when `styleId` is missing. An existing
 * style is NEVER modified (not even to make it match `spec`).
 */
export function applyEnsureStyleToXml(
  stylesXml: string,
  styleId: string,
  spec: EnsureStyleOptions,
): { xml: string; created: boolean } {
  if (styleBlockRe(styleId).test(stylesXml) || !/<\/w:styles>\s*$/.test(stylesXml)) {
    return { xml: stylesXml, created: false };
  }
  const fresh =
    `<w:style w:type="${spec.type}" w:styleId="${styleId}"><w:name w:val="${spec.name}"/>` +
    (spec.basedOn ? `<w:basedOn w:val="${spec.basedOn}"/>` : "") +
    (spec.next ? `<w:next w:val="${spec.next}"/>` : "") +
    (spec.uiPriority !== undefined ? `<w:uiPriority w:val="${spec.uiPriority}"/>` : "") +
    (spec.qFormat ? "<w:qFormat/>" : "") +
    `</w:style>`;
  return { xml: stylesXml.replace(/<\/w:styles>\s*$/, `${fresh}</w:styles>`), created: true };
}

export interface DocDefaultsInput {
  fontLatin?: string;
  fontCs?: string;
  sizePt?: number;
  sizeCsPt?: number;
  lang?: { val?: string; bidi?: string };
  /** Multiple of single line spacing. */
  lineSpacing?: number;
  afterPt?: number;
}

/**
 * Pure transform: write document defaults into `<w:docDefaults>` (created as the
 * first child of `<w:styles>` when absent), creating `rPrDefault` / `pPrDefault`
 * as needed. Only the named properties change.
 */
export function applyDocDefaultsToXml(
  stylesXml: string,
  d: DocDefaultsInput,
): { xml: string; updated: boolean } {
  const runProps: RunProps = {};
  if (d.fontLatin !== undefined) runProps.fontLatin = d.fontLatin;
  if (d.fontCs !== undefined) runProps.fontCs = d.fontCs;
  if (d.sizePt !== undefined) runProps.sizePt = d.sizePt;
  if (d.sizeCsPt !== undefined) runProps.sizeCsPt = d.sizeCsPt;
  const wantsRun = Object.keys(runProps).length > 0 || d.lang !== undefined;
  const paraProps: ParaStyleProps = {};
  if (d.lineSpacing !== undefined) paraProps.lineSpacing = d.lineSpacing;
  if (d.afterPt !== undefined) paraProps.afterPt = d.afterPt;
  const wantsPara = Object.keys(paraProps).length > 0;
  if (!wantsRun && !wantsPara) return { xml: stylesXml, updated: false };

  // Locate (or plan to create) docDefaults.
  const dd = /<w:docDefaults\s*\/>|<w:docDefaults(?:\s[^>]*)?>([\s\S]*?)<\/w:docDefaults>/.exec(stylesXml);
  let inner = dd?.[1] ?? "";

  const upsertDefault = (wrapper: "w:rPrDefault" | "w:pPrDefault", leaf: "w:rPr" | "w:pPr", edit: (leafInner: string) => string) => {
    const wRe = new RegExp(`<${wrapper}\\s*/>|<${wrapper}(?:\\s[^>]*)?>([\\s\\S]*?)</${wrapper}>`);
    const w = wRe.exec(inner);
    const wInner = w?.[1] ?? "";
    const lRe = new RegExp(`<${leaf}\\s*/>|<${leaf}(?:\\s[^>]*)?>([\\s\\S]*?)</${leaf}>`);
    const l = lRe.exec(wInner);
    const newLeaf = `<${leaf}>${edit(l?.[1] ?? "")}</${leaf}>`;
    const newWrapperInner = l ? wInner.replace(l[0], newLeaf) : wInner + newLeaf;
    const newWrapper = `<${wrapper}>${newWrapperInner}</${wrapper}>`;
    if (w) inner = inner.replace(w[0], newWrapper);
    else inner = wrapper === "w:rPrDefault" ? newWrapper + inner : inner + newWrapper;
  };

  if (wantsRun) {
    upsertDefault("w:rPrDefault", "w:rPr", (rpr) => {
      let out = Object.keys(runProps).length ? mergeRunProps(rpr, runProps) : rpr;
      if (d.lang) {
        const m = /<w:lang(?=[\s/>])([^>]*?)\/?>/.exec(out);
        const attrs = new Map<string, string>();
        if (m) for (const a of m[1].matchAll(/\s(w:\w+)=["']([^"']*)["']/g)) attrs.set(a[1], a[2]);
        if (d.lang.val !== undefined) attrs.set("w:val", d.lang.val);
        if (d.lang.bidi !== undefined) attrs.set("w:bidi", d.lang.bidi);
        const tag = `<w:lang${[...attrs].map(([k, v]) => ` ${k}="${v}"`).join("")}/>`;
        out = canonicalizeRunProps((m ? out.replace(m[0], "") : out) + tag);
      }
      return out;
    });
  }
  if (wantsPara) upsertDefault("w:pPrDefault", "w:pPr", (ppr) => applyParaStyleProps(ppr, paraProps));

  const block = `<w:docDefaults>${inner}</w:docDefaults>`;
  let xml: string;
  if (dd) xml = stylesXml.replace(dd[0], block);
  else xml = stylesXml.replace(/(<w:styles\b[^>]*>)/, `$1${block}`);
  return { xml, updated: xml !== stylesXml };
}

export class StylesManager {
  private zip: AdmZip;

  constructor(zip: AdmZip) {
    this.zip = zip;
  }

  private async readStyles(): Promise<any> {
    const xml = this.zip.readAsText(STYLES_PATH);
    if (!xml) return { "w:styles": { "w:style": [] } };
    return XmlUtils.parseXml(xml);
  }

  private async writeStyles(obj: any): Promise<void> {
    const xml = XmlUtils.buildXml(obj["w:styles"], {
      rootName: "w:styles",
      headless: false,
      pretty: true,
    });
    this.zip.addFile(STYLES_PATH, Buffer.from(xml, "utf-8"));
  }

  private normalizeStylesArray(obj: any): any[] {
    const raw = obj?.["w:styles"]?.["w:style"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  /**
   * Lists all styles in the document.
   */
  public async listStyles(): Promise<StyleEntry[]> {
    const obj = await this.readStyles();
    return this.normalizeStylesArray(obj).map((s: any) => ({
      id: s.$?.["w:styleId"] ?? "",
      name: s["w:name"]?.$?.["w:val"] ?? "",
      type: s.$?.["w:type"] ?? "",
    }));
  }

  /**
   * Returns the raw style object for a given style ID, or null.
   */
  public async getStyle(styleId: string): Promise<any | null> {
    const obj = await this.readStyles();
    return (
      this.normalizeStylesArray(obj).find((s: any) => s.$?.["w:styleId"] === styleId) ?? null
    );
  }

  /**
   * Adds a new style. Does nothing if a style with the same ID already exists.
   */
  public async addStyle(styleObj: any): Promise<void> {
    const obj = await this.readStyles();
    const styles = this.normalizeStylesArray(obj);
    const id = styleObj.$?.["w:styleId"];

    if (id && styles.some((s: any) => s.$?.["w:styleId"] === id)) return;

    styles.push(styleObj);
    obj["w:styles"]["w:style"] = styles;
    await this.writeStyles(obj);
  }

  /**
   * Removes a style by ID.
   */
  public async removeStyle(styleId: string): Promise<void> {
    const obj = await this.readStyles();
    const styles = this.normalizeStylesArray(obj);
    obj["w:styles"]["w:style"] = styles.filter(
      (s: any) => s.$?.["w:styleId"] !== styleId,
    );
    await this.writeStyles(obj);
  }

  /**
   * Forces `formatting` (font size / bold / color) onto the `Heading{level}`
   * styles for the given `levels` (default 1–6) — a STYLE-level change, so it
   * applies uniformly to every heading using that level, present and future,
   * rather than one paragraph's run. See {@link applyHeadingStyleToXml}.
   */
  public async setHeadingStyle(
    levels: number[] = [1, 2, 3, 4, 5, 6],
    formatting: HeadingRunFormatting,
  ): Promise<HeadingStyleResult> {
    const xml = this.zip.readAsText(STYLES_PATH) ?? "";
    const { xml: out, result } = applyHeadingStyleToXml(xml, levels, formatting);
    this.zip.addFile(STYLES_PATH, Buffer.from(out, "utf-8"));
    return result;
  }

  /**
   * Ensure `styleId` exists (creating it from `ensure` when missing), then apply
   * `props` to its `<w:rPr>`. A STYLE-level change: it reaches every paragraph
   * using the style, present and future.
   */
  public async setStyleRunProps(
    styleId: string,
    props: RunProps,
    ensure?: EnsureStyleSpec,
  ): Promise<{ created: boolean; updated: boolean }> {
    const xml = this.zip.readAsText(STYLES_PATH) ?? "";
    const { xml: out, created, updated } = applyStyleRunPropsToXml(xml, styleId, props, ensure);
    if (created || updated) this.zip.addFile(STYLES_PATH, Buffer.from(out, "utf-8"));
    return { created, updated };
  }

  /** Write paragraph properties onto a style's `pPr` (string surgery). */
  public async setStyleParagraphProps(
    styleId: string,
    props: ParaStyleProps,
    ensure?: EnsureStyleSpec,
  ): Promise<{ created: boolean; updated: boolean }> {
    const xml = this.zip.readAsText(STYLES_PATH) ?? "";
    const r = applyStyleParagraphPropsToXml(xml, styleId, props, ensure);
    if (r.created || r.updated) this.zip.addFile(STYLES_PATH, Buffer.from(r.xml, "utf-8"));
    return { created: r.created, updated: r.updated };
  }

  /** Write `w:docDefaults` (fonts, sizes, language, line spacing). */
  public async setDocDefaults(d: DocDefaultsInput): Promise<{ updated: boolean }> {
    const xml = this.zip.readAsText(STYLES_PATH) ?? "";
    const r = applyDocDefaultsToXml(xml, d);
    if (r.updated) this.zip.addFile(STYLES_PATH, Buffer.from(r.xml, "utf-8"));
    return { updated: r.updated };
  }

  /** Create `styleId` if missing. NEVER modifies an existing style. */
  public async ensureStyle(styleId: string, spec: EnsureStyleOptions): Promise<{ created: boolean }> {
    const xml = this.zip.readAsText(STYLES_PATH) ?? "";
    const r = applyEnsureStyleToXml(xml, styleId, spec);
    if (r.created) this.zip.addFile(STYLES_PATH, Buffer.from(r.xml, "utf-8"));
    return { created: r.created };
  }
}
