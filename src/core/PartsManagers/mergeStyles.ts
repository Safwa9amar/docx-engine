/**
 * How a merge reconciles two documents' styles. Pure: strings and parsed numbering in, a plan out.
 *
 * Word's identity for a style is its NAME. Two styles sharing a `w:name` are merged at load and
 * the last one wins, whatever their ids, so appending a source's `a` (name "Normal") redefines the
 * target's Normal for the whole document, in a way a later edit of `Normal` cannot undo. An id, on
 * the other hand, means nothing across documents: Arabic Word hands out `a3`, `a4`, `ac` per file,
 * so one id can be a header paragraph here and a table grid there.
 *
 * So each style the copied content uses gets one decision:
 *  - the target has that name and the same meaning: bind to the target's style (`merge`);
 *  - the target lacks the name: copy it, under a new id if its own means something else here (`import`);
 *  - the target has the name with another look, on a style nothing rebuilds by role (a custom style,
 *    a character or table style): copy it under a new name and id, so both looks stay (`import`, forked);
 *  - Word's own default of a type, with no default of that type in the target, or a style the source
 *    never defined: nothing to bind to (`skip`).
 *
 * A built-in paragraph style with another look binds to the target's on purpose: Regenerate rebuilds
 * those by role. What such a style carries beyond its look is its numbering, the words Word prints
 * before the paragraph, and that does not come back. So when binding would change a paragraph's list,
 * the paragraph takes it directly (a `stamp`), the way Word itself keeps source formatting on a paste.
 */

import { canonicalizeParagraphProps } from "@/core/ooxml/canonicalOrder";
import type { RawNumbering } from "./NumberingManager";

export interface StyleInfo {
  id: string;
  type: string;
  name: string;
  /** Word's identity: the name, without regard to case or spacing. */
  key: string;
  xml: string;
  isDefault: boolean;
  custom: boolean;
  basedOn?: string;
  link?: string;
  next?: string;
  /** From the style's own `w:numPr`; `0` turns numbering off. */
  numId?: string;
  ilvl?: string;
}

export type StyleDecision =
  /** The caller's styleMap decides. */
  | { kind: "alias"; id: string }
  | { kind: "merge"; id: string; key: string }
  | { kind: "import"; id: string; name: string; key: string }
  | { kind: "skip" };

export interface StylePlan {
  decisions: Map<string, StyleDecision>;
  /** Definitions to append, references rewritten. Their numIds are still the source's. */
  imports: { id: string; xml: string }[];
  /** Source paragraph style id → the numbering its paragraphs must take directly. */
  stamps: Map<string, { numId: string; ilvl: string }>;
  /** Every source list something copied uses. */
  numIds: Set<string>;
  styleMap: Record<string, string>;
}

export interface StylePlanInput {
  sourceStyles: string | null | undefined;
  targetStyles: string | null | undefined;
  sourceLists: RawNumbering;
  targetLists: RawNumbering;
  styleMap: Record<string, string>;
  /** The copied blocks and notes. */
  contentXml: string;
}

/**
 * Word's own default of each type, by name. A style named Normal restyles every plain paragraph of
 * the document even without `w:default`, in any case (measured in real Word, a 20 pt one over a
 * 12 pt body: 19.9; one named Standard: 12). So none of these is copied in beside the target's text.
 */
const IMPLICIT_DEFAULTS = new Set(["normal", "defaultparagraphfont", "normaltable", "nolist"]);
const STYLE_ELEMENT = /<w:style\b[^>]*?\/>|<w:style\b[^>]*>[\s\S]*?<\/w:style>/g;
const isOn = (v: string | undefined) => v !== undefined && /^(1|true|on)$/i.test(v);
const nameKey = (name: string) => name.toLowerCase().replace(/\s+/g, "");
const childVal = (xml: string, tag: string) => new RegExp(`<${tag}\\b[^>]*\\bw:val="([^"]*)"`).exec(xml)?.[1];
const asArray = <T>(x: T | T[] | undefined | null): T[] => (Array.isArray(x) ? x : x == null ? [] : [x]);
const nodeVal = (node: any, key: string): string | undefined => node?.[key]?.$?.["w:val"];

export function readStyles(stylesXml: string | null | undefined): StyleInfo[] {
  const out: StyleInfo[] = [];
  for (const m of (stylesXml ?? "").matchAll(STYLE_ELEMENT)) {
    const xml = m[0];
    const open = xml.slice(0, xml.indexOf(">") + 1);
    const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(open)?.[1];
    const id = attr("w:styleId");
    if (!id) continue;
    const name = childVal(xml, "w:name") ?? id;
    const numPr = /<w:numPr\b[^>]*>([\s\S]*?)<\/w:numPr>/.exec(xml)?.[1] ?? "";
    out.push({
      id,
      type: attr("w:type") ?? "paragraph",
      name,
      key: nameKey(name),
      xml,
      isDefault: isOn(attr("w:default")),
      custom: isOn(attr("w:customStyle")),
      basedOn: childVal(xml, "w:basedOn"),
      link: childVal(xml, "w:link"),
      next: childVal(xml, "w:next"),
      numId: childVal(numPr, "w:numId"),
      ilvl: childVal(numPr, "w:ilvl"),
    });
  }
  return out;
}

// ─── Lists ───────────────────────────────────────────────────────────────────

interface Side {
  byId: Map<string, StyleInfo>;
  lists: RawNumbering;
}

const numOf = (side: Side, numId: string) => side.lists.nums.find((n) => n?.$?.["w:numId"] === numId);

/** The definition holding a list's levels. A `w:numStyleLink` holds none: it names the list
 *  style whose own list does. */
function levelsDefinition(side: Side, numId: string, seen = new Set<string>()): any {
  const abs = side.lists.abstractNums.find((a) => a?.$?.["w:abstractNumId"] === nodeVal(numOf(side, numId), "w:abstractNumId"));
  const link = nodeVal(abs, "w:numStyleLink");
  const via = link && !seen.has(link) ? side.byId.get(link)?.numId : undefined;
  if (link) seen.add(link);
  return (via && levelsDefinition(side, via, seen)) ?? abs;
}

/**
 * What a list prints: two lists with the same print count on as one. A numId with no definition is
 * no absence: Word gives it a default "1." of its own (measured, from a paragraph and from a style),
 * so it prints something too.
 */
function listPrint(side: Side, numId: string): string {
  const num = numOf(side, numId);
  const abs = levelsDefinition(side, numId);
  if (!num || !abs) return "no definition";
  const level = (l: any) => [l?.$?.["w:ilvl"], nodeVal(l, "w:start"), nodeVal(l, "w:numFmt"), nodeVal(l, "w:lvlText"), nodeVal(l, "w:suff")];
  return JSON.stringify([
    asArray(abs["w:lvl"]).map(level),
    asArray(num["w:lvlOverride"]).map((o: any) => [o?.$?.["w:ilvl"], nodeVal(o, "w:startOverride"), o?.["w:lvl"] ? level(o["w:lvl"]) : null]),
  ]);
}

interface Numbered {
  numId: string;
  /** The style in the chain that names the list. */
  owner: string;
  ilvl?: string;
}

/** The list a paragraph style gives its paragraphs, through its chain. Word's own `Heading2` shape
 *  names only a level and takes the list from `Heading1`, so a level is read before the list. */
function effectiveList(side: Side, id: string): Numbered | null {
  let ilvl: string | undefined;
  const seen = new Set<string>();
  for (let cur: string | undefined = id; cur && !seen.has(cur); cur = side.byId.get(cur)?.basedOn) {
    seen.add(cur);
    const s = side.byId.get(cur);
    if (!s) break;
    ilvl ??= s.ilvl;
    if (s.numId !== undefined) return s.numId === "0" ? null : { numId: s.numId, owner: cur, ilvl };
  }
  return null;
}

/** The level a style's paragraphs sit at. A level linked to the style outranks a style's `w:ilvl`. */
function levelOf(side: Side, id: string, n: Numbered): string {
  const levels = asArray<any>(levelsDefinition(side, n.numId)?.["w:lvl"]);
  const linkedTo = (styleId: string) => levels.find((l) => nodeVal(l, "w:pStyle") === styleId)?.$?.["w:ilvl"];
  return linkedTo(id) ?? n.ilvl ?? linkedTo(n.owner) ?? "0";
}

// ─── The plan ────────────────────────────────────────────────────────────────

/** The part of a definition that says how text looks, in a form two documents can compare. `parentKey`
 *  says where its `basedOn` lands; a list compares by what it prints, since numIds are per file. */
function comparable(s: StyleInfo, side: Side, parentKey: (id: string) => string): string {
  const inner = s.xml.endsWith("/>") ? "" : s.xml.slice(s.xml.indexOf(">") + 1, s.xml.lastIndexOf("</w:style>"));
  return inner
    .replace(/<w:(name|aliases|next|link|uiPriority|semiHidden|unhideWhenUsed|qFormat|locked|personal|personalCompose|personalReply|rsid|autoRedefine|hidden)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g, "")
    .replace(/>\s+</g, "><")
    .trim()
    .replace(/<w:basedOn\b[^>]*\bw:val="([^"]*)"[^>]*?(?:\/>|>\s*<\/w:basedOn>)/, (_m, id) => `<basedOn ${parentKey(id)}/>`)
    .replace(/<w:numId\b[^>]*\bw:val="([^"]*)"[^>]*?(?:\/>|>\s*<\/w:numId>)/g, (_m, id) => `<list ${id === "0" ? "none" : listPrint(side, id)}/>`);
}

export function planStyles(input: StylePlanInput): StylePlan {
  const { styleMap } = input;
  const sourceAll = readStyles(input.sourceStyles);
  const targetAll = readStyles(input.targetStyles);
  const src: Side = { byId: new Map(sourceAll.map((s) => [s.id, s])), lists: input.sourceLists };
  const tgt: Side = { byId: new Map(targetAll.map((s) => [s.id, s])), lists: input.targetLists };
  // A file can hold one name twice, and then the last one is the style.
  const targetByKey = new Map(targetAll.map((s) => [s.key, s]));
  const sourceKeys = new Set(sourceAll.map((s) => s.key));
  const targetDefaults = new Map(targetAll.filter((s) => s.isDefault).map((s) => [s.type, s]));
  const canWrite = !input.targetStyles || input.targetStyles.includes("</w:styles>");

  const takenIds = new Set([...src.byId.keys(), ...tgt.byId.keys()]);
  const plannedKeys = new Set<string>();
  const decisions = new Map<string, StyleDecision>();
  const deciding = new Set<string>();

  const importAs = (s: StyleInfo, name: string): StyleDecision => {
    let id = s.id;
    if (tgt.byId.has(id)) {
      let k = 1;
      while (takenIds.has(`${s.id}_m${k}`)) k++;
      id = `${s.id}_m${k}`;
      takenIds.add(id);
    }
    plannedKeys.add(nameKey(name));
    return { kind: "import", id, name, key: nameKey(name) };
  };
  const parentLanding = (id: string): string => {
    const d = decide(id);
    return d.kind === "merge" || d.kind === "import" ? d.key : "";
  };
  const sameLook = (s: StyleInfo, t: StyleInfo) =>
    comparable(s, src, parentLanding) === comparable(t, tgt, (id) => tgt.byId.get(id)?.key ?? "");

  /** Beside the target's, under a name of its own; or on what an earlier part brought with this look. */
  const fork = (s: StyleInfo): StyleDecision => {
    for (let k = 2; ; k++) {
      const name = `${s.name} (${k})`;
      const key = nameKey(name);
      const t = targetByKey.get(key);
      if (t) {
        if (t.type === s.type && sameLook(s, t)) return { kind: "merge", id: t.id, key };
        continue;
      }
      if (!plannedKeys.has(key) && !sourceKeys.has(key)) return importAs(s, name);
    }
  };

  function decide(id: string): StyleDecision {
    const known = decisions.get(id);
    if (known) return known;
    // A basedOn cycle is invalid, and only has to end here.
    if (deciding.has(id)) return { kind: "skip" };
    deciding.add(id);
    const d = decideFresh(id);
    deciding.delete(id);
    decisions.set(id, d);
    return d;
  }

  function decideFresh(id: string): StyleDecision {
    if (styleMap[id]) return { kind: "alias", id: styleMap[id] };
    const s = src.byId.get(id);
    if (!s || !canWrite) return { kind: "skip" };
    const t = targetByKey.get(s.key);
    if (!t) {
      if (!IMPLICIT_DEFAULTS.has(s.key)) return importAs(s, s.name);
      const own = targetDefaults.get(s.type);
      return own ? { kind: "merge", id: own.id, key: own.key } : { kind: "skip" };
    }
    if (t.type !== s.type) return fork(s);
    if (s.isDefault || t.isDefault) return { kind: "merge", id: t.id, key: t.key };
    if (s.type === "paragraph" && !s.custom && !t.custom) return { kind: "merge", id: t.id, key: t.key };
    return sameLook(s, t) ? { kind: "merge", id: t.id, key: t.key } : fork(s);
  }

  // What the copied content needs, closed over chains and over the styles a list names.
  const visited = new Set<string>();
  const imported: string[] = [];
  const numIds = new Set<string>();
  const visit = (id: string) => {
    if (visited.has(id)) return;
    visited.add(id);
    if (decide(id).kind !== "import") return;
    imported.push(id);
    const s = src.byId.get(id)!;
    for (const r of [s.basedOn, s.link, s.next]) if (r) visit(r);
    if (s.numId) useList(s.numId);
  };
  const useList = (numId: string) => {
    if (numIds.has(numId)) return;
    numIds.add(numId);
    const abs = input.sourceLists.abstractNums.find((a) => a?.$?.["w:abstractNumId"] === nodeVal(numOf(src, numId), "w:abstractNumId"));
    if (!abs) return;
    for (const key of ["w:styleLink", "w:numStyleLink"]) {
      const v = nodeVal(abs, key);
      if (v) visit(v);
    }
    for (const lvl of asArray<any>(abs["w:lvl"])) {
      const v = nodeVal(lvl?.["w:rPr"], "w:rStyle");
      if (v) visit(v);
    }
  };

  const refs = (re: RegExp) => new Set([...input.contentXml.matchAll(re)].map((m) => m[1]));
  const paragraphRefs = refs(/<w:pStyle\b[^>]*\bw:val="([^"]+)"/g);
  for (const id of refs(/<w:numId\b[^>]*\bw:val="([^"]+)"/g)) useList(id);
  for (const id of paragraphRefs) visit(id);
  for (const id of refs(/<w:(?:rStyle|tblStyle)\b[^>]*\bw:val="([^"]+)"/g)) visit(id);

  /** The list this style's paragraphs get here: as they had it ("as-is": a copied definition carries it,
   *  or a chain through the caller's styleMap, whose template owns the numbering), from the target, or none. */
  const landedList = (id: string): "as-is" | { numId: string; style: string; n: Numbered } | null => {
    const seen = new Set<string>();
    for (let cur: string | undefined = id; cur && !seen.has(cur); cur = src.byId.get(cur)?.basedOn) {
      seen.add(cur);
      const d = decide(cur);
      if (d.kind === "skip") return null;
      if (d.kind === "alias") return "as-is";
      if (d.kind === "merge") {
        const n = effectiveList(tgt, d.id);
        return n ? { numId: n.numId, style: d.id, n } : null;
      }
      const numId = src.byId.get(cur)?.numId;
      if (numId !== undefined) return numId === "0" ? null : "as-is";
    }
    return null;
  };

  const stamps = new Map<string, { numId: string; ilvl: string }>();
  for (const id of paragraphRefs) {
    if (!src.byId.has(id)) continue;
    const list = effectiveList(src, id);
    const here = landedList(id);
    if (here === "as-is") continue;
    if (!list) {
      if (here) stamps.set(id, { numId: "0", ilvl: "0" });
      continue;
    }
    const level = levelOf(src, id, list);
    if (here && listPrint(src, list.numId) === listPrint(tgt, here.numId) && level === levelOf(tgt, here.style, here.n)) continue;
    stamps.set(id, { numId: list.numId, ilvl: level });
    useList(list.numId);
  }

  const imports = imported.map((id) => buildImport(src.byId.get(id)!, decisions.get(id) as Extract<StyleDecision, { kind: "import" }>, decide));
  return { decisions, imports, stamps, numIds, styleMap };
}

function buildImport(
  s: StyleInfo,
  d: Extract<StyleDecision, { kind: "import" }>,
  decide: (id: string) => StyleDecision,
): { id: string; xml: string } {
  const gt = s.xml.indexOf(">") + 1;
  // Only one default style per type: the target keeps its own.
  let open = s.xml.slice(0, gt).replace(/\s+w:default="(?:1|true|on)"/i, "");
  if (d.id !== s.id) open = open.replace(/(\bw:styleId=")[^"]*(")/, (_m, a, c) => a + d.id + c);
  let inner = s.xml.slice(gt);
  if (d.name !== s.name) inner = inner.replace(/(<w:name\b[^>]*\bw:val=")[^"]*(")/, (_m, a, c) => a + d.name + c);
  inner = inner.replace(/<w:(basedOn|next|link)\b[^>]*\bw:val="([^"]+)"[^>]*?(?:\/>|>\s*<\/w:\1>)/g, (tag, which, val) => {
    const to = decide(val);
    // Word's links run both ways, and the target's own style links to something else.
    if (to.kind === "skip" || (which === "link" && to.kind !== "import")) return "";
    return tag.replace(`w:val="${val}"`, `w:val="${to.id}"`);
  });
  return { id: d.id, xml: open + inner };
}

// ─── Applying it ─────────────────────────────────────────────────────────────

/** Point `pStyle`, `rStyle` and `tblStyle` at what each style became. A styleMap entry retargets
 *  paragraphs only, as it always has. One pass, so a new id is never mapped again. */
export function retargetStyleRefs(xml: string, plan: StylePlan): string {
  return xml.replace(/<w:(pStyle|rStyle|tblStyle)\b[^>]*\bw:val="([^"]+)"[^>]*\/?>/g, (tag, which, val) => {
    const d = plan.decisions.get(val);
    let to: string | undefined;
    if (d?.kind === "alias") to = which === "pStyle" ? d.id : undefined;
    else if (d?.kind === "merge" || d?.kind === "import") to = d.id;
    return to && to !== val ? tag.replace(`w:val="${val}"`, `w:val="${to}"`) : tag;
  });
}

/** Give each paragraph of a stamped style its list directly. A paragraph with its own numbering
 *  already overrides the style, and a tracked change's record of the past is left as found. */
export function stampNumbering(xml: string, plan: StylePlan): string {
  if (plan.stamps.size === 0 || !xml.includes("<w:pStyle")) return xml;
  const open = /<w:pPr(?:\s[^>]*?)?(?<!\/)>/g;
  let out = "";
  let at = 0;
  for (let m = open.exec(xml); m; m = open.exec(xml)) {
    const start = m.index + m[0].length;
    const end = paragraphPropsEnd(xml, start);
    if (end === -1) break;
    open.lastIndex = end;
    const inner = xml.slice(start, end);
    const own = inner.split("<w:pPrChange")[0];
    const stamp = plan.stamps.get(childVal(own, "w:pStyle") ?? "");
    if (!stamp || own.includes("<w:numPr")) continue;
    const numPr = `<w:numPr><w:ilvl w:val="${stamp.ilvl}"/><w:numId w:val="${stamp.numId}"/></w:numPr>`;
    out += xml.slice(at, start) + canonicalizeParagraphProps(inner + numPr);
    at = end;
  }
  return out + xml.slice(at);
}

// ─── What a paragraph's style chain gives it ────────────────────────────────

/** What a paragraph takes from its styles and that a merge can lose: its direction and where its lines stand. */
export interface ParagraphLook {
  bidi: boolean;
  jc?: string;
  line: string;
  lineRule: string;
  before: string;
  after: string;
}

const SPACING_ATTRS = ["line", "lineRule", "before", "after"] as const;
const OFF = /\bw:val="(?:0|false|off)"/;

function applyPPr(look: ParagraphLook, pPr: string): void {
  const own = pPr.split("<w:pPrChange")[0];
  const bidi = /<w:bidi\b[^>]*>/.exec(own)?.[0];
  if (bidi) look.bidi = !OFF.test(bidi);
  const jc = childVal(own, "w:jc");
  if (jc) look.jc = jc;
  const spacing = /<w:spacing\b[^>]*>/.exec(own)?.[0];
  if (spacing) for (const a of SPACING_ATTRS) {
    const v = new RegExp(`\\bw:${a}="([^"]*)"`).exec(spacing)?.[1];
    if (v !== undefined) look[a] = v;
  }
}

/** The look of a paragraph of this style, through its `basedOn` chain and the document's defaults. `null`: the default style. */
export function effectiveParagraphLook(stylesXml: string | null | undefined, styleId: string | null | undefined): ParagraphLook {
  const look: ParagraphLook = { bidi: false, line: "240", lineRule: "auto", before: "0", after: "0" };
  const xml = stylesXml ?? "";
  const defaults = /<w:pPrDefault\b[^>]*>([\s\S]*?)<\/w:pPrDefault>/.exec(xml)?.[1];
  if (defaults) applyPPr(look, defaults);
  const styles = readStyles(xml).filter((s) => s.type === "paragraph");
  const start = styles.find((s) => s.id === styleId) ?? styles.find((s) => s.isDefault);
  const chain: StyleInfo[] = [];
  const seen = new Set<string>();
  for (let at = start; at && !seen.has(at.id); at = styles.find((s) => s.id === at!.basedOn)) {
    seen.add(at.id);
    chain.unshift(at);
  }
  for (const s of chain) {
    const pPr = /<w:pPr\b[^>]*>([\s\S]*?)<\/w:pPr>/.exec(s.xml)?.[1];
    if (pPr) applyPPr(look, pPr);
  }
  return look;
}

/**
 * Write on every paragraph of a copied block, nested ones included, what its style gave it in the source
 * and the style it lands on would not. What the paragraph says itself stands. `lookOf` takes the source's
 * style id, `null` for none, and returns the properties to write, or null when nothing differs.
 */
export function stampParagraphLook(xml: string, lookOf: (styleId: string | null) => Partial<ParagraphLook> | null): string {
  const open = /<w:p(?:\s[^>]*?)?(\/)?>/g;
  let out = "";
  let at = 0;
  for (let m = open.exec(xml); m; m = open.exec(xml)) {
    const after = m.index + m[0].length;
    const selfClosed = !!m[1];
    const hasPPr = !selfClosed && /^\s*<w:pPr[\s/>]/.test(xml.slice(after, after + 12));
    let innerStart = after;
    let innerEnd = after;
    let inner = "";
    let replaceEnd = after;
    if (hasPPr) {
      const pPrOpen = /^\s*<w:pPr(?:\s[^>]*?)?(\/)?>/.exec(xml.slice(after))!;
      if (pPrOpen[1]) replaceEnd = after + pPrOpen[0].length;
      else {
        innerStart = after + pPrOpen[0].length;
        innerEnd = paragraphPropsEnd(xml, innerStart);
        if (innerEnd === -1) break;
        inner = xml.slice(innerStart, innerEnd);
        replaceEnd = innerEnd + "</w:pPr>".length;
      }
    }
    const own = inner.split("<w:pPrChange")[0];
    const wanted = lookOf(childVal(own, "w:pStyle") ?? null);
    if (!wanted) continue;
    let next = inner;
    if (wanted.bidi !== undefined && !own.includes("<w:bidi")) next += wanted.bidi ? "<w:bidi/>" : '<w:bidi w:val="0"/>';
    if (wanted.jc !== undefined && !own.includes("<w:jc")) next += `<w:jc w:val="${wanted.jc}"/>`;
    const attrs = SPACING_ATTRS.filter((a) => wanted[a] !== undefined);
    if (attrs.length) {
      const had = /<w:spacing\b[^>]*?\/?>/.exec(own)?.[0];
      const missing = attrs.filter((a) => !(had && new RegExp(`\\bw:${a}=`).test(had))).map((a) => ` w:${a}="${wanted[a]}"`).join("");
      if (had && missing) next = next.replace(had, had.replace(/\s*(\/?)>$/, `${missing}$1>`));
      else if (!had && missing) next += `<w:spacing${missing}/>`;
    }
    if (next === inner) continue;
    const pPr = `<w:pPr>${canonicalizeParagraphProps(next)}</w:pPr>`;
    if (selfClosed) {
      out += xml.slice(at, m.index) + m[0].replace(/\s*\/>$/, ">") + pPr + "</w:p>";
      at = after;
    } else {
      out += xml.slice(at, after) + pPr;
      at = replaceEnd;
    }
    open.lastIndex = Math.max(open.lastIndex, at);
  }
  return out + xml.slice(at);
}

/** Where the `w:pPr` whose content starts at `from` closes; a `w:pPrChange` nests another. */
function paragraphPropsEnd(xml: string, from: number): number {
  const tag = /<w:pPr(?:\s[^>]*?)?(\/)?>|<\/w:pPr>/g;
  tag.lastIndex = from;
  let depth = 1;
  for (let m = tag.exec(xml); m; m = tag.exec(xml)) {
    if (m[0] === "</w:pPr>") {
      if (--depth === 0) return m.index;
    } else if (!m[1]) {
      depth++;
    }
  }
  return -1;
}

/** Rewrite, in a copied (parsed) list definition, what it says about styles: a list style's link
 *  follows the style, and a level links only to a style that came along with this list. */
export function retargetListRefs(node: any, plan: StylePlan): void {
  if (Array.isArray(node)) {
    for (const n of node) retargetListRefs(n, plan);
    return;
  }
  if (!node || typeof node !== "object") return;
  for (const key of Object.keys(node)) {
    if (key === "w:styleLink" || key === "w:numStyleLink" || key === "w:rStyle") {
      const d = plan.decisions.get(node[key]?.$?.["w:val"]);
      if (d?.kind === "merge" || d?.kind === "import" || d?.kind === "alias") node[key].$["w:val"] = d.id;
      else if (key === "w:rStyle") delete node[key];
    } else if (key === "w:pStyle") {
      const d = plan.decisions.get(node[key]?.$?.["w:val"]);
      if (d?.kind === "import") node[key].$["w:val"] = d.id;
      else delete node[key];
    } else if (key !== "$") {
      retargetListRefs(node[key], plan);
    }
  }
}
