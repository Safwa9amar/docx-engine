/**
 * Where a later part's sections go in a merge.
 *
 * A section's properties govern the text BEFORE them. So a part's first section break, copied after
 * part 1's text, also governs part 1's closing text, and the part's last section, which has no break
 * to travel with, is simply lost: its text ends up under part 1's last section.
 *
 * The answer is a seam: part 1's last section becomes a break on part 1's last paragraph, and the last
 * section of the merged file takes the later part's page. The seam starts no page of its own (a
 * section with another page size does, whatever its type says). When the later part has no break
 * and the same page, none of this is needed and nothing is added.
 */
import { upsertParagraphSectPr } from "@/core/files/body/sectPr";
import type { BodyBlock } from "@/core/files/body/OrderedBody";

const SECT = /<w:sectPr\b/;

/** Where the first section in this XML ends. A tracked change nests the old section inside the new. */
function firstSection(xml: string): { start: number; end: number } | null {
  const tag = /<w:sectPr\b[^>]*?(\/?)>|<\/w:sectPr>/g;
  let depth = 0;
  let start = -1;
  for (let m = tag.exec(xml); m; m = tag.exec(xml)) {
    if (m[0].startsWith("</")) {
      if (--depth === 0) return { start, end: tag.lastIndex };
    } else {
      if (depth === 0) start = m.index;
      if (m[1]) {
        if (depth === 0) return { start, end: tag.lastIndex };
      } else depth++;
    }
  }
  return null;
}

/** CT_SectPr is a sequence. */
const ORDER = [
  "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
  "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
];
/** What makes a section's page what it is. */
const PAGE = new Set(["pgSz", "pgMar", "cols", "textDirection", "bidi", "rtlGutter", "docGrid"]);
/** Marks where a section begins, which the part of it that carries on past a seam does not do. */
const BEGINNING = new Set(["type", "titlePg", "sectPrChange"]);

interface Child {
  name: string;
  xml: string;
}

function childrenOf(sect: string | null): Child[] {
  const open = sect && /^<w:sectPr\b[^>]*?(\/?)>/.exec(sect);
  if (!sect || !open || open[1]) return [];
  const inner = sect.slice(open[0].length, sect.lastIndexOf("</w:sectPr>"));
  return [...inner.matchAll(/<w:([A-Za-z]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g)].map((m) => ({ name: m[1], xml: m[0] }));
}

const attrs = (xml: string | undefined, names: string[]) => names.map((n) => (xml && new RegExp(`\\bw:${n}="([^"]*)"`).exec(xml)?.[1]) ?? "").join(",");

/** The same page, as far as `source` states one. A paper code, a text grid or a header's distance is no difference worth a break. */
function samePage(target: string | null, source: string | null): boolean {
  const t = childrenOf(target);
  const s = childrenOf(source);
  const get = (list: Child[], name: string) => list.find((c) => c.name === name)?.xml;
  const size = get(s, "pgSz");
  if (size) {
    const facts = (xml: string | undefined) => {
      const [w, h, orient] = attrs(xml, ["w", "h", "orient"]).split(",");
      return `${w},${h},${orient || "portrait"}`;
    };
    if (facts(size) !== facts(get(t, "pgSz"))) return false;
  }
  const margins = get(s, "pgMar");
  if (margins && attrs(margins, ["top", "right", "bottom", "left", "gutter"]) !== attrs(get(t, "pgMar"), ["top", "right", "bottom", "left", "gutter"])) return false;
  if (!size && !margins) return true;
  const columns = (list: Child[]) => attrs(get(list, "cols"), ["num"]) || "1";
  const flag = (list: Child[], name: string) => {
    const xml = get(list, name);
    return xml ? !/\bw:val="(?:0|false|off)"/.test(xml) : false;
  };
  return columns(t) === columns(s) && flag(t, "bidi") === flag(s, "bidi") && attrs(get(t, "textDirection"), ["val"]) === attrs(get(s, "textDirection"), ["val"]);
}

function build(sect: string | null, children: Child[]): string {
  const open = (sect && /^<w:sectPr\b[^>]*?(\/?)>/.exec(sect)?.[0].replace(/\s*\/>$/, ">")) || "<w:sectPr>";
  const rank = (c: Child) => {
    const i = ORDER.indexOf(c.name);
    return i < 0 ? ORDER.length : i;
  };
  const sorted = children.map((c, i) => ({ c, i })).sort((a, b) => rank(a.c) - rank(b.c) || a.i - b.i);
  return `${open}${sorted.map(({ c }) => c.xml).join("")}</w:sectPr>`;
}

/** This section's chrome and notes, carrying on, on the source's page. It does not begin again: no new first page, no restart of the numbers. */
function carryingOn(target: string | null, source: string | null, type: string | undefined): string {
  const s = childrenOf(source);
  const statesPage = s.some((c) => c.name === "pgSz" || c.name === "pgMar");
  const kept: Child[] = [];
  for (const c of childrenOf(target)) {
    if (BEGINNING.has(c.name)) continue;
    if (statesPage && PAGE.has(c.name) && !(c.name === "pgSz" && !s.some((x) => x.name === "pgSz")) && !(c.name === "pgMar" && !s.some((x) => x.name === "pgMar"))) continue;
    kept.push(c.name === "pgNumType" ? { name: c.name, xml: c.xml.replace(/\s+w:start="[^"]*"/, "") } : c);
  }
  if (statesPage) kept.push(...s.filter((c) => PAGE.has(c.name)));
  if (type) kept.push({ name: "type", xml: type });
  return build(target, kept);
}

const CONTINUOUS = `<w:type w:val="continuous"/>`;

function continuous(sect: string): string {
  const children = childrenOf(sect).filter((c) => c.name !== "type");
  return build(sect, [...children, { name: "type", xml: CONTINUOUS }]);
}

const isBlank = (b: BodyBlock) => b.kind === "other" && !b.xml.trim();

export interface Seam {
  existing: BodyBlock[];
  incoming: BodyBlock[];
  /** The merged file's last section. Null: it stays what it is. */
  final: string | null;
}

export function seamSections(existing: BodyBlock[], incoming: BodyBlock[], targetFinal: string | null, sourceFinal: string | null): Seam {
  const firstBreak = incoming.findIndex((b) => SECT.test(b.xml));
  if (firstBreak < 0 && samePage(targetFinal, sourceFinal)) return { existing, incoming, final: null };

  const sourceType = childrenOf(sourceFinal).find((c) => c.name === "type")?.xml;
  const blank = existing.every(isBlank);
  // Nothing of part 1 to keep a page for; its section goes on to carry the source's text.
  if (blank) return { existing, incoming, final: carryingOn(targetFinal, sourceFinal, firstBreak < 0 ? undefined : sourceType) };

  const seam = targetFinal ?? "<w:sectPr/>";
  const out = [...existing];
  let last = out.length - 1;
  while (last >= 0 && isBlank(out[last])) last--;
  const carrier = out[last];
  if (carrier.kind === "paragraph" && carrier.tag === "w:p" && !SECT.test(carrier.xml)) {
    out[last] = { ...carrier, xml: upsertParagraphSectPr(carrier.xml, seam) };
  } else {
    out.splice(last + 1, 0, { kind: "paragraph", tag: "w:p", xml: `<w:p><w:pPr>${seam}</w:pPr></w:p>` });
  }

  const next = [...incoming];
  if (firstBreak >= 0) {
    const { xml } = next[firstBreak];
    const at = firstSection(xml);
    if (at) next[firstBreak] = { ...next[firstBreak], xml: xml.slice(0, at.start) + continuous(xml.slice(at.start, at.end)) + xml.slice(at.end) };
  }
  return { existing: out, incoming: next, final: carryingOn(targetFinal, sourceFinal, firstBreak < 0 ? CONTINUOUS : sourceType) };
}
