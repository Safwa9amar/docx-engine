/**
 * Paragraph-property writers for STYLES and docDefaults, by string surgery.
 *
 * Pure: they take and return XML fragments (a `<w:pPr>` inner string), never
 * touch a zip, and never go through xml2js (which regroups and trims). Every
 * result is re-sorted into `CT_PPr` order, because Word rejects a sequence
 * violation.
 */

import { canonicalizeParagraphProps } from "./canonicalOrder";

export interface ParaStyleProps {
  /** Line spacing as a MULTIPLE of single (1.5 → `w:line=360 w:lineRule=auto`). */
  lineSpacing?: number;
  beforePt?: number;
  afterPt?: number;
  jc?: "left" | "right" | "center" | "both" | "start" | "end";
  /** First-line indent. Replaces any hanging indent. */
  firstLineCm?: number;
  /** Hanging indent. Replaces any first-line indent. */
  hangingCm?: number;
  /** Start-side indent (the right side in an RTL paragraph). */
  startCm?: number;
  /** `false` removes the element. */
  keepNext?: boolean;
  keepLines?: boolean;
  /** `null` removes the element. */
  outlineLvl?: number | null;
  bidi?: boolean;
  pageBreakBefore?: boolean;
}

const JC = new Set(["left", "right", "center", "both", "start", "end"]);
const TWIPS_PER_CM = 567;

const twips = (cm: number, what: string): number => {
  if (!Number.isFinite(cm)) throw new Error(`styleProps: ${what} must be a finite number, got ${cm}`);
  return Math.round(cm * TWIPS_PER_CM);
};
const ptTwips = (pt: number, what: string): number => {
  if (!Number.isFinite(pt) || pt < 0) throw new Error(`styleProps: ${what} must be a non-negative number, got ${pt}`);
  return Math.round(pt * 20);
};

/** Remove every `<tag …/>` / `<tag …>…</tag>` from a pPr inner string. */
function strip(inner: string, tag: string): string {
  return inner
    .replace(new RegExp(`<${tag}(?=[\\s/>])[^>]*/>`, "g"), "")
    .replace(new RegExp(`<${tag}(?=[\\s>])[^>]*>[\\s\\S]*?</${tag}>`, "g"), "");
}

/** The attributes of the first `<tag …>` in `inner`, in order. */
function attrsOf(inner: string, tag: string): Map<string, string> {
  const m = new RegExp(`<${tag}(?=[\\s/>])([^>]*?)/?>`).exec(inner);
  const out = new Map<string, string>();
  if (m) for (const a of m[1].matchAll(/\s(w:\w+)=["']([^"']*)["']/g)) out.set(a[1], a[2]);
  return out;
}

const elem = (tag: string, attrs: Map<string, string>) =>
  `<${tag}${[...attrs].map(([k, v]) => ` ${k}="${v}"`).join("")}/>`;

/** Replace `tag` with one carrying `attrs` (existing attributes kept unless overwritten). */
function upsert(inner: string, tag: string, set: Record<string, string>, drop: string[] = []): string {
  const attrs = attrsOf(inner, tag);
  for (const d of drop) attrs.delete(d);
  for (const [k, v] of Object.entries(set)) attrs.set(k, v);
  return strip(inner, tag) + elem(tag, attrs);
}

/** Apply `props` to a `<w:pPr>` inner string; the result is canonical. */
export function applyParaStyleProps(pPrInner: string, props: ParaStyleProps): string {
  let out = pPrInner;

  if (props.lineSpacing !== undefined || props.beforePt !== undefined || props.afterPt !== undefined) {
    const set: Record<string, string> = {};
    if (props.lineSpacing !== undefined) {
      if (!Number.isFinite(props.lineSpacing) || props.lineSpacing <= 0) {
        throw new Error(`styleProps: lineSpacing must be a positive number, got ${props.lineSpacing}`);
      }
      set["w:line"] = String(Math.round(240 * props.lineSpacing));
      set["w:lineRule"] = "auto";
    }
    if (props.beforePt !== undefined) set["w:before"] = String(ptTwips(props.beforePt, "beforePt"));
    if (props.afterPt !== undefined) set["w:after"] = String(ptTwips(props.afterPt, "afterPt"));
    // Line-based spacing would override the twip values just written.
    const drop = [
      ...(set["w:before"] ? ["w:beforeLines", "w:beforeAutospacing"] : []),
      ...(set["w:after"] ? ["w:afterLines", "w:afterAutospacing"] : []),
    ];
    out = upsert(out, "w:spacing", set, drop);
  }

  if (props.jc !== undefined) {
    if (!JC.has(props.jc)) throw new Error(`styleProps: jc must be one of ${[...JC].join(", ")}, got ${props.jc}`);
    out = strip(out, "w:jc") + `<w:jc w:val="${props.jc}"/>`;
  }

  if (props.firstLineCm !== undefined || props.hangingCm !== undefined || props.startCm !== undefined) {
    const set: Record<string, string> = {};
    const drop: string[] = [];
    if (props.startCm !== undefined) {
      set["w:start"] = String(twips(props.startCm, "startCm"));
      drop.push("w:left"); // two spellings of the same thing: never write both
    }
    if (props.firstLineCm !== undefined) {
      set["w:firstLine"] = String(twips(props.firstLineCm, "firstLineCm"));
      drop.push("w:hanging", "w:firstLineChars", "w:hangingChars");
    }
    if (props.hangingCm !== undefined) {
      set["w:hanging"] = String(twips(props.hangingCm, "hangingCm"));
      drop.push("w:firstLine", "w:firstLineChars", "w:hangingChars");
    }
    out = upsert(out, "w:ind", set, drop);
  }

  const toggle = (tag: string, v: boolean | undefined) => {
    if (v === undefined) return;
    out = strip(out, tag) + (v ? `<${tag}/>` : "");
  };
  toggle("w:keepNext", props.keepNext);
  toggle("w:keepLines", props.keepLines);
  toggle("w:bidi", props.bidi);
  toggle("w:pageBreakBefore", props.pageBreakBefore);

  if (props.outlineLvl !== undefined) {
    out = strip(out, "w:outlineLvl");
    if (props.outlineLvl !== null) {
      if (!Number.isInteger(props.outlineLvl) || props.outlineLvl < 0 || props.outlineLvl > 9) {
        throw new Error(`styleProps: outlineLvl must be an integer 0-9 or null, got ${props.outlineLvl}`);
      }
      out += `<w:outlineLvl w:val="${props.outlineLvl}"/>`;
    }
  }

  return canonicalizeParagraphProps(out);
}
