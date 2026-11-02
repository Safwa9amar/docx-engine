/**
 * Presentation wipe: strip DIRECT formatting from one body block so a regenerated
 * style cascade shows through, by string surgery only.
 *
 * Rescue rebuilds a thesis's formatting inside the student's own file. The
 * xml2js-backed paths regroup `<w:body>` and trim spaces, so nothing here
 * touches them: a block's XML goes in, a block's XML comes out, and every byte
 * that is not a stripped property is copied through.
 *
 * What is kept is what carries MEANING (bold, italic, underline, script, text
 * direction, language, character style, hidden text); what is removed is what
 * carries LOOK (fonts, sizes, colours, spacing, alignment, indents).
 *
 * Never edited, whatever they contain: text boxes (`w:txbxContent`),
 * `mc:AlternateContent`, `w:drawing`, math (`m:oMath*`) and field instructions
 * (`w:instrText`). Those spans are lifted out before any rewrite and put back
 * verbatim.
 */

import { canonicalizeParagraphProps, canonicalizeRunProps, elementName, splitTopLevelElements } from "./canonicalOrder";

export interface WipeOptions {
  /** Remove `<w:br w:type="page"/>` and `<w:lastRenderedPageBreak/>`. Default true. */
  removePageBreaks?: boolean;
  /** Also remove a paragraph's own `w:sectPr` (a section break). Default false. */
  removeParagraphSectPr?: boolean;
}

export interface WipeResult {
  xml: string;
  runsStripped: number;
  paragraphsStripped: number;
  pageBreaksRemoved: number;
  sectPrRemoved: number;
}

/** Run properties that carry meaning rather than look. Without the `w:` prefix. */
export const KEPT_RUN_PROPS: readonly string[] = [
  "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vertAlign", "rtl", "cs",
  "lang", "rStyle", "vanish", "specVanish", "noProof",
];

/** Paragraph properties that carry look. Without the `w:` prefix. */
export const REMOVED_PARA_PROPS: readonly string[] = [
  "spacing", "jc", "ind", "keepNext", "keepLines", "pageBreakBefore", "widowControl",
  "framePr", "shd", "pBdr", "contextualSpacing", "textAlignment", "snapToGrid",
  "adjustRightInd", "mirrorIndents",
];

/** Tracked-change markers inside an rPr are revision data, not formatting. */
const REVISION_MARKS = new Set(["w:ins", "w:del", "w:moveFrom", "w:moveTo"]);

const KEPT_RUN = new Set(KEPT_RUN_PROPS.map((p) => `w:${p}`));
const REMOVED_PARA = new Set(REMOVED_PARA_PROPS.map((p) => `w:${p}`));

/** Containers whose content is never rewritten. */
const PROTECTED_TAGS = [
  "w:txbxContent", "mc:AlternateContent", "w:drawing", "m:oMathPara", "m:oMath", "w:instrText",
];

// ─── protected spans ─────────────────────────────────────────────────────────

/** Index just past the close of the element opened at `openAt`, or -1. */
function findElementEnd(xml: string, openAt: number, tag: string): number {
  const re = new RegExp(`<(/?)${tag}(?=[\\s/>])[^>]*?(/?)>`, "g");
  re.lastIndex = openAt;
  let depth = 0;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[1]) {
      if (--depth === 0) return re.lastIndex;
    } else if (m[2]) {
      if (depth === 0) return re.lastIndex; // self-closing root
    } else {
      depth++;
    }
  }
  return -1;
}

/** Replace every protected span with a markup-free token; return the spans. */
function lift(xml: string): { text: string; spans: string[] } {
  const spans: string[] = [];
  const opener = new RegExp(`<(?:${PROTECTED_TAGS.join("|")})(?=[\\s/>])`, "g");
  let out = "";
  let at = 0;
  for (let m = opener.exec(xml); m; m = opener.exec(xml)) {
    if (m.index < at) continue;
    const tag = m[0].slice(1);
    const end = findElementEnd(xml, m.index, tag);
    if (end === -1) continue; // unclosed: leave it to the normal rules
    out += xml.slice(at, m.index) + `\u0001${spans.length}\u0001`;
    spans.push(xml.slice(m.index, end));
    at = end;
    opener.lastIndex = end;
  }
  return { text: out + xml.slice(at), spans };
}

const restore = (text: string, spans: string[]) =>
  text.replace(/\u0001(\d+)\u0001/g, (_m, i) => spans[Number(i)]);

// ─── property filters ────────────────────────────────────────────────────────

/** Filter an rPr's inner XML to the kept set. */
function filterRunProps(inner: string): string {
  if (!inner.trim()) return "";
  const keep = splitTopLevelElements(inner).filter((el) => {
    const n = elementName(el);
    return KEPT_RUN.has(n) || REVISION_MARKS.has(n);
  });
  return canonicalizeRunProps(keep.join(""));
}

const BREAK_RE = /<w:br\b(?=[^>]*\bw:type="page")[^>]*?(?:\/>|>\s*<\/w:br>)|<w:lastRenderedPageBreak\b[^>]*?(?:\/>|>\s*<\/w:lastRenderedPageBreak>)/g;

/** A run with nothing left in it but (optionally) properties. */
const EMPTY_RUN_RE = /^<w:r(?:\s[^>]*?)?(?<!\/)>\s*(?:<w:rPr\s*\/>|<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>)?\s*<\/w:r>$/;

const RUN_RE = /<w:r(?:\s[^>]*?)?(?<!\/)>[\s\S]*?<\/w:r>/g;
const RUN_RPR_RE = /^(<w:r(?:\s[^>]*?)?(?<!\/)>)(\s*)(<w:rPr\s*\/>|<w:rPr(?:\s[^>]*)?>([\s\S]*?)<\/w:rPr>)/;

/** An opening `<w:p …>` (not self-closing) followed directly by its `<w:pPr`. */
const P_OPEN_PPR_RE = /<w:p(?:\s[^>]*?)?(?<!\/)>(\s*)<w:pPr(?=[\s/>])/g;

export function wipeBlockPresentation(blockXml: string, opts: WipeOptions = {}): WipeResult {
  const removeBreaks = opts.removePageBreaks !== false;
  const removeSect = opts.removeParagraphSectPr === true;
  const res: WipeResult = { xml: blockXml, runsStripped: 0, paragraphsStripped: 0, pageBreaksRemoved: 0, sectPrRemoved: 0 };

  const { text, spans } = lift(blockXml);

  // Runs: strip rPr, drop page breaks, drop a run left empty.
  let work = text.replace(RUN_RE, (run) => {
    let r = run;
    let changed = false;

    const rp = RUN_RPR_RE.exec(r);
    if (rp) {
      const inner = rp[4] ?? "";
      const next = filterRunProps(inner);
      const replacement = next ? `<w:rPr>${next}</w:rPr>` : "";
      if (replacement !== rp[3]) {
        r = rp[1] + (next ? rp[2] : "") + replacement + r.slice(rp[0].length);
        changed = true;
        res.runsStripped++;
      }
    }

    if (removeBreaks) {
      let n = 0;
      const stripped = r.replace(BREAK_RE, () => (n++, ""));
      if (n) {
        res.pageBreaksRemoved += n;
        r = stripped;
        if (EMPTY_RUN_RE.test(r)) return "";
        changed = true;
      }
    }
    return changed ? r : run;
  });

  // Paragraphs: strip the paragraph's OWN pPr (the one right after its open tag).
  let out = "";
  let at = 0;
  P_OPEN_PPR_RE.lastIndex = 0;
  for (let m = P_OPEN_PPR_RE.exec(work); m; m = P_OPEN_PPR_RE.exec(work)) {
    const pPrStart = m.index + m[0].length - "<w:pPr".length;
    const selfClosing = /^<w:pPr\s*\/>/.exec(work.slice(pPrStart));
    let pPrEnd: number;
    let inner: string;
    if (selfClosing) {
      pPrEnd = pPrStart + selfClosing[0].length;
      inner = "";
    } else {
      pPrEnd = findElementEnd(work, pPrStart, "w:pPr");
      if (pPrEnd === -1) continue;
      const openEnd = work.indexOf(">", pPrStart) + 1;
      inner = work.slice(openEnd, pPrEnd - "</w:pPr>".length);
    }
    if (!inner.trim()) continue;

    const kept = splitTopLevelElements(inner).filter((el) => {
      const n = elementName(el);
      if (REMOVED_PARA.has(n)) return false;
      if (n === "w:sectPr" && removeSect) {
        res.sectPrRemoved++;
        return false;
      }
      return true;
    }).map((el) => {
      if (elementName(el) !== "w:rPr") return el;
      const mark = /^<w:rPr(?:\s[^>]*)?>([\s\S]*)<\/w:rPr>$/.exec(el);
      if (!mark) return ""; // self-closing mark rPr
      const f = filterRunProps(mark[1]);
      return f ? `<w:rPr>${f}</w:rPr>` : "";
    }).filter(Boolean);

    const next = kept.length ? `<w:pPr>${canonicalizeParagraphProps(kept.join(""))}</w:pPr>` : "";
    const before = work.slice(pPrStart, pPrEnd);
    if (next === before) continue;
    res.paragraphsStripped++;
    // Keep the whitespace between the open tag and pPr only when a pPr remains.
    const headEnd = m.index + m[0].length - "<w:pPr".length - m[1].length;
    out += work.slice(at, headEnd) + (next ? m[1] + next : "");
    at = pPrEnd;
    P_OPEN_PPR_RE.lastIndex = pPrEnd;
  }
  work = out + work.slice(at);

  res.xml = restore(work, spans);
  return res;
}
