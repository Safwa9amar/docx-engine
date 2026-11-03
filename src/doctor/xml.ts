/**
 * Depth-aware XML scanning primitives for the document doctor.
 *
 * Deliberately NOT an XML parser. Every repair the doctor performs rewrites at
 * most one element's children, and the whole point is that untouched bytes stay
 * byte-identical — a parse/rebuild round-trip is itself one of the corruption
 * causes this module exists to detect (see ../doctor/index.ts).
 *
 * The scanner is quote-, comment- and CDATA-aware, and every walk counts depth on
 * the element's OWN tag name so nesting cannot fool it.
 */

/** Index just past the `>` that closes the markup starting at `lt`, or -1.
 *  Handles comments, CDATA, processing instructions and `>` inside attributes. */
export function markupEnd(xml: string, lt: number): number {
  if (xml.startsWith("<!--", lt)) {
    const end = xml.indexOf("-->", lt + 4);
    return end === -1 ? -1 : end + 3;
  }
  if (xml.startsWith("<![CDATA[", lt)) {
    const end = xml.indexOf("]]>", lt + 9);
    return end === -1 ? -1 : end + 3;
  }
  let quote = "";
  for (let i = lt + 1; i < xml.length; i++) {
    const c = xml[i];
    if (quote) {
      if (c === quote) quote = "";
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === ">") {
      return i + 1;
    }
  }
  return -1;
}

/** The tag name of the element starting at `lt` ("" for comments/CDATA/PIs). */
export function tagNameAt(xml: string, lt: number): string {
  let i = lt + 1;
  if (xml[i] === "/") i++;
  if (xml[i] === "!" || xml[i] === "?") return "";
  const start = i;
  while (i < xml.length && !/[\s/>]/.test(xml[i])) i++;
  return xml.slice(start, i);
}

/** True when `<tag` at `lt` is the whole tag name and not a longer one
 *  (`<w:tblPr` must not match `<w:tblPrEx`). */
export function isTag(xml: string, lt: number, tag: string): boolean {
  if (!xml.startsWith(`<${tag}`, lt)) return false;
  const after = xml[lt + 1 + tag.length];
  return after === ">" || after === "/" || after === " " || after === "\t" || after === "\n" || after === "\r";
}

/** The index just past the END of the complete element that starts at `lt`
 *  — its whole subtree, not just its open tag. -1 when it never closes.
 *
 *  Depth-counted on the element's OWN tag name, so `<w:pPr>` inside a
 *  `<w:pPrChange>` inside a `<w:pPr>` closes the right one. Getting this wrong
 *  is not a subtle bug: a scanner that walks tag-by-tag instead of
 *  subtree-by-subtree treats a nested `<w:rPr>` as a child of the outer element
 *  and will happily hoist it out of the element it belongs to. */
export function elementEnd(xml: string, lt: number): number {
  const openEnd = markupEnd(xml, lt);
  if (openEnd === -1) return -1;
  if (xml[openEnd - 2] === "/") return openEnd; // self-closing
  const tag = tagNameAt(xml, lt);
  if (!tag) return openEnd; // comment / CDATA / PI — already whole

  const close = `</${tag}>`;
  let depth = 1;
  let i = openEnd;
  while (i < xml.length && depth > 0) {
    const next = xml.indexOf("<", i);
    if (next === -1) return -1;
    const nEnd = markupEnd(xml, next);
    if (nEnd === -1) return -1;
    if (xml.startsWith(close, next)) {
      if (--depth === 0) return nEnd;
    } else if (isTag(xml, next, tag) && xml[nEnd - 2] !== "/") {
      depth++;
    }
    i = nEnd;
  }
  return -1;
}

export interface ElementRange {
  /** `<` of the open tag. */ start: number;
  /** just past `>` of the close tag. */ end: number;
  /** just past `>` of the open tag (=== innerEnd when self-closing). */ innerStart: number;
  /** `<` of the close tag. */ innerEnd: number;
  selfClosing: boolean;
}

/** Every OUTERMOST, non-overlapping occurrence of `tag` in `xml`, in order.
 *  Outermost-only means a `w:pPr` nested inside a `w:pPrChange` (tracked-change
 *  history) is skipped — reordering a revision record would be pointless anyway. */
export function findElements(xml: string, tag: string): ElementRange[] {
  const out: ElementRange[] = [];
  const open = `<${tag}`;
  let from = 0;

  for (;;) {
    let lt = xml.indexOf(open, from);
    while (lt !== -1 && !isTag(xml, lt, tag)) lt = xml.indexOf(open, lt + 1);
    if (lt === -1) return out;

    const openEnd = markupEnd(xml, lt);
    if (openEnd === -1) return out;
    if (xml[openEnd - 2] === "/") {
      out.push({ start: lt, end: openEnd, innerStart: openEnd, innerEnd: openEnd, selfClosing: true });
      from = openEnd;
      continue;
    }

    const end = elementEnd(xml, lt);
    if (end === -1) return out; // unbalanced — leave the rest alone
    out.push({ start: lt, end, innerStart: openEnd, innerEnd: end - tag.length - 3, selfClosing: false });
    from = end;
  }
}

/** One child of an element: `xml` is its COMPLETE subtree, and `lead` is the
 *  text/whitespace that preceded it (kept glued so re-ordering a pretty-printed
 *  part keeps its indentation). */
export interface ChildItem { lead: string; tag: string; xml: string }

/** Split an element's inner XML into child SUBTREES. Returns null when the
 *  content is mixed (real text between elements) or unbalanced — neither is
 *  something we are willing to reorder. */
export function splitChildren(inner: string): { items: ChildItem[]; tail: string } | null {
  const items: ChildItem[] = [];
  let lead = "";
  let i = 0;

  while (i < inner.length) {
    const lt = inner.indexOf("<", i);
    if (lt === -1) {
      const rest = inner.slice(i);
      if (rest.trim()) return null;
      return { items, tail: lead + rest };
    }
    if (lt > i) {
      const text = inner.slice(i, lt);
      if (text.trim()) return null; // mixed content
      lead += text;
    }
    const tag = tagNameAt(inner, lt);
    if (!tag) return null; // comment / CDATA / PI — don't shuffle around it
    if (inner[lt + 1] === "/") return null; // a close tag here means we mis-scanned
    const end = elementEnd(inner, lt);
    if (end === -1) return null;
    items.push({ lead, tag, xml: inner.slice(lt, end) });
    lead = "";
    i = end;
  }
  return { items, tail: lead };
}

/** Reorder an element's children into schema order. Returns the new inner XML,
 *  or null when it was already in order (or is not safely reorderable). */
export function reorderInner(inner: string, rank: Map<string, number>): string | null {
  if (!inner.trim()) return null;
  const split = splitChildren(inner);
  if (!split || split.items.length < 2) return null;

  // Unknown children glue to the last known one (`cursor`), so a vendor
  // extension keeps the position it was authored in.
  let cursor = -1;
  const keyed = split.items.map((item, i) => {
    const known = rank.get(item.tag);
    if (known !== undefined) cursor = known;
    return { item, rank: cursor, i };
  });

  const sorted = [...keyed].sort((a, b) => (a.rank - b.rank) || (a.i - b.i));
  if (sorted.every((k, i) => k.i === i)) return null; // already in order

  const rebuilt = sorted.map((k) => k.item.lead + k.item.xml).join("") + split.tail;
  return rebuilt === inner ? null : rebuilt;
}

/** Apply `transform` to every outermost `tag` element's inner XML. Edits are
 *  applied back-to-front so earlier offsets stay valid. */
export function rewriteElements(
  xml: string,
  tag: string,
  transform: (inner: string) => string | null,
): { xml: string; changed: number } {
  const ranges = findElements(xml, tag);
  let out = xml;
  let changed = 0;
  for (let i = ranges.length - 1; i >= 0; i--) {
    const r = ranges[i];
    if (r.selfClosing) continue;
    const next = transform(xml.slice(r.innerStart, r.innerEnd));
    if (next == null) continue;
    out = out.slice(0, r.innerStart) + next + out.slice(r.innerEnd);
    changed++;
  }
  return { xml: out, changed };
}

const ATTRIBUTE = /\s([^\s=<>/]+)\s*=\s*(?:"[^"]*"|'[^']*')/g;

/** The name of an attribute an open tag states twice; null when none does. Values are skipped whole. */
function repeatedAttribute(tag: string): string | null {
  const first = tag.indexOf("=");
  if (first < 0 || tag.indexOf("=", first + 1) < 0) return null;
  const seen = new Set<string>();
  for (const m of tag.matchAll(ATTRIBUTE)) {
    if (seen.has(m[1])) return m[1];
    seen.add(m[1]);
  }
  return null;
}

/** A cheap well-formedness check: tag balance with quote/comment awareness.
 *  Not a validator — its job is to catch a part we (or a previous writer) tore. */
export function firstXmlError(xml: string): string | null {
  const stack: string[] = [];
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    if (lt === -1) break;
    const end = markupEnd(xml, lt);
    if (end === -1) return "unterminated tag";
    if (!xml.startsWith("<!", lt) && !xml.startsWith("<?", lt)) {
      const tag = tagNameAt(xml, lt);
      if (!tag) return "unreadable tag name";
      if (xml[lt + 1] === "/") {
        const open = stack.pop();
        if (open !== tag) return `</${tag}> closes <${open ?? "nothing"}>`;
      } else {
        // XML forbids it, and Word refuses the whole file for it: a tag matcher alone would pass this.
        const repeated = repeatedAttribute(xml.slice(lt, end));
        if (repeated) return `<${tag}> repeats attribute ${repeated}`;
        if (xml[end - 2] !== "/") stack.push(tag);
      }
    }
    i = end;
  }
  return stack.length ? `unclosed <${stack[stack.length - 1]}>` : null;
}
