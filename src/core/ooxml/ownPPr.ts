/**
 * Where a paragraph's OWN `<w:pPr>` is. `w:pPr` is by schema the paragraph's first child, and a
 * tracked property change nests the OLD properties in a `w:pPr` of its own
 * (`<w:pPr>…<w:pPrChange><w:pPr>…</w:pPr></w:pPrChange></w:pPr>`), so neither "the first `<w:pPr`
 * in the paragraph" (it may be a text box's) nor "the first `</w:pPr>`" (it may be the nested one)
 * is safe to splice at. Every reader and writer of a paragraph's properties goes through this.
 */

const PPR_TAG = /<w:pPr(?=[\s/>])[^>]*?(\/?)>|<\/w:pPr>/g;

export interface OwnPPr {
  /** Offset of `<w:pPr` in the string. */
  start: number;
  /** Offset just past the opening tag. */
  innerStart: number;
  /** Offset of the matching `</w:pPr>`; equals innerStart for a self-closing `<w:pPr/>`. */
  innerEnd: number;
  /** Offset just past the whole element. */
  end: number;
  selfClosing: boolean;
}

/**
 * The `w:pPr` that starts at `from`, whitespace aside; null when something else stands there or the
 * element never closes.
 */
export function pPrAt(xml: string, from: number): OwnPPr | null {
  let start = from;
  while (start < xml.length && /\s/.test(xml[start])) start++;
  const re = new RegExp(PPR_TAG.source, "g");
  re.lastIndex = start;
  const first = re.exec(xml);
  if (!first || first.index !== start || first[0].startsWith("</")) return null;
  const innerStart = re.lastIndex;
  if (first[1] === "/") return { start, innerStart, innerEnd: innerStart, end: innerStart, selfClosing: true };
  let depth = 1;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[0].startsWith("</")) {
      if (--depth === 0) return { start, innerStart, innerEnd: m.index, end: re.lastIndex, selfClosing: false };
    } else if (m[1] !== "/") depth++;
  }
  return null;
}

/** The own `w:pPr` of a paragraph element; null when it has none. */
export function ownPPr(paragraphXml: string): OwnPPr | null {
  const open = /^\s*<w:p(?=[\s/>])[^>]*?(\/?)>/.exec(paragraphXml);
  if (!open || open[1] === "/") return null;
  return pPrAt(paragraphXml, open[0].length);
}
