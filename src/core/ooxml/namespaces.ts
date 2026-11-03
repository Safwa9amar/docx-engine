/**
 * Namespace prefixes a part uses but never declares. Such a part is well-formed to a tag matcher, and
 * Word refuses the whole file: a root that declares only `xmlns:w` (some generators write that) plus a
 * section reference written as `r:id` is enough. `declareUsedPrefixes` is string surgery on the root
 * tag only, so nothing else in the part moves.
 */

/** The prefixes Word itself writes, with what they stand for. */
const KNOWN: Readonly<Record<string, string>> = {
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  o: "urn:schemas-microsoft-com:office:office",
  v: "urn:schemas-microsoft-com:vml",
  w10: "urn:schemas-microsoft-com:office:word",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  a14: "http://schemas.microsoft.com/office/drawing/2010/main",
  w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
  w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
};

/** Markup only: comments, CDATA and instructions are skipped whole, text never reaches the pattern. */
const MARKUP = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<!(?:[^>"']|"[^"]*"|'[^']*')*>|<(\/?)([A-Za-z_][\w.\-]*:)?[A-Za-z_][\w.\-]*((?:\s+[^\s=<>/]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
const ATTR = /([^\s=<>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

export interface UndeclaredPrefix { prefix: string; count: number }

export function undeclaredPrefixes(xml: string): UndeclaredPrefix[] {
  const counts = new Map<string, number>();
  const note = (prefix: string) => counts.set(prefix, (counts.get(prefix) ?? 0) + 1);
  /** Each open element's own declarations; null when it declares none, which is nearly always. */
  const scope: (Set<string> | null)[] = [];
  const declared = (prefix: string) => {
    for (let i = scope.length - 1; i >= 0; i--) if (scope[i]?.has(prefix)) return true;
    return false;
  };
  for (const m of xml.matchAll(MARKUP)) {
    if (m[0].startsWith("<!") || m[0].startsWith("<?")) continue;
    if (m[1] === "/") { scope.pop(); continue; }
    const attrs = m[3] ?? "";
    let own: Set<string> | null = null;
    const used: string[] = [];
    let ignorable = "";
    if (attrs) {
      for (const a of attrs.matchAll(ATTR)) {
        const name = a[1];
        const colon = name.indexOf(":");
        if (colon < 0) continue;
        const prefix = name.slice(0, colon);
        if (prefix === "xmlns") (own ??= new Set()).add(name.slice(colon + 1));
        else if (prefix !== "xml") used.push(prefix);
        if (name === "mc:Ignorable") ignorable = a[2] ?? a[3] ?? "";
      }
    }
    scope.push(own);
    if (m[2]) {
      const p = m[2].slice(0, -1);
      if (p !== "xml" && !declared(p)) note(p);
    }
    for (const p of used) if (!declared(p)) note(p);
    for (const p of ignorable.split(/\s+/).filter(Boolean)) if (!declared(p)) note(p);
    if (m[4] === "/") scope.pop();
  }
  return [...counts].map(([prefix, count]) => ({ prefix, count }));
}

/** Declare on the root every undeclared prefix this module knows. An unknown one stays for the doctor. */
export function declareUsedPrefixes(xml: string): string {
  const add = undeclaredPrefixes(xml).filter((u) => KNOWN[u.prefix]);
  if (!add.length) return xml;
  // The root tag, past the XML declaration and anything else that is not an element.
  const root = /<[A-Za-z_][^<>]*?(\/?)>/.exec(xml.replace(/<\?[\s\S]*?\?>|<!--[\s\S]*?-->/g, (s) => " ".repeat(s.length)));
  if (!root) return xml;
  const at = root.index + root[0].length - 1 - root[1].length;
  return xml.slice(0, at) + add.map((u) => ` xmlns:${u.prefix}="${KNOWN[u.prefix]}"`).join("") + xml.slice(at);
}

/** Whether a prefix can be repaired by `declareUsedPrefixes`. */
export const isKnownPrefix = (prefix: string): boolean => Object.prototype.hasOwnProperty.call(KNOWN, prefix);
