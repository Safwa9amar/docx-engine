/**
 * Root-element namespace helpers for merged WordprocessingML parts.
 *
 * A part's body is stored as exact XML substrings, so concatenating blocks from
 * a source whose ROOT declared `w14`/`mc`/`wps`/`wp14` onto a target whose root
 * did not leaves those prefixes unbound — well-formed XML forbids that and Word
 * refuses to open the file ("-10000"). `carryRootNamespaces` copies the
 * declarations across, editing ONLY the target's root open tag.
 */

export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** The namespace set a Word-authored part root declares. */
export const WORD_NAMESPACES: Record<string, string> = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  mc: MC_NS,
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w10: "urn:schemas-microsoft-com:office:word",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
  wne: "http://schemas.microsoft.com/office/word/2006/wordml",
};

/** Root attributes (xmlns:* + mc:Ignorable) for a freshly created Word part. */
export function wordRootAttributes(): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [p, uri] of Object.entries(WORD_NAMESPACES)) out[`xmlns:${p}`] = uri;
  out["mc:Ignorable"] = "w14 w15 w16cid wp14";
  return out;
}

function rootOpen(xml: string): { start: number; end: number; tag: string } | null {
  const m = /<(?![?!])[A-Za-z_][\w.-]*(?::[\w.-]+)?\b[^>]*>/.exec(xml);
  return m ? { start: m.index, end: m.index + m[0].length, tag: m[0] } : null;
}

const declsOf = (tag: string) =>
  new Map([...tag.matchAll(/\sxmlns:([\w.-]+)="([^"]*)"/g)].map((m) => [m[1], m[2]] as [string, string]));

function ignorableOf(tag: string, ns: Map<string, string>) {
  for (const m of tag.matchAll(/\s([\w.-]+):Ignorable="([^"]*)"/g)) {
    if (ns.get(m[1]) === MC_NS) {
      return { attr: m[0], prefix: m[1], tokens: m[2].split(/\s+/).filter(Boolean) };
    }
  }
  return null;
}

export interface CarryResult {
  xml: string;
  added: string[];
  /** Prefixes bound to a DIFFERENT uri in the source than in the target. */
  conflicts: string[];
}

/**
 * Add every `xmlns:p` the source root declares and the target root lacks, and
 * union the source's `mc:Ignorable` tokens (only prefixes now declared). Reuses
 * the prefix the target already binds to the MC namespace (legacy roots use
 * `ve`). Never touches anything but the target's root open tag.
 */
export function carryRootNamespaces(targetXml: string, sourceXml: string | null | undefined): CarryResult {
  const none: CarryResult = { xml: targetXml, added: [], conflicts: [] };
  if (!sourceXml) return none;
  const t = rootOpen(targetXml);
  const s = rootOpen(sourceXml);
  if (!t || !s) return none;

  const tNs = declsOf(t.tag);
  const sNs = declsOf(s.tag);
  const added: string[] = [];
  const conflicts: string[] = [];
  const close = t.tag.endsWith("/>") ? "/>" : ">";
  let body = t.tag.slice(0, -close.length);

  for (const [p, uri] of sNs) {
    if (!tNs.has(p)) {
      body += ` xmlns:${p}="${uri}"`;
      tNs.set(p, uri);
      added.push(p);
    } else if (tNs.get(p) !== uri) {
      conflicts.push(p);
    }
  }

  const sIgn = ignorableOf(s.tag, sNs);
  if (sIgn) {
    const tIgn = ignorableOf(body, tNs);
    const have = new Set(tIgn?.tokens ?? []);
    const extra = sIgn.tokens.filter((p) => !have.has(p) && tNs.has(p));
    if (extra.length) {
      if (tIgn) {
        body = body.replace(tIgn.attr, ` ${tIgn.prefix}:Ignorable="${[...tIgn.tokens, ...extra].join(" ")}"`);
      } else {
        // The source declared its MC prefix (that is how it has an Ignorable), so
        // the loop above has bound the MC namespace on the target by now.
        const mcp = [...tNs].find(([, u]) => u === MC_NS)![0];
        body += ` ${mcp}:Ignorable="${extra.join(" ")}"`;
      }
    }
  }

  return {
    xml: targetXml.slice(0, t.start) + body + close + targetXml.slice(t.end),
    added,
    conflicts,
  };
}
