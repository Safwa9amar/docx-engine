import { parseStringPromise, Builder } from "xml2js";

export async function parseXml<T = any>(
  xml: string,
  opts?: {
    /**
     * ⚠️ `trim: true` (the default, kept for compatibility) trims EVERY text
     * node — which DELETES significant whitespace from `<w:t>` runs: a
     * space-only run (`<w:t xml:space="preserve"> </w:t>`, how Word carries
     * the space between two styled words) becomes empty, and a trailing space
     * inside a run vanishes. On a parse→mutate→rebuild round-trip that is
     * permanent DATA LOSS — words glue together in the saved document.
     * Read-or-rewrite paths over run text must pass `trim: false`.
     */
    trim?: boolean;
  },
): Promise<T> {
  return parseStringPromise(xml, {
    explicitArray: false, // prefer objects to arrays where possible
    preserveChildrenOrder: true,
    attrkey: "$",
    charkey: "_",
    trim: opts?.trim ?? true,
    // Companion to trim:false for whitespace-only text in UNattributed nodes.
    // ⚠️ It is not sufficient: for a node WITH attributes — exactly Word's
    // `<w:t xml:space="preserve"> </w:t>` — xml2js drops whitespace-only text
    // regardless of this flag (the parse returns only the `$` attrs, no text
    // key). Callers that must keep such spaces have to protect them BEFORE the
    // parse — see Table.fromXml's sentinel encoding.
    includeWhiteChars: opts?.trim === false,
  }) as T;
}

export function buildXml(
  obj: any,
  options?: {
    rootName?: string;
    headless?: boolean;
    pretty?: boolean;
    indent?: string;
    newline?: string;
    standalone?: boolean;
  }
) {
  const builder = new Builder({
    rootName: options?.rootName,
    headless: options?.headless ?? false,
    xmldec: { version: "1.0", encoding: "UTF-8", standalone: options?.standalone ?? true },
    renderOpts: {
      pretty: options?.pretty ?? true,
      indent: options?.indent ?? "  ",
      newline: options?.newline ?? "\n",
    },
  });
  return builder.buildObject(obj);
}
