/**
 * Paragraph-level OOXML surgery: the `<w:pPr>` counterpart to `runProps.ts`.
 *
 * Two capabilities live here, both of which act on ONE paragraph's XML and
 * leave every other byte of it untouched:
 *
 *  - PAGINATION (`w:keepNext`, `w:keepLines`, `w:widowControl`,
 *    `w:pageBreakBefore`) — Word's "keep with next" / "keep lines together" /
 *    "widow and orphan control" / "page break before" paragraph properties.
 *    These are the ones that stop a heading being stranded alone at the foot of
 *    a page, which no tool could express before.
 *
 *  - DIRECT RUN FORMATTING for a SINGLE paragraph ({@link applyRunPropsToParagraph}).
 *    `TextStyleManager` styles named PARTS of a document through their Word
 *    style; this styles one paragraph and nothing else, which is what "make THIS
 *    heading 48pt" means — restyling Heading1 would resize every chapter title.
 *
 * Everything routes through {@link updateParagraphProps}, which re-canonicalises
 * `<w:pPr>` into `CT_PPr` order on the way out. That is not a nicety: `w:keepNext`
 * belongs directly after `w:pStyle`, near the FRONT of the sequence, while an
 * Arabic thesis paragraph already carries `w:bidi` and `w:jc` from much later in
 * it. Appending would emit a sequence violation and Word would refuse the file.
 */

import { canonicalizeParagraphProps } from "./canonicalOrder";
import { pPrAt } from "./ownPPr";
import { mergeRunProps, type RunProps } from "./runProps";

/** Word's paragraph pagination toggles. `undefined` = leave as it is. */
export interface ParagraphPagination {
  /** Keep with next — the paragraph never ends a page alone, ahead of its body text. */
  keepWithNext?: boolean;
  /** Keep lines together — the paragraph is never split across two pages. */
  keepLines?: boolean;
  /** Widow/orphan control — never leave a single line of it on a page by itself. */
  widowControl?: boolean;
  /** Always start this paragraph on a new page. */
  pageBreakBefore?: boolean;
}

/** Which `w:pPr` child each pagination flag owns. */
const PAGINATION_TAGS: Record<keyof ParagraphPagination, string> = {
  keepWithNext: "w:keepNext",
  keepLines: "w:keepLines",
  widowControl: "w:widowControl",
  pageBreakBefore: "w:pageBreakBefore",
};

/**
 * The paragraph's opening tag. The `(\/)?` is load-bearing: `[^>]*` alone would
 * swallow the slash of a SELF-CLOSING `<w:p w:rsidR="00A1"/>` — a perfectly
 * ordinary empty paragraph — and every writer below would then splice children
 * in after a tag that had already closed, producing markup Word cannot open.
 * Attribute values may not contain a raw `>` in well-formed XML, so `[^>]` is a
 * safe attribute-run.
 */
const P_OPEN_RE = /^(\s*)<w:p((?:\s[^>]*?)?)(\/)?>/;

interface ParagraphSplit {
  /** Leading whitespace + the opening tag, always in PAIRED form. */
  head: string;
  /** Everything after it, `</w:p>` included. Synthesised for a self-closing `<w:p/>`. */
  rest: string;
}

/** Split a paragraph into its opening tag and the rest, normalising `<w:p/>`. */
function splitParagraph(xml: string): ParagraphSplit | null {
  const m = P_OPEN_RE.exec(xml);
  if (!m) return null;
  const head = `${m[1]}<w:p${m[2] ?? ""}>`;
  const after = xml.slice(m[0].length);
  return { head, rest: m[3] ? `</w:p>${after}` : after };
}

/** True when `xml` is a `<w:p>` element this module can rewrite. */
export function isParagraphXml(xml: string): boolean {
  return P_OPEN_RE.test(xml);
}

/**
 * Splice `fragment` in immediately after a paragraph's OPENING TAG.
 *
 * Use this — never a hand-rolled `xml.replace(/<w:p\b[^>]*>/, …)` — whenever
 * anything is inserted at the head of a paragraph. `[^>]*` swallows the slash of
 * a SELF-CLOSING `<w:p w:rsidR="00A1"/>`, which is exactly what Word writes for
 * an empty paragraph, so the naive replace appends after a tag that has ALREADY
 * CLOSED: the fragment lands as the paragraph's SIBLING — a direct child of
 * `<w:body>`. A body holding a bare `<w:pPr>` is not openable in Word
 * (`body.illegal-child`, fatal), and the app draws it as an unknown-block chip.
 * {@link splitParagraph} normalises `<w:p/>` to paired form first, so the
 * fragment always lands INSIDE the paragraph.
 *
 * Anchored to the opening tag, so a paragraph nested in a text box
 * (`w:txbxContent`) can never be the one that gets written to.
 *
 * Returns `xml` unchanged when it does not begin with a `<w:p>` element, or when
 * `fragment` is empty.
 */
export function insertAfterParagraphOpen(xml: string, fragment: string): string {
  if (!fragment) return xml;
  const split = splitParagraph(xml);
  if (!split) return xml;
  return `${split.head}${fragment}${split.rest}`;
}

/**
 * A leading `<w:pPr>` — self-closing or paired — at the start of a fragment, matched by depth: a
 * `w:pPrChange` nests a `w:pPr`, and stopping at the first close would cut the element in two.
 */
function leadingPPr(fragment: string): { inner: string; end: number } | null {
  const own = pPrAt(fragment, 0);
  return own && { inner: fragment.slice(own.innerStart, own.innerEnd), end: own.end };
}

/**
 * Read a paragraph's `<w:pPr>` inner XML.
 *
 * Anchored to the paragraph's OPENING TAG rather than searching the whole
 * string, because `w:pPr` — when present — is by schema the first child of
 * `w:p`. A free search would happily find the `<w:pPr>` of a paragraph nested
 * inside a text box (`w:txbxContent`) and rewrite the wrong one.
 *
 * Returns `null` when `xml` is not a paragraph element, and `""` when it is a
 * paragraph with no properties of its own.
 */
export function readParagraphProps(xml: string): string | null {
  const split = splitParagraph(xml);
  if (!split) return null;
  return leadingPPr(split.rest)?.inner ?? "";
}

/**
 * Rewrite a paragraph's `<w:pPr>` through `mutate`, which receives the current
 * inner XML ("" when the paragraph has no `w:pPr` yet) and returns the new one.
 * The result is canonicalised into `CT_PPr` order and written back; a `w:pPr`
 * is created (or dropped, if `mutate` empties it) as needed.
 *
 * The rest of the paragraph — every run, bookmark, hyperlink, field and
 * drawing — is copied through byte for byte.
 *
 * @throws {Error} if `xml` is not a `<w:p>` element, or if the resulting
 * `w:pPr` cannot be cleanly parsed (see `canonicalOrder.splitTopLevelElements`).
 */
export function updateParagraphProps(xml: string, mutate: (pPrInner: string) => string): string {
  const split = splitParagraph(xml);
  if (!split) {
    throw new Error("paragraphProps.updateParagraphProps: not a <w:p> element");
  }
  const existing = leadingPPr(split.rest);
  const current = existing?.inner ?? "";
  const body = existing ? split.rest.slice(existing.end) : split.rest;

  const next = mutate(current);
  // Byte-identical when nothing changed — including the no-pPr/nothing-to-add
  // case, so a paragraph is never gratuitously rewritten (a self-closing
  // `<w:p/>` must not come back as `<w:p></w:p>` for no reason).
  if (next === current) return xml;
  const pPr = next.trim() ? `<w:pPr>${canonicalizeParagraphProps(next)}</w:pPr>` : "";
  return `${split.head}${pPr}${body}`;
}

/** Remove every occurrence of `tag` (self-closing or paired) from a `w:pPr` fragment. */
function stripPPrTag(pPrInner: string, tag: string): string {
  return pPrInner
    // `(?=[\s/>])` rather than `\b`, so "w:keepNext" can never match "w:keepNextFoo".
    .replace(new RegExp(`<${tag}(?=[\\s/>])[^>]*/>`, "g"), "")
    .replace(new RegExp(`<${tag}(?=[\\s>])[^>]*>[\\s\\S]*?</${tag}>`, "g"), "");
}

/**
 * Apply pagination toggles to one paragraph's XML.
 *
 * `true` writes the bare element (`<w:keepNext/>`); `false` writes an EXPLICIT
 * `w:val="0"` rather than deleting the element. Deleting would only fall back to
 * whatever the paragraph's style says — and for a Heading style that is very
 * often `keepNext` ON, so "stop keeping this heading with the next paragraph"
 * would silently do nothing. `undefined` leaves the property alone.
 */
export function applyParagraphPagination(xml: string, opts: ParagraphPagination): string {
  return updateParagraphProps(xml, (inner) => {
    let out = inner;
    for (const key of Object.keys(PAGINATION_TAGS) as (keyof ParagraphPagination)[]) {
      const want = opts[key];
      if (want === undefined) continue;
      const tag = PAGINATION_TAGS[key];
      out = stripPPrTag(out, tag);
      out += want ? `<${tag}/>` : `<${tag} w:val="0"/>`;
    }
    return out;
  });
}

/** Read back the effective (directly-set) pagination flags of a paragraph. */
export function readParagraphPagination(xml: string): ParagraphPagination {
  const inner = readParagraphProps(xml) ?? "";
  const out: ParagraphPagination = {};
  for (const key of Object.keys(PAGINATION_TAGS) as (keyof ParagraphPagination)[]) {
    const m = new RegExp(`<${PAGINATION_TAGS[key]}(?=[\\s/>])([^>]*)/?>`).exec(inner);
    if (m) out[key] = !/w:val="(?:0|false|off)"/.test(m[1] ?? "");
  }
  return out;
}

/** `<w:r …> … </w:r>` — real runs only; `\b` alone would also match `<w:rPr>`. */
const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;

/** Merge `props` into ONE run's `<w:rPr>`, creating it as the run's first child. */
function styleRun(runXml: string, props: RunProps): string {
  const selfClosing = /<w:rPr\s*\/>/.exec(runXml);
  if (selfClosing) {
    return runXml.replace(selfClosing[0], `<w:rPr>${mergeRunProps("", props)}</w:rPr>`);
  }
  const paired = /<w:rPr(?:\s[^>]*)?>([\s\S]*?)<\/w:rPr>/.exec(runXml);
  if (paired) {
    return runXml.replace(paired[0], `<w:rPr>${mergeRunProps(paired[1], props)}</w:rPr>`);
  }
  // CT_R requires run properties BEFORE the run's content.
  return runXml.replace(/^(<w:r(?:\s[^>]*)?>)/, `$1<w:rPr>${mergeRunProps("", props)}</w:rPr>`);
}

/** Merge `props` into the paragraph MARK's run properties (`<w:pPr><w:rPr>`). */
function styleParagraphMark(xml: string, props: RunProps): string {
  return updateParagraphProps(xml, (inner) => {
    const selfClosing = /<w:rPr\s*\/>/.exec(inner);
    if (selfClosing) return inner.replace(selfClosing[0], `<w:rPr>${mergeRunProps("", props)}</w:rPr>`);
    const paired = /<w:rPr(?:\s[^>]*)?>([\s\S]*?)<\/w:rPr>/.exec(inner);
    if (paired) return inner.replace(paired[0], `<w:rPr>${mergeRunProps(paired[1], props)}</w:rPr>`);
    return `${inner}<w:rPr>${mergeRunProps("", props)}</w:rPr>`;
  });
}

export interface ParagraphRunStyleResult {
  xml: string;
  /** How many runs were restyled. Zero means the paragraph holds no text runs. */
  runs: number;
}

/**
 * Apply DIRECT run formatting to every run of ONE paragraph, plus its paragraph
 * mark.
 *
 * Direct, not style-level, on purpose. Two reasons, both learned the hard way:
 * an imported thesis carries its formatting on the RUNS, and direct run
 * formatting beats a style in the OOXML cascade — so a style patch alone is
 * invisible on exactly the documents students bring us; and the whole point of
 * this function is to hit ONE paragraph, which a shared style cannot do.
 *
 * The paragraph mark is included so the size sticks to the ¶ itself: Word draws
 * the empty line after a heading at the mark's size, and text typed at the end
 * of the paragraph inherits it.
 *
 * @throws {Error} if `xml` is not a `<w:p>` element, or if `props` carries an
 * invalid size or colour (see `runProps.mergeRunProps`).
 */
export function applyRunPropsToParagraph(xml: string, props: RunProps): ParagraphRunStyleResult {
  if (!isParagraphXml(xml)) {
    throw new Error("paragraphProps.applyRunPropsToParagraph: not a <w:p> element");
  }

  // Style the mark first: it goes through updateParagraphProps, which is
  // anchored to the opening tag, so it cannot stray into the body's runs.
  const withMark = styleParagraphMark(xml, props);

  // Then the runs — but only in the BODY, never inside `<w:pPr>`. The paragraph
  // mark's `w:rPr` is not a `w:r`, so RUN_RE could not match it anyway; keeping
  // `w:pPr` out of the replaced span makes that independent of the regex.
  const split = splitParagraph(withMark)!;
  const pPr = leadingPPr(split.rest);
  const head = split.head + (pPr ? split.rest.slice(0, pPr.end) : "");
  const body = pPr ? split.rest.slice(pPr.end) : split.rest;

  let runs = 0;
  const styledBody = body.replace(RUN_RE, (run) => {
    runs++;
    return styleRun(run, props);
  });
  return { xml: head + styledBody, runs };
}
