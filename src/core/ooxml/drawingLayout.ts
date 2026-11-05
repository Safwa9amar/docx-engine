/**
 * Picture PLACEMENT surgery: where a `<w:drawing>` sits on the page.
 *
 * Word stores a picture in one of two containers, and which one it is decides
 * everything about where it can go:
 *
 *  - `<wp:inline>` — the picture IS a character in the text flow. It can only be
 *    where the line is; "put it in the middle of the page" is not expressible.
 *  - `<wp:anchor>` — the picture FLOATS, with a `wp:positionH`/`wp:positionV`
 *    pair naming a frame of reference (`page`, `margin`, …) and either an offset
 *    or a named alignment. This is Word's Layout ▸ Position dialog, and
 *    "Position in Middle Center relative to Page" is exactly
 *    `<wp:align>center</wp:align>` on both axes relative to `page`.
 *
 * So vertical placement is not a property you can set on a picture — it is a
 * container conversion. That is what {@link applyDrawingLayout} does, in both
 * directions, copying every other byte of the drawing through untouched (the
 * `a:graphic` subtree, which holds the image reference, is never even read).
 *
 * TWO RULES the schema enforces, both of which would produce a file Word refuses
 * to open if broken:
 *
 *  1. `CT_Anchor` is an ORDERED sequence — simplePos, positionH, positionV,
 *     extent, effectExtent?, <wrap>, docPr, cNvGraphicFramePr?, graphic. The
 *     wrap element in particular goes immediately BEFORE `wp:docPr`, not at the
 *     end where appending would put it.
 *  2. Geometry is read and written on the PREFIX before `<a:graphic` only. A
 *     shape's text box can contain a whole nested drawing, so a free search for
 *     `wp:positionV` can find a child's and move the wrong picture.
 */

/** Frames of reference this module will write. Word accepts more; these are the
 *  two that mean anything to a student ("on the page" / "inside the margins"). */
export type DrawingRelativeTo = "page" | "margin";

/** Text wrapping around a floating picture, in Word's UI vocabulary. */
export type DrawingWrap = "none" | "square" | "topAndBottom";

/** Where a floating picture sits. `undefined` = leave that axis alone. */
export interface DrawingLayout {
  horizontal?: "left" | "center" | "right";
  vertical?: "top" | "center" | "bottom";
  /** Frame of reference for BOTH axes. Default `page`. */
  relativeTo?: DrawingRelativeTo;
  /**
   * Default `topAndBottom`, not Word's own `square`. A thesis figure is close to
   * the full column width, and square wrapping round one leaves a two-word
   * sliver of text down its side; top-and-bottom reserves the band the picture
   * occupies and never does. On the page this default was built for — a lone
   * picture centred on an otherwise empty page — the two are identical.
   */
  wrap?: DrawingWrap;
  /** `false` returns the picture to the text flow (anchor → inline). */
  float?: boolean;
}

/** What a drawing's placement currently is. */
export interface DrawingPlacement {
  /** `wp:anchor` rather than `wp:inline`. */
  floating: boolean;
  horizontal: { relativeTo: string; align: string | null; offsetEmu: number | null } | null;
  vertical: { relativeTo: string; align: string | null; offsetEmu: number | null } | null;
  wrap: DrawingWrap | "tight" | "through" | "inline";
  /** Painted behind the text (Word's "Send Behind Text"). */
  behindDoc: boolean;
  widthEmu: number;
  heightEmu: number;
}

export interface DrawingLayoutResult {
  xml: string;
  changed: boolean;
  placement: DrawingPlacement;
  /** The drawing ships with an `mc:Fallback` VML twin that was left untouched —
   *  see {@link applyDrawingLayout}. Informational; nothing needs to act on it. */
  legacyTwin: boolean;
}

const WRAP_TAGS: Record<string, DrawingPlacement["wrap"]> = {
  wrapNone: "none",
  wrapSquare: "square",
  wrapTight: "tight",
  wrapThrough: "through",
  wrapTopAndBottom: "topAndBottom",
};

/** The wrap element for a wrap kind. `square` needs its `wrapText` attribute. */
function wrapElement(wrap: DrawingWrap): string {
  if (wrap === "none") return "<wp:wrapNone/>";
  if (wrap === "topAndBottom") return "<wp:wrapTopAndBottom/>";
  return '<wp:wrapSquare wrapText="bothSides"/>';
}

/**
 * The span of the OUTERMOST `<wp:inline>` or `<wp:anchor>` in `drawingXml`.
 *
 * Found by walking from the container's opening tag with a depth counter over
 * BOTH tag names, so a nested drawing inside a text box cannot end the outer
 * element early.
 */
interface ContainerSpan {
  kind: "inline" | "anchor";
  /** Index of `<wp:inline`/`<wp:anchor` in the source string. */
  start: number;
  /** Index just past `</wp:inline>`/`</wp:anchor>`. */
  end: number;
  /** The opening tag itself, e.g. `<wp:anchor behindDoc="1" …>`. */
  open: string;
  /** Everything between the opening and closing tags. */
  inner: string;
}

const CONTAINER_TAG_RE = /<wp:(inline|anchor)(?=[\s/>])([^>]*?)(\/?)>|<\/wp:(inline|anchor)>/g;

function findContainer(xml: string): ContainerSpan | null {
  CONTAINER_TAG_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  let depth = 0;
  let start = -1;
  let kind: "inline" | "anchor" = "inline";
  let openTag = "";
  while ((m = CONTAINER_TAG_RE.exec(xml))) {
    if (m[4]) {
      // A closing tag.
      depth--;
      if (depth === 0 && start >= 0) {
        return {
          kind,
          start,
          end: m.index + m[0].length,
          open: openTag,
          inner: xml.slice(start + openTag.length, m.index),
        };
      }
      continue;
    }
    // A self-closing `<wp:inline/>` holds no picture; it is not a container.
    if (m[3]) continue;
    if (depth === 0) {
      start = m.index;
      kind = m[1] as "inline" | "anchor";
      openTag = m[0];
    }
    depth++;
  }
  return null;
}

/** Attributes of the container's opening tag, as a raw string. */
function openAttrs(open: string): string {
  return /^<wp:(?:inline|anchor)((?:\s[^>]*?)?)>$/.exec(open)?.[1] ?? "";
}

/** A namespace declaration on a tag: `xmlns:x="…"` or the default `xmlns="…"`, leading space included. */
const NS_DECL_RE = /\sxmlns(?::[\w.-]+)?="[^"]*"/g;

/**
 * The namespace declarations on the container's own opening tag. `addImage` writes
 * `xmlns:wp` and `xmlns:a` there, not on a root above, and the `a:graphic` that
 * is copied through still uses `a:`: a container rebuilt without them leaves a
 * prefix nothing declares, which Word refuses to open.
 */
function namespaceDeclarations(open: string): string {
  return open.match(NS_DECL_RE)?.join("") ?? "";
}

/** The opening tag's attributes with its namespace declarations taken out, so they are written once. */
function withoutNamespaceDeclarations(attrs: string): string {
  return attrs.replace(NS_DECL_RE, "");
}

/** The part of a container's inner XML that holds its geometry — everything
 *  before `<a:graphic`, which is where a nested drawing could start. */
function geometryPrefix(inner: string): { geo: string; rest: string } {
  const cut = inner.indexOf("<a:graphic");
  return cut < 0 ? { geo: inner, rest: "" } : { geo: inner.slice(0, cut), rest: inner.slice(cut) };
}

function readAxis(geo: string, axis: "positionH" | "positionV"): DrawingPlacement["horizontal"] {
  const m = new RegExp(`<wp:${axis}\\b[^>]*\\brelativeFrom="([^"]+)"[^>]*>([\\s\\S]*?)</wp:${axis}>`).exec(geo);
  if (!m) return null;
  const align = /<wp:align>\s*(\w+)\s*<\/wp:align>/.exec(m[2])?.[1] ?? null;
  const offset = /<wp:posOffset>\s*(-?\d+)\s*<\/wp:posOffset>/.exec(m[2])?.[1];
  return { relativeTo: m[1], align, offsetEmu: offset === undefined ? null : Number(offset) };
}

function placementOf(span: ContainerSpan): DrawingPlacement {
  const { geo } = geometryPrefix(span.inner);
  const extent = /<wp:extent\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(geo);
  const wrapTag = /<wp:(wrapNone|wrapSquare|wrapTight|wrapThrough|wrapTopAndBottom)\b/.exec(geo)?.[1];
  return {
    floating: span.kind === "anchor",
    horizontal: span.kind === "anchor" ? readAxis(geo, "positionH") : null,
    vertical: span.kind === "anchor" ? readAxis(geo, "positionV") : null,
    wrap: span.kind === "inline" ? "inline" : (wrapTag && WRAP_TAGS[wrapTag]) || "none",
    behindDoc: /\bbehindDoc="(?:1|true)"/.test(span.open),
    widthEmu: extent ? Number(extent[1]) : 0,
    heightEmu: extent ? Number(extent[2]) : 0,
  };
}

/** Read where the first drawing in `xml` is placed. `null` when there is none. */
export function readDrawingLayout(xml: string): DrawingPlacement | null {
  const span = findContainer(xml);
  return span ? placementOf(span) : null;
}

/**
 * True when the drawing at `at` sits inside an `<mc:Fallback>` — the legacy VML
 * copy of a picture that Word 2007 and later never render. Scanning backwards
 * for the nearest open tag is enough: `mc:Fallback` does not nest inside itself,
 * and its own `</mc:Fallback>` closes it.
 */
function insideLegacyFallback(xml: string, at: number): boolean {
  const open = xml.lastIndexOf("<mc:Fallback", at);
  if (open < 0) return false;
  const close = xml.indexOf("</mc:Fallback>", open);
  return close < 0 || close > at;
}

/** An anchor's wrap element, self-closing or paired (wrapTight/wrapThrough carry
 *  a `wp:wrapPolygon`). Matched on the geometry prefix only. */
const WRAP_ELEMENT_RE =
  /<wp:wrap(?:None|Square|TopAndBottom|Tight|Through)\b[^>]*\/>|<wp:wrap(?:Tight|Through)\b[^>]*>[\s\S]*?<\/wp:wrap(?:Tight|Through)>/;

/** Strip the positioning children an anchor owns and an inline must not have. */
function stripAnchorGeometry(geo: string): string {
  return geo
    .replace(/<wp:simplePos\b[^>]*\/>/g, "")
    .replace(/<wp:positionH\b[\s\S]*?<\/wp:positionH>/g, "")
    .replace(/<wp:positionV\b[\s\S]*?<\/wp:positionV>/g, "")
    .replace(/<wp:wrap(?:None|Square|Tight|Through|TopAndBottom)\b[^>]*\/>/g, "")
    .replace(/<wp:wrap(?:Tight|Through)\b[^>]*>[\s\S]*?<\/wp:wrap(?:Tight|Through)>/g, "");
}

/**
 * Insert `fragment` immediately before `<wp:docPr`, which is where the wrap
 * element belongs in `CT_Anchor`'s sequence. A drawing always has a `wp:docPr`
 * (it carries the picture's name and id); if one is somehow missing, the
 * fragment goes last, which is still ahead of `a:graphic`.
 */
function insertBeforeDocPr(geo: string, fragment: string): string {
  const at = geo.indexOf("<wp:docPr");
  return at < 0 ? geo + fragment : geo.slice(0, at) + fragment + geo.slice(at);
}

/** The `wp:positionH`/`wp:positionV` pair for an anchor, in schema order. */
function positionElements(
  horizontal: NonNullable<DrawingLayout["horizontal"]>,
  vertical: NonNullable<DrawingLayout["vertical"]>,
  relativeTo: DrawingRelativeTo,
): string {
  return (
    `<wp:positionH relativeFrom="${relativeTo}"><wp:align>${horizontal}</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="${relativeTo}"><wp:align>${vertical}</wp:align></wp:positionV>`
  );
}

/**
 * Word's own defaults for a picture it floats out of the flow. `distL`/`distR`
 * of 114300 EMU (0.125") is what Word writes for a square-wrapped picture, so a
 * paragraph beside it does not touch it.
 */
const ANCHOR_ATTRS =
  'distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"';

/**
 * Place the first picture in a paragraph, converting its container as needed.
 *
 * - Asking for a `vertical` position (or `float: true`) FLOATS an inline
 *   picture: `wp:inline` → `wp:anchor` with both axes written as named
 *   alignments relative to `relativeTo` (default `page`). This is the only way
 *   "centre it on the page" can be expressed in OOXML.
 * - `float: false` returns a floating picture to the text flow — the positioning
 *   children are dropped and the container becomes `wp:inline` again.
 * - Passing ONLY `horizontal` on an inline picture changes nothing here: it is
 *   the carrier PARAGRAPH's `w:jc` that moves an in-flow picture sideways, and
 *   floating a picture just to nudge it left would change how text flows around
 *   it. The caller handles that case (`changed: false` with `floating: false`
 *   is the signal). On an already-floating picture, `horizontal` IS applied.
 *
 * Everything outside the container's geometry prefix — the whole `a:graphic`
 * subtree, so the image itself, its crop, its effects — is copied byte for byte.
 *
 * ON `mc:AlternateContent`. A great many real pictures ship as a pair: an
 * `mc:Choice` holding the modern DrawingML, and an `mc:Fallback` holding a VML
 * twin of the same picture for readers older than Word 2007. This function used
 * to refuse the whole paragraph on sight of one, on the theory that rewriting a
 * single branch leaves the file inconsistent. That was the wrong trade, and a
 * student's بسم الله page proved it within the hour: the tool refused, the
 * assistant reported the capability as missing, and the student got nothing —
 * over a compatibility branch that no Word since 2007 has read.
 *
 * So: the `mc:Choice` drawing IS the picture, and it is repositioned. The
 * `mc:Fallback` twin is left exactly as it was — Word never renders it, our own
 * reader already discards it (`stripAltContentFallback`), and rewriting VML
 * geometry is a second, unrelated language. `legacyTwin` reports that it is
 * there. Only a drawing found INSIDE the fallback is refused, because moving the
 * copy nobody renders would look, to the student, like nothing happened.
 *
 * @throws {Error} when `xml` holds no drawing this can move.
 */
export function applyDrawingLayout(xml: string, opts: DrawingLayout): DrawingLayoutResult {
  const span = findContainer(xml);
  if (!span) throw new Error("drawingLayout: no <wp:inline>/<wp:anchor> drawing in this block");
  if (insideLegacyFallback(xml, span.start)) {
    throw new Error(
      "drawingLayout: the only drawing here is inside an mc:Fallback (the legacy copy Word does not render), so moving it would change nothing on the page",
    );
  }
  const legacyTwin = /<mc:Fallback[\s>]/.test(xml);

  const current = placementOf(span);
  const wantsFloat =
    opts.float === true || (opts.float !== false && opts.vertical !== undefined);
  const wantsInline = opts.float === false;

  // ── Back into the text flow ───────────────────────────────────────────────
  if (wantsInline) {
    if (span.kind === "inline") return { xml, changed: false, placement: current, legacyTwin };
    const { geo, rest } = geometryPrefix(span.inner);
    const inner = stripAnchorGeometry(geo) + rest;
    const next =
      xml.slice(0, span.start) +
      `<wp:inline distT="0" distB="0" distL="0" distR="0"${namespaceDeclarations(span.open)}>${inner}</wp:inline>` +
      xml.slice(span.end);
    return { xml: next, changed: true, placement: readDrawingLayout(next)!, legacyTwin };
  }

  // ── Float it (or re-place one that already floats) ────────────────────────
  if (!wantsFloat && span.kind === "inline") {
    // Horizontal-only on an in-flow picture: nothing to do at the drawing level.
    return { xml, changed: false, placement: current, legacyTwin };
  }

  const relativeTo: DrawingRelativeTo = opts.relativeTo ?? "page";
  const horizontal =
    opts.horizontal ??
    (current.horizontal?.align as DrawingLayout["horizontal"]) ??
    "center";
  const vertical =
    opts.vertical ?? (current.vertical?.align as DrawingLayout["vertical"]) ?? "center";
  const { geo, rest } = geometryPrefix(span.inner);
  // An already-floating picture keeps its OWN wrap element, verbatim, unless the
  // caller names a new one — including the tight/through kinds this module will
  // not write. Re-deriving it from `current.wrap` would quietly flatten a tight
  // wrap (and its whole wrapPolygon) into a rectangle nobody asked for.
  const wrapXml =
    opts.wrap !== undefined
      ? wrapElement(opts.wrap)
      : (span.kind === "anchor" ? WRAP_ELEMENT_RE.exec(geo)?.[0] : undefined) ?? wrapElement("topAndBottom");

  let nextGeo = stripAnchorGeometry(geo);
  nextGeo = insertBeforeDocPr(nextGeo, wrapXml);
  nextGeo =
    '<wp:simplePos x="0" y="0"/>' + positionElements(horizontal, vertical, relativeTo) + nextGeo;

  // An anchor that already exists keeps its own attributes (z-order, behindDoc,
  // layoutInCell) — they are the student's, not ours to reset. A freshly floated
  // picture gets Word's defaults.
  const attrs =
    span.kind === "anchor"
      ? withoutNamespaceDeclarations(openAttrs(span.open)).trim() || ANCHOR_ATTRS
      : ANCHOR_ATTRS;
  // The declarations travel with the container (written once, after the attributes): anything in scope at
  // the old position is still in scope at the new one, so the copied a:graphic keeps resolving its prefixes.
  const next =
    xml.slice(0, span.start) +
    `<wp:anchor ${attrs}${namespaceDeclarations(span.open)}>${nextGeo}${rest}</wp:anchor>` +
    xml.slice(span.end);
  return { xml: next, changed: next !== xml, placement: readDrawingLayout(next)!, legacyTwin };
}
