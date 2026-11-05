import { describe, expect, it } from "vitest";

import { applyDrawingLayout, readDrawingLayout } from "./drawingLayout";
import { undeclaredPrefixes } from "./namespaces";

/** A picture in the text flow — what insert/import writes, and what an imported
 *  thesis's بسم الله page holds. */
const INLINE =
  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="4572000" cy="1600200"/>` +
  `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:docPr id="1" name="Picture 1"/>` +
  `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic>` +
  `</a:graphicData></a:graphic>` +
  `</wp:inline></w:drawing></w:r></w:p>`;

/** A shape whose TEXT BOX contains another drawing — the trap `geometryPrefix`
 *  and the depth-counting container walk exist for. */
const NESTED =
  `<w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="5">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="margin"><wp:posOffset>100</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>200</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="100" cy="200"/><wp:wrapNone/><wp:docPr id="3" name="Box"/>` +
  `<a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>` +
  `<w:p><w:r><w:drawing><wp:inline distT="0"><wp:extent cx="9" cy="9"/>` +
  `<wp:docPr id="4" name="Inner"/><a:graphic/></wp:inline></w:drawing></w:r></w:p>` +
  `</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>` +
  `</wp:anchor></w:drawing></w:r></w:p>`;

describe("readDrawingLayout", () => {
  it("reads an inline picture as in-flow", () => {
    expect(readDrawingLayout(INLINE)).toEqual({
      floating: false,
      horizontal: null,
      vertical: null,
      wrap: "inline",
      behindDoc: false,
      widthEmu: 4572000,
      heightEmu: 1600200,
    });
  });

  it("returns null for a paragraph with no drawing", () => {
    expect(readDrawingLayout("<w:p><w:r><w:t>plain</w:t></w:r></w:p>")).toBeNull();
  });

  it("reads the OUTER anchor's geometry, not the nested drawing's", () => {
    const p = readDrawingLayout(NESTED)!;
    expect(p.floating).toBe(true);
    expect(p.horizontal).toEqual({ relativeTo: "margin", align: null, offsetEmu: 100 });
    expect(p.widthEmu).toBe(100);
  });
});

describe("applyDrawingLayout — floating an inline picture", () => {
  const centred = applyDrawingLayout(INLINE, { horizontal: "center", vertical: "center" });

  it("converts wp:inline into wp:anchor", () => {
    expect(centred.changed).toBe(true);
    expect(centred.xml).toContain("<wp:anchor ");
    expect(centred.xml).not.toContain("<wp:inline");
  });

  it("writes both axes as named alignments relative to the page", () => {
    expect(centred.xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>',
    );
    expect(centred.xml).toContain(
      '<wp:positionV relativeFrom="page"><wp:align>center</wp:align></wp:positionV>',
    );
    expect(centred.placement.vertical).toEqual({ relativeTo: "page", align: "center", offsetEmu: null });
  });

  it("keeps CT_Anchor's sequence — simplePos, positions, extent, wrap, docPr, graphic", () => {
    const order = ["wp:simplePos", "wp:positionH", "wp:positionV", "wp:extent", "wp:wrapTopAndBottom", "wp:docPr", "a:graphic"]
      .map((t) => centred.xml.indexOf(`<${t}`));
    expect(order.every((i) => i > 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });

  it("defaults to top-and-bottom wrapping, not Word's square", () => {
    // A thesis figure is nearly column-wide; square wrapping round one leaves a
    // sliver of text down its side.
    expect(centred.placement.wrap).toBe("topAndBottom");
  });

  it("copies the picture itself through byte for byte", () => {
    expect(centred.xml).toContain('<a:blip r:embed="rId7"/>');
    expect(centred.xml).toContain('<wp:extent cx="4572000" cy="1600200"/>');
  });

  it("honours an explicit frame of reference and wrap", () => {
    const r = applyDrawingLayout(INLINE, { horizontal: "right", vertical: "top", relativeTo: "margin", wrap: "none" });
    expect(r.xml).toContain('<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>');
    expect(r.xml).toContain("<wp:wrapNone/>");
  });

  it("leaves an in-flow picture alone when only `horizontal` is asked for", () => {
    const r = applyDrawingLayout(INLINE, { horizontal: "center" });
    expect(r.changed).toBe(false);
    expect(r.xml).toBe(INLINE);
    expect(r.placement.floating).toBe(false);
  });
});

describe("applyDrawingLayout — re-placing and un-floating", () => {
  const floated = applyDrawingLayout(INLINE, { vertical: "center" }).xml;

  it("moves an already-floating picture without a second conversion", () => {
    const r = applyDrawingLayout(floated, { vertical: "bottom" });
    expect(r.xml).toContain("<wp:align>bottom</wp:align>");
    expect(r.xml.match(/<wp:anchor /g)).toHaveLength(1);
    // The horizontal axis it already had is kept, not reset.
    expect(r.placement.horizontal?.align).toBe("center");
  });

  it("float:false returns it to the text flow with no positioning left behind", () => {
    const r = applyDrawingLayout(floated, { float: false });
    expect(r.changed).toBe(true);
    expect(r.xml).toContain("<wp:inline ");
    expect(r.xml).not.toContain("wp:positionV");
    expect(r.xml).not.toContain("wp:wrapSquare");
    expect(r.xml).not.toContain("wp:simplePos");
    expect(r.placement.floating).toBe(false);
  });

  it("keeps a wrap it will not itself write, rather than flattening it", () => {
    const tight =
      `<w:p><w:r><w:drawing><wp:anchor behindDoc="0">` +
      `<wp:simplePos x="0" y="0"/>` +
      `<wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>` +
      `<wp:extent cx="100" cy="100"/>` +
      `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>` +
      `<wp:docPr id="9" name="Tight"/><a:graphic/></wp:anchor></w:drawing></w:r></w:p>`;
    const r = applyDrawingLayout(tight, { vertical: "center" });
    expect(r.xml).toContain("<wp:wrapPolygon edited=\"0\">");
    expect(r.placement.wrap).toBe("tight");
    expect(r.placement.vertical?.align).toBe("center");
  });

  it("float:false on an already-inline picture is a no-op", () => {
    const r = applyDrawingLayout(INLINE, { float: false });
    expect(r.changed).toBe(false);
    expect(r.xml).toBe(INLINE);
  });

  it("keeps the nested drawing untouched when the outer shape moves", () => {
    const r = applyDrawingLayout(NESTED, { vertical: "center" });
    expect(r.xml).toContain('<wp:docPr id="4" name="Inner"/>');
    expect(r.xml).toContain('<wp:inline distT="0"><wp:extent cx="9" cy="9"/>');
    expect(r.placement.vertical?.align).toBe("center");
  });
});

// A picture that ships as a PAIR: the modern DrawingML in mc:Choice, and a VML
// twin in mc:Fallback for readers older than Word 2007. Extremely common in real
// theses — the shape this module refused outright on its first day, which is how
// a student's بسم الله page went unmoved.
const PAIRED =
  `<w:p><w:r><mc:AlternateContent>` +
  `<mc:Choice Requires="wps"><w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="4572000" cy="1600200"/><wp:docPr id="1" name="Modern"/>` +
  `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
  `</wp:inline></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><v:shape id="_x0000_s1026" style="width:360pt;height:126pt">` +
  `<v:imagedata r:id="rId7"/></v:shape></w:pict></mc:Fallback>` +
  `</mc:AlternateContent></w:r></w:p>`;

describe("applyDrawingLayout — a picture with a legacy VML twin", () => {
  const r = applyDrawingLayout(PAIRED, { horizontal: "center", vertical: "center" });

  it("moves it — the mc:Choice drawing IS the picture Word renders", () => {
    expect(r.changed).toBe(true);
    expect(r.placement.vertical).toEqual({ relativeTo: "page", align: "center", offsetEmu: null });
    expect(r.legacyTwin).toBe(true);
  });

  it("leaves the fallback twin byte for byte — Word never reads it", () => {
    expect(r.xml).toContain(`<mc:Fallback><w:pict><v:shape id="_x0000_s1026" style="width:360pt;height:126pt">`);
    expect(r.xml).toContain(`<v:imagedata r:id="rId7"/>`);
    expect(r.xml).toContain("</mc:AlternateContent>");
    expect(r.xml).toContain(`<mc:Choice Requires="wps">`);
  });
});

describe("applyDrawingLayout — refusals", () => {
  it("refuses a paragraph with no drawing", () => {
    expect(() => applyDrawingLayout("<w:p><w:r><w:t>hi</w:t></w:r></w:p>", { vertical: "center" })).toThrow(
      /no <wp:inline>/,
    );
  });

  it("refuses when the ONLY drawing is the copy Word does not render", () => {
    const fallbackOnly =
      `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:pict/></mc:Choice>` +
      `<mc:Fallback><w:drawing><wp:inline distT="0"><wp:extent cx="1" cy="1"/>` +
      `<wp:docPr id="2" name="Legacy"/><a:graphic/></wp:inline></w:drawing></mc:Fallback>` +
      `</mc:AlternateContent></w:r></w:p>`;
    expect(() => applyDrawingLayout(fallbackOnly, { vertical: "center" })).toThrow(/mc:Fallback/);
  });
});

// The picture as `addImage` writes it: the wp and a prefixes are declared on the
// drawing's OWN tag, not on a root above it. A block fragment carries nothing
// else, so a container that loses those declarations leaves its a:graphic
// pointing at a prefix nothing declares, and Word refuses the file.
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const PIC_NS = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SELF_DECLARED =
  `<w:p ${W_NS} ${R_NS}><w:r><w:drawing>` +
  `<wp:inline ${WP_NS} ${A_NS} ${PIC_NS} distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="4572000" cy="1600200"/>` +
  `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:docPr id="1" name="Picture 1"/>` +
  `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic>` +
  `</a:graphicData></a:graphic>` +
  `</wp:inline></w:drawing></w:r></w:p>`;

describe("applyDrawingLayout — the namespaces a picture declares on its own tag", () => {
  it("the fixture is clean before any change", () => {
    expect(undeclaredPrefixes(SELF_DECLARED)).toEqual([]);
  });

  it("floating it keeps every prefix it uses declared", () => {
    const r = applyDrawingLayout(SELF_DECLARED, { horizontal: "center", vertical: "center" });
    expect(r.changed).toBe(true);
    expect(undeclaredPrefixes(r.xml)).toEqual([]);
  });

  it("carries the declarations onto the wp:anchor that replaces the wp:inline", () => {
    const r = applyDrawingLayout(SELF_DECLARED, { vertical: "center" });
    const open = /<wp:anchor\b[^>]*>/.exec(r.xml)![0];
    expect(open).toContain(A_NS);
    expect(open).toContain(WP_NS);
    expect(open).toContain(PIC_NS);
    // Each declaration appears once: a repeated attribute is not well-formed.
    expect(open.match(/xmlns:a=/g)).toHaveLength(1);
  });

  it("returning it to the flow keeps every prefix declared, and nothing is declared twice", () => {
    const floated = applyDrawingLayout(SELF_DECLARED, { horizontal: "center", vertical: "center" }).xml;
    const back = applyDrawingLayout(floated, { float: false });
    expect(back.changed).toBe(true);
    expect(undeclaredPrefixes(back.xml)).toEqual([]);
    const open = /<wp:inline\b[^>]*>/.exec(back.xml)![0];
    expect(open.match(/xmlns:a=/g)).toHaveLength(1);
  });

  it("moving an already-floating picture does not repeat its declarations", () => {
    const floated = applyDrawingLayout(SELF_DECLARED, { vertical: "center" }).xml;
    const moved = applyDrawingLayout(floated, { vertical: "bottom" });
    expect(undeclaredPrefixes(moved.xml)).toEqual([]);
    expect(/<wp:anchor\b[^>]*>/.exec(moved.xml)![0].match(/xmlns:a=/g)).toHaveLength(1);
  });
});
