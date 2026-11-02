import { describe, test, expect } from "vitest";
import { Mdocxengine, inspectDocx, type DocxZip } from "../index";
import { fixture, P, png, relsPart, type FixtureOptions } from "./docx-fixture";

/**
 * Whatever a copied block points at through a relationship has to reach the same thing
 * afterwards. Pictures, links, charts and OLE objects have their own paths; everything else
 * used to keep the source's rId, which here names something else or nothing.
 */

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="${type.startsWith("http") ? type : `${R}/${type}`}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`;
const FILLER = REL("rId1", "image", "media/filler.png") + REL("rId2", "image", "media/filler.png");

const merged = async (source: FixtureOptions) => {
  const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("target"), docRels: FILLER, parts: { "word/media/filler.png": png(9) } }));
  await t.merge.appendDocument(fixture(source));
  const rels = t.zip.readAsText("word/_rels/document.xml.rels")!;
  const relOf = (id: string) => {
    const tag = new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*`).exec(rels)?.[0] ?? "";
    return { type: /Type="([^"]+)"/.exec(tag)?.[1], target: /Target="([^"]+)"/.exec(tag)?.[1], external: /TargetMode="External"/.test(tag) };
  };
  return { t, doc: t.zip.readAsText("word/document.xml")!, rels, relOf, fatals: inspectDocx(t.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal") };
};

const picture = (attrs: string, docPrInner = "") =>
  `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="p">${docPrInner}</wp:docPr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip ${attrs}/><a:stretch/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

describe("merge: every relationship a copied block uses comes along", () => {
  test("a link on a picture", async () => {
    const { doc, relOf, fatals } = await merged({
      body: picture('r:embed="rId5"', '<a:hlinkClick r:id="rId6"/>'),
      docRels: REL("rId5", "image", "media/image1.png") + REL("rId6", "hyperlink", "https://example.org/a?b=1&amp;c=2", true),
      parts: { "word/media/image1.png": png(1) },
    });
    const id = /<a:hlinkClick r:id="([^"]+)"/.exec(doc)![1];
    expect(relOf(id)).toMatchObject({ target: "https://example.org/a?b=1&amp;c=2", external: true });
    expect(fatals).toEqual([]);
  });

  test("an external picture link stays external, even when its file name matches a copied image", async () => {
    const { t, doc, relOf } = await merged({
      body: picture('r:embed="rId5" r:link="rId6"'),
      docRels: REL("rId5", "image", "media/image1.png") + REL("rId6", "image", "file:///C:/pictures/image1.png", true),
      parts: { "word/media/image1.png": png(1) },
    });
    const embed = /r:embed="([^"]+)"/.exec(doc)![1];
    const link = /r:link="([^"]+)"/.exec(doc)![1];
    expect(relOf(link)).toMatchObject({ target: "file:///C:/pictures/image1.png", external: true });
    expect(t.zip.getEntry(`word/${relOf(embed).target}`)!.getData().equals(png(1))).toBe(true);
  });

  test("a picture effect's data keeps its own relationship type", async () => {
    const HD = "http://schemas.microsoft.com/office/2007/relationships/hdphoto";
    const body = picture('r:embed="rId5"').replace(
      "<a:stretch/>",
      `<a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId7"/></a14:imgProps></a:ext></a:extLst><a:stretch/>`,
    );
    const { t, doc, relOf } = await merged({
      body,
      docRels: REL("rId5", "image", "media/image1.png") + REL("rId7", HD, "media/hdphoto1.wdp"),
      parts: { "word/media/image1.png": png(1), "word/media/hdphoto1.wdp": Buffer.from("wdp-bytes") },
      contentTypes: `<Default Extension="wdp" ContentType="image/vnd.ms-photo"/>`,
    });
    const id = /<a14:imgLayer r:embed="([^"]+)"/.exec(doc)![1];
    expect(relOf(id).type).toBe(HD);
    expect(t.zip.getEntry(`word/${relOf(id).target}`)!.getData().toString()).toBe("wdp-bytes");
    expect(t.zip.readAsText("[Content_Types].xml")).toContain('Extension="wdp"');
  });

  test("a section break's printer settings are dropped: a deposit has no use for them", async () => {
    const { t, doc, fatals } = await merged({
      body: `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:printerSettings r:id="rId9"/></w:sectPr></w:pPr></w:p>`,
      docRels: REL("rId9", "printerSettings", "printerSettings/printerSettings1.bin"),
      parts: { "word/printerSettings/printerSettings1.bin": Buffer.from([1, 2, 3]) },
    });
    const copied = doc.slice(doc.indexOf("target"));
    expect(copied).not.toContain("printerSettings");
    expect(copied).toContain('<w:pgSz w:w="11906"');
    expect(t.zip.getEntry("word/printerSettings/printerSettings1.bin")).toBeNull();
    expect(fatals).toEqual([]);
  });

  test("SmartArt: all four parts and the drawing the data part names, under ids of this document", async () => {
    const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    const data = `<dgm:dataModel xmlns:dgm="${DGM}"><dgm:extLst><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId14" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`;
    const body = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="${DGM}"><dgm:relIds r:dm="rId10" r:lo="rId11" r:qs="rId12" r:cs="rId13"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const REL_BASE = `${R}/`;
    const { t, doc, relOf, fatals } = await merged({
      body,
      docRels:
        REL("rId10", `${REL_BASE}diagramData`, "diagrams/data1.xml") + REL("rId11", `${REL_BASE}diagramLayout`, "diagrams/layout1.xml") +
        REL("rId12", `${REL_BASE}diagramQuickStyle`, "diagrams/quickStyle1.xml") + REL("rId13", `${REL_BASE}diagramColors`, "diagrams/colors1.xml") +
        REL("rId14", "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing", "diagrams/drawing1.xml"),
      parts: {
        "word/diagrams/data1.xml": data, "word/diagrams/layout1.xml": "<layout/>", "word/diagrams/quickStyle1.xml": "<qs/>",
        "word/diagrams/colors1.xml": "<colors/>", "word/diagrams/drawing1.xml": "<drawing/>",
      },
      contentTypes:
        `<Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>` +
        `<Override PartName="/word/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>` +
        `<Override PartName="/word/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>` +
        `<Override PartName="/word/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>` +
        `<Override PartName="/word/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>`,
    });
    const ids = /<dgm:relIds r:dm="([^"]+)" r:lo="([^"]+)" r:qs="([^"]+)" r:cs="([^"]+)"/.exec(doc)!.slice(1);
    const dataPath = `word/${relOf(ids[0]).target}`;
    expect(relOf(ids[0]).type).toContain("diagramData");
    expect(relOf(ids[3]).type).toContain("diagramColors");
    const copied = t.zip.readAsText(dataPath)!;
    const drawingId = /relId="([^"]+)"/.exec(copied)![1];
    expect(drawingId).not.toBe("rId14");
    expect(relOf(drawingId).type).toContain("diagramDrawing");
    expect(t.zip.readAsText(`word/${relOf(drawingId).target}`)).toBe("<drawing/>");
    expect(t.zip.readAsText("[Content_Types].xml")).toContain("diagramDrawing+xml");
    expect(fatals).toEqual([]);
  });

  test("an embedded chunk and an ink part: the part and its type come along", async () => {
    const AF = `${R}/aFChunk`;
    const { t, doc, relOf } = await merged({
      body: `<w:altChunk r:id="rId20"/><w:p/>`,
      docRels: REL("rId20", AF, "afchunk.htm"),
      parts: { "word/afchunk.htm": "<html>x</html>" },
      contentTypes: `<Default Extension="htm" ContentType="application/xhtml+xml"/>`,
    });
    const id = /<w:altChunk r:id="([^"]+)"/.exec(doc)![1];
    expect(id).not.toBe("rId20");
    expect(relOf(id).type).toBe(AF);
    expect(t.zip.readAsText(`word/${relOf(id).target}`)).toBe("<html>x</html>");
  });

  test("a part outside word/ keeps a target that resolves", async () => {
    const { t, doc, relOf } = await merged({
      body: `<w:altChunk r:id="rId20"/><w:p/>`,
      docRels: REL("rId20", "aFChunk", "../extras/a.htm"),
      parts: { "extras/a.htm": "<html>y</html>" },
      contentTypes: `<Default Extension="htm" ContentType="application/xhtml+xml"/>`,
    });
    const target = relOf(/<w:altChunk r:id="([^"]+)"/.exec(doc)![1]).target!;
    expect(target.startsWith("../extras/")).toBe(true);
    expect(t.zip.readAsText(`extras/${target.split("/").pop()}`)).toBe("<html>y</html>");
  });

  test("a legacy r:pict on a VML picture is remapped with the rest", async () => {
    const { doc, relOf } = await merged({
      body: `<w:p><w:r><w:pict><v:shape><v:imagedata r:pict="rId5" r:id="rId5"/></v:shape></w:pict></w:r></w:p>`,
      docRels: REL("rId5", "image", "media/image1.png"),
      parts: { "word/media/image1.png": png(1) },
    });
    const m = /r:pict="([^"]+)" r:id="([^"]+)"/.exec(doc)!;
    expect(m[1]).toBe(m[2]);
    expect(relOf(m[1]).target).toMatch(/media\//);
  });

  test("a reference with no relationship in the source stays what it was: nothing to bring", async () => {
    const { doc } = await merged({ body: `<w:p><w:r><w:drawing><a:hlinkClick r:id="rId77"/></w:drawing></w:r></w:p>` });
    expect(doc).toContain('r:id="rId77"');
  });
});
