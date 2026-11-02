import { describe, test, expect } from "vitest";
import { Mdocxengine, inspectDocx, type BodyBlock, type DocxZip } from "../index";
import { fixture, note, notesPart, numberingPart, P, png, relsPart, withNormal } from "./docx-fixture";

/**
 * A merge can be the first thing to give a document a numbering part: part 1 has no
 * list, a later part does. The definitions were copied, but the new part was never
 * registered, and a part with no relationship and no content type is one Word never
 * reads: every list of the merged thesis came out with no bullet or number.
 */

const NUM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUM_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const para = (xml: string): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml });

/** The sample with its numbering part gone, entirely, the way a list-free Word file ships. */
async function noNumbering(body: string[]): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  e.zip.deleteFile("word/numbering.xml");
  const rels = e.zip.readAsText("word/_rels/document.xml.rels")!;
  e.zip.addFile("word/_rels/document.xml.rels", Buffer.from(rels.replace(/<Relationship\b[^>]*numbering\.xml[^>]*\/>/, ""), "utf8"));
  const ct = e.zip.readAsText("[Content_Types].xml")!;
  e.zip.addFile("[Content_Types].xml", Buffer.from(ct.replace(/<Override\b[^>]*numbering\.xml[^>]*\/>/, ""), "utf8"));
  await e.document.saveBlocks(body.map(para));
  return e;
}

async function withList(): Promise<Buffer> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  const numId = /<w:num\b[^>]*w:numId="(\d+)"/.exec(e.zip.readAsText("word/numbering.xml")!)![1];
  await e.document.saveBlocks([para(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`)]);
  return e.zip.toBuffer();
}

describe("merge: the numbering part a merge creates is one Word can see", () => {
  test("it gets its relationship and its content type, and the document points into it", async () => {
    const t = await noNumbering(["<w:p><w:r><w:t>no list here</w:t></w:r></w:p>"]);
    expect(t.zip.getEntry("word/numbering.xml")).toBeNull();
    await t.merge.appendDocument(await withList());

    const part = t.zip.readAsText("word/numbering.xml")!;
    const rels = t.zip.readAsText("word/_rels/document.xml.rels")!;
    const ct = t.zip.readAsText("[Content_Types].xml")!;
    expect(rels).toContain(NUM_REL);
    expect(new RegExp(`<Relationship\\b[^>]*Type="${NUM_REL}"[^>]*`).exec(rels)![0]).toContain('Target="numbering.xml"');
    expect(new RegExp(`<Override\\b[^>]*PartName="/word/numbering.xml"[^>]*`).exec(ct)![0]).toContain(NUM_CT);
    const body = (await t.document.getBlocks()).map((b) => b.xml).join("");
    expect(part).toContain(`w:numId="${/<w:numId w:val="(\d+)"/.exec(body)![1]}"`);
    expect(inspectDocx(t.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal")).toEqual([]);
  });

  test("merging twice registers it once", async () => {
    const t = await noNumbering(["<w:p><w:r><w:t>no list here</w:t></w:r></w:p>"]);
    await t.merge.appendDocument(await withList());
    await t.merge.appendDocument(await withList());
    expect(t.zip.readAsText("word/_rels/document.xml.rels")!.match(/relationships\/numbering"/g)).toHaveLength(1);
    expect(t.zip.readAsText("[Content_Types].xml")!.match(/\/word\/numbering\.xml/g)).toHaveLength(1);
  });

  test("a document that already has the part is not registered again", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    const rels = t.zip.readAsText("word/_rels/document.xml.rels")!;
    const ct = t.zip.readAsText("[Content_Types].xml")!;
    await t.merge.appendDocument(await withList());
    expect(t.zip.readAsText("word/_rels/document.xml.rels")).toBe(rels);
    expect(t.zip.readAsText("[Content_Types].xml")).toBe(ct);
  });
});

/**
 * A level can use a picture as its bullet: `w:lvlPicBulletId` names a `w:numPicBullet`,
 * whose image is reached through the numbering part's own relationships. Copying the level
 * alone leaves an id naming nothing, and Word refuses the whole file.
 */
describe("merge: a picture bullet travels with its level", () => {
  const V_PICT = (rId: string) =>
    `<w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:11pt;height:11pt" o:bullet="t"><v:imagedata r:id="${rId}" o:title="mso1"/></v:shape></w:pict>`;
  const BULLET = (id: number, rId: string) => `<w:numPicBullet w:numPicBulletId="${id}">${V_PICT(rId)}</w:numPicBullet>`;
  const LEVEL = (pic: number | null) =>
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uF0B7"/>${pic === null ? "" : `<w:lvlPicBulletId w:val="${pic}"/>`}<w:lvlJc w:val="left"/></w:lvl>`;
  const LIST = (abs: number, numId: number, pic: number | null) => `<w:abstractNum w:abstractNumId="${abs}"><w:multiLevelType w:val="hybridMultilevel"/>${LEVEL(pic)}</w:abstractNum><w:num w:numId="${numId}"><w:abstractNumId w:val="${abs}"/></w:num>`;
  const ITEM = (numId: number, t = "item") => P(t, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
  const WITH_BULLET = (seed = 0, numId = 1) =>
    fixture({
      body: ITEM(numId),
      numbering: numberingPart(BULLET(0, "rId1") + LIST(0, numId, 0)),
      parts: {
        "word/media/image1.png": png(seed),
        "word/_rels/numbering.xml.rels": relsPart(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`),
      },
    });

  /** The image a bullet id reaches in the merged package. */
  const bulletImage = (t: Mdocxengine, picId: string) => {
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    const bullet = new RegExp(`<w:numPicBullet\\b[^>]*w:numPicBulletId="${picId}"[\\s\\S]*?</w:numPicBullet>`).exec(numbering)![0];
    const rId = /<v:imagedata\b[^>]*r:id="([^"]+)"/.exec(bullet)![1];
    const target = new RegExp(`<Relationship\\b[^>]*Id="${rId}"[^>]*`).exec(t.zip.readAsText("word/_rels/numbering.xml.rels")!)![0];
    return t.zip.getEntry(`word/${/Target="([^"]+)"/.exec(target)![1]}`)!.getData();
  };
  const fatal = (t: Mdocxengine) => inspectDocx(t.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal");

  test("a target with no numbering part: bullet, relationship and image all land, in schema order", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("no list here") }));
    await t.merge.appendDocument(WITH_BULLET(1));
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    expect(numbering.indexOf("<w:numPicBullet")).toBeGreaterThan(-1);
    expect(numbering.indexOf("<w:numPicBullet")).toBeLessThan(numbering.indexOf("<w:abstractNum "));
    expect(numbering.indexOf("<w:abstractNum ")).toBeLessThan(numbering.indexOf("<w:num "));
    const id = /<w:lvlPicBulletId w:val="(\d+)"/.exec(numbering)![1];
    expect(bulletImage(t, id).equals(png(1))).toBe(true);
    expect(t.zip.readAsText("[Content_Types].xml")).toMatch(/Extension="png"/);
    expect(fatal(t)).toEqual([]);
  });

  test("a target whose numbering has no bullet keeps its own list and order", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: ITEM(1, "mine"), numbering: numberingPart(LIST(0, 1, null)) }));
    await t.merge.appendDocument(WITH_BULLET(1));
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    expect(numbering.match(/<w:abstractNum\b/g)).toHaveLength(2);
    expect(numbering.indexOf("<w:numPicBullet")).toBeLessThan(numbering.indexOf("<w:abstractNum"));
    expect(numbering.lastIndexOf("<w:abstractNum ")).toBeLessThan(numbering.indexOf("<w:num "));
  });

  test("a target with its own bullet 0 keeps it, and each list reaches its own picture", async () => {
    const t = await Mdocxengine.loadFromBuffer(WITH_BULLET(1));
    await t.merge.appendDocument(WITH_BULLET(2));
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    const ids = [...numbering.matchAll(/<w:numPicBullet w:numPicBulletId="(\d+)"/g)].map((m) => m[1]);
    expect(ids).toEqual(["0", "1"]);
    const used = [...numbering.matchAll(/<w:lvlPicBulletId w:val="(\d+)"/g)].map((m) => m[1]);
    expect(used).toEqual(["0", "1"]);
    expect(bulletImage(t, "0").equals(png(1))).toBe(true);
    expect(bulletImage(t, "1").equals(png(2))).toBe(true);
  });

  test("a bullet only a list nobody copied uses stays behind", async () => {
    const src = fixture({ body: ITEM(2), numbering: numberingPart(BULLET(0, "rId1") + LIST(0, 1, 0) + LIST(1, 2, null)), parts: { "word/media/image1.png": png(1), "word/_rels/numbering.xml.rels": relsPart(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`) } });
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("x") }));
    await t.merge.appendDocument(src);
    expect(t.zip.readAsText("word/numbering.xml")).not.toContain("numPicBullet");
    expect(t.zip.getEntry("word/_rels/numbering.xml.rels")).toBeNull();
  });

  test("a level naming a bullet the source does not have loses the reference, not the level", async () => {
    const src = fixture({ body: ITEM(1), numbering: numberingPart(LIST(0, 1, 7)) });
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("x") }));
    await t.merge.appendDocument(src);
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    expect(numbering).not.toContain("lvlPicBulletId");
    expect(numbering).toContain('w:val="bullet"');
  });
});

describe("addNumberingDefinition on a document with no numbering part", () => {
  test("the new part declares its namespaces, or Word refuses it now that it is registered", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("x") }));
    await t.numbering.addNumberingDefinition("1", "0");
    const numbering = t.zip.readAsText("word/numbering.xml")!;
    expect(/<w:numbering\b[^>]*>/.exec(numbering)![0]).toContain(`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`);
    expect(t.zip.readAsText("word/_rels/document.xml.rels")).toContain("numbering.xml");
  });
});

/**
 * A numId with no definition is no absence. Measured in real Word: it prints a default "1.",
 * each missing id counting on its own, with or without a numbering part, from a paragraph or
 * from a style. So one has to stay missing on both sides of a merge: the target's must not be
 * filled by a copied list, and the source's must not fall into a list of that id here.
 */
describe("merge: a numId with no definition stays one", () => {
  const LIST = `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>`;
  const ITEM = (numId: number, t: string) => P(t, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
  const numIdOf = (doc: string, text: string) => new RegExp(`<w:numId w:val="(\\d+)"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">${text}`).exec(doc)?.[1];
  const defined = (numbering: string) => [...numbering.matchAll(/<w:num\b[^>]*\bw:numId="(\d+)"/g)].map((m) => m[1]);

  test("a target's: a copied list takes an id above it", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: ITEM(3, "missing here") + ITEM(5, "this one too") }));
    await t.merge.appendDocument(fixture({ body: ITEM(3, "item"), numbering: numberingPart(LIST) }));
    const doc = t.zip.readAsText("word/document.xml")!;
    expect(numIdOf(doc, "missing here")).toBe("3");
    expect(numIdOf(doc, "this one too")).toBe("5");
    expect(Number(numIdOf(doc, "item"))).toBeGreaterThan(5);
    expect(defined(t.zip.readAsText("word/numbering.xml")!)).toEqual([numIdOf(doc, "item")]);
  });

  test("the same for an id only the target's styles or notes name", async () => {
    const t = await Mdocxengine.loadFromBuffer(
      fixture({
        body: P("x"),
        styles: withNormal(`<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>`),
        footnotes: notesPart("footnote", note("footnote", 1, `<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:p>`)),
      }),
    );
    await t.merge.appendDocument(fixture({ body: ITEM(3, "item"), numbering: numberingPart(LIST) }));
    expect(Number(defined(t.zip.readAsText("word/numbering.xml")!)[0])).toBeGreaterThan(7);
  });

  const merged = async (source: Parameters<typeof fixture>[0]) => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: ITEM(3, "mine"), numbering: numberingPart(LIST) }));
    await t.merge.appendDocument(fixture(source));
    return { doc: t.zip.readAsText("word/document.xml")!, numbering: t.zip.readAsText("word/numbering.xml")!, styles: t.zip.readAsText("word/styles.xml")! };
  };

  test("a source's: it does not fall into the target's list of that id, and two ids stay two", async () => {
    const { doc, numbering } = await merged({ body: ITEM(3, "theirs") + ITEM(3, "theirs too") + ITEM(4, "another") });
    expect(numIdOf(doc, "mine")).toBe("3");
    const id = numIdOf(doc, "theirs")!;
    expect(["0", "3"]).not.toContain(id);
    expect(numIdOf(doc, "theirs too")).toBe(id);
    expect(numIdOf(doc, "another")).not.toBe(id);
    expect(defined(numbering)).toEqual(["3"]);
  });
  test("a num whose abstractNum is not there is the same, and no such num is copied", async () => {
    const { doc, numbering } = await merged({ body: ITEM(1, "theirs"), numbering: numberingPart(`<w:num w:numId="1"><w:abstractNumId w:val="44"/></w:num>`) });
    expect(["0", "1", "3"]).not.toContain(numIdOf(doc, "theirs"));
    expect(defined(numbering)).toEqual(["3"]);
  });
  test("a style naming no list: the same", async () => {
    const { styles } = await merged({ body: P("theirs", '<w:pStyle w:val="S"/>'), styles: withNormal(`<w:style w:type="paragraph" w:customStyle="1" w:styleId="S"><w:name w:val="S"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style>`) });
    const id = /w:styleId="S"[\s\S]*?<w:numId w:val="(\d+)"/.exec(styles)![1];
    expect(["0", "3"]).not.toContain(id);
  });
  test("a target with no numbering part gets none for it", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("mine") }));
    await t.merge.appendDocument(fixture({ body: ITEM(3, "theirs") }));
    expect(t.zip.getEntry("word/numbering.xml")).toBeNull();
  });
  test("a built-in style's missing list against the target's unnumbered one: the paragraph keeps its default label", async () => {
    const heading = (extra: string) => `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>${extra}</w:style>`;
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("mine"), styles: withNormal(heading("")) }));
    await t.merge.appendDocument(fixture({ body: P("theirs", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading('<w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr>')) }));
    const id = /<w:numId w:val="(\d+)"\/><\/w:numPr><\/w:pPr><w:r><w:t xml:space="preserve">theirs/.exec(t.zip.readAsText("word/document.xml")!)?.[1];
    expect(id).toBeDefined();
    expect(id).not.toBe("0");
  });
  test("the target's style naming a missing list, the part's none: the paragraph turns it off", async () => {
    const heading = (extra: string) => `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>${extra}</w:style>`;
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("mine"), styles: withNormal(heading('<w:pPr><w:numPr><w:numId w:val="9"/></w:numPr></w:pPr>')) }));
    await t.merge.appendDocument(fixture({ body: P("theirs", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading("")) }));
    expect(t.zip.readAsText("word/document.xml")).toMatch(/<w:numId w:val="0"\/><\/w:numPr><\/w:pPr><w:r><w:t xml:space="preserve">theirs/);
  });
});
