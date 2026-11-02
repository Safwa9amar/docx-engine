import { describe, test, expect } from "vitest";
import { Mdocxengine, inspectDocx, type BodyBlock, type DocxZip } from "../index";

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
