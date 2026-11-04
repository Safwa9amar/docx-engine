import { describe, test, expect } from "vitest";
import { Mdocxengine, effectiveParagraphLook, type BodyBlock } from "../index";

/**
 * A merge keeps the first part's Normal. A later part whose right-to-left direction, alignment or line
 * spacing came from its own Normal would lose them, every paragraph of it, with every word in place.
 * All text is invented.
 */

const stylesOf = (ppr: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}</w:style></w:styles>`;

async function partOf(normalPPr: string, blocks: string[]): Promise<Buffer> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  e.zip.addFile("word/styles.xml", Buffer.from(stylesOf(normalPPr), "utf8"));
  await e.document.saveBlocks(blocks.map((xml): BodyBlock => ({ kind: xml.startsWith("<w:tbl") ? "table" : "paragraph", tag: xml.startsWith("<w:tbl") ? "w:tbl" : "w:p", xml })));
  return e.zip.toBuffer();
}
const para = (t: string, ppr = "") => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}<w:r><w:t>${t}</w:t></w:r></w:p>`;

async function merged(first: Buffer, later: Buffer) {
  const e = await Mdocxengine.loadFromBuffer(first);
  await e.merge.appendDocument(later, {});
  return (await e.document.getBlocks()).map((b) => b.xml);
}

describe("a later part keeps the look its own styles gave its paragraphs", () => {
  test("right to left from its own Normal: each paragraph says so, a table's and a text box's too", async () => {
    const a = await partOf("", [para("un")]);
    const b = await partOf("<w:bidi/>", [para("deux"), `<w:tbl><w:tr><w:tc>${para("cellule")}</w:tc></w:tr></w:tbl>`]);
    const out = await merged(a, b);
    const later = out.filter((x) => /deux|cellule/.test(x));
    expect(later).toHaveLength(2);
    for (const x of later) expect(x).toMatch(/<w:bidi\/>/);
    expect(out.find((x) => x.includes(">un<"))).not.toContain("<w:bidi");
  });

  test("a paragraph that says its own direction keeps it", async () => {
    const out = await merged(await partOf("", [para("un")]), await partOf("<w:bidi/>", [para("latin", '<w:bidi w:val="0"/>')]));
    expect(out.find((x) => x.includes("latin"))).toContain('<w:bidi w:val="0"/>');
    expect(out.find((x) => x.includes("latin"))!.match(/<w:bidi\b/g)).toHaveLength(1);
  });

  test("the other way: the first part is right to left, the later one is not", async () => {
    const out = await merged(await partOf("<w:bidi/>", [para("un")]), await partOf("", [para("deux")]));
    expect(out.find((x) => x.includes("deux"))).toContain('<w:bidi w:val="0"/>');
  });

  test("line spacing and alignment too, attribute by attribute, in schema order", async () => {
    const out = await merged(
      await partOf('<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>', [para("un")]),
      await partOf('<w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/>', [para("deux"), para("trois", '<w:spacing w:line="240" w:lineRule="auto"/>')]),
    );
    const deux = out.find((x) => x.includes("deux"))!;
    expect(deux).toMatch(/<w:spacing\b[^>]*w:line="360"/);
    expect(deux).toMatch(/<w:spacing\b[^>]*w:after="0"/);
    expect(deux).toMatch(/<w:spacing\b[^>]*\/><w:jc w:val="both"\/>/);
    const trois = out.find((x) => x.includes("trois"))!;
    expect(trois).toMatch(/w:line="240"/);
    expect(trois).toMatch(/w:after="0"/);
  });

  test("the same look on both sides adds nothing", async () => {
    const out = await merged(await partOf("<w:bidi/>", [para("un")]), await partOf("<w:bidi/>", [para("deux")]));
    expect(out.find((x) => x.includes("deux"))).toBe(para("deux"));
  });

  test("it reads through basedOn and the document defaults", () => {
    const styles = `<w:styles><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160"/><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="N"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="H"><w:name w:val="h"/><w:basedOn w:val="N"/><w:pPr><w:bidi w:val="0"/><w:jc w:val="center"/></w:pPr></w:style></w:styles>`;
    expect(effectiveParagraphLook(styles, null)).toMatchObject({ bidi: true, after: "160", line: "240" });
    expect(effectiveParagraphLook(styles, "H")).toMatchObject({ bidi: false, jc: "center", after: "160" });
  });
});
