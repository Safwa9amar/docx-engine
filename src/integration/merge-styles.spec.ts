import { describe, test, expect } from "vitest";
import { Mdocxengine, type BodyBlock } from "../index";

/**
 * Styles across a merge: the blocks reference `pStyle`/`rStyle`/`tblStyle` ids
 * whose DEFINITIONS live in the source's styles.xml. Without importing them the
 * references dangle.
 */

const styleXml = (id: string, extra = "", attrs = 'w:type="paragraph"') =>
  `<w:style ${attrs} w:styleId="${id}"><w:name w:val="${id}"/>${extra}</w:style>`;

async function srcDoc(styles: string, blocks: string[]): Promise<Buffer> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  const sx = e.zip.readAsText("word/styles.xml")!;
  e.zip.addFile("word/styles.xml", Buffer.from(sx.replace("</w:styles>", styles + "</w:styles>"), "utf8"));
  await e.document.saveBlocks(blocks.map((xml): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml })));
  return e.zip.toBuffer();
}
const p = (inner: string, ppr = "") => `<w:p>${ppr}<w:r>${inner}<w:t>x</w:t></w:r></w:p>`;
const ids = (xml: string) => [...xml.matchAll(/<w:style\b[^>]*\bw:styleId="([^"]+)"/g)].map((m) => m[1]);

describe("merge imports the styles the target lacks", () => {
  test("a source-only style and its basedOn chain resolve; a colliding id keeps the target's", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    const before = t.zip.readAsText("word/styles.xml")!;
    expect(ids(before)).toContain("Heading1");
    const h1Before = new RegExp(`<w:style\\b[^>]*w:styleId="Heading1"[\\s\\S]*?</w:style>`).exec(before)![0];

    const src = await srcDoc(
      styleXml("MyBase", "<w:rPr><w:i/></w:rPr>") +
        styleXml("MyQuote", '<w:basedOn w:val="MyBase"/><w:next w:val="MyAfter"/><w:link w:val="MyQuoteChar"/>', 'w:type="paragraph" w:default="1"') +
        styleXml("MyAfter") +
        styleXml("MyQuoteChar", '<w:link w:val="MyQuote"/>', 'w:type="character"') +
        styleXml("MyEmph", "<w:rPr><w:b/></w:rPr>", 'w:type="character"') +
        styleXml("MyTable", "<w:tblPr/>", 'w:type="table"') +
        styleXml("Heading1", '<w:rPr><w:color w:val="FF0000"/></w:rPr>'),
      [
        p("", '<w:pPr><w:pStyle w:val="MyQuote"/></w:pPr>'),
        p('<w:rPr><w:rStyle w:val="MyEmph"/></w:rPr>'),
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>h</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="NeverDefined"/></w:pPr><w:r><w:t>z</w:t></w:r></w:p>',
      ],
    );
    // a table-style reference lives in tblPr — use a raw block
    const e2 = await Mdocxengine.loadFromBuffer(src);
    await e2.document.saveBlocks([
      ...(await e2.document.getBlocks()),
      { kind: "table", tag: "w:tbl", xml: '<w:tbl><w:tblPr><w:tblStyle w:val="MyTable"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' },
    ]);

    await t.merge.appendDocument(e2.zip.toBuffer(), { styleMap: {} });
    const after = t.zip.readAsText("word/styles.xml")!;
    const got = ids(after);
    for (const id of ["MyBase", "MyQuote", "MyAfter", "MyQuoteChar", "MyEmph", "MyTable"]) {
      expect(got.filter((x) => x === id), id).toHaveLength(1);
    }
    expect(got).not.toContain("NeverDefined");
    // the colliding Heading1 is byte-identical to the target's own
    expect(got.filter((x) => x === "Heading1")).toHaveLength(1);
    expect(after).toContain(h1Before);
    expect(after).not.toContain('w:val="FF0000"');
    // only the target keeps its default paragraph style
    expect(after).not.toMatch(/w:styleId="MyQuote"[^>]*w:default|w:default="1"[^>]*w:styleId="MyQuote"/);
    // every style reference in the merged body now resolves (except the undefined one)
    const body = (await t.document.getBlocks()).map((b) => b.xml).join("");
    for (const m of body.matchAll(/<w:(?:pStyle|rStyle|tblStyle)\b[^>]*\bw:val="([^"]+)"/g)) {
      if (m[1] !== "NeverDefined") expect(got, m[1]).toContain(m[1]);
    }
  });

  test("a style's numbering is copied and remapped; styleMap'd ids are not imported", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    const s0 = await Mdocxengine.loadFromFile("samples/example.docx");
    const numId = /<w:num\b[^>]*w:numId="(\d+)"/.exec(s0.zip.readAsText("word/numbering.xml")!)![1];
    const numBefore = (t.zip.readAsText("word/numbering.xml")!.match(/<w:num\b/g) ?? []).length;
    const src = await srcDoc(
      styleXml("MyList", `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`) +
        styleXml("Mapped"),
      [
        '<w:p><w:pPr><w:pStyle w:val="MyList"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="Mapped"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
      ],
    );
    await t.merge.appendDocument(src, { styleMap: { Mapped: "Normal" } });
    const styles = t.zip.readAsText("word/styles.xml")!;
    expect(ids(styles)).toContain("MyList");
    expect(ids(styles)).not.toContain("Mapped");
    const newNum = /w:styleId="MyList"[\s\S]*?<w:numId w:val="(\d+)"/.exec(styles)![1];
    const numAfter = t.zip.readAsText("word/numbering.xml")!;
    expect((numAfter.match(/<w:num\b/g) ?? []).length).toBeGreaterThan(numBefore);
    expect(numAfter).toContain(`w:numId="${newNum}"`);
    expect(newNum).not.toBe(numId); // copied with a fresh id, not left pointing at the target's own list
  });

  test("a target whose styles part has no closing tag is left alone", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    t.zip.addFile("word/styles.xml", Buffer.from("<w:styles/>", "utf8"));
    const src = await srcDoc(styleXml("Only"), ['<w:p><w:pPr><w:pStyle w:val="Only"/></w:pPr></w:p>']);
    await t.merge.appendDocument(src, {});
    expect(ids(t.zip.readAsText("word/styles.xml")!)).toEqual([]);
  });
});
