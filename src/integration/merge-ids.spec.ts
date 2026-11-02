import { describe, test, expect } from "vitest";
import { Mdocxengine, type BodyBlock } from "../index";

/**
 * Ids that must stay unique across a merge: `wp:docPr/@id`, `w:bookmarkStart/@w:id`
 * (paired with its `w:bookmarkEnd`), and bookmark NAMES (with the references to a
 * renamed name rewritten in the copied blocks only).
 */

const drawing = (id: number) =>
  `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/><wp:docPr id="${id}" name="Picture ${id}"/></wp:inline></w:drawing></w:r></w:p>`;
const bm = (id: number, name: string, text: string) =>
  `<w:p><w:bookmarkStart w:id="${id}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`;
const field = (instr: string) =>
  `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

async function doc(xmls: string[]): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  await e.document.saveBlocks(xmls.map((xml): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml })));
  return e;
}
const body = async (e: Mdocxengine) => (await e.document.getBlocks()).map((b) => b.xml).join("");
const all = (xml: string, re: RegExp) => [...xml.matchAll(re)].map((m) => m[1]);

describe("merge keeps ids unique", () => {
  test("docPr ids are renumbered above the target's max and unique", async () => {
    const t = await doc([drawing(5), drawing(9)]);
    const s = await doc([drawing(1), drawing(1), drawing(2)]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const ids = all(await body(t), /<wp:docPr\b[^>]*\bid="(\d+)"/g);
    expect(ids).toHaveLength(8);
    expect(new Set(ids).size).toBe(8);
    expect(ids.slice(0, 2)).toEqual(["5", "9"]);
    expect(ids.slice(2).every((i) => Number(i) > 9)).toBe(true);
  });

  test("bookmark ids are renumbered above the target's max, each pair consistent", async () => {
    const t = await doc([bm(3, "tA", "t")]);
    const s = await doc([bm(1, "sA", "a"), bm(2, "sB", "b")]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const xml = await body(t);
    const starts = all(xml, /<w:bookmarkStart\b[^>]*\bw:id="(\d+)"/g);
    const ends = all(xml, /<w:bookmarkEnd\b[^>]*\bw:id="(\d+)"/g);
    expect(starts).toEqual(["3", "4", "5"]);
    expect(ends).toEqual(starts);
  });

  test("a colliding bookmark name is renamed and its references follow, in copied blocks only", async () => {
    const t = await doc([
      bm(1, "_Toc1", "target heading"),
      field("PAGEREF _Toc1 \\h"),
      `<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>t-link</w:t></w:r></w:hyperlink></w:p>`,
    ]);
    const s = await doc([
      bm(1, "_Toc1", "source heading"),
      bm(2, "_Toc1_m1", "source clash"),
      field("PAGEREF _Toc1 \\h"),
      field("REF _Toc1 \\r \\h"),
      field('HYPERLINK \\l "_Toc1"'),
      field("NOTEREF _Toc12 \\h"),
      `<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>s-link</w:t></w:r></w:hyperlink></w:p>`,
      `<w:p><w:fldSimple w:instr=" PAGEREF _Toc1 \\h "><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>`,
    ]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const xml = await body(t);
    const names = all(xml, /<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g);
    expect(names).toEqual(["_Toc1", "_Toc1_m2", "_Toc1_m1"]); // _m1 is the source's own name
    expect(new Set(names).size).toBe(3);
    // target's own references untouched
    expect(xml).toContain("PAGEREF _Toc1 \\h </w:instrText>");
    expect(xml).toContain('<w:hyperlink w:anchor="_Toc1"><w:r><w:t>t-link');
    // source references follow the rename
    expect(xml).toContain("PAGEREF _Toc1_m2 \\h");
    expect(xml).toContain("REF _Toc1_m2 \\r \\h");
    expect(xml).toContain('HYPERLINK \\l "_Toc1_m2"');
    expect(xml).toContain('<w:hyperlink w:anchor="_Toc1_m2"><w:r><w:t>s-link');
    expect(xml).toContain('w:instr=" PAGEREF _Toc1_m2 \\h "');
    // a longer name sharing the prefix is not touched
    expect(xml).toContain("NOTEREF _Toc12 \\h");
  });

  test("quoted and regex-special names are handled", async () => {
    const t = await doc([bm(1, "a.b", "t")]);
    const s = await doc([bm(1, "a.b", "s"), bm(2, "aXb", "s2"), field('REF "a.b" \\h'), field("REF aXb \\h")]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const xml = await body(t);
    expect(xml).toContain('REF "a.b_m1" \\h');
    expect(xml).toContain("REF aXb \\h");
  });
});
