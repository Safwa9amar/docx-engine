import { describe, test, expect } from "vitest";
import { Mdocxengine, type BodyBlock } from "../index";
import { fixture, note, notesPart, P, type FixtureOptions } from "./docx-fixture";

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

describe("a renamed bookmark's references follow it everywhere", () => {
  const FIELD = (...pieces: string[]) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    pieces.map((t) => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`).join("") +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const MARK = (t: string) => `<w:p><w:bookmarkStart w:id="1" w:name="_Ref100"/><w:r><w:t>${t}</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
  const NOTE_REF = `<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;

  const mergedWith = async (source: FixtureOptions) => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: MARK("first") }));
    await t.merge.appendDocument(fixture(source));
    return { doc: t.zip.readAsText("word/document.xml")!, notes: t.zip.readAsText("word/footnotes.xml") ?? "" };
  };
  const instr = (xml: string) => [...xml.matchAll(/<w:instrText[^>]*>([^<]*)/g)].map((m) => m[1]).join("|");

  test("into a copied footnote: a REF field and a link to the renamed bookmark", async () => {
    const { doc, notes } = await mergedWith({
      body: MARK("second") + NOTE_REF,
      footnotes: notesPart("footnote", note("footnote", 1, `<w:p>${FIELD(" REF _Ref100 \\h ")}<w:hyperlink w:anchor="_Ref100"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>`)),
    });
    expect([...doc.matchAll(/w:name="([^"]+)"/g)].map((m) => m[1])).toEqual(["_Ref100", "_Ref100_m1"]);
    expect(instr(notes)).toBe(" REF _Ref100_m1 \\h ");
    expect(notes).toContain('w:anchor="_Ref100_m1"');
  });

  test("into an endnote and a fldSimple, with the quotes an attribute must escape", async () => {
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: MARK("first") }));
    await t.merge.appendDocument(
      fixture({
        body: MARK("second") + `<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>`,
        endnotes: notesPart("endnote", note("endnote", 1, `<w:p><w:fldSimple w:instr=" HYPERLINK \\l &quot;_Ref100&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" PAGEREF &quot;_Ref100&quot; "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>`)),
      }),
    );
    const notes = t.zip.readAsText("word/endnotes.xml")!;
    expect(notes).toContain("HYPERLINK \\l &quot;_Ref100_m1&quot;");
    expect(notes).toContain("PAGEREF &quot;_Ref100_m1&quot;");
  });

  test("a field code split across runs, in the body and in a note", async () => {
    const split = FIELD(" REF ", "_Ref", "100 \\h ");
    const { doc, notes } = await mergedWith({
      body: MARK("second") + `<w:p>${split}</w:p>` + NOTE_REF,
      footnotes: notesPart("footnote", note("footnote", 1, `<w:p>${split}</w:p>`)),
    });
    const copied = doc.slice(doc.indexOf("second"));
    expect(instr(copied).replace(/\|/g, "")).toBe(" REF _Ref100_m1 \\h ");
    expect(instr(notes).replace(/\|/g, "")).toBe(" REF _Ref100_m1 \\h ");
    // Each run stays: only the words moved.
    expect(copied.match(/<w:instrText/g)).toHaveLength(3);
  });

  test("a nested field's code is its own, and an unrelated field is left alone", async () => {
    const nested =
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r><w:r><w:instrText xml:space="preserve">_Ref100 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> = 1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const { doc } = await mergedWith({ body: MARK("second") + `<w:p>${nested}</w:p><w:p>${FIELD(" SEQ Figure \\* ARABIC ")}</w:p>` });
    const copied = doc.slice(doc.indexOf("second"));
    expect(instr(copied).split("|").join("")).toContain(" REF _Ref100_m1 ");
    expect(instr(copied)).toContain(" IF ");
    expect(instr(copied)).toContain(" SEQ Figure \\* ARABIC ");
  });

  test("a bookmark that lives inside a note is renamed too, when its name is taken", async () => {
    const { notes } = await mergedWith({
      body: NOTE_REF + `<w:p>${FIELD(" NOTEREF _Ref100 \\h ")}</w:p>`,
      footnotes: notesPart("footnote", note("footnote", 1, `<w:p><w:bookmarkStart w:id="5" w:name="_Ref100"/><w:r><w:t>n</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>`)),
    });
    expect(notes).toContain('w:name="_Ref100_m1"');
  });
});

describe("Word reads bookmark names without regard to case", () => {
  const code = (xml: string) => all(xml, /<w:instrText[^>]*>([^<]*)<\/w:instrText>/g).map((c) => c.trim());
  const names = (xml: string) => all(xml, /<w:bookmarkStart\b[^>]*\bw:name="([^"]*)"/g);
  const merged = async (first: string[], second: string[]) => {
    const t = await doc(first);
    await t.merge.appendDocument((await doc(second)).zip.toBuffer(), {});
    return body(t);
  };

  test("names differing only by case collide, and the later one's references follow", async () => {
    const xml = await merged([bm(1, "Resultat", "a"), field("REF Resultat \\h")], [bm(1, "RESULTAT", "b"), field("REF RESULTAT \\h")]);
    expect(names(xml)).toEqual(["Resultat", "RESULTAT_m1"]);
    expect(code(xml)).toEqual(["REF Resultat \\h", "REF RESULTAT_m1 \\h"]);
  });

  test("a reference written in another case than its bookmark follows it", async () => {
    const xml = await merged([bm(1, "res1", "a")], [bm(1, "res1", "b"), field("REF RES1 \\h"), `<w:p><w:hyperlink w:anchor="RES1"><w:r><w:t>lien</w:t></w:r></w:hyperlink></w:p>`]);
    expect(code(xml)).toEqual(["REF res1_m1 \\h"]);
    expect(xml).toContain('w:anchor="res1_m1"');
  });

  test("a free name is looked for without regard to case too", async () => {
    const xml = await merged([bm(1, "res", "a"), bm(2, "RES_M1", "a2")], [bm(1, "res", "b")]);
    expect(names(xml)).toEqual(["res", "RES_M1", "res_m2"]);
  });

  test("a part's own two spellings of one name stay one name", async () => {
    const xml = await merged([bm(1, "res", "a")], [bm(1, "res", "b"), bm(2, "RES", "c"), field("REF RES")]);
    expect(names(xml)).toEqual(["res", "res_m1", "res_m1"]);
    expect(code(xml)).toEqual(["REF res_m1"]);
  });

  test("a new name stays within Word's forty characters", async () => {
    const long = "Chapitre_resultats_de_terrain_annee_2024";
    expect(long).toHaveLength(40);
    const xml = await merged([bm(1, long, "a")], [bm(1, long, "b"), field(`REF ${long} \\h`)]);
    const [, renamed] = names(xml);
    expect(renamed.length).toBeLessThanOrEqual(40);
    expect(renamed.toLowerCase()).not.toBe(long.toLowerCase());
    expect(code(xml)).toEqual([`REF ${renamed} \\h`]);
  });

  test("a bookmark in this document's notes takes its name too", async () => {
    const t = await Mdocxengine.loadFromBuffer(
      fixture({ body: P("Un") + `<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`, footnotes: notesPart("footnote", note("footnote", 1, bm(7, "res", "note"))) }),
    );
    await t.merge.appendDocument(fixture({ body: bm(1, "RES", "b") }), {});
    expect(names(await body(t))).toEqual(["RES_m1"]);
    const ids = all(await body(t), /<w:bookmarkStart\b[^>]*\bw:id="(\d+)"/g);
    expect(Number(ids[0])).toBeGreaterThan(7);
  });
});

describe("a field names a bookmark in more ways than REF", () => {
  const code = (xml: string) => all(xml, /<w:instrText[^>]*>([^<]*)<\/w:instrText>/g).map((c) => c.trim());
  const after = async (...codes: string[]) => {
    const t = await doc([bm(1, "res1", "a")]);
    await t.merge.appendDocument((await doc([bm(1, "res1", "b"), ...codes.map(field)])).zip.toBuffer(), {});
    return code(await body(t));
  };

  test.each([
    ["ref res1 \\h", "ref res1_m1 \\h"],
    ["Ref res1", "Ref res1_m1"],
    ["pageref res1", "pageref res1_m1"],
    ["noteref res1 \\f", "noteref res1_m1 \\f"],
    ['REF "res1"', 'REF "res1_m1"'],
    ["REF &quot;res1&quot; \\h", "REF &quot;res1_m1&quot; \\h"],
    ["res1", "res1_m1"],
    ["RES1", "res1_m1"],
    ["res1 \\* Upper", "res1_m1 \\* Upper"],
    ['HYPERLINK \\l "res1"', 'HYPERLINK \\l "res1_m1"'],
    ["hyperlink \\l res1", "hyperlink \\l res1_m1"],
    ['HYPERLINK "https://exemple.invalid/" \\l "res1"', 'HYPERLINK "https://exemple.invalid/" \\l "res1_m1"'],
    ["TOC \\b res1", "TOC \\b res1_m1"],
    ["toc \\o 1-3 \\b res1 \\h", "toc \\o 1-3 \\b res1_m1 \\h"],
    ["INDEX \\b res1", "INDEX \\b res1_m1"],
    ["SEQ Figure res1", "SEQ Figure res1_m1"],
    ["=res1*2", "=res1_m1*2"],
    ["= SUM(res1, 3)", "= SUM(res1_m1, 3)"],
    ['IF res1 > 0 "res1" "non"', 'IF res1_m1 > 0 "res1" "non"'],
  ])("%s", async (from, to) => {
    expect(await after(from)).toEqual([to]);
  });

  test.each([
    ["PAGE"],
    ['REF other \\h'],
    ['HYPERLINK "res1"'],
    ["TOC \\o 1-3"],
    ["SEQ res1"],
    ["TA \\l res1 \\b"],
    ['=otherres1*2'],
  ])("%s stays as it is", async (c) => {
    expect(await after(c)).toEqual([c]);
  });

  test("a bookmark that shares its name with a field keeps the field a field", async () => {
    const t = await doc([bm(1, "page", "a")]);
    await t.merge.appendDocument((await doc([bm(1, "page", "b"), field("PAGE"), field("REF page")])).zip.toBuffer(), {});
    expect(code(await body(t))).toEqual(["PAGE", "REF page_m1"]);
  });

  test("a code split across runs, in another case, is one code", async () => {
    const split = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ref re</w:instrText></w:r><w:r><w:instrText>s1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const t = await doc([bm(1, "res1", "a")]);
    await t.merge.appendDocument((await doc([bm(1, "res1", "b"), split])).zip.toBuffer(), {});
    expect(code(await body(t)).join("")).toBe("ref res1_m1 \\h");
  });

  test("a simple field too", async () => {
    const t = await doc([bm(1, "res1", "a")]);
    await t.merge.appendDocument((await doc([bm(1, "res1", "b"), `<w:p><w:fldSimple w:instr=" ref res1 "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>`])).zip.toBuffer(), {});
    expect(await body(t)).toContain('w:instr=" ref res1_m1 "');
  });
});
