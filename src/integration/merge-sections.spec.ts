import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, checkDocxBuffer, splitDocument } from "../index";
import { fixture, P, PAGE, SECT } from "./docx-fixture";

/**
 * A section's properties govern the text before them, so a later part's first section break also
 * governs part 1's closing text, and its last section, which has no break, has nothing left to
 * govern. Each part's text keeps its own page. When nothing about the pages differs, nothing is added.
 */

const LAND = `<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>`;
const LANDSCAPE = `<w:sectPr>${LAND}</w:sectPr>`;
const NARROW = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;

const merge = async (...parts: Buffer[]) => {
  const e = await Mdocxengine.loadFromBuffer(parts[0]);
  for (const p of parts.slice(1)) await e.merge.appendDocument(p, { keepPages: true });
  return e.zip.toBuffer();
};
const docOf = (d: Buffer) => new AdmZip(d).readAsText("word/document.xml");

/** Each section's page width and the words it governs, in order. */
function sections(d: Buffer): { width: string | undefined; words: string[] }[] {
  const blocks = splitDocument(docOf(d)).blocks;
  const out: { width: string | undefined; words: string[] }[] = [];
  let words: string[] = [];
  for (const b of blocks) {
    if (b.kind === "sectPr") {
      out.push({ width: /<w:pgSz\b[^>]*\bw:w="(\d+)"/.exec(b.xml)?.[1], words });
      continue;
    }
    words.push(...[...b.xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]));
    const inner = /<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/.exec(b.xml)?.[0];
    if (inner) {
      out.push({ width: /<w:pgSz\b[^>]*\bw:w="(\d+)"/.exec(inner)?.[1], words });
      words = [];
    }
  }
  return out;
}
const sectCount = (d: Buffer) => (docOf(d).match(/<w:sectPr\b/g) ?? []).length;

describe("each part's text keeps its page", () => {
  test("a section break in a later part does not reach back into part 1", async () => {
    const out = await merge(fixture({ body: P("Premiere.") }), fixture({ body: P("Large.", LANDSCAPE) + P("Suite.") }));
    expect(sections(out)).toEqual([
      { width: "11906", words: ["Premiere."] },
      { width: "16838", words: ["Large."] },
      { width: "11906", words: ["Suite."] },
    ]);
  });

  test("a later part wholly landscape stays so", async () => {
    const out = await merge(fixture({ body: P("Premiere.") }), fixture({ body: P("Large."), sect: LANDSCAPE }));
    expect(sections(out)).toEqual([
      { width: "11906", words: ["Premiere."] },
      { width: "16838", words: ["Large."] },
    ]);
  });

  test("part 1 landscape stays so, and does not take the later part with it", async () => {
    const out = await merge(fixture({ body: P("Large."), sect: LANDSCAPE }), fixture({ body: P("Suite.") }));
    expect(sections(out)).toEqual([
      { width: "16838", words: ["Large."] },
      { width: "11906", words: ["Suite."] },
    ]);
  });

  test("three parts, each its own", async () => {
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: P("Deux."), sect: LANDSCAPE }), fixture({ body: P("Trois.") }));
    expect(sections(out)).toEqual([
      { width: "11906", words: ["Un."] },
      { width: "16838", words: ["Deux."] },
      { width: "11906", words: ["Trois."] },
    ]);
  });

  test("margins differ: each part keeps its own", async () => {
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: P("Deux."), sect: NARROW }));
    expect(sectCount(out)).toBe(2);
    expect([...docOf(out).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)][0][0]).toContain('w:top="1440"');
    const last = splitDocument(docOf(out)).blocks.filter((b) => b.kind === "sectPr").pop()!.xml;
    expect(last).toContain('w:top="1000"');
  });
});

describe("when nothing about the pages differs, nothing is added", () => {
  test("no section break, the same words, the same last section", async () => {
    const a = fixture({ body: P("Un.") });
    const out = await merge(a, fixture({ body: P("Deux.") }));
    expect(sectCount(out)).toBe(1);
    expect(docOf(out)).toContain(SECT);
  });

  test("a part with no section at all states no page", async () => {
    expect(sectCount(await merge(fixture({ body: P("Un.") }), fixture({ body: P("Deux."), sect: null })))).toBe(1);
  });

  test("a paper code or a grid is no difference", async () => {
    const coded = `<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440" w:header="708" w:footer="708" w:gutter="0"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    expect(sectCount(await merge(fixture({ body: P("Un.") }), fixture({ body: P("Deux."), sect: coded })))).toBe(1);
  });
});

describe("a caller whose template owns the pages", () => {
  test("gets no seam", async () => {
    const e = await Mdocxengine.loadFromBuffer(fixture({ body: P("Un.") }));
    await e.merge.appendDocument(fixture({ body: P("Deux."), sect: LANDSCAPE }), {});
    expect(sectCount(e.zip.toBuffer())).toBe(1);
    expect(docOf(e.zip.toBuffer())).toContain(SECT);
  });
});

describe("the seam itself", () => {
  test("it does not start a new page of its own: the next section runs on", async () => {
    const out = await merge(fixture({ body: P("Premiere.") }), fixture({ body: P("Meme page.", `<w:sectPr>${PAGE}</w:sectPr>`) + P("Suite.") }));
    const xml = docOf(out);
    const second = [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)][1][0];
    expect(second).toContain('<w:type w:val="continuous"/>');
    expect(second.indexOf("<w:type")).toBeLessThan(second.indexOf("<w:pgSz"));
  });

  test("a tracked change on that first break stays whole", async () => {
    const tracked = `<w:sectPr>${PAGE}<w:sectPrChange w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z"><w:sectPr>${LAND}</w:sectPr></w:sectPrChange></w:sectPr>`;
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: P("Deux.", tracked) + P("Trois.") }));
    expect(docOf(out)).toContain(`<w:sectPrChange w:id="5" w:author="A" w:date="2026-01-01T00:00:00Z"><w:sectPr>${LAND}</w:sectPr></w:sectPrChange></w:sectPr></w:pPr>`);
    expect((await checkDocxBuffer(out)).findings.filter((f) => f.severity === "fatal")).toEqual([]);
  });

  test("part 1's last section keeps what it was, and the new last one does not start over", async () => {
    const own = `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:titlePg/></w:sectPr>`;
    const out = await merge(fixture({ body: P("Un."), sect: own }), fixture({ body: P("Deux."), sect: LANDSCAPE }));
    const all = [...docOf(out).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]);
    expect(all[0]).toBe(own);
    expect(all[1]).toContain('w:orient="landscape"');
    expect(all[1]).toContain('<w:type w:val="continuous"/>');
    expect(all[1]).not.toContain("<w:titlePg");
    expect(all[1]).not.toContain("w:start=");
    expect(all[1]).toContain('w:fmt="lowerRoman"');
  });

  test("the break goes on part 1's last paragraph, inside its properties, ahead of a tracked change", async () => {
    const last = `<w:p><w:pPr><w:jc w:val="left"/><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>Un.</w:t></w:r></w:p>`;
    const xml = docOf(await merge(fixture({ body: last }), fixture({ body: P("Deux."), sect: LANDSCAPE })));
    expect(xml.indexOf("<w:sectPr")).toBeLessThan(xml.indexOf("<w:pPrChange"));
    expect(xml.indexOf('<w:jc w:val="left"/>')).toBeLessThan(xml.indexOf("<w:sectPr"));
  });

  test("a part 1 ending on a table or on a break of its own takes a paragraph for it", async () => {
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cellule</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    for (const body of [table, P("Fin.", `<w:sectPr>${PAGE}</w:sectPr>`)]) {
      const out = await merge(fixture({ body }), fixture({ body: P("Deux."), sect: LANDSCAPE }));
      expect(sections(out).at(-1)).toEqual({ width: "16838", words: ["Deux."] });
      expect((await checkDocxBuffer(out)).findings.filter((f) => f.severity === "fatal")).toEqual([]);
    }
  });

  test("a part 1 with nothing in it has no page to keep", async () => {
    const out = await merge(fixture({ body: "" }), fixture({ body: P("Deux."), sect: LANDSCAPE }));
    expect(sections(out)).toEqual([{ width: "16838", words: ["Deux."] }]);
  });

  test("what comes out passes the doctor", async () => {
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: P("Large.", LANDSCAPE) + P("Suite."), sect: NARROW }));
    expect((await checkDocxBuffer(out)).findings.filter((f) => f.severity === "fatal")).toEqual([]);
  });
});

describe("what Word itself keeps at a seam", () => {
  const TBL = (t: string, cols = 2) =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(cols)}</w:tblGrid><w:tr>${Array.from({ length: cols }, (_, i) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}${i}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr></w:tbl>`;
  const rules = async (d: Buffer) => (await checkDocxBuffer(d)).findings.map((f) => f.rule).filter((r) => r.startsWith("story."));

  test("a part ending on a table gets its closing paragraph", async () => {
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: TBL("B") }));
    expect(await rules(out)).toEqual([]);
    expect(docOf(out)).toMatch(/<\/w:tbl><w:p\/>/);
  });

  test("two tables meeting at a seam stay two", async () => {
    const out = await merge(fixture({ body: TBL("A", 2) }), fixture({ body: TBL("B", 3) }));
    expect(await rules(out)).toEqual([]);
    expect(docOf(out)).not.toMatch(/<\/w:tbl>\s*<w:tbl/);
  });

  test("a paragraph is added nowhere else", async () => {
    const out = await merge(fixture({ body: TBL("A") + "<w:p/>" }), fixture({ body: TBL("B") + "<w:p/>" }));
    expect((docOf(out).match(/<w:p\/>/g) ?? []).length).toBe(2);
  });

  test("an Arabic table's spacer reads right to left", async () => {
    const out = await merge(fixture({ body: P("Un.") }), fixture({ body: TBL("B").replace("<w:tblPr>", "<w:tblPr><w:bidiVisual/>") }));
    expect(docOf(out)).toContain("<w:p><w:pPr><w:bidi/></w:pPr></w:p>");
  });
});
