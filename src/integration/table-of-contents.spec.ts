import { describe, test, expect } from "vitest";
import { Mdocxengine, makeParagraphNode, type BodyBlock } from "@/index";
import { paragraphText } from "@/core/files/body/OrderedBody";

/**
 * The Table of Contents — Word's References → Table of Contents.
 *
 * Two things are being held down here. First, Word parity: a real `TOC` field
 * with `_Toc…` anchors and PAGEREF page numbers, so it renumbers itself, and
 * pre-populated entries so it READS correctly before Word ever updates a field.
 * Second, and the reason this manager was rewritten: inserting it must not
 * disturb the body. The old implementation rebuilt document.xml through an XML
 * object model, which regroups the body by tag — every table in the thesis
 * hoisted away from the paragraphs around it.
 */

const BASE = "samples/example.docx";

/** An engine whose body is exactly `blocks`. */
async function docOf(blocks: Array<string | BodyBlock>): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile(BASE);
  await e.document.saveBlocks(blocks.map((b) => (typeof b === "string" ? makeParagraphNode(b) : b)));
  return e;
}

const heading = (text: string, level = 1): BodyBlock => makeParagraphNode(text, `Heading${level}`);

const table = (cell: string): BodyBlock => ({
  kind: "table",
  tag: "w:tbl",
  xml: `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
});

/** Visible text of every body block, in order. */
async function texts(e: Mdocxengine): Promise<string[]> {
  const blocks = await e.document.getBlocks();
  return blocks.map((b) => paragraphText(b.xml).trim());
}

const xmlOf = (e: Mdocxengine): string => e.zip.readAsText("word/document.xml")!;

describe("insertTOC — a title Word cannot read off the page", () => {
  const tc = (code: string) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const holder = (code: string, style = ""): BodyBlock => ({
    kind: "paragraph",
    tag: "w:p",
    xml: `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}${tc(code)}</w:p>`,
  });

  test("a TC field gets its entry, in document order, and the field collects TC fields", async () => {
    const e = await docOf([heading("Introduction"), holder(' TC "Chapitre &amp; un" \\l "1" '), heading("Section", 2)]);
    const r = await e.toc.insertTOC({ title: "", headingDepth: 3 }, 0);
    expect(r.entries).toBe(3);
    const lines = (await texts(e)).filter((t) => t.includes("\t"));
    expect(lines.map((l) => l.split("\t")[0])).toEqual(["Introduction", "Chapitre & un", "Section"]);
    expect(xmlOf(e)).toMatch(/TOC \\o "1-3" \\h \\z \\u \\f /);
  });

  test("a heading holding a TC field is listed once, by what the field says", async () => {
    const e = await docOf([holder(' TC "Titre \\"cit\u00e9\\" ici" \\l "1" ', "Heading1")]);
    const r = await e.toc.insertTOC({ title: "", headingDepth: 3 }, 0);
    expect(r.entries).toBe(1);
    expect((await texts(e)).filter((t) => t.includes("\t")).map((l) => l.split("\t")[0])).toEqual(['Titre "cit\u00e9" ici']);
  });

  test("the level comes from the field, and one too deep is left out", async () => {
    const e = await docOf([holder(' TC "Profond" \\l "4" '), holder(' TC "Moyen" \\l "2" ')]);
    const r = await e.toc.insertTOC({ title: "", headingDepth: 3 }, 0);
    expect(r.entries).toBe(1);
    expect(xmlOf(e)).toContain('<w:pStyle w:val="TOC2"/>');
  });

  test("no TC field: the instruction is what it always was", async () => {
    const e = await docOf([heading("Introduction")]);
    await e.toc.insertTOC({ title: "" }, 0);
    expect(xmlOf(e)).not.toMatch(/TOC [^<]*\\f/);
  });
});

describe("insertTOC — the field", () => {
  test("writes a real TOC field with the requested depth", async () => {
    const e = await docOf([heading("Introduction"), "body", heading("Method")]);
    await e.toc.insertTOC({ title: "Table of Contents", headingDepth: 3 });

    const xml = xmlOf(e);
    expect(xml).toContain('TOC \\o "1-3"');
    expect(xml).toContain("\\h");                       // hyperlinks on by default
    expect(xml).toContain('w:fldCharType="begin"');
    expect(xml).toContain('w:fldCharType="separate"');
    expect(xml).toContain('w:fldCharType="end"');
    expect(xml).toContain('w:val="TOCHeading"');
  });

  // The engine emits no dirty fields anywhere (caption-list.spec, hanachi spec).
  // Pinned here too, because a table of contents is the most tempting place to
  // add one — it is how you would make Word recompute the page numbers on open.
  test("emits no dirty field", async () => {
    const e = await docOf([heading("Introduction"), "body"]);
    await e.toc.insertTOC({ title: "Contents" });
    expect(xmlOf(e)).not.toContain("w:dirty");
  });

  test("collects every heading down to the depth, and no deeper", async () => {
    const e = await docOf([
      heading("Chapter One", 1),
      heading("Section 1.1", 2),
      heading("Sub 1.1.1", 3),
      "body text",
    ]);
    const res = await e.toc.insertTOC({ headingDepth: 2, title: "" });

    expect(res.entries).toBe(2);
    const xml = xmlOf(e);
    expect(xml).toContain('TOC \\o "1-2"');
    expect(xml).toContain('<w:pStyle w:val="TOC1"/>');
    expect(xml).toContain('<w:pStyle w:val="TOC2"/>');
    expect(xml).not.toContain('<w:pStyle w:val="TOC3"/>');
  });

  test("entries are pre-populated, anchored and paginated", async () => {
    const e = await docOf([heading("Introduction"), "body"]);
    await e.toc.insertTOC({ title: "Contents" });

    const xml = xmlOf(e);
    // The heading got an anchor …
    const anchor = /<w:bookmarkStart\b[^>]*w:name="(_Toc\d+)"/.exec(xml)?.[1];
    expect(anchor).toBeTruthy();
    // … and the entry both links to it and paginates off it.
    expect(xml).toContain(`<w:hyperlink w:anchor="${anchor}"`);
    expect(xml).toContain(`PAGEREF ${anchor}`);
    // The entry text is really in the document, not just promised by the field.
    expect(await texts(e)).toContain("Introduction");
  });

  test("reuses an anchor the heading already has instead of stacking a second", async () => {
    const anchored: BodyBlock = {
      kind: "paragraph",
      tag: "w:p",
      xml:
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` +
        `<w:bookmarkStart w:id="7" w:name="_Toc900000001"/>` +
        `<w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>`,
    };
    const e = await docOf([anchored, "body"]);
    await e.toc.insertTOC({ title: "" });

    const xml = xmlOf(e);
    expect(xml).toContain("PAGEREF _Toc900000001");
    expect([...xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="_Toc/g)].length).toBe(1);
  });

  test("distinct anchors per heading — entries must not all point at the first", async () => {
    const e = await docOf([heading("One"), heading("Two"), heading("Three")]);
    await e.toc.insertTOC({ title: "" });

    const xml = xmlOf(e);
    const names = [...xml.matchAll(/<w:bookmarkStart\b[^>]*w:name="(_Toc\d+)"/g)].map((m) => m[1]);
    expect(names.length).toBe(3);
    expect(new Set(names).size).toBe(3);

    const ids = [...xml.matchAll(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test("a document with no headings still gets a usable field", async () => {
    const e = await docOf(["just", "body", "text"]);
    const res = await e.toc.insertTOC({ title: "Contents" });

    expect(res.entries).toBe(0);
    expect(xmlOf(e)).toContain("No headings found");
  });

  test("a deeper re-run registers the entry styles the first one didn't", async () => {
    const e = await docOf([heading("One", 1), heading("Two", 2), heading("Three", 3), heading("Four", 4)]);
    await e.toc.insertTOC({ headingDepth: 2, title: "" });
    // The engine is cached per thesis, so this SECOND call reuses the same
    // manager — the one that already thinks it registered its styles.
    const res = await e.toc.insertTOC({ headingDepth: 4, title: "" });

    expect(res.entries).toBe(4);
    const styles = e.zip.readAsText("word/styles.xml")!;
    for (const lvl of [1, 2, 3, 4]) expect(styles, `TOC${lvl}`).toContain(`w:styleId="TOC${lvl}"`);
  });

  test("headingDepth is clamped to Word's 1–9", async () => {
    const e = await docOf([heading("One")]);
    const res = await e.toc.insertTOC({ headingDepth: 42, title: "" });
    expect(res.headingDepth).toBe(9);
    expect(xmlOf(e)).toContain('TOC \\o "1-9"');
  });
});

describe("insertTOC — placement and body order", () => {
  test("a table between paragraphs stays between them", async () => {
    const e = await docOf([heading("Chapter"), "before", table("cell"), "after"]);
    await e.toc.insertTOC({ title: "Contents" });

    const blocks = await e.document.getBlocks();
    const kinds  = blocks.map((b) => b.kind);
    const tbl    = kinds.indexOf("table");
    expect(tbl).toBeGreaterThan(-1);

    const order = await texts(e);
    expect(order.indexOf("before")).toBeLessThan(tbl);
    expect(order.indexOf("after")).toBeGreaterThan(tbl);
  });

  test("inserts at the given block index, not always at the top", async () => {
    const e = await docOf(["cover page", heading("Introduction"), "body"]);
    const res = await e.toc.insertTOC({ title: "Contents" }, 1);

    expect(res.atIndex).toBe(1);
    const order = await texts(e);
    expect(order[0]).toBe("cover page");
    expect(order[1]).toBe("Contents");
  });

  test("the title itself never becomes an entry of its own table", async () => {
    const e = await docOf([heading("Introduction")]);
    await e.toc.insertTOC({ title: "Contents" });

    // Re-running proves our own scan skips it: still one entry, not two.
    const again = await e.toc.insertTOC({ title: "Contents" });
    expect(again.entries).toBe(1);

    // And Word's own field update can't collect it either: TOCHeading is based on
    // Heading1 for its look, so without an explicit outlineLvl 9 it would inherit
    // outline level 1 and list itself.
    const styles = e.zip.readAsText("word/styles.xml")!;
    const tocHeading = /<w:style\b[^>]*w:styleId="TOCHeading"[\s\S]*?<\/w:style>/.exec(styles)?.[0];
    expect(tocHeading).toContain('<w:outlineLvl w:val="9"/>');
  });
});

describe("insertTOC — RTL", () => {
  test("an Arabic table of contents is written right-to-left", async () => {
    const e = await docOf([heading("المقدمة"), "نص"]);
    await e.toc.insertTOC({ title: "الفهرس", rtl: true });

    const xml = xmlOf(e);
    expect(xml).toContain("<w:bidi/>");
    expect(xml).toContain("<w:rtl/>");
    expect(await texts(e)).toContain("الفهرس");
    expect(await texts(e)).toContain("المقدمة");
  });
});

describe("findTypedTOC — the table the student typed by hand", () => {
  // The production failure: the thesis carried a hand-typed contents list, which
  // is not a field, so removeTOC never saw it. The model was left guessing a
  // block range, apologised for "صعوبة تقنية في تحديد نهاية الفهرس اليدوي", and
  // asked to delete 14 blocks it had not located. The span is computed here.
  const typedEntry = (text: string): BodyBlock => makeParagraphNode(`${text} .......... 5`);

  test("finds title + entries and stops at the body text", async () => {
    const e = await docOf([
      "Page de garde",
      "Table des matières",
      typedEntry("Introduction"),
      typedEntry("Chapitre 1"),
      typedEntry("Chapitre 2"),
      typedEntry("Conclusion"),
      "Ceci est un vrai paragraphe du mémoire qui continue le texte normalement.",
    ]);
    const span = (await e.toc.findTypedTOC())!;

    expect(span).toBeTruthy();
    expect(span.title).toBe("Table des matières");
    expect(span.startIndex).toBe(1);       // the title, not the first entry
    expect(span.endIndex).toBe(5);         // last entry — the body text is excluded
    expect(span.entries).toBe(4);
    expect(span.sample[0]).toContain("Introduction");
  });

  test("a tab-separated typed table counts too (a tab carries no text)", async () => {
    const tabbed = (text: string): BodyBlock => ({
      kind: "paragraph",
      tag: "w:p",
      xml: `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>7</w:t></w:r></w:p>`,
    });
    const e = await docOf(["الفهرس", tabbed("المقدمة"), tabbed("الفصل الأول"), tabbed("الخاتمة")]);
    const span = (await e.toc.findTypedTOC())!;

    expect(span.title).toBe("الفهرس");
    expect(span.entries).toBe(3);
  });

  test("a GENERATED table is never mistaken for a typed one", async () => {
    const e = await docOf([heading("One"), heading("Two"), heading("Three"), "body"]);
    await e.toc.insertTOC({ title: "Contents" });
    expect(await e.toc.findTypedTOC()).toBeNull();
  });

  test("a list of figures is never mistaken for one either", async () => {
    const e = await docOf(["[fig]"]);
    await e.captions.insertFigureCaption(0, "Schéma");
    await e.captions.insertListOfFigures("List of Figures", 0);
    expect(await e.toc.findTypedTOC()).toBeNull();
  });

  test("ordinary numbered prose is not a table of contents", async () => {
    const e = await docOf([
      "Le chiffre d'affaires a progressé de 12",
      "Les effectifs sont passés à 42",
      "La marge nette atteint 8",
    ]);
    expect(await e.toc.findTypedTOC()).toBeNull();   // no leader, no title
  });

  test("removeTypedTOC deletes exactly the span and nothing else", async () => {
    const e = await docOf([
      "Page de garde",
      "Table des matières",
      typedEntry("Introduction"),
      typedEntry("Chapitre 1"),
      typedEntry("Conclusion"),
      "Le corps du mémoire commence ici et doit survivre intact.",
    ]);
    const removal = await e.toc.removeTypedTOC();

    expect(removal.removed).toBe(4);
    expect(removal.at).toBe(1);
    expect(await texts(e)).toEqual([
      "Page de garde",
      "Le corps du mémoire commence ici et doit survivre intact.",
    ]);
  });
});

describe("removeTOC / replaceExisting", () => {
  test("re-inserting refreshes rather than stacking a second table", async () => {
    const e = await docOf([heading("One"), "body"]);
    await e.toc.insertTOC({ title: "Contents" });

    // A heading is added after the fact — exactly why a student re-runs it.
    const blocks = await e.document.getBlocks();
    blocks.push(heading("Two"));
    await e.document.saveBlocks(blocks);

    const res = await e.toc.insertTOC({ title: "Contents" });
    expect(res.replaced).toBeGreaterThan(0);
    expect(res.entries).toBe(2);

    const xml = xmlOf(e);
    expect([...xml.matchAll(/TOC \\o/g)].length).toBe(1);
    expect([...xml.matchAll(/w:val="TOCHeading"/g)].length).toBe(1);
  });

  test("replaceExisting:false leaves the old one alone", async () => {
    const e = await docOf([heading("One")]);
    await e.toc.insertTOC({ title: "Contents" });
    await e.toc.insertTOC({ title: "Contents", replaceExisting: false });

    expect([...xmlOf(e).matchAll(/TOC \\o/g)].length).toBe(2);
  });

  test("removeTOC deletes title, field and every entry", async () => {
    const e = await docOf([heading("One"), "body"]);
    await e.toc.insertTOC({ title: "Contents" });
    expect(await e.toc.hasTOC()).toBe(true);

    const removal = await e.toc.removeTOC();
    expect(removal.removed).toBeGreaterThan(0);
    expect(await e.toc.hasTOC()).toBe(false);

    const xml = xmlOf(e);
    expect(xml).not.toContain("TOC \\o");
    expect(xml).not.toContain("Contents");
    // The document's own content survives untouched.
    expect(await texts(e)).toEqual(["One", "body"]);
  });

  test("a List of Figures is NOT a table of contents — replacing must not eat it", async () => {
    const e = await docOf([heading("One"), "body"]);
    await e.captions.insertListOfFigures("List of Figures", 0);
    await e.toc.insertTOC({ title: "Contents" });

    const xml = xmlOf(e);
    expect(xml).toContain('\\c "Figure"');
    expect(xml).toContain("List of Figures");
    expect(xml).toContain("TOC \\o");

    await e.toc.removeTOC();
    const after = xmlOf(e);
    expect(after).toContain('\\c "Figure"');
    expect(after).toContain("List of Figures");
    expect(after).not.toContain("TOC \\o");
  });
});
