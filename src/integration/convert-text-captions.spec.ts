import { describe, test, expect } from "vitest";
import { Mdocxengine, makeParagraphNode, parseTextCaption, type BodyBlock } from "@/index";

/**
 * Converting HAND-TYPED captions ("Figure 1 : Organigramme") into real Word
 * captions. This is the only path by which an imported thesis — which is nearly
 * every thesis — can ever get a working List of Figures, cross-references or
 * automatic renumbering.
 *
 * The bar is Word parity plus the product rule: the student's wording,
 * separator, alignment and run formatting survive the conversion untouched.
 */

const BASE = "samples/example.docx";

/** An engine whose body is exactly `blocks`. */
async function docOf(blocks: Array<string | BodyBlock>): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile(BASE);
  await e.document.saveBlocks(blocks.map((b) => (typeof b === "string" ? makeParagraphNode(b) : b)));
  return e;
}

const heading = (text: string): BodyBlock => makeParagraphNode(text, "Heading1");

/** Visible text of every body paragraph, in order. */
async function texts(e: Mdocxengine): Promise<string[]> {
  const blocks = await e.document.getBlocks();
  const { paragraphText } = await import("@/core/files/body/OrderedBody");
  return blocks.map((b) => paragraphText(b.xml).trim());
}

describe("parseTextCaption — detection", () => {
  test("recognises the shapes a thesis actually types", () => {
    expect(parseTextCaption("Figure 1 : Organigramme")).toMatchObject({
      kind: "figure", rawLabel: "Figure", rawNumber: "1", text: "Organigramme",
    });
    expect(parseTextCaption("Tableau 3. Répartition de l'échantillon")).toMatchObject({
      kind: "table", rawNumber: "3", text: "Répartition de l'échantillon",
    });
    expect(parseTextCaption("Fig. I-2 – Schéma du processus")).toMatchObject({
      kind: "figure", rawLabel: "Fig.", rawNumber: "I-2",
    });
    expect(parseTextCaption("الجدول رقم 3: توزيع العينة")).toMatchObject({
      kind: "table", rawLabel: "الجدول رقم", rawNumber: "3", text: "توزيع العينة",
    });
    expect(parseTextCaption("Équation 2 : E = mc²")).toMatchObject({ kind: "equation" });
    // Arabic-Indic digits and a parenthesised number.
    expect(parseTextCaption("الشكل رقم (٤) : الهيكل التنظيمي")).toMatchObject({
      kind: "figure", rawNumber: "٤",
    });
  });

  test("refuses prose, headings and false friends", () => {
    // A cross-reference inside a sentence — the classic false positive.
    expect(parseTextCaption("Figure 1 shows that the sample was representative")).toBeNull();
    expect(parseTextCaption("Table of contents")).toBeNull();
    expect(parseTextCaption("Table des matières")).toBeNull();
    expect(parseTextCaption("Figured out the answer")).toBeNull();
    expect(parseTextCaption("Les résultats de la Figure 1 : voir plus bas")).toBeNull();
  });
});

describe("parseTextCaption — a label with its adjective", () => {
  test("الشكل البياني رقم (7): is a figure's label and number, not wording", () => {
    expect(parseTextCaption("الشكل البياني رقم (7): وصف قصير")).toMatchObject({
      kind: "figure", rawLabel: "الشكل البياني رقم", rawNumber: "7", text: "وصف قصير",
    });
    expect(parseTextCaption("شكل بياني 3 - وصف قصير")).toMatchObject({ kind: "figure", rawLabel: "شكل بياني", rawNumber: "3" });
  });

  test("the words alone, with nothing marking a caption, stay a sentence", () => {
    expect(parseTextCaption("الشكل البياني يوضح أن النتائج تحسنت كثيرا في هذه المرحلة من العمل الميداني")).toBeNull();
  });
});

describe("CaptionManager.convertTextCaptions", () => {
  test("rewrites a typed caption as a real SEQ field, in place", async () => {
    const e = await docOf(["Intro", "Figure 1 : Organigramme de l'entreprise", "Suite"]);
    const before = (await e.document.getBlocks()).length;

    const res = await e.captions.convertTextCaptions();

    expect(res.converted).toHaveLength(1);
    expect(res.converted[0]).toMatchObject({ blockIndex: 1, kind: "figure", label: "Figure" });
    // In place: no block added or removed, so caller-held indices stay valid.
    expect((await e.document.getBlocks()).length).toBe(before);

    const [caption] = await e.captions.getCaptions();
    expect(caption).toMatchObject({
      label: "Figure", displayLabel: "Figure", number: "1", blockIndex: 1,
      // Everything after the number IS the caption text, separator included —
      // the same thing Word stores when you type ": Organigramme" in its dialog.
      text: ": Organigramme de l'entreprise",
      fullText: "Figure 1 : Organigramme de l'entreprise",
    });
    const xml = (await e.document.getBlocks())[1].xml;
    expect(xml).toContain("SEQ Figure");
    expect(xml).toContain('<w:pStyle w:val="Caption"/>');
    expect(xml).toMatch(/<w:bookmarkStart\b/);
  });

  test("keeps the student's own separator, alignment and run formatting", async () => {
    const typed: BodyBlock = {
      kind: "paragraph",
      tag: "w:p",
      xml:
        '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240"/></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>Figure 1 : Organigramme</w:t></w:r></w:p>',
    };
    const e = await docOf([typed]);
    await e.captions.convertTextCaptions();

    const xml = (await e.document.getBlocks())[0].xml;
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('<w:spacing w:before="240"/>');
    expect(xml).toContain('<w:sz w:val="20"/>');
    // pPr is an ORDERED sequence — pStyle must come first or Word rejects the file.
    expect(xml).toMatch(/<w:pPr><w:pStyle w:val="Caption"\/><w:jc/);
    // The caption still READS the way it was typed.
    expect((await texts(e))[0]).toBe("Figure 1 : Organigramme");
  });

  test("renumbers in DOCUMENT order, whatever the student typed", async () => {
    const e = await docOf([
      "Figure 7 : Septième",
      "Figure 3 : Troisième",
      "Figure 12 : Douzième",
    ]);
    await e.captions.convertTextCaptions();

    expect((await e.captions.getCaptions()).map((c) => c.number)).toEqual(["1", "2", "3"]);
    expect(await texts(e)).toEqual([
      "Figure 1 : Septième",
      "Figure 2 : Troisième",
      "Figure 3 : Douzième",
    ]);
  });

  test("plain sequential numbering is the default — no chapter number", async () => {
    const e = await docOf([
      heading("Chapitre 1"),
      "Figure 1 : A",
      heading("Chapitre 2"),
      "Figure 1 : B",
    ]);
    await e.captions.convertTextCaptions();

    expect(await texts(e)).toEqual(["Chapitre 1", "Figure 1 : A", "Chapitre 2", "Figure 2 : B"]);
    expect((await e.document.getBlocks())[1].xml).not.toContain("STYLEREF");
  });

  test("per-chapter numbering gives Figure I-1, I-2, II-1", async () => {
    const e = await docOf([
      heading("Chapitre 1"),
      "Figure 1 : A",
      "Figure 2 : B",
      heading("Chapitre 2"),
      "Figure 1 : C",
    ]);
    await e.captions.convertTextCaptions({
      numbering: { includeChapterNumber: true, chapterStyle: "Heading1", chapterFormat: "ROMAN" },
    });

    expect(await texts(e)).toEqual([
      "Chapitre 1", "Figure I-1 : A", "Figure I-2 : B", "Chapitre 2", "Figure II-1 : C",
    ]);
    const xml = (await e.document.getBlocks())[1].xml;
    // The live field, not just the cached text Word will recompute.
    expect(xml).toContain("STYLEREF 1 \\s \\* ROMAN");
    expect(xml).toContain("SEQ Figure \\* ARABIC \\s 1");
  });

  test("unifies mixed spellings so one sequence collects them all", async () => {
    const e = await docOf(["Fig. 2 : A", "Figure 3 : B", "Figure 9 : C"]);
    const res = await e.captions.convertTextCaptions();

    expect(res.labels.figure).toBe("Figure");
    const labels = await e.captions.listLabels();
    expect(labels).toHaveLength(1);
    expect(labels[0]).toMatchObject({ label: "Figure", count: 3 });
  });

  test("figures and tables get their own independent sequences", async () => {
    const e = await docOf([
      "Figure 1 : A", "Tableau 1 : T1", "Figure 2 : B", "Tableau 2 : T2",
    ]);
    await e.captions.convertTextCaptions();

    expect(await texts(e)).toEqual([
      "Figure 1 : A", "Tableau 1 : T1", "Figure 2 : B", "Tableau 2 : T2",
    ]);
    expect((await e.captions.listLabels()).map((l) => l.label).sort()).toEqual(["Figure", "Tableau"]);
  });

  test("kind restricts the conversion", async () => {
    const e = await docOf(["Figure 1 : A", "Tableau 1 : T1"]);
    const res = await e.captions.convertTextCaptions({ kind: "table" });

    expect(res.converted.map((c) => c.blockIndex)).toEqual([1]);
    expect((await e.captions.getCaptions()).map((c) => c.label)).toEqual(["Tableau"]);
  });

  test("fromIndex/toIndex bound the scan", async () => {
    const e = await docOf(["Figure 1 : A", "Figure 2 : B", "Figure 3 : C"]);
    const res = await e.captions.convertTextCaptions({ fromIndex: 1, toIndex: 1 });

    expect(res.converted.map((c) => c.blockIndex)).toEqual([1]);
    expect(await e.captions.getCaptions()).toHaveLength(1);
  });

  test("reuses the label the document's REAL captions already use", async () => {
    const e = await docOf(["Une figure", "الشكل 4 : صورة ثانية"]);
    await e.captions.insertCaption(0, { label: "الشكل رقم", text: "صورة أولى" });

    const res = await e.captions.convertTextCaptions();
    // Not a second, competing "الشكل" sequence.
    expect(res.labels.figure).toBe("الشكل رقم");
    expect(await e.captions.listLabels()).toHaveLength(1);
  });

  test("leaves an existing real caption alone", async () => {
    const e = await docOf(["Une figure"]);
    await e.captions.insertCaption(0, { label: "Figure", text: "Déjà une vraie légende" });

    const res = await e.captions.convertTextCaptions();
    expect(res.converted).toHaveLength(0);
    expect(res.skipped).toEqual([
      expect.objectContaining({ reason: "already-a-caption", blockIndex: 1 }),
    ]);
    expect(await e.captions.getCaptions()).toHaveLength(1);
  });

  test("never touches a paragraph that carries the picture itself", async () => {
    const withImage: BodyBlock = {
      kind: "paragraph",
      tag: "w:p",
      xml: '<w:p><w:r><w:t>Figure 1 : Vue</w:t></w:r><w:r><w:drawing/></w:r></w:p>',
    };
    const e = await docOf([withImage]);
    const res = await e.captions.convertTextCaptions();

    expect(res.converted).toHaveLength(0);
    expect(res.skipped[0]).toMatchObject({ reason: "contains-image" });
    expect((await e.document.getBlocks())[0].xml).toContain("<w:drawing/>");
  });

  test("dryRun reports the plan and writes nothing", async () => {
    const e = await docOf(["Figure 7 : A", "Figure 9 : B"]);
    const before = e.zip.readAsText("word/document.xml")!;

    const res = await e.captions.convertTextCaptions({ dryRun: true });

    expect(res.dryRun).toBe(true);
    expect(res.converted).toHaveLength(2);
    // The preview shows the numbers the real run would write.
    expect(res.converted.map((c) => c.after)).toEqual(["Figure 1 : A", "Figure 2 : B"]);
    expect(e.zip.readAsText("word/document.xml")).toBe(before);
    expect(await e.captions.getCaptions()).toHaveLength(0);
  });

  test("converts a caption parked inside a one-cell table", async () => {
    const cellCaption: BodyBlock = {
      kind: "table",
      tag: "w:tbl",
      xml:
        "<w:tbl><w:tr><w:tc>" +
        "<w:p><w:r><w:t>Figure 1 : Dans une cellule</w:t></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl>",
    };
    const e = await docOf([cellCaption]);
    const res = await e.captions.convertTextCaptions();

    expect(res.converted).toHaveLength(1);
    const [caption] = await e.captions.getCaptions();
    expect(caption).toMatchObject({ inTable: true, number: "1", fullText: "Figure 1 : Dans une cellule" });
    // The table survived — only the paragraph inside it was rewritten.
    expect((await e.document.getBlocks())[0].kind).toBe("table");
  });

  test("an Arabic caption stays right-to-left", async () => {
    const rtlTyped: BodyBlock = {
      kind: "paragraph",
      tag: "w:p",
      xml:
        '<w:p><w:pPr><w:bidi/></w:pPr>' +
        '<w:r><w:rPr><w:rtl/></w:rPr><w:t>الجدول رقم 3: توزيع العينة</w:t></w:r></w:p>',
    };
    const e = await docOf([rtlTyped]);
    await e.captions.convertTextCaptions();

    const xml = (await e.document.getBlocks())[0].xml;
    expect(xml).toContain("<w:bidi/>");
    expect(xml).toContain("<w:rtl/>");
    expect((await texts(e))[0]).toBe("الجدول رقم 1: توزيع العينة");
  });

  test("converted captions feed a real List of Figures", async () => {
    const e = await docOf(["Figure 1 : A", "Figure 2 : B"]);
    await e.captions.convertTextCaptions();
    await e.captions.insertCaptionList("Figure", "Liste des figures", 0);

    const xml = e.zip.readAsText("word/document.xml")!;
    expect(xml).toContain('TOC \\h \\z \\c "Figure"');
    // Each converted caption anchors its OWN bookmark, so the list's PAGEREFs
    // point at three different places rather than all at the first one.
    const refs = new Set([...xml.matchAll(/PAGEREF (_Ref\d+)/g)].map((m) => m[1]));
    expect(refs.size).toBe(2);
  });
});
