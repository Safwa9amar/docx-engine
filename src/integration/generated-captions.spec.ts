import { describe, test, expect } from "vitest";
import {
  Mdocxengine, makeParagraphNode, parseTextCaption, parseCaptionParagraph, removeLeadingWord, paragraphText,
  inspectDocx, type BodyBlock,
} from "@/index";

/**
 * `addGeneratedCaption`: an UNNUMBERED caption gets a real label and number as
 * a field, in front of the student's runs — which stay byte-identical. And the
 * leading typed number ("2- جدول …") is a caption number, not wording.
 */

const BASE = "samples/example.docx";
const rulesOf = (e: Mdocxengine) => inspectDocx(e.zip as any).findings.map((f) => f.rule).sort();

async function docOf(blocks: Array<string | BodyBlock>): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile(BASE);
  await e.document.saveBlocks(blocks.map((b) => (typeof b === "string" ? makeParagraphNode(b) : b)));
  return e;
}
const para = (inner: string, pPr = ""): BodyBlock => ({
  kind: "paragraph", tag: "w:p", xml: `<w:p>${pPr}${inner}</w:p>`,
});
const run = (t: string, rPr = "") => `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>`;

describe("parseTextCaption — leading typed number", () => {
  test("a number in front of the label is the caption number", () => {
    expect(parseTextCaption("2- جدول البيان المالي")).toMatchObject({
      kind: "table", rawNumber: "2", leadingNumber: "2", text: "البيان المالي",
    });
    expect(parseTextCaption("3) Figure : Plan du site")).toMatchObject({
      kind: "figure", rawNumber: "3", text: "Plan du site",
    });
    expect(parseTextCaption("١. شكل س")).toMatchObject({ kind: "figure", rawNumber: "١", text: "س" });
  });

  test("a label's own number wins; plain captions are unchanged", () => {
    expect(parseTextCaption("4- Tableau 7 : Résultats")).toMatchObject({ rawNumber: "7", leadingNumber: "4" });
    expect(parseTextCaption("Figure 1 : Organigramme")).toMatchObject({ rawNumber: "1", leadingNumber: "" });
  });

  test("lists, sections and prose stay out", () => {
    expect(parseTextCaption("2- Table of contents")).toBeNull();
    expect(parseTextCaption("2- Table des matières")).toBeNull();
    expect(parseTextCaption("2.1 Tableau synthétique")).toBeNull();
    expect(parseTextCaption("1. Image quality matters")).toBeNull();
    expect(parseTextCaption("2- Figure")).toBeNull();
    expect(parseTextCaption("civil- Tableau des bilans")).toBeNull();
    expect(parseTextCaption("Table")).toBeNull();
  });
});

describe("convertTextCaptions — leading number", () => {
  test("drops the typed lead and numbers it for real", async () => {
    const e = await docOf(["Intro", "7- جدول البيان المالي", "8- جدول آخر"]);
    const res = await e.captions.convertTextCaptions();
    expect(res.converted.map((c) => [c.typedNumber, c.number])).toEqual([["7", "1"], ["8", "2"]]);
    const caps = await e.captions.getCaptions();
    expect(caps.map((c) => c.fullText)).toEqual(["جدول 1 البيان المالي", "جدول 2 آخر"]);
    expect(caps[0].text).toBe("البيان المالي");
  });
});

describe("CaptionManager.addGeneratedCaption", () => {
  test("label + SEQ number in front; student's words untouched; reads back", async () => {
    const e = await docOf(["Intro", "Plan masse du site", "Suite"]);
    const r = await e.captions.addGeneratedCaption(1, { kind: "figure", label: "Figure", textSeparator: " : " });
    expect(r).toEqual({ ok: true, label: "Figure", number: "1", droppedWord: false });

    const [cap] = await e.captions.getCaptions();
    expect(cap).toMatchObject({ label: "Figure", displayLabel: "Figure", number: "1", text: "Plan masse du site" });
    const xml = (await e.document.getBlocks())[1].xml;
    expect(xml).toContain('<w:pStyle w:val="Caption"/>');
    expect(xml).toContain("SEQ Figure");
    expect(xml).toContain("w:bookmarkStart");
    // It now reads as a typed caption too, but convert must not double-number it.
    expect((await e.captions.convertTextCaptions()).converted).toEqual([]);
    expect(rulesOf(e)).toEqual(rulesOf(await docOf(["Intro", "Plan masse du site", "Suite"])));
  });

  test("student's runs stay byte-identical: footnote ref, hyperlink, field, formatting, bookmark", async () => {
    const student =
      '<w:bookmarkStart w:id="900" w:name="mine"/>' +
      run("Budget ", "<w:rPr><w:b/></w:rPr>") +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' +
      '<w:hyperlink r:id="rId9"><w:r><w:t>lien</w:t></w:r></w:hyperlink>' +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple>' +
      '<w:bookmarkEnd w:id="900"/>';
    const e = await docOf(["Intro", para(student, '<w:pPr><w:jc w:val="center"/></w:pPr>')]);
    const r = await e.captions.addGeneratedCaption(1, { kind: "table", label: "Tableau", textSeparator: " : " });
    expect(r.ok).toBe(true);
    const xml = (await e.document.getBlocks())[1].xml;
    expect(xml.endsWith(student + "</w:p>")).toBe(true);
    // pPr is ordered: pStyle first, the student's alignment kept.
    expect(xml).toMatch(/<w:pPr><w:pStyle w:val="Caption"\/><w:jc w:val="center"\/><\/w:pPr>/);
    const [cap] = parseCaptionParagraph(xml, 1);
    expect(cap.text).toBe("Budget lien4");
  });

  test("numbers in document order and joins the typed sequence", async () => {
    const e = await docOf([
      "Tableau 5 : Typé", "Corps", "Sans numéro A", "Sans numéro B",
    ]);
    await e.captions.convertTextCaptions();
    expect((await e.captions.addGeneratedCaption(3, { kind: "table", label: "Tableau" })).ok).toBe(true);
    expect((await e.captions.addGeneratedCaption(2, { kind: "table", label: "Tableau" })).ok).toBe(true);
    const caps = await e.captions.getCaptions();
    expect(caps.map((c) => [c.label, c.number, c.text])).toEqual([
      ["Tableau", "1", ": Typé"],
      ["Tableau", "2", "Sans numéro A"],
      ["Tableau", "3", "Sans numéro B"],
    ]);
    expect(rulesOf(e).filter((x) => /dup|bookmark/i.test(x))).toEqual([]);
  });

  test("refuses with a typed reason and changes nothing", async () => {
    const withImage = para(
      '<w:r><w:drawing><wp:inline/></w:drawing></w:r>' + run("Légende"),
    );
    const e = await docOf([
      makeParagraphNode("Titre", "Heading1"),
      "",
      withImage,
      "Figure 1 : déjà",
      "Texte",
    ]);
    await e.captions.convertTextCaptions();
    const before = (await e.document.getBlocks()).map((b) => b.xml);
    const opts = { kind: "figure" as const, label: "Figure" };
    expect(await e.captions.addGeneratedCaption(0, opts)).toEqual({ ok: false, reason: "is-a-heading" });
    expect(await e.captions.addGeneratedCaption(1, opts)).toEqual({ ok: false, reason: "empty-text" });
    expect(await e.captions.addGeneratedCaption(2, opts)).toEqual({ ok: false, reason: "contains-image" });
    expect(await e.captions.addGeneratedCaption(3, opts)).toEqual({ ok: false, reason: "already-a-caption" });
    expect(await e.captions.addGeneratedCaption(99, opts)).toEqual({ ok: false, reason: "no-such-block" });
    expect((await e.document.getBlocks()).map((b) => b.xml)).toEqual(before);
  });

  test("a table block is refused as not-a-paragraph", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const blocks = await e.document.getBlocks();
    const t = blocks.findIndex((b) => b.kind === "table");
    if (t < 0) return;
    expect(await e.captions.addGeneratedCaption(t, { kind: "table", label: "Table" })).toEqual({
      ok: false, reason: "not-a-paragraph",
    });
  });
});


// ─── dropLeadingWord: the typed label word becomes the generated label ────────

const AR_TABLE = "جدول";
const arPara = (inner: string, pPr = "<w:pPr><w:bidi/></w:pPr>"): BodyBlock => para(inner, pPr);
const rtlRun = (t: string, extra = "") => run(t, `<w:rPr>${extra}<w:rtl/></w:rPr>`);

describe("removeLeadingWord — run-level surgery", () => {
  test("drops the word and the whitespace after it from the first text run only", () => {
    const first = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">جدول الميزانية المالية</w:t></w:r>';
    const second = run(" و ملحق", "<w:rPr><w:i/></w:rPr>");
    const r = removeLeadingWord(`<w:p><w:pPr><w:bidi/></w:pPr>${first}${second}</w:p>`, AR_TABLE);
    expect(r.dropped).toBe(true);
    expect(r.xml).toBe(
      `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">الميزانية المالية</w:t></w:r>${second}</w:p>`,
    );
  });

  test("a run that held only the word and its space disappears; the next runs are byte-identical", () => {
    const rest = rtlRun("الميزانية ", "<w:b/>") +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' + rtlRun("المالية");
    const r = removeLeadingWord(`<w:p>${rtlRun("جدول ")}${rest}</w:p>`, AR_TABLE);
    expect(r.dropped).toBe(true);
    expect(r.xml).toBe(`<w:p>${rest}</w:p>`);
  });

  test("the word alone in its run, the space in the next: the next run is not touched", () => {
    const next = rtlRun(" الميزانية");
    const r = removeLeadingWord(`<w:p>${rtlRun("جدول")}${next}</w:p>`, AR_TABLE);
    expect(r.dropped).toBe(true);
    expect(r.xml).toBe(`<w:p>${next}</w:p>`);
  });

  test("keeps the run's other children, and a leading bookmark", () => {
    const xml = '<w:p><w:bookmarkStart w:id="9" w:name="mine"/><w:r><w:t>جدول الميزانية</w:t><w:tab/></w:r><w:bookmarkEnd w:id="9"/></w:p>';
    const r = removeLeadingWord(xml, AR_TABLE);
    expect(r.dropped).toBe(true);
    expect(r.xml).toBe('<w:p><w:bookmarkStart w:id="9" w:name="mine"/><w:r><w:t>الميزانية</w:t><w:tab/></w:r><w:bookmarkEnd w:id="9"/></w:p>');
  });

  test("leading spaces, a no-break space and a tatweel in the word are tolerated", () => {
    expect(removeLeadingWord(`<w:p>${rtlRun("  جدول\u00A0الميزانية")}</w:p>`, AR_TABLE).xml)
      .toBe(`<w:p>${rtlRun("الميزانية")}</w:p>`);
    expect(removeLeadingWord(`<w:p>${rtlRun("جـدول الميزانية")}</w:p>`, AR_TABLE).dropped).toBe(true);
  });

  test("XML-special characters after the word stay escaped", () => {
    const r = removeLeadingWord(`<w:p>${run("جدول الربح &amp; الخسارة")}</w:p>`, AR_TABLE);
    expect(r.xml).toBe(`<w:p>${run("الربح &amp; الخسارة")}</w:p>`);
  });

  test.each([
    ["split across runs", `<w:p>${rtlRun("جد")}${rtlRun("ول الميزانية")}</w:p>`],
    ["not at the start", `<w:p>${rtlRun("الميزانية جدول")}</w:p>`],
    ["a longer word", `<w:p>${rtlRun("جدولة المصاريف")}</w:p>`],
    ["the word with an article", `<w:p>${rtlRun("الجدول المالي")}</w:p>`],
    ["the word alone", `<w:p>${rtlRun("جدول")}</w:p>`],
    ["the word and spaces alone", `<w:p>${rtlRun("جدول  ")}</w:p>`],
    ["a letter in the next run", `<w:p>${rtlRun("جدول")}${rtlRun("ة")}</w:p>`],
    ["a run before it", `<w:p><w:r><w:tab/></w:r>${rtlRun("جدول الميزانية")}</w:p>`],
    ["inside a hyperlink", `<w:p><w:hyperlink r:id="rId9">${rtlRun("جدول الميزانية")}</w:hyperlink></w:p>`],
    ["inside tracked insertion", `<w:p><w:ins w:id="1" w:author="a">${rtlRun("جدول الميزانية")}</w:ins></w:p>`],
    ["after a field", `<w:p><w:fldSimple w:instr=" PAGE ">${run("4")}</w:fldSimple>${rtlRun("جدول الميزانية")}</w:p>`],
    ["a run holding a text box (a nested run)", '<w:p><w:r><w:t>جدول الميزانية</w:t><w:drawing><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>'],
    ["a revision of the run properties", '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="a"><w:rPr/></w:rPrChange></w:rPr><w:t>جدول الميزانية</w:t></w:r></w:p>'],
    ["an empty paragraph", "<w:p/>"],
  ])("leaves everything alone: %s", (_name, xml) => {
    expect(removeLeadingWord(xml, AR_TABLE)).toEqual({ xml, dropped: false });
  });

  test("an empty word drops nothing", () => {
    const xml = `<w:p>${rtlRun("جدول الميزانية")}</w:p>`;
    expect(removeLeadingWord(xml, "  ")).toEqual({ xml, dropped: false });
  });
});

describe("CaptionManager.addGeneratedCaption — dropLeadingWord", () => {
  const AR = { kind: "table" as const, label: AR_TABLE, rtl: true, textSeparator: ": " };

  test("the word becomes the label: no doubled label", async () => {
    const e = await docOf(["مقدمة", arPara(rtlRun("جدول الميزانية المالية")), "ختام"]);
    const r = await e.captions.addGeneratedCaption(1, { ...AR, dropLeadingWord: AR_TABLE });
    expect(r).toEqual({ ok: true, label: AR_TABLE, number: "1", droppedWord: true });
    const [cap] = await e.captions.getCaptions();
    expect(cap).toMatchObject({ displayLabel: AR_TABLE, number: "1", text: "الميزانية المالية" });
    const xml = (await e.document.getBlocks())[1].xml;
    expect(paragraphText(xml)).toBe("جدول 1: الميزانية المالية");
    expect(xml).toMatch(/<w:pPr><w:pStyle w:val="Caption"\/><w:bidi\/><\/w:pPr>/);
  });

  test("the generated label stays the thesis's own spelling; the student's word is not repeated", async () => {
    const e = await docOf([arPara(rtlRun("جدول المصاريف"))]);
    const r = await e.captions.addGeneratedCaption(0, { kind: "table", label: "الجدول رقم", rtl: true, dropLeadingWord: AR_TABLE });
    expect(r).toMatchObject({ ok: true, droppedWord: true });
    expect(paragraphText((await e.document.getBlocks())[0].xml)).toBe("الجدول رقم 1: المصاريف");
  });

  test("the student's other runs stay byte for byte, in order", async () => {
    const student =
      rtlRun("الميزانية ", "<w:b/>") +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>' +
      '<w:hyperlink r:id="rId9"><w:r><w:t>lien</w:t></w:r></w:hyperlink>' +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple>';
    const e = await docOf([arPara(rtlRun("جدول ", "<w:b/>") + student)]);
    expect((await e.captions.addGeneratedCaption(0, { ...AR, dropLeadingWord: AR_TABLE })).droppedWord).toBe(true);
    const xml = (await e.document.getBlocks())[0].xml;
    expect(xml.endsWith(student + "</w:p>")).toBe(true);
    // the label and number keep the formatting of the first run (bold)
    expect(xml).toContain('<w:r><w:rPr><w:b/><w:rtl/></w:rPr><w:t xml:space="preserve">جدول</w:t></w:r>');
  });

  test("the word alone in its run: the separator gives up its own space, so nothing is doubled", async () => {
    const e = await docOf([arPara(rtlRun("جدول") + rtlRun(" الميزانية"))]);
    expect((await e.captions.addGeneratedCaption(0, { ...AR, dropLeadingWord: AR_TABLE })).droppedWord).toBe(true);
    const xml = (await e.document.getBlocks())[0].xml;
    expect(xml.endsWith(rtlRun(" الميزانية") + "</w:p>")).toBe(true);
    expect(paragraphText(xml)).toBe("جدول 1: الميزانية");
  });

  test.each([
    ["split across runs", rtlRun("جد") + rtlRun("ول الميزانية")],
    ["not at the start", rtlRun("الميزانية جدول")],
    ["a longer word", rtlRun("جدولة المصاريف")],
    ["the word alone", rtlRun("جدول")],
    ["inside a hyperlink", `<w:hyperlink r:id="rId9">${rtlRun("جدول الميزانية")}</w:hyperlink>`],
  ])("cannot drop (%s): returns normally, droppedWord false, nothing else different", async (_name, inner) => {
    const withOption = await docOf([arPara(inner)]);
    const without = await docOf([arPara(inner)]);
    const r = await withOption.captions.addGeneratedCaption(0, { ...AR, dropLeadingWord: AR_TABLE });
    await without.captions.addGeneratedCaption(0, AR);
    expect(r).toMatchObject({ ok: true, droppedWord: false });
    expect((await withOption.document.getBlocks())[0].xml).toBe((await without.document.getBlocks())[0].xml);
  });

  test("without the option nothing is dropped", async () => {
    const e = await docOf([arPara(rtlRun("جدول الميزانية"))]);
    const r = await e.captions.addGeneratedCaption(0, AR);
    expect(r).toMatchObject({ ok: true, droppedWord: false });
    expect(paragraphText((await e.document.getBlocks())[0].xml)).toBe("جدول 1: جدول الميزانية");
  });

  test("the result reads back as one sequence with the typed captions, and the doctor stays quiet", async () => {
    const e = await docOf([
      "جدول 5: مكتوب", arPara(rtlRun("جدول بدون رقم")), arPara(rtlRun("آخر بدون رقم")),
    ]);
    await e.captions.convertTextCaptions();
    await e.captions.addGeneratedCaption(1, { ...AR, dropLeadingWord: AR_TABLE });
    await e.captions.addGeneratedCaption(2, AR);
    const caps = await e.captions.getCaptions();
    expect(caps.map((c) => [c.number, c.text])).toEqual([["1", ": مكتوب"], ["2", "بدون رقم"], ["3", "آخر بدون رقم"]]);
    expect(rulesOf(e).filter((x) => /dup|bookmark/i.test(x))).toEqual([]);
  });
});

// ─── convertTextCaptions: one separator for every converted caption ──────────

describe("convertTextCaptions — textSeparator override", () => {
  const typed = [
    "Tableau 1 - Alpha", "Tableau 2. Beta", "Tableau 3 : Gamma", "Tableau 4 Delta", "Tableau 5، Epsilon",
  ];

  test("without it each caption keeps the punctuation the student typed", async () => {
    const e = await docOf(["Intro", ...typed]);
    await e.captions.convertTextCaptions();
    const texts = (await e.document.getBlocks()).slice(1).map((b) => paragraphText(b.xml));
    expect(texts).toEqual(["Tableau 1 - Alpha", "Tableau 2. Beta", "Tableau 3 : Gamma", "Tableau 4 Delta", "Tableau 5، Epsilon"]);
  });

  test.each([[" : "], [": "], [" - "]])("with %j every converted caption gets exactly that", async (sep) => {
    const e = await docOf(["Intro", ...typed]);
    const res = await e.captions.convertTextCaptions({ textSeparator: sep });
    expect(res.converted).toHaveLength(5);
    const blocks = (await e.document.getBlocks()).slice(1);
    expect(blocks.map((b) => paragraphText(b.xml))).toEqual([
      `Tableau 1${sep}Alpha`, `Tableau 2${sep}Beta`, `Tableau 3${sep}Gamma`, `Tableau 4${sep}Delta`, `Tableau 5${sep}Epsilon`,
    ]);
    // and the sequence still numbers 1..5 under one label
    expect((await e.captions.getCaptions()).map((c) => c.number)).toEqual(["1", "2", "3", "4", "5"]);
  });

  test("a caption with no wording gets no dangling separator", async () => {
    const e = await docOf(["Intro", "Tableau 7"]);
    await e.captions.convertTextCaptions({ textSeparator: " : " });
    expect(paragraphText((await e.document.getBlocks())[1].xml)).toBe("Tableau 1");
  });
});

describe("addGeneratedCaption — spaces typed in front of the wording", () => {
  test("they do not stand between the separator and the words", async () => {
    const e = await docOf([para(run("          وصف قصير") + run(" آخر"))]);
    const r = await e.captions.addGeneratedCaption(0, { kind: "figure", label: "شكل", rtl: true, textSeparator: ": " });
    expect(r.ok).toBe(true);
    expect(paragraphText((await e.document.getBlocks())[0].xml)).toBe("شكل 1: وصف قصير آخر");
  });

  test("a space further on in the wording stays: only what stands in front of it goes", async () => {
    const rPr = "<w:rPr><w:rtl/></w:rPr>";
    const e = await docOf([para(run("mots premiers ( ", rPr) + run("12 )", rPr) + run(" suite ", rPr) + run("fin", rPr))]);
    const r = await e.captions.addGeneratedCaption(0, { kind: "table", label: "Tableau" });
    expect(r.ok).toBe(true);
    expect(paragraphText((await e.document.getBlocks())[0].xml)).toContain("mots premiers ( 12 ) suite fin");
  });

  test("a run of nothing but spaces goes, and the run after it loses its own", async () => {
    const e = await docOf([para(run("   ") + run("  وصف"))]);
    await e.captions.addGeneratedCaption(0, { kind: "figure", label: "شكل", rtl: true, textSeparator: ": " });
    expect(paragraphText((await e.document.getBlocks())[0].xml)).toBe("شكل 1: وصف");
  });
});
