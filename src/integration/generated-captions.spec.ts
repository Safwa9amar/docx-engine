import { describe, test, expect } from "vitest";
import {
  Mdocxengine, makeParagraphNode, parseTextCaption, parseCaptionParagraph, inspectDocx, type BodyBlock,
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
    expect(r).toEqual({ ok: true, label: "Figure", number: "1" });

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
