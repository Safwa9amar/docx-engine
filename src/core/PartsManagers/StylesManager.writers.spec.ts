import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, inspectDocx } from "../../index";
import { mergeRunProps } from "../ooxml/runProps";
import {
  applyStyleParagraphPropsToXml,
  applyDocDefaultsToXml,
  applyEnsureStyleToXml,
  StylesManager,
} from "./StylesManager";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const styles = (inner: string) => `<?xml version="1.0"?><w:styles ${W}>${inner}</w:styles>`;
const NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const H1 =
  '<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="H1"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="100" w:line="300" w:lineRule="exact" w:beforeLines="5"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>';

const apply = (xml: string, id: string, props: Parameters<typeof applyStyleParagraphPropsToXml>[2], ensure?: any) =>
  applyStyleParagraphPropsToXml(xml, id, props, ensure);
const pPrOf = (xml: string, id: string) =>
  new RegExp(`w:styleId="${id}"[\\s\\S]*?(<w:pPr>[\\s\\S]*?</w:pPr>)`).exec(xml)?.[1] ?? "";

describe("mergeRunProps — cs/latin split", () => {
  test("fontLatin touches ascii/hAnsi only; fontCs touches cs only", () => {
    const base = '<w:rFonts w:ascii="A" w:hAnsi="A" w:cs="C" w:eastAsia="E"/>';
    expect(mergeRunProps(base, { fontLatin: "L" })).toBe('<w:rFonts w:ascii="L" w:hAnsi="L" w:eastAsia="E" w:cs="C"/>');
    expect(mergeRunProps(base, { fontCs: "Q" })).toBe('<w:rFonts w:ascii="A" w:hAnsi="A" w:eastAsia="E" w:cs="Q"/>');
    expect(mergeRunProps("", { fontLatin: "L", fontCs: "Q" })).toBe('<w:rFonts w:ascii="L" w:hAnsi="L" w:cs="Q"/>');
  });

  test("a theme attribute shadowing an overwritten font is dropped; the others stay", () => {
    const base = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi" w:eastAsiaTheme="minorEastAsia"/>';
    const out = mergeRunProps(base, { fontCs: "Q" });
    expect(out).toContain('w:cs="Q"');
    expect(out).not.toContain("cstheme");
    expect(out).toContain('w:asciiTheme="minorHAnsi"');
    expect(mergeRunProps(base, { fontLatin: "L" })).not.toContain("asciiTheme");
  });

  test("font plus fontCs: font fills ascii/hAnsi, fontCs wins for cs", () => {
    expect(mergeRunProps("", { font: "F", fontCs: "Q" })).toBe('<w:rFonts w:ascii="F" w:hAnsi="F" w:cs="Q"/>');
    expect(mergeRunProps("", { font: "F", fontLatin: "L" })).toBe('<w:rFonts w:ascii="L" w:hAnsi="L" w:cs="F"/>');
  });

  test("sizes: sizePt alone = sz+szCs (unchanged); sizeCsPt alone = szCs; both = separate", () => {
    expect(mergeRunProps("", { sizePt: 12 })).toBe('<w:sz w:val="24"/><w:szCs w:val="24"/>');
    expect(mergeRunProps('<w:sz w:val="20"/><w:szCs w:val="20"/>', { sizeCsPt: 14 })).toBe('<w:sz w:val="20"/><w:szCs w:val="28"/>');
    expect(mergeRunProps("", { sizePt: 12, sizeCsPt: 14 })).toBe('<w:sz w:val="24"/><w:szCs w:val="28"/>');
    expect(() => mergeRunProps("", { sizeCsPt: 0 })).toThrow();
    expect(() => mergeRunProps("", { sizePt: 12, sizeCsPt: -1 })).toThrow();
  });

  test("font alone is unchanged", () => {
    expect(mergeRunProps("", { font: "Arial" })).toBe('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>');
  });
});

describe("setStyleParagraphProps (pure)", () => {
  test("every property round-trips, in schema order", () => {
    const r = apply(styles(NORMAL), "Normal", {
      lineSpacing: 1.5, beforePt: 6, afterPt: 12, jc: "both", firstLineCm: 1, startCm: 2,
      keepNext: true, keepLines: true, outlineLvl: 2, bidi: true, pageBreakBefore: true,
    });
    expect(r).toMatchObject({ created: false, updated: true });
    expect(pPrOf(r.xml, "Normal")).toBe(
      '<w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/><w:bidi/><w:spacing w:line="360" w:lineRule="auto" w:before="120" w:after="240"/>' +
        '<w:ind w:start="1134" w:firstLine="567"/><w:jc w:val="both"/><w:outlineLvl w:val="2"/></w:pPr>',
    );
  });

  test("null/false remove; unnamed props are untouched; spacing keeps its other attributes", () => {
    const r = apply(styles(H1), "H1", { keepNext: false, outlineLvl: null, afterPt: 3 });
    const p = pPrOf(r.xml, "H1");
    expect(p).not.toContain("keepNext");
    expect(p).not.toContain("outlineLvl");
    expect(p).toContain('w:before="100"');
    expect(p).toContain('w:line="300" w:lineRule="exact"');
    expect(p).toContain('w:after="60"');
    expect(r.xml).toContain("<w:rPr><w:b/></w:rPr>"); // the style's run props untouched
  });

  test("a before-in-points write drops the line-based twin that would override it", () => {
    const p = pPrOf(apply(styles(H1), "H1", { beforePt: 2 }).xml, "H1");
    expect(p).toContain('w:before="40"');
    expect(p).not.toContain("beforeLines");
  });

  test("hanging replaces firstLine and vice versa; start replaces left", () => {
    const mk = (ind: string) =>
      styles(`<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/><w:pPr><w:ind ${ind}/></w:pPr></w:style>`);
    expect(pPrOf(apply(mk('w:left="9" w:firstLine="5"'), "S", { hangingCm: 1, startCm: 1 }).xml, "S")).toBe(
      '<w:pPr><w:ind w:start="567" w:hanging="567"/></w:pPr>',
    );
    expect(pPrOf(apply(mk('w:hanging="9"'), "S", { firstLineCm: 1 }).xml, "S")).toBe(
      '<w:pPr><w:ind w:firstLine="567"/></w:pPr>',
    );
  });

  test("a missing style: no-op without ensure; created with ensure", () => {
    expect(apply(styles(NORMAL), "X", { jc: "left" })).toMatchObject({ created: false, updated: false });
    const r = apply(styles(NORMAL), "X", { jc: "left" }, { name: "X", basedOn: "Normal" });
    expect(r).toMatchObject({ created: true, updated: true });
    expect(pPrOf(r.xml, "X")).toBe('<w:pPr><w:jc w:val="left"/></w:pPr>');
    expect(r.xml).toContain('<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr>');
  });

  test("emptying a pPr removes it; an unchanged result reports no update", () => {
    const only = styles('<w:style w:type="paragraph" w:styleId="S"><w:name w:val="S"/><w:pPr><w:keepNext/></w:pPr></w:style>');
    const r = apply(only, "S", { keepNext: false });
    expect(r.xml).not.toContain("<w:pPr");
    expect(r.updated).toBe(true);
    expect(apply(r.xml, "S", { keepNext: false })).toMatchObject({ updated: false });
    expect(apply(styles(H1), "H1", { keepNext: true })).toMatchObject({ updated: false });
  });

  test("invalid values throw", () => {
    const x = styles(NORMAL);
    for (const bad of [
      { lineSpacing: 0 }, { lineSpacing: NaN }, { beforePt: -1 }, { afterPt: Infinity },
      { jc: "bogus" as any }, { firstLineCm: NaN }, { outlineLvl: 12 }, { outlineLvl: 1.5 },
    ]) expect(() => apply(x, "Normal", bad), JSON.stringify(bad)).toThrow();
  });
});

describe("setDocDefaults (pure)", () => {
  const DD = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="fr-FR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';

  test("fonts, sizes, lang and spacing are written; separate cs/latin values", () => {
    const r = applyDocDefaultsToXml(styles(DD + NORMAL), {
      fontLatin: "Times New Roman", fontCs: "Simplified Arabic", sizePt: 12, sizeCsPt: 14,
      lang: { val: "fr-FR", bidi: "ar-DZ" }, lineSpacing: 1.5, afterPt: 6,
    });
    expect(r.updated).toBe(true);
    expect(r.xml).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Simplified Arabic"/>');
    expect(r.xml).toContain('<w:sz w:val="24"/>');
    expect(r.xml).toContain('<w:szCs w:val="28"/>');
    expect(r.xml).toContain('<w:lang w:val="fr-FR" w:eastAsia="en-US" w:bidi="ar-DZ"/>');
    expect(r.xml).toContain('<w:spacing w:line="360" w:lineRule="auto" w:after="120"/>');
    expect(r.xml.endsWith(NORMAL + "</w:styles>")).toBe(true); // the rest untouched
  });

  test("creates docDefaults, rPrDefault and pPrDefault as needed (first child of styles)", () => {
    const r = applyDocDefaultsToXml(styles(NORMAL), { fontCs: "Q", lineSpacing: 2, lang: { val: "ar" } });
    expect(r.xml).toContain(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:cs="Q"/><w:lang w:val="ar"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style',
    );
    const onlyP = applyDocDefaultsToXml(styles("<w:docDefaults/>" + NORMAL), { afterPt: 1 });
    expect(onlyP.xml).toContain('<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="20"/></w:pPr></w:pPrDefault></w:docDefaults>');
    const onlyR = applyDocDefaultsToXml(styles("<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"), { sizePt: 10 });
    expect(onlyR.xml).toContain("<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr></w:rPrDefault><w:pPrDefault>");
  });

  test("lang on a docDefaults with no lang adds one; nothing named is a no-op", () => {
    const r = applyDocDefaultsToXml(styles("<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"2\"/></w:rPr></w:rPrDefault></w:docDefaults>"), { lang: { bidi: "ar-SA" } });
    expect(r.xml).toContain('<w:lang w:bidi="ar-SA"/>');
    expect(applyDocDefaultsToXml(styles(NORMAL), {})).toMatchObject({ updated: false });
  });
});

describe("ensureStyle (pure)", () => {
  test("creates a minimal style; never modifies an existing one", () => {
    const r = applyEnsureStyleToXml(styles(NORMAL), "Q", { name: "Q", type: "character", basedOn: "Normal", next: "Normal", uiPriority: 9, qFormat: true });
    expect(r.created).toBe(true);
    expect(r.xml).toBe(
      styles(NORMAL + '<w:style w:type="character" w:styleId="Q"><w:name w:val="Q"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/></w:style>'),
    );
    const again = applyEnsureStyleToXml(r.xml, "Q", { name: "Different", type: "paragraph" });
    expect(again).toEqual({ xml: r.xml, created: false });
    expect(applyEnsureStyleToXml("<not-styles/>", "Z", { name: "Z", type: "paragraph" }).created).toBe(false);
    expect(applyEnsureStyleToXml(styles(""), "Z", { name: "Z", type: "paragraph" }).xml).toContain('<w:style w:type="paragraph" w:styleId="Z"><w:name w:val="Z"/></w:style>');
  });
});

describe("StylesManager methods on samples/2.docx", () => {
  const rulesOf = (e: Mdocxengine) => inspectDocx(e.zip as any).findings.map((f) => f.rule).sort();

  test("creating a style leaves every other byte of styles.xml unchanged; doctor unchanged", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const before = e.zip.readAsText("word/styles.xml")!;
    const rulesBefore = rulesOf(e);
    const r = await e.styles.ensureStyle("RescueBody", { name: "Rescue Body", type: "paragraph", basedOn: "Normal", qFormat: true });
    expect(r.created).toBe(true);
    const after = e.zip.readAsText("word/styles.xml")!;
    const idx = after.lastIndexOf("</w:styles>");
    expect(after.slice(0, before.lastIndexOf("</w:styles>"))).toBe(before.slice(0, before.lastIndexOf("</w:styles>")));
    expect(after.slice(idx)).toBe("</w:styles>");
    expect(after).toContain('w:styleId="RescueBody"');
    expect(rulesOf(e)).toEqual(rulesBefore);
    expect((await e.styles.ensureStyle("RescueBody", { name: "x", type: "paragraph" })).created).toBe(false);
  }, 120_000);

  test("paragraph props + docDefaults written through the manager reload and stay doctor-clean", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const rulesBefore = rulesOf(e);
    await e.styles.ensureStyle("RescueBody", { name: "Rescue Body", type: "paragraph", basedOn: "Normal" });
    const s = await e.styles.setStyleParagraphProps("RescueBody", { lineSpacing: 1.5, jc: "both", bidi: true });
    expect(s).toEqual({ created: false, updated: true });
    expect(await e.styles.setStyleParagraphProps("Nope", { jc: "left" })).toEqual({ created: false, updated: false });
    expect(await e.styles.setDocDefaults({ fontCs: "Simplified Arabic", sizeCsPt: 14 })).toEqual({ updated: true });
    expect(await e.styles.setDocDefaults({})).toEqual({ updated: false });
    const re = await Mdocxengine.loadFromBuffer(e.zip.toBuffer());
    const sx = re.zip.readAsText("word/styles.xml")!;
    expect(pPrOf(sx, "RescueBody")).toBe('<w:pPr><w:bidi/><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>');
    expect(sx).toContain('w:cs="Simplified Arabic"');
    expect(rulesOf(re)).toEqual(rulesBefore);
  }, 120_000);

  test("a manager over a zip with no styles part is a no-op", async () => {
    const m = new StylesManager(new AdmZip());
    expect(await m.setStyleParagraphProps("X", { jc: "left" })).toEqual({ created: false, updated: false });
    expect(await m.ensureStyle("X", { name: "X", type: "paragraph" })).toEqual({ created: false });
    expect(await m.setDocDefaults({ sizePt: 12 })).toEqual({ updated: false });
  });
});
