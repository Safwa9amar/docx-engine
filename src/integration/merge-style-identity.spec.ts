import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, inspectDocx, type DocxZip } from "../index";
import {
  fixture, notesPart, note, noteText, numberingPart, P, refP, stylesDoc, style, styleIds, styleNames, styleOf, withNormal, NORMAL,
  type FixtureOptions,
} from "./docx-fixture";

/**
 * A style's identity in Word is its NAME: two styles sharing a `w:name` are merged and the
 * last one wins, whatever their ids. An id means nothing across documents (Arabic Word hands
 * out `a3`, `a4`, `ac` per file). So a merge reconciles by name, and what the copied blocks
 * point at is rewritten to the style that means the same thing here.
 */

const merged = async (target: FixtureOptions, ...sources: FixtureOptions[]) => {
  const t = await Mdocxengine.loadFromBuffer(fixture(target));
  for (const s of sources) await t.merge.appendDocument(fixture(s));
  const part = (name: string) => t.zip.readAsText(name) ?? "";
  return {
    t,
    styles: part("word/styles.xml"),
    doc: part("word/document.xml"),
    footnotes: part("word/footnotes.xml"),
    numbering: part("word/numbering.xml"),
    fatals: inspectDocx(t.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal").map((f) => f.rule),
  };
};
const pStyles = (xml: string) => [...xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)].map((m) => m[1]);
const pStyleOf = (doc: string, text: string) => new RegExp(`<w:p>(?:(?!</w:p>).)*?<w:pStyle w:val="([^"]+)"(?:(?!</w:p>).)*?${text}`).exec(doc)?.[1];
const heading = (id: string, size = 28, extra = "") => style("paragraph", id, "heading 1", `<w:basedOn w:val="Normal"/><w:qFormat/>${extra}<w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr>`);

describe("a style the target already has by name is never added beside it", () => {
  const FRENCH = stylesDoc(NORMAL.replace("</w:style>", '<w:rPr><w:sz w:val="24"/></w:rPr></w:style>') + heading("Titre1", 28));
  const ARABIC = stylesDoc(
    style("paragraph", "a", "Normal", '<w:rPr><w:sz w:val="40"/></w:rPr>', 'w:default="1"') +
      style("paragraph", "1", "heading 1", '<w:basedOn w:val="a"/><w:rPr><w:b/><w:sz w:val="56"/></w:rPr>'),
  );

  test("another UI's ids bind to this document's styles, whose definitions stay byte for byte", async () => {
    const { styles, doc, fatals } = await merged(
      { body: P("french title", '<w:pStyle w:val="Titre1"/>') + P("french plain"), styles: FRENCH },
      { body: P("second title", '<w:pStyle w:val="1"/>') + P("second explicit", '<w:pStyle w:val="a"/>') + P("second plain"), styles: ARABIC },
    );
    expect(styles).toBe(FRENCH);
    expect(styleIds(styles)).toEqual(["Normal", "Titre1"]);
    expect(pStyleOf(doc, "second title")).toBe("Titre1");
    expect(pStyleOf(doc, "second explicit")).toBe("Normal");
    expect(fatals).toEqual([]);
  });

  test("names compare without regard to case or spacing", async () => {
    const { styles, doc } = await merged(
      { body: P("first", '<w:pStyle w:val="Titre1"/>'), styles: FRENCH },
      { body: P("second", '<w:pStyle w:val="H"/>'), styles: withNormal(style("paragraph", "H", "Heading  1")) },
    );
    expect(styleIds(styles)).toEqual(["Normal", "Titre1"]);
    expect(pStyleOf(doc, "second")).toBe("Titre1");
  });

  test("character and table references follow too, in the notes as well as the body", async () => {
    const target = withNormal(
      style("character", "Appelnotedebasdep", "footnote reference", '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>') +
        style("table", "Grilledutableau", "Table Grid", "<w:tblPr><w:tblBorders/></w:tblPr>"),
    );
    const source = withNormal(
      style("character", "ac", "footnote reference", '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>') +
        style("table", "a3", "Table Grid", "<w:tblPr><w:tblBorders/></w:tblPr>"),
    );
    const { styles, doc, footnotes } = await merged(
      { body: P("first"), styles: target },
      {
        body: `<w:tbl><w:tblPr><w:tblStyle w:val="a3"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr><w:rStyle w:val="ac"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`,
        footnotes: notesPart("footnote", `<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="ac"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t>n</w:t></w:r></w:p></w:footnote>`),
        styles: source,
      },
    );
    expect(styleIds(styles)).toEqual(["Normal", "Appelnotedebasdep", "Grilledutableau"]);
    expect(doc).toContain('<w:tblStyle w:val="Grilledutableau"/>');
    expect(doc).toContain('<w:rStyle w:val="Appelnotedebasdep"/>');
    expect(footnotes).toContain('<w:rStyle w:val="Appelnotedebasdep"/>');
    expect(doc).not.toMatch(/w:val="(?:ac|a3)"/);
  });

  test("a style based on the source's default imports based on this document's", async () => {
    const { styles } = await merged(
      { body: P("first"), styles: withNormal() },
      {
        body: P("second", '<w:pStyle w:val="Mine"/>'),
        styles: stylesDoc(style("paragraph", "a", "Normal", "", 'w:default="1"') + style("paragraph", "Mine", "Mine", '<w:basedOn w:val="a"/><w:next w:val="a"/>', 'w:customStyle="1"')),
      },
    );
    expect(styleIds(styles)).toEqual(["Normal", "Mine"]);
    expect(styleOf(styles, "Mine")).toContain('<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>');
  });

  test("a target whose default goes by another name: the source's Normal binds to it", async () => {
    const { styles, doc } = await merged(
      { body: P("first"), styles: stylesDoc(style("paragraph", "Std", "Standard", "", 'w:default="1"')) },
      { body: P("second", '<w:pStyle w:val="Normal"/>') + P("third", '<w:pStyle w:val="H"/>'), styles: withNormal(style("paragraph", "H", "heading 1", '<w:basedOn w:val="Normal"/>')) },
    );
    expect(styleIds(styles)).toEqual(["Std", "H"]);
    expect(pStyleOf(doc, "second")).toBe("Std");
    expect(styleOf(styles, "H")).toContain('<w:basedOn w:val="Std"/>');
  });

  test("a target with no default of that type does not get the source's: it would restyle the target's text", async () => {
    const { styles, doc, fatals } = await merged(
      { body: P("first"), styles: stylesDoc("") },
      { body: P("second", '<w:pStyle w:val="H"/>') + P("plain", '<w:pStyle w:val="Normal"/>'), styles: stylesDoc(NORMAL.replace("</w:style>", '<w:rPr><w:sz w:val="40"/></w:rPr></w:style>') + style("paragraph", "H", "heading 1", '<w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr>')) },
    );
    expect(styleIds(styles)).toEqual(["H"]);
    expect(styleOf(styles, "H")).not.toContain("w:basedOn");
    expect(pStyleOf(doc, "plain")).toBe("Normal");
    expect(fatals).toEqual([]);
  });
});

describe("an id that means something else here travels under a new id", () => {
  test("the same id on another name: the copied blocks follow the new id, the target's stays", async () => {
    const target = withNormal(style("paragraph", "a3", "header", "<w:rPr><w:i/></w:rPr>"));
    const big = '<w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="56"/></w:rPr>';
    const { styles, doc, fatals } = await merged(
      { body: P("first", '<w:pStyle w:val="a3"/>'), styles: target },
      { body: P("second", '<w:pStyle w:val="a3"/>'), styles: withNormal(style("paragraph", "a3", "Big Para", big, 'w:customStyle="1"')) },
    );
    expect(styleOf(styles, "a3")).toBe(style("paragraph", "a3", "header", "<w:rPr><w:i/></w:rPr>"));
    expect(pStyleOf(doc, "first")).toBe("a3");
    const fresh = pStyleOf(doc, "second")!;
    expect(fresh).not.toBe("a3");
    expect(styleOf(styles, fresh)).toContain(big);
    expect(styleOf(styles, fresh)).toContain('<w:name w:val="Big Para"/>');
    expect(fatals).toEqual([]);
  });

  test("the same id on another type", async () => {
    const { styles, doc } = await merged(
      { body: P("first"), styles: withNormal(style("character", "a3", "Quiet Char", "<w:rPr><w:i/></w:rPr>")) },
      { body: P("second", '<w:pStyle w:val="a3"/>'), styles: withNormal(style("paragraph", "a3", "Big Para", "<w:rPr><w:b/></w:rPr>")) },
    );
    const fresh = pStyleOf(doc, "second")!;
    expect(fresh).not.toBe("a3");
    expect(styleOf(styles, fresh)).toMatch(/^<w:style w:type="paragraph"/);
    expect(styleOf(styles, "a3")).toMatch(/^<w:style w:type="character"/);
  });

  test("the same name on another type stays apart too: the name is taken, so it gets its own", async () => {
    const { styles, doc } = await merged(
      { body: P("first"), styles: withNormal(style("character", "T", "Shared Name", "<w:rPr><w:i/></w:rPr>")) },
      { body: P("second", '<w:pStyle w:val="S"/>'), styles: withNormal(style("paragraph", "S", "Shared Name", "<w:rPr><w:b/></w:rPr>")) },
    );
    expect(pStyleOf(doc, "second")).toBe("S");
    expect(styleOf(styles, "S")).toContain('<w:name w:val="Shared Name (2)"/>');
    expect(styleOf(styles, "T")).toContain('<w:name w:val="Shared Name"/>');
  });

  test("a chain renames as a whole: basedOn, next and link follow, and a link stays two-way", async () => {
    const source = withNormal(
      style("paragraph", "a3", "My Base", "<w:rPr><w:i/></w:rPr>") +
        style("paragraph", "a4", "My Head", '<w:basedOn w:val="a3"/><w:next w:val="a3"/><w:link w:val="a5"/>') +
        style("character", "a5", "My Head Char", '<w:link w:val="a4"/><w:rPr><w:b/></w:rPr>'),
    );
    const target = withNormal(style("paragraph", "a3", "x", "") + style("paragraph", "a4", "y", "") + style("paragraph", "a5", "z", ""));
    const { styles, doc } = await merged({ body: P("first"), styles: target }, { body: P("second", '<w:pStyle w:val="a4"/>'), styles: source });
    const head = pStyleOf(doc, "second")!;
    const def = styleOf(styles, head);
    const base = /<w:basedOn w:val="([^"]+)"/.exec(def)![1];
    const link = /<w:link w:val="([^"]+)"/.exec(def)![1];
    expect(/<w:next w:val="([^"]+)"/.exec(def)![1]).toBe(base);
    expect(styleOf(styles, base)).toContain('<w:name w:val="My Base"/>');
    expect(styleOf(styles, link)).toContain('<w:name w:val="My Head Char"/>');
    expect(styleOf(styles, link)).toContain(`<w:link w:val="${head}"/>`);
    expect(new Set(styleIds(styles)).size).toBe(styleIds(styles).length);
  });

  test("a link whose partner is this document's own is dropped: Word's links are two-way", async () => {
    const target = withNormal(style("character", "TC", "My Head Char", "<w:rPr><w:b/></w:rPr>"));
    const source = withNormal(style("paragraph", "H", "My Head", '<w:link w:val="HC"/>') + style("character", "HC", "My Head Char", '<w:link w:val="H"/><w:rPr><w:b/></w:rPr>'));
    const { styles } = await merged({ body: P("first"), styles: target }, { body: P("second", '<w:pStyle w:val="H"/>'), styles: source });
    expect(styleOf(styles, "H")).not.toContain("w:link");
  });
});

describe("a custom or non-paragraph style that differs is kept beside the target's under its own name", () => {
  const SRC = (inner: string, type = "paragraph") => style(type, "Mine", "My Style", inner, 'w:customStyle="1"');

  test("a custom paragraph style with another look: both looks stay", async () => {
    const { styles, doc } = await merged(
      { body: P("first", '<w:pStyle w:val="Mine"/>'), styles: withNormal(SRC("<w:rPr><w:i/></w:rPr>")) },
      { body: P("second", '<w:pStyle w:val="Mine"/>'), styles: withNormal(SRC("<w:rPr><w:b/></w:rPr>")) },
    );
    const fresh = pStyleOf(doc, "second")!;
    expect(fresh).not.toBe("Mine");
    expect(styleOf(styles, fresh)).toContain("<w:b/>");
    expect(styleOf(styles, "Mine")).toContain("<w:i/>");
    expect(new Set(styleNames(styles)).size).toBe(styleNames(styles).length);
  });

  test("the same look comes to the same style, whatever the housekeeping around it", async () => {
    const a = style("paragraph", "Mine", "My Style", '<w:basedOn w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="00AA"/><w:rPr><w:b/></w:rPr>', 'w:customStyle="1"');
    const b = style("paragraph", "Yours", "My Style", '<w:basedOn w:val="a"/>\n  <w:next w:val="a"/><w:uiPriority w:val="1"/><w:rsid w:val="00BB"/><w:rPr><w:b/></w:rPr>', 'w:customStyle="1"');
    const { styles, doc } = await merged(
      { body: P("first", '<w:pStyle w:val="Mine"/>'), styles: withNormal(a) },
      { body: P("second", '<w:pStyle w:val="Yours"/>'), styles: stylesDoc(style("paragraph", "a", "Normal", "", 'w:default="1"') + b) },
    );
    expect(styleIds(styles)).toEqual(["Normal", "Mine"]);
    expect(pStyleOf(doc, "second")).toBe("Mine");
  });

  test("a third part with the second part's look reuses what the second part brought", async () => {
    const second = { body: P("second", '<w:pStyle w:val="Mine"/>'), styles: withNormal(SRC("<w:rPr><w:b/></w:rPr>")) };
    const { styles, doc } = await merged({ body: P("first", '<w:pStyle w:val="Mine"/>'), styles: withNormal(SRC("<w:rPr><w:i/></w:rPr>")) }, second, { ...second, body: P("third", '<w:pStyle w:val="Mine"/>') });
    expect(styleIds(styles)).toHaveLength(3);
    expect(pStyleOf(doc, "third")).toBe(pStyleOf(doc, "second"));
  });

  test.each(["character", "table"])("a %s style differing under one name stays apart, built-in or not", async (type) => {
    const ref = type === "table" ? "tblStyle" : "rStyle";
    const def = (inner: string) => style(type, "X", "Hyperlink", inner);
    const body = type === "table"
      ? `<w:tbl><w:tblPr><w:tblStyle w:val="X"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>`
      : `<w:p><w:r><w:rPr><w:rStyle w:val="X"/></w:rPr><w:t>second</w:t></w:r></w:p>`;
    const { styles, doc } = await merged({ body: P("first"), styles: withNormal(def("<w:rPr><w:i/></w:rPr>")) }, { body, styles: withNormal(def("<w:rPr><w:b/></w:rPr>")) });
    const fresh = new RegExp(`<w:${ref} w:val="([^"]+)"`).exec(doc)![1];
    expect(fresh).not.toBe("X");
    expect(styleOf(styles, fresh)).toContain("<w:b/>");
    expect(styleOf(styles, "X")).toContain("<w:i/>");
  });

  test("a built-in paragraph style keeps the target's look: Regenerate rebuilds those by role", async () => {
    const { styles, doc } = await merged(
      { body: P("first", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading("Heading1", 28)) },
      { body: P("second", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading("Heading1", 56)) },
    );
    expect(styleIds(styles)).toEqual(["Normal", "Heading1"]);
    expect(styleOf(styles, "Heading1")).toContain('w:val="28"');
    expect(pStyleOf(doc, "second")).toBe("Heading1");
  });

  test("a default style is never kept apart, however it differs", async () => {
    const { styles } = await merged(
      { body: P("first"), styles: withNormal(style("character", "DPF", "Default Paragraph Font", "", 'w:default="1"')) },
      { body: `<w:p><w:r><w:rPr><w:rStyle w:val="DPF"/></w:rPr><w:t>x</w:t></w:r></w:p>`, styles: withNormal(style("character", "DPF", "Default Paragraph Font", "<w:rPr><w:b/></w:rPr>", 'w:default="1"')) },
    );
    expect(styleIds(styles)).toEqual(["Normal", "DPF"]);
  });

  test("a child comes along when the parent it depends on is kept apart", async () => {
    const child = style("paragraph", "Kid", "My Kid", '<w:basedOn w:val="Mine"/><w:rPr><w:u w:val="single"/></w:rPr>', 'w:customStyle="1"');
    const { styles, doc } = await merged(
      { body: P("first"), styles: withNormal(SRC("<w:rPr><w:i/></w:rPr>") + child) },
      { body: P("second", '<w:pStyle w:val="Kid"/>'), styles: withNormal(SRC("<w:rPr><w:b/></w:rPr>") + child) },
    );
    const kid = pStyleOf(doc, "second")!;
    expect(kid).not.toBe("Kid");
    const parent = /<w:basedOn w:val="([^"]+)"/.exec(styleOf(styles, kid))![1];
    expect(parent).not.toBe("Mine");
    expect(styleOf(styles, parent)).toContain("<w:b/>");
  });
});

describe("numbering a style carries survives a merge that keeps the target's style", () => {
  const list = (id: number, text: string, opts: { start?: number; linkAt?: Record<number, string>; numId?: number } = {}) =>
    `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/>` +
    [0, 1].map((l) => `<w:lvl w:ilvl="${l}"><w:start w:val="${opts.start ?? 1}"/><w:numFmt w:val="decimal"/>${opts.linkAt?.[l] ? `<w:pStyle w:val="${opts.linkAt[l]}"/>` : ""}<w:lvlText w:val="${text} %${l + 1} :"/><w:lvlJc w:val="left"/></w:lvl>`).join("") +
    `</w:abstractNum>`;
  const num = (numId: number, abs: number) => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abs}"/></w:num>`;
  const numbered = (id = "Heading1", numId = 1, ilvl?: number) =>
    heading(id, 28, `<w:pPr><w:numPr>${ilvl === undefined ? "" : `<w:ilvl w:val="${ilvl}"/>`}<w:numId w:val="${numId}"/></w:numPr></w:pPr>`);
  const CHAPTER = numberingPart(list(0, "Chapitre", { start: 2, linkAt: { 0: "Heading1" } }) + num(1, 0));
  const newNumId = (doc: string) => /<w:numId w:val="(\d+)"/.exec(doc)?.[1];
  const numPrOf = (doc: string, text: string) => new RegExp(`<w:p>((?:(?!</w:p>).)*?)${text}`).exec(doc)?.[1].match(/<w:numPr>.*?<\/w:numPr>/)?.[0];

  test("the target's style has none: the paragraphs carry the list, the target's own stay unnumbered", async () => {
    const { styles, doc, numbering, fatals } = await merged(
      { body: P("front", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading("Heading1")) },
      { body: P("methods", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: CHAPTER },
    );
    expect(styleOf(styles, "Heading1")).not.toContain("numPr");
    expect(numPrOf(doc, "front")).toBeUndefined();
    const id = newNumId(doc)!;
    expect(numPrOf(doc, "methods")).toBe(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>`);
    expect(numbering).toContain(`w:numId="${id}"`);
    expect(numbering).toContain("Chapitre %1 :");
    // The copied level is no longer the style's: nothing here points back at it.
    expect(numbering).not.toContain("<w:pStyle");
    expect(fatals).toEqual([]);
  });

  test("the list's level comes from the level linked to the style, not from an ilvl", async () => {
    const numberingXml = numberingPart(list(0, "Sec", { linkAt: { 1: "Heading2" } }) + num(1, 0));
    const h2 = style("paragraph", "Heading2", "heading 2", '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>');
    const { doc } = await merged(
      { body: P("front"), styles: withNormal(style("paragraph", "Heading2", "heading 2", "")) },
      { body: P("sub", '<w:pStyle w:val="Heading2"/>'), styles: withNormal(h2), numbering: numberingXml },
    );
    expect(numPrOf(doc, "sub")).toContain('<w:ilvl w:val="1"/>');
  });

  test("a level with the same words: the parts share the target's list, nothing changes", async () => {
    const same = numberingPart(list(0, "Chapitre", { linkAt: { 0: "Heading1" } }) + num(1, 0));
    const { doc, numbering } = await merged(
      { body: P("one", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: same },
      { body: P("two", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: same },
    );
    expect(numPrOf(doc, "two")).toBeUndefined();
    expect(numbering.match(/<w:abstractNum\b/g)).toHaveLength(1);
  });

  test("another label or another start: the part keeps its own list", async () => {
    const mine = numberingPart(list(0, "Partie", { linkAt: { 0: "Heading1" } }) + num(1, 0));
    const { doc, numbering } = await merged(
      { body: P("one", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: mine },
      { body: P("two", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: CHAPTER },
    );
    const ref = numPrOf(doc, "two")!;
    expect(ref).toBeDefined();
    const id = /<w:numId w:val="(\d+)"/.exec(ref)![1];
    expect(id).not.toBe("1");
    expect(numbering).toContain("Partie %1 :");
    expect(numbering).toContain("Chapitre %1 :");
    // The target's list keeps its own link to the target's own style.
    expect(numbering.match(/<w:pStyle/g)).toHaveLength(1);
  });

  test("the target's style is numbered and the part's is not: the part's paragraphs stay unnumbered", async () => {
    const { doc } = await merged(
      { body: P("one", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(numbered()), numbering: CHAPTER },
      { body: P("typed by hand", '<w:pStyle w:val="Heading1"/>'), styles: withNormal(heading("Heading1")) },
    );
    expect(numPrOf(doc, "typed by hand")).toBe('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>');
  });

  test("numbering inherited through the chain counts", async () => {
    const kid = style("paragraph", "Kid", "My Kid", '<w:basedOn w:val="Heading1"/>', 'w:customStyle="1"');
    const { doc, styles } = await merged(
      { body: P("front"), styles: withNormal(heading("Heading1")) },
      { body: P("inherits", '<w:pStyle w:val="Kid"/>'), styles: withNormal(numbered() + kid), numbering: CHAPTER },
    );
    expect(styleOf(styles, "Kid")).toContain('<w:basedOn w:val="Heading1"/>');
    expect(numPrOf(doc, "inherits")).toContain("<w:numId");
  });

  test("a style-level ilvl with the numId on the parent: Word's own Heading2 shape", async () => {
    const numberingXml = numberingPart(list(0, "Sec") + num(1, 0));
    const h2 = style("paragraph", "Heading2", "heading 2", '<w:basedOn w:val="Heading1"/><w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr>');
    const { doc } = await merged(
      { body: P("front"), styles: withNormal(heading("Heading1") + style("paragraph", "Heading2", "heading 2", "")) },
      { body: P("sub", '<w:pStyle w:val="Heading2"/>'), styles: withNormal(numbered() + h2), numbering: numberingXml },
    );
    expect(numPrOf(doc, "sub")).toContain('<w:ilvl w:val="1"/>');
  });

  test("a paragraph with its own numbering is left alone; a note's paragraph is treated like a body one", async () => {
    const { doc, footnotes } = await merged(
      { body: P("front"), styles: withNormal(heading("Heading1")) },
      {
        body: P("own", '<w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>') + `<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
        footnotes: notesPart("footnote", `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>in note</w:t></w:r></w:p></w:footnote>`),
        styles: withNormal(numbered()),
        numbering: CHAPTER,
      },
    );
    expect(numPrOf(doc, "own")).toContain('<w:ilvl w:val="1"/>');
    expect(new RegExp("<w:p>((?:(?!</w:p>).)*?)own").exec(doc)![1].match(/<w:numPr>/g)).toHaveLength(1);
    expect(footnotes).toMatch(/<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>/);
  });

  test("the pPr stays in schema order, and a tracked change's history is not stamped", async () => {
    const { doc } = await merged(
      { body: P("front"), styles: withNormal(heading("Heading1")) },
      {
        body: P("tracked", `<w:pStyle w:val="Heading1"/><w:keepNext/><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:pPrChange>`),
        styles: withNormal(numbered()),
        numbering: CHAPTER,
      },
    );
    expect(doc).toMatch(/<w:pStyle w:val="Heading1"\/><w:keepNext\/><w:numPr>.*?<\/w:numPr><w:jc w:val="center"\/><w:pPrChange [^>]*><w:pPr><w:pStyle w:val="Heading1"\/><\/w:pPr><\/w:pPrChange>/);
  });

  test("what a tracked change records of the past does not decide: only the paragraph's own properties do", async () => {
    const change = (inner: string) => `<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr>${inner}</w:pPr></w:pPrChange>`;
    const { doc } = await merged(
      { body: P("front"), styles: withNormal(heading("Heading1")) },
      {
        body:
          P("was numbered", `<w:pStyle w:val="Heading1"/>${change('<w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>')}`) +
          P("was a heading", change('<w:pStyle w:val="Heading1"/>')),
        styles: withNormal(numbered()),
        numbering: CHAPTER,
      },
    );
    expect(numPrOf(doc, "was numbered")).toMatch(/<w:ilvl w:val="0"\/>/);
    expect(numPrOf(doc, "was a heading")).toBeUndefined();
  });

  test("an imported style keeps its list, and the level's link follows a new id", async () => {
    const source = withNormal(style("paragraph", "a3", "Chap", '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>', 'w:customStyle="1"'));
    const numberingXml = numberingPart(list(0, "Chap", { linkAt: { 0: "a3" } }) + num(1, 0));
    const { styles, doc, numbering } = await merged(
      { body: P("front"), styles: withNormal(style("paragraph", "a3", "other", "")) },
      { body: P("second", '<w:pStyle w:val="a3"/>'), styles: source, numbering: numberingXml },
    );
    const fresh = pStyleOf(doc, "second")!;
    expect(fresh).not.toBe("a3");
    expect(numPrOf(doc, "second")).toBeUndefined();
    expect(numbering).toContain(`<w:pStyle w:val="${fresh}"/>`);
    const numId = /<w:numId w:val="(\d+)"/.exec(styleOf(styles, fresh))![1];
    expect(numbering).toContain(`w:numId="${numId}"`);
  });

  test("a level's link to a style nothing here uses is dropped: it would dangle", async () => {
    const numberingXml = numberingPart(list(0, "Sec", { linkAt: { 0: "Heading1", 1: "NeverUsed" } }) + num(1, 0));
    const { numbering } = await merged({ body: P("first") }, { body: P("second", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'), styles: withNormal(style("paragraph", "NeverUsed", "Never Used", "")), numbering: numberingXml });
    expect(numbering).not.toContain("<w:pStyle");
  });
});

describe("a list defined through a list style", () => {
  const LEVELS = `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="(%1)"/><w:lvlJc w:val="left"/></w:lvl>`;
  const NUMBERING = numberingPart(
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="MyList"/>${LEVELS}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="MyList"/></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`,
  );
  const LIST_STYLE = style("numbering", "MyList", "My List", '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>', 'w:customStyle="1"');
  const ITEM = P("item", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');

  test("the list style and the definition holding its levels come along, wired to each other", async () => {
    const { styles, numbering, doc, fatals } = await merged({ body: P("first") }, { body: ITEM, styles: withNormal(LIST_STYLE), numbering: NUMBERING });
    const used = /<w:numId w:val="(\d+)"/.exec(doc)![1];
    expect(styleIds(styles)).toContain("MyList");
    expect(numbering).toContain(`w:numId="${used}"`);
    expect(numbering.match(/<w:num\b/g)).toHaveLength(2);
    expect(numbering.match(/<w:abstractNum\b/g)).toHaveLength(2);
    expect(numbering).toContain("<w:styleLink");
    expect(numbering).toContain("<w:numStyleLink");
    expect(numbering).toContain('w:val="(%1)"');
    const listNum = /<w:numId w:val="(\d+)"/.exec(styleOf(styles, "MyList"))![1];
    expect(numbering).toContain(`w:numId="${listNum}"`);
    expect(fatals).toEqual([]);
  });

  test("the same list style with other levels stays apart, links rewritten", async () => {
    const other = numberingPart(
      `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="MyList"/>${LEVELS.replace("(%1)", "[%1]")}</w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
    );
    const { styles, numbering } = await merged({ body: P("first", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'), styles: withNormal(LIST_STYLE), numbering: other }, { body: ITEM, styles: withNormal(LIST_STYLE), numbering: NUMBERING });
    const links = [...numbering.matchAll(/<w:(?:numStyleLink|styleLink) w:val="([^"]+)"/g)].map((m) => m[1]);
    expect(links).toContain("MyList");
    const fresh = links.find((l) => l !== "MyList")!;
    expect(fresh).toBeDefined();
    expect(styleOf(styles, fresh)).toContain("<w:numPr>");
    expect(links.filter((l) => l === fresh)).toHaveLength(2);
  });

  test("the same list style with the same levels binds to the target's", async () => {
    const { styles, numbering } = await merged({ body: P("first", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'), styles: withNormal(LIST_STYLE), numbering: NUMBERING }, { body: ITEM, styles: withNormal(LIST_STYLE), numbering: NUMBERING });
    expect(styleIds(styles)).toEqual(["Normal", "MyList"]);
    expect([...numbering.matchAll(/<w:(?:numStyleLink|styleLink) w:val="([^"]+)"/g)].every((m) => m[1] === "MyList")).toBe(true);
  });
});

describe("a target with no styles part", () => {
  test("gets one, registered, holding what the copied text uses and no default", async () => {
    const source = withNormal(style("paragraph", "Heading1", "heading 1", '<w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="48"/></w:rPr>'));
    const { t, styles, fatals } = await merged({ body: P("first"), styles: null }, { body: P("title", '<w:pStyle w:val="Heading1"/>'), styles: source });
    expect(styleIds(styles)).toEqual(["Heading1"]);
    expect(styles).not.toContain("w:docDefaults");
    expect(t.zip.readAsText("[Content_Types].xml")).toContain('PartName="/word/styles.xml"');
    expect(t.zip.readAsText("word/_rels/document.xml.rels")).toContain('Target="styles.xml"');
    expect(fatals).toEqual([]);
    expect(new AdmZip(t.zip.toBuffer()).readAsText("word/styles.xml")).toBe(styles);
  });

  test("a source with no styles to bring creates nothing", async () => {
    const { t } = await merged({ body: P("first"), styles: null }, { body: P("second"), styles: withNormal() });
    expect(t.zip.getEntry("word/styles.xml")).toBeNull();
  });
});

describe("a caller's styleMap outranks all of it", () => {
  test("mapped ids are not imported, stamped or renamed, and a chain through one follows the map", async () => {
    const source = withNormal(
      style("paragraph", "Titre1", "heading 1", '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>') +
        style("paragraph", "Kid", "My Kid", '<w:basedOn w:val="Titre1"/>', 'w:customStyle="1"'),
    );
    const numberingXml = numberingPart(`<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`);
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("first"), styles: withNormal(style("paragraph", "Heading1", "heading 1", "")) }));
    await t.merge.appendDocument(fixture({ body: P("mapped", '<w:pStyle w:val="Titre1"/>') + P("kid", '<w:pStyle w:val="Kid"/>'), styles: source, numbering: numberingXml }), { styleMap: { Titre1: "Heading1" } });
    const styles = t.zip.readAsText("word/styles.xml")!;
    const doc = t.zip.readAsText("word/document.xml")!;
    expect(styleIds(styles)).toEqual(["Normal", "Heading1", "Kid"]);
    expect(styleOf(styles, "Kid")).toContain('<w:basedOn w:val="Heading1"/>');
    expect(pStyles(doc)).toEqual(["Heading1", "Kid"]);
    expect(doc).not.toContain("<w:numPr>");
  });
});
