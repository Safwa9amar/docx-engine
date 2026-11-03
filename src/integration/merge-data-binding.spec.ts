import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine } from "../index";
import { fixture, P, note, notesPart, refP } from "./docx-fixture";

/**
 * A content control bound to a store takes its text from that store when Word opens the file. The
 * store (core properties, a custom XML part) stays behind in the part's package, so a later part's
 * control would print part 1's title, or nothing. It keeps its own words by losing the binding.
 */

const TITLE_BINDING = `w:prefixMappings="xmlns:ns0='http://purl.org/dc/elements/1.1/' xmlns:ns1='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' " w:xpath="/ns1:coreProperties[1]/ns0:title[1]" w:storeItemID="{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"`;
const bound = (t: string, binding = `<w:dataBinding ${TITLE_BINDING}/>`) =>
  `<w:sdt><w:sdtPr><w:alias w:val="Titre"/><w:id w:val="1234"/>${binding}<w:text/></w:sdtPr><w:sdtContent><w:r><w:t>${t}</w:t></w:r></w:sdtContent></w:sdt>`;
const boundP = (t: string, binding?: string) => `<w:p>${bound(t, binding)}</w:p>`;

const merge = async (a: Buffer, b: Buffer) => {
  const e = await Mdocxengine.loadFromBuffer(a);
  const r = await e.merge.appendDocument(b, {});
  return { result: r, zip: new AdmZip(e.zip.toBuffer()) };
};
const bindings = (xml: string) => (xml.match(/<w(?:15)?:dataBinding\b/g) ?? []).length;

describe("a later part's bound control keeps its own words", () => {
  test("the binding goes, the control and its words stay", async () => {
    const { result, zip } = await merge(fixture({ body: P("Un.") }), fixture({ body: boundP("Le Titre Du Chapitre Deux") }));
    const xml = zip.readAsText("word/document.xml");
    expect(bindings(xml)).toBe(0);
    expect(xml).toContain("Le Titre Du Chapitre Deux");
    expect(xml).toContain('<w:alias w:val="Titre"/>');
    expect(xml).toContain("<w:sdtContent>");
    expect(result.bindingsDropped).toBe(1);
  });

  test("the first part's keep theirs: its store is still there", async () => {
    const { result, zip } = await merge(fixture({ body: boundP("Titre Un") }), fixture({ body: P("Deux.") }));
    expect(bindings(zip.readAsText("word/document.xml"))).toBe(1);
    expect(result.bindingsDropped).toBe(0);
  });

  test("the later part's go even when the first keeps its own", async () => {
    const { zip } = await merge(fixture({ body: boundP("Titre Un") }), fixture({ body: boundP("Titre Deux") }));
    expect(bindings(zip.readAsText("word/document.xml"))).toBe(1);
  });

  test("in a table, a block-level control, a repeating section's binding, and with a > in the path", async () => {
    const cell = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${bound("Cellule")}</w:p></w:tc></w:tr></w:tbl>`;
    const block = `<w:sdt><w:sdtPr><w:dataBinding ${TITLE_BINDING}></w:dataBinding></w:sdtPr><w:sdtContent>${P("Bloc")}</w:sdtContent></w:sdt>`;
    const w15 = boundP("Quinze", `<w15:dataBinding xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w:prefixMappings="" w:xpath="/a[1]" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/>`);
    const arrow = boundP("Fleche", `<w:dataBinding w:prefixMappings="" w:xpath="/a[b>1]/c" w:storeItemID="{00000000-0000-0000-0000-000000000002}"/>`);
    const { result, zip } = await merge(fixture({ body: P("Un.") }), fixture({ body: cell + block + w15 + arrow + P("Fin") }));
    const xml = zip.readAsText("word/document.xml");
    expect(bindings(xml)).toBe(0);
    expect(result.bindingsDropped).toBe(4);
    for (const w of ["Cellule", "Bloc", "Quinze", "Fleche", "Fin"]) expect(xml).toContain(w);
    expect(xml).not.toContain("/a[b>1]");
  });

  test("in a note that comes across", async () => {
    const part = fixture({ body: refP("footnote", 1), footnotes: notesPart("footnote", note("footnote", 1, boundP("Dans la note"))) });
    const { result, zip } = await merge(fixture({ body: P("Un.") }), part);
    const notes = zip.readAsText("word/footnotes.xml");
    expect(bindings(notes)).toBe(0);
    expect(notes).toContain("Dans la note");
    expect(result.bindingsDropped).toBe(1);
  });
});
