import { describe, test, expect } from "vitest";
import { declareUsedPrefixes, undeclaredPrefixes } from "./namespaces";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

describe("undeclaredPrefixes", () => {
  test("a root that declares only w and a reference that writes r:id", () => {
    const xml = `<?xml version="1.0"?><w:document ${W}><w:body><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`;
    expect(undeclaredPrefixes(xml)).toEqual([{ prefix: "r", count: 1 }]);
  });

  test("words that look like prefixes are text, not markup", () => {
    const xml = `<w:document ${W}><w:body><w:p><w:r><w:t>r:id foo:bar &lt;a:b/&gt;</w:t></w:r></w:p></w:body></w:document>`;
    expect(undeclaredPrefixes(xml)).toEqual([]);
  });

  test("a declaration on the element itself or an ancestor counts; a sibling's does not", () => {
    const local = `<w:document ${W}><w:body><w:p><a:blip xmlns:a="urn:a" r:embed="rId1" xmlns:r="${R}"/></w:p></w:body></w:document>`;
    expect(undeclaredPrefixes(local)).toEqual([]);
    const sibling = `<w:document ${W}><w:body><w:p xmlns:a="urn:a"><a:x/></w:p><w:p><a:x/></w:p></w:body></w:document>`;
    expect(undeclaredPrefixes(sibling)).toEqual([{ prefix: "a", count: 1 }]);
  });

  test("xml: needs no declaration", () => {
    expect(undeclaredPrefixes(`<w:document ${W}><w:t xml:space="preserve">x</w:t></w:document>`)).toEqual([]);
  });

  test("a prefix mc:Ignorable names must be declared too", () => {
    const mc = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    expect(undeclaredPrefixes(`<w:document ${W} ${mc} mc:Ignorable="w14 wp14"><w:body/></w:document>`).map((u) => u.prefix).sort()).toEqual(["w14", "wp14"]);
  });
});

describe("declareUsedPrefixes", () => {
  test("adds the known namespace to the root, and nothing else moves", () => {
    const xml = `<?xml version="1.0"?><w:document ${W}><w:body><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`;
    const out = declareUsedPrefixes(xml);
    expect(out).toBe(xml.replace(`${W}>`, `${W} xmlns:r="${R}">`));
    expect(undeclaredPrefixes(out)).toEqual([]);
  });

  test("a file with nothing undeclared comes back as the same string", () => {
    const xml = `<w:document ${W} xmlns:r="${R}"><w:body><w:p r:id="rId1"/></w:body></w:document>`;
    expect(declareUsedPrefixes(xml)).toBe(xml);
  });

  test("a prefix it does not know is left for the doctor to report", () => {
    const xml = `<w:document ${W}><foo:x/></w:document>`;
    expect(declareUsedPrefixes(xml)).toBe(xml);
    expect(undeclaredPrefixes(xml)).toEqual([{ prefix: "foo", count: 1 }]);
  });
});
