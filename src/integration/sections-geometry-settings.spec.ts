import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, Doc, inspectDocx, applySettingToXml, SettingsManager } from "../index";
import {
  upsertSectPrGeometry,
  removeSectPrTitlePg,
  removeSectPrHeaderFooterRefs,
  extractSectPrGeometry,
} from "../core/files/body/sectPr";

const rulesOf = (e: Mdocxengine) => inspectDocx(e.zip as any).findings.map((f) => f.rule).sort();
const doc = (e: Mdocxengine) => e.zip.readAsText("word/document.xml")!;
const sectPrs = (xml: string) => [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]);
const bodyGeometry = (xml: string) => extractSectPrGeometry(sectPrs(xml).at(-1)!);
const A4_PGSZ = '<w:pgSz w:w="11907" w:h="16839" w:code="9"/>';
const A4_PGMAR = '<w:pgMar w:top="1418" w:right="1275" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>';

describe("upsertSectPrGeometry (pure)", () => {
  test("changes only the named attributes and keeps the rest (w:code, header/footer)", () => {
    const s = `<w:sectPr><w:type w:val="nextPage"/>${A4_PGSZ}${A4_PGMAR}<w:cols w:space="708"/></w:sectPr>`;
    const out = upsertSectPrGeometry(s, { widthTw: 12240, orient: "landscape", leftTw: 2000, gutterTw: 5 });
    expect(out).toBe(
      '<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="16839" w:code="9" w:orient="landscape"/>' +
        '<w:pgMar w:top="1418" w:right="1275" w:bottom="1418" w:left="2000" w:header="709" w:footer="709" w:gutter="5"/><w:cols w:space="708"/></w:sectPr>',
    );
    expect(upsertSectPrGeometry(s, { widthTw: 11907 })).toBe(s); // no change → identical
    expect(upsertSectPrGeometry(s, {})).toBe(s);
  });

  test("creates pgSz/pgMar in schema order; a new pgMar is completed with required attributes", () => {
    const s = '<w:sectPr><w:type w:val="continuous"/><w:cols w:space="1"/></w:sectPr>';
    const out = upsertSectPrGeometry(s, { widthTw: 11906, heightTw: 16838, topTw: 1000 });
    expect(out).toBe(
      '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1000" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="1"/></w:sectPr>',
    );
    expect(upsertSectPrGeometry("<w:sectPr/>", { widthTw: 1, heightTw: 2, orient: "portrait" })).toBe(
      '<w:sectPr><w:pgSz w:w="1" w:h="2" w:orient="portrait"/></w:sectPr>',
    );
  });

  test("a missing pgSz needs both dimensions; bad numbers throw", () => {
    expect(() => upsertSectPrGeometry("<w:sectPr/>", { widthTw: 1 })).toThrow(/pgSz/);
    expect(() => upsertSectPrGeometry("<w:sectPr/>", { topTw: -1 })).toThrow(/topTw/);
    expect(() => upsertSectPrGeometry("<w:sectPr/>", { headerTw: NaN })).toThrow(/headerTw/);
  });

  test("not a sectPr or unclosed: returned untouched", () => {
    expect(upsertSectPrGeometry("<w:p/>", { topTw: 1 })).toBe("<w:p/>");
    expect(upsertSectPrGeometry("<w:sectPr><w:cols/>", { topTw: 1 })).toBe("<w:sectPr><w:cols/>");
  });
});

describe("titlePg and header/footer references (pure)", () => {
  test("removeSectPrTitlePg / removeSectPrHeaderFooterRefs", () => {
    const s = '<w:sectPr><w:headerReference w:type="default" r:id="a"/><w:footerReference w:type="first" r:id="b"/><w:type w:val="nextPage"/><w:titlePg/></w:sectPr>';
    expect(removeSectPrTitlePg(s)).toBe(s.replace("<w:titlePg/>", ""));
    expect(removeSectPrHeaderFooterRefs(s)).toBe('<w:sectPr><w:type w:val="nextPage"/><w:titlePg/></w:sectPr>');
    expect(removeSectPrTitlePg("<w:sectPr/>")).toBe("<w:sectPr/>");
    expect(removeSectPrTitlePg("<w:sectPr><w:titlePg w:val=\"1\"></w:titlePg></w:sectPr>")).toBe("<w:sectPr></w:sectPr>");
  });

  test("extractSectPrGeometry returns pgSz + pgMar, or nothing", () => {
    expect(extractSectPrGeometry(`<w:sectPr>${A4_PGSZ}${A4_PGMAR}</w:sectPr>`)).toBe(A4_PGSZ + A4_PGMAR);
    expect(extractSectPrGeometry("<w:sectPr/>")).toBe("");
  });
});

describe("section breaks keep the page size", () => {
  test("a break inserted on samples/2.docx has the body's A4 pgSz + pgMar", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const before = doc(e);
    const count = sectPrs(before).length;
    await e.sections.addSectionBreak(3, "nextPage");
    const added = sectPrs(doc(e)).length;
    expect(added).toBeGreaterThanOrEqual(count); // paragraph 3 may already hold one (replaced)
    const geo = bodyGeometry(before);
    expect(geo).toMatch(/<w:pgSz w:w="1190\d" w:h="1683\d"/); // A4
    expect(sectPrs(doc(e))).toContain(`<w:sectPr><w:type w:val="nextPage"/>${geo}</w:sectPr>`);
  });

  test("copyBodyGeometry:false keeps the bare form; a document with no body geometry stays bare", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    await e.sections.addSectionBreak(2, "continuous", { copyBodyGeometry: false });
    expect(doc(e)).toContain('<w:sectPr><w:type w:val="continuous"/></w:sectPr>');
    const bare = await Mdocxengine.loadFromFile("samples/example.docx");
    const x = doc(bare).replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>\s*(?=<\/w:body>)/, "");
    bare.zip.addFile("word/document.xml", Buffer.from(x, "utf8"));
    await bare.sections.addSectionBreak(0, "nextPage");
    expect(doc(bare)).toContain('<w:sectPr><w:type w:val="nextPage"/></w:sectPr>');
  });

  test("Doc.startOnNewPage and Doc.setSectionGeometry go through the same byte-safe path", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const d = Doc.from(e);
    const blocks = await e.document.getBlocks();
    const at = blocks.findIndex((b, i) => i > 5 && b.kind === "paragraph");
    await d.startOnNewPage(at);
    expect(sectPrs(doc(e))).toContain(`<w:sectPr><w:type w:val="nextPage"/>${bodyGeometry(doc(e))}</w:sectPr>`);
    const r = await d.setSectionGeometry(at + 1, { topTw: 777 });
    expect(r.totalSections).toBeGreaterThan(1);
    expect(doc(e)).toContain('w:top="777"');
  });
});

describe("setSectionGeometry on samples/2.docx", () => {
  test("round-trips on the final section; every other byte unchanged; doctor unchanged", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const rulesBefore = rulesOf(e);
    const before = doc(e);
    const total = (await e.sections.getSections()).length;
    const body = sectPrs(before).at(-1)!;
    const pgSz = /<w:pgSz\b[^>]*\/>/.exec(body)![0];
    const pgMar = /<w:pgMar\b[^>]*\/>/.exec(body)![0];
    await e.sections.setSectionGeometry(total - 1, { widthTw: 12000, heightTw: 17000, leftTw: 1500, headerTw: 600 });
    const after = doc(e);
    const newBody = body
      .replace(pgSz, pgSz.replace(/w:w="\d+"/, 'w:w="12000"').replace(/w:h="\d+"/, 'w:h="17000"'))
      .replace(pgMar, pgMar.replace(/w:left="\d+"/, 'w:left="1500"').replace(/w:header="\d+"/, 'w:header="600"'));
    expect(after).toBe(before.slice(0, before.lastIndexOf(body)) + newBody + before.slice(before.lastIndexOf(body) + body.length));
    const s = (await e.sections.getSections()).at(-1)!;
    expect(s.pageSize).toMatchObject({ width: 12000, height: 17000 });
    expect(s.margins).toMatchObject({ left: 1500, header: 600 });
    expect(rulesOf(e)).toEqual(rulesBefore);
    await expect(e.sections.setSectionGeometry(99, { topTw: 1 })).rejects.toThrow(/out of range/);
  }, 120_000);

  test("an intermediate section is edited in place", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const before = doc(e);
    const target = (await e.sections.getSections())[0];
    expect(target.isFinal).toBe(false);
    await e.sections.setSectionGeometry(0, { topTw: 1234 });
    const first = sectPrs(doc(e))[0];
    expect(first).toContain('w:top="1234"');
    expect(doc(e).length).toBe(before.length - (sectPrs(before)[0].length - first.length) );
  });
});

describe("removeTitlePg / removeSectionHeaderFooterRefs on samples/2.docx", () => {
  test("only the named things go", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    // give section 2 a titlePg first (the sample has none), by string surgery
    const first = sectPrs(doc(e))[2];
    expect(first).not.toContain("titlePg");
    const withTitle = first.replace(/(<w:cols\b|<w:docGrid\b)/, "<w:titlePg/>$1");
    expect(withTitle).not.toBe(first);
    e.zip.addFile("word/document.xml", Buffer.from(doc(e).replace(first, withTitle), "utf8"));
    const before = doc(e);
    await e.sections.removeTitlePg(2);
    expect(doc(e)).toBe(before.replace("<w:titlePg/>", ""));

    const ridx = sectPrs(doc(e)).findIndex((s) => /<w:footerReference/.test(s));
    const prev = sectPrs(doc(e))[ridx];
    await e.sections.removeSectionHeaderFooterRefs(ridx);
    const now = sectPrs(doc(e))[ridx];
    expect(now).toBe(prev.replace(/<w:(?:header|footer)Reference\b[^>]*\/>/g, ""));
    expect(now).not.toMatch(/Reference/);
  });
});

describe("settings writer", () => {
  const S = (inner: string) =>
    `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:settings>`;

  test("added at its schema position, changed in place, removed", () => {
    const base = S('<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat/>');
    const a = applySettingToXml(base, "mirrorMargins", {});
    expect(a).toEqual({ xml: S('<w:zoom w:percent="100"/><w:mirrorMargins/><w:defaultTabStop w:val="720"/><w:compat/>'), changed: true });
    const b = applySettingToXml(a.xml, "updateFields", { val: "true" });
    expect(b.xml).toBe(S('<w:zoom w:percent="100"/><w:mirrorMargins/><w:defaultTabStop w:val="720"/><w:updateFields w:val="true"/><w:compat/>'));
    const c = applySettingToXml(b.xml, "updateFields", { "w:val": "false" });
    expect(c.xml).toContain('<w:updateFields w:val="false"/>');
    expect(applySettingToXml(c.xml, "updateFields", { "w:val": "false" })).toEqual({ xml: c.xml, changed: false });
    const d = applySettingToXml(c.xml, "updateFields", null);
    expect(d.xml).toBe(a.xml);
    expect(applySettingToXml(d.xml, "updateFields", null)).toEqual({ xml: d.xml, changed: false });
    expect(applySettingToXml(S(""), "evenAndOddHeaders", {}).xml).toBe(S("<w:evenAndOddHeaders/>"));
    expect(applySettingToXml(S('<w:evenAndOddHeaders></w:evenAndOddHeaders>'), "evenAndOddHeaders", null).xml).toBe(S(""));
  });

  test("a self-closing root is expanded; bad input throws", () => {
    const sc = '<w:settings xmlns:w="x"/>';
    expect(applySettingToXml(sc, "mirrorMargins", {}).xml).toBe('<w:settings xmlns:w="x"><w:mirrorMargins/></w:settings>');
    expect(() => applySettingToXml(S(""), "notASetting", {})).toThrow(/not a CT_Settings/);
    expect(() => applySettingToXml("<x/>", "zoom", {})).toThrow(/no <w:settings>/);
  });

  test("on samples/2.docx: add, change, remove; other bytes unchanged; doctor unchanged", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const rulesBefore = rulesOf(e);
    const before = e.zip.readAsText("word/settings.xml")!;
    expect(await e.settings.setSetting("updateFields", { val: "true" })).toEqual({ changed: true });
    const added = e.zip.readAsText("word/settings.xml")!;
    expect(added.length).toBe(before.length + '<w:updateFields w:val="true"/>'.length);
    expect(added.replace('<w:updateFields w:val="true"/>', "")).toBe(before);
    expect(await e.settings.setSetting("updateFields", { val: "true" })).toEqual({ changed: false });
    expect(await e.settings.setSetting("updateFields", null)).toEqual({ changed: true });
    expect(e.zip.readAsText("word/settings.xml")).toBe(before);
    expect(rulesOf(e)).toEqual(rulesBefore);
  }, 120_000);

  test("a package with no settings part gets one (content type + relationship)", async () => {
    const zip = new AdmZip();
    zip.addFile("[Content_Types].xml", Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>'));
    zip.addFile("word/_rels/document.xml.rels", Buffer.from('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'));
    const m = new SettingsManager(zip);
    expect(await m.setSetting("mirrorMargins", null)).toEqual({ changed: false });
    expect(await m.setSetting("mirrorMargins", {})).toEqual({ changed: true });
    expect(zip.readAsText("word/settings.xml")).toContain("<w:mirrorMargins/>");
    expect(zip.readAsText("[Content_Types].xml")).toContain("/word/settings.xml");
    expect(zip.readAsText("word/_rels/document.xml.rels")).toContain('Target="settings.xml"');
  });
});
