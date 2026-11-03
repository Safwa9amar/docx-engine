import { describe, it, expect, vi } from "vitest";
import AdmZip from "adm-zip";
import { ZipManager } from "@/utils/ZipManager";
import {
  inspectDocx,
  firstXmlError,
  clearFalseDataDescriptors,
  checkDocxBuffer,
  repairDocxBuffer,
  inspectZipFlags,
  type DocxZip,
} from "./index";

// ── Fixture builders ─────────────────────────────────────────────────────────
//
// A minimal but REAL package: the doctor's package-integrity rules run on every
// inspection, so a fixture missing [Content_Types].xml or _rels/.rels would drown
// the assertion under unrelated fatals.

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const CONTENT_TYPES = (extra = "") =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  extra +
  `</Types>`;

const ROOT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

const doc = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`;

const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
const P = (t = "hello") => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const TBL = (cell = "x") =>
  `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p><w:r><w:t>${cell}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

/** Build a package. `parts` are added on top of the minimum viable trio. */
function pkg(documentBody: string, parts: Record<string, string> = {}, contentTypeExtra = ""): AdmZip & DocxZip {
  const zip = new AdmZip();
  zip.addFile("[Content_Types].xml", Buffer.from(CONTENT_TYPES(contentTypeExtra), "utf8"));
  zip.addFile("_rels/.rels", Buffer.from(ROOT_RELS, "utf8"));
  zip.addFile("word/document.xml", Buffer.from(doc(documentBody), "utf8"));
  for (const [name, xml] of Object.entries(parts)) zip.addFile(name, Buffer.from(xml, "utf8"));
  return zip as AdmZip & DocxZip;
}

const rulesOf = (z: DocxZip, opts = {}) => inspectDocx(z, opts).findings.map((f) => f.rule);
const findRule = (z: DocxZip, rule: string, opts = {}) =>
  inspectDocx(z, opts).findings.find((f) => f.rule === rule);

// ─────────────────────────────────────────────────────────────────────────────

describe("docx-doctor: a healthy package", () => {
  it("reports nothing and rewrites nothing", () => {
    const report = inspectDocx(pkg(P() + P("world") + SECT));
    expect(report.findings).toEqual([]);
    expect(report.ok).toBe(true);
    expect(report.repairedParts).toEqual([]);
  });

  it("leaves a healthy package byte-identical even when asked to fix", () => {
    const zip = pkg(P() + TBL() + P() + SECT);
    const before = zip.readAsText("word/document.xml");
    const report = inspectDocx(zip, { fix: true });
    expect(report.repairedParts).toEqual([]);
    expect(zip.readAsText("word/document.xml")).toBe(before);
  });
});

describe("docx-doctor: schema sequence", () => {
  // The defect that made Word refuse a real Arabic thesis: CT_TblPrBase puts
  // w:bidiVisual BEFORE w:tblW, and we emitted them the other way round.
  const BAD_TBL =
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:bidiVisual/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;

  it("flags a tblPr with bidiVisual after tblW as FATAL", () => {
    const f = findRule(pkg(BAD_TBL + P() + SECT), "sequence.out-of-order");
    expect(f).toMatchObject({ severity: "fatal", part: "word/document.xml", detail: "w:tblPr", fixable: true });
  });

  it("moves bidiVisual in front of tblW when fixing", () => {
    const zip = pkg(BAD_TBL + P() + SECT);
    expect(inspectDocx(zip, { fix: true }).repairedParts).toContain("word/document.xml");
    const tblPr = /<w:tblPr>[\s\S]*?<\/w:tblPr>/.exec(zip.readAsText("word/document.xml"))![0];
    expect(tblPr.indexOf("<w:bidiVisual")).toBeLessThan(tblPr.indexOf("<w:tblW"));
  });

  it("is idempotent — a second pass finds nothing to do", () => {
    const zip = pkg(BAD_TBL + P() + SECT);
    inspectDocx(zip, { fix: true });
    expect(inspectDocx(zip, { fix: true }).repairedParts).toEqual([]);
  });

  it("reorders pPr children but only calls it a warning (Word tolerates it)", () => {
    // LibreOffice really does emit <w:pPr><w:sectPr/><w:pStyle/></w:pPr>.
    const body = `<w:p><w:pPr>${SECT}<w:pStyle w:val="Normal"/></w:pPr></w:p>${P()}${SECT}`;
    const zip = pkg(body);
    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({ severity: "warning" });
    inspectDocx(zip, { fix: true });
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(zip.readAsText("word/document.xml"))![0];
    expect(pPr.indexOf("<w:pStyle")).toBeLessThan(pPr.indexOf("<w:sectPr"));
  });

  // Regression: an earlier scanner walked tag-by-tag instead of subtree-by-subtree
  // and hoisted this nested w:rPr out of its w:tblStylePr and up into the w:style.
  it("never hoists a nested element out of the child it belongs to", () => {
    const styles =
      `<w:styles ${NS}><w:style w:type="table" w:styleId="T"><w:name w:val="T"/>` +
      `<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tblPr/></w:tblStylePr>` +
      `</w:style></w:styles>`;
    const zip = pkg(P() + SECT, { "word/styles.xml": styles });
    const report = inspectDocx(zip, { fix: true });
    expect(report.findings.filter((f) => f.part === "word/styles.xml")).toEqual([]);
    expect(zip.readAsText("word/styles.xml")).toBe(styles);
  });

  it("does NOT touch w:rPr — EG_RPrBase is a repeated choice, so order is free", () => {
    // Word writes rStyle first; LibreOffice writes it after webHidden. Both open.
    const body = `<w:p><w:r><w:rPr><w:webHidden/><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>a</w:t></w:r></w:p>${SECT}`;
    const zip = pkg(body);
    const before = zip.readAsText("word/document.xml");
    expect(inspectDocx(zip, { fix: true }).findings).toEqual([]);
    expect(zip.readAsText("word/document.xml")).toBe(before);
  });

  // CT_Row is `tblPrEx?, trPr?, EG_ContentCellContent*`. The FIRST_CHILD rule
  // hoisted w:trPr to index 0 unconditionally, so the doctor took rows Word had
  // written correctly and inverted them — its repair WAS the corruption. Found on
  // 2026-08-15 in a combined thesis Word refused to open.
  const rowTbl = (row: string) =>
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
    `<w:tr>${row}<w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
  const TBL_PR_EX = `<w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPrEx>`;
  const TR_PR = `<w:trPr><w:trHeight w:val="448"/></w:trPr>`;
  const rowKids = (z: DocxZip) =>
    [...(/<w:tr>[\s\S]*?<\/w:tr>/.exec(z.readAsText("word/document.xml"))![0])
      .matchAll(/<(w:tblPrEx|w:trPr|w:tc)[ >]/g)].map((m) => m[1]);

  it("leaves a row alone when tblPrEx legally precedes trPr", () => {
    const zip = pkg(rowTbl(TBL_PR_EX + TR_PR) + P() + SECT);
    const before = zip.readAsText("word/document.xml");
    inspectDocx(zip, { fix: true });
    expect(rowKids(zip)).toEqual(["w:tblPrEx", "w:trPr", "w:tc"]);
    expect(zip.readAsText("word/document.xml")).toBe(before);
  });

  it("puts an inverted trPr/tblPrEx pair back in schema order", () => {
    const zip = pkg(rowTbl(TR_PR + TBL_PR_EX) + P() + SECT);
    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({ severity: "fatal", detail: "w:tblPrEx" });
    inspectDocx(zip, { fix: true });
    expect(rowKids(zip)).toEqual(["w:tblPrEx", "w:trPr", "w:tc"]);
  });

  it("still hoists a trPr that has fallen behind the cells", () => {
    const zip = pkg(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc>${TR_PR}</w:tr></w:tbl>` + P() + SECT,
    );
    inspectDocx(zip, { fix: true });
    expect(rowKids(zip)).toEqual(["w:trPr", "w:tc"]);
  });

  // CT_Lvl seats w:lvlJc LATE — just before w:pPr. Exporters emit it right after
  // w:start, which puts w:numFmt and w:lvlText out of sequence behind it.
  it("reorders a numbering level whose lvlJc came early (warning — the seed asset has 9 and Word opens it)", () => {
    const numbering =
      `<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>` +
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:lvlJc w:val="left"/><w:numFmt w:val="bullet"/>` +
      `<w:lvlText w:val="●"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
      `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const zip = pkg(P() + SECT, { "word/numbering.xml": numbering });
    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({
      severity: "warning", part: "word/numbering.xml", detail: "w:lvl",
    });
    inspectDocx(zip, { fix: true });
    const lvl = /<w:lvl [\s\S]*?<\/w:lvl>/.exec(zip.readAsText("word/numbering.xml"))![0];
    expect([...lvl.matchAll(/<(w:start|w:numFmt|w:lvlText|w:lvlJc|w:pPr)[ >]/g)].map((m) => m[1]))
      .toEqual(["w:start", "w:numFmt", "w:lvlText", "w:lvlJc", "w:pPr"]);
  });

  // The seed thesis-base.docx emits displayBackgroundShape after w:compat, which
  // CT_Settings forbids. Word tolerates it, so it is only a warning — but it made
  // every thesis the product ever produced fail schema validation.
  it("moves displayBackgroundShape ahead of compat, as a warning", () => {
    const settings =
      `<w:settings ${NS}><w:compat><w:compatSetting w:val="15" w:uri="u" w:name="compatibilityMode"/></w:compat>` +
      `<w:displayBackgroundShape/></w:settings>`;
    const zip = pkg(P() + SECT, { "word/settings.xml": settings });
    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({
      severity: "warning", part: "word/settings.xml", detail: "w:settings",
    });
    inspectDocx(zip, { fix: true });
    const out = zip.readAsText("word/settings.xml");
    expect(out.indexOf("<w:displayBackgroundShape")).toBeLessThan(out.indexOf("<w:compat>"));
  });

  it("keeps headerReference/footerReference in their authored order", () => {
    // EG_HdrFtrReferences is a repeated choice inside an ordered sequence, and
    // Word genuinely emits footerReference before headerReference.
    const sect =
      `<w:sectPr><w:footerReference w:type="default" r:id="rId9"/>` +
      `<w:headerReference w:type="default" r:id="rId8"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;
    const zip = pkg(P() + sect);
    const before = zip.readAsText("word/document.xml");
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toBe(before);
  });
});

describe("docx-doctor: story shape", () => {
  it("flags and separates two touching tables", () => {
    const zip = pkg(P() + TBL("a") + TBL("b") + P() + SECT);
    expect(findRule(zip, "story.adjacent-tables")).toMatchObject({ severity: "fatal", count: 1 });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).not.toMatch(/<\/w:tbl>\s*<w:tbl/);
    expect(inspectDocx(zip).findings.map((f) => f.rule)).not.toContain("story.adjacent-tables");
  });

  it("flags and pads a body that ends on a table", () => {
    const zip = pkg(P() + TBL() + SECT);
    expect(findRule(zip, "story.ends-with-table")).toMatchObject({ severity: "fatal" });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toMatch(/<\/w:tbl>\s*<w:p\/?>/);
    expect(inspectDocx(zip).findings.map((f) => f.rule)).not.toContain("story.ends-with-table");
  });

  it("uses a bidi-tagged spacer in an RTL document", () => {
    const rtlTable = TBL().replace("<w:tblPr>", "<w:tblPr><w:bidiVisual/>");
    const zip = pkg(P() + rtlTable + SECT);
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain("<w:p><w:pPr><w:bidi/></w:pPr></w:p>");
  });

  it("applies the same two rules to header and footer parts", () => {
    const hdr = `<w:hdr ${NS}>${TBL("h")}</w:hdr>`;
    const zip = pkg(P() + SECT, { "word/header1.xml": hdr }, `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
    expect(findRule(zip, "story.ends-with-table")).toMatchObject({ part: "word/header1.xml" });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/header1.xml")).toMatch(/<\/w:tbl><w:p\/><\/w:hdr>/);
  });
});

describe("docx-doctor: body order", () => {
  it("moves a stray body-level sectPr to the end", () => {
    const zip = pkg(P("one") + SECT + P("two"));
    expect(findRule(zip, "body.sectpr-not-last")).toMatchObject({ severity: "fatal", fixable: true });
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml.indexOf("<w:sectPr")).toBeGreaterThan(xml.indexOf("two"));
    expect(inspectDocx(zip).findings.map((f) => f.rule)).not.toContain("body.sectpr-not-last");
  });

  it("notices a body with no sectPr at all, as a warning", () => {
    expect(findRule(pkg(P()), "body.no-sectpr")).toMatchObject({ severity: "warning" });
  });

  // The xml2js regroup signature: every table hoisted above every paragraph. The
  // original order is gone, so this is detect-only.
  it("detects a body regrouped by tag and refuses to pretend it can fix it", () => {
    const body = Array.from({ length: 4 }, (_, i) => TBL(`t${i}`)).join("") +
      Array.from({ length: 25 }, (_, i) => P(`p${i}`)).join("") + SECT;
    const f = findRule(pkg(body), "body.tag-grouped");
    expect(f).toMatchObject({ severity: "fatal", fixable: false });
    expect(f!.message).toContain("restore the thesis from document history");
  });

  it("does not cry wolf on a document that merely interleaves tables", () => {
    const body = Array.from({ length: 25 }, (_, i) => P(`p${i}`) + (i % 6 === 0 ? TBL(`t${i}`) + P() : "")).join("") + SECT;
    expect(rulesOf(pkg(body))).not.toContain("body.tag-grouped");
  });
});

describe("docx-doctor: text and bookmarks", () => {
  it("protects an edge space that would otherwise be trimmed away", () => {
    const zip = pkg(`<w:p><w:r><w:t>معهد </w:t></w:r></w:p>${SECT}`);
    expect(findRule(zip, "text.unprotected-space")).toMatchObject({ severity: "warning", count: 1 });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain('<w:t xml:space="preserve">معهد </w:t>');
  });

  it("leaves a newline-indented run alone (that whitespace is a pretty-printer's)", () => {
    const zip = pkg(`<w:p><w:r><w:t>\n  hello\n</w:t></w:r></w:p>${SECT}`);
    expect(rulesOf(zip)).not.toContain("text.unprotected-space");
  });

  it("closes an orphaned bookmarkStart and drops an orphaned bookmarkEnd", () => {
    const body =
      `<w:p><w:bookmarkStart w:id="7" w:name="_Ref1"/><w:r><w:t>a</w:t></w:r></w:p>` +
      `<w:p><w:bookmarkEnd w:id="99"/></w:p>${SECT}`;
    const zip = pkg(body);
    expect(findRule(zip, "bookmark.unmatched")).toMatchObject({ count: 2, fixable: true });
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml).toContain('<w:bookmarkEnd w:id="7"/>');
    expect(xml).not.toContain('w:id="99"');
    expect(rulesOf(zip)).not.toContain("bookmark.unmatched");
  });
});

describe("docx-doctor: package integrity", () => {
  it("resolves the package-root _rels/.rels correctly", () => {
    // Regression: an over-eager path regex resolved every root relationship
    // under `_rels/`, so a perfectly good document.xml looked dangling.
    expect(rulesOf(pkg(P() + SECT))).not.toContain("rels.dangling-target");
  });

  it("flags a relationship pointing at a part that is not in the package", () => {
    const rels =
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId5" Type="t" Target="media/image9.png"/></Relationships>`;
    const zip = pkg(P() + SECT, { "word/_rels/document.xml.rels": rels });
    expect(findRule(zip, "rels.dangling-target")).toMatchObject({ severity: "fatal", count: 1 });
    // Unreferenced, so removable — but only when explicitly asked.
    expect(findRule(zip, "rels.dangling-target", { fix: true })!.fixed).toBe(false);
    expect(findRule(zip, "rels.dangling-target", { fix: true, aggressive: true })!.fixed).toBe(true);
  });

  // The real case from a live thesis: word/_rels/header5.xml.rels ships
  // Target="NULL" on an image rel that the header's <w:drawing> still points at.
  const DEAD_IMAGE_RELS =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="NULL"/></Relationships>`;

  it("removes the broken picture that keeps a dangling relationship alive", () => {
    const body = `<w:p><w:r><w:drawing><wp:anchor r:embed="rId5"/></w:drawing></w:r></w:p>${P("keep me")}${SECT}`;
    const zip = pkg(body, { "word/_rels/document.xml.rels": DEAD_IMAGE_RELS });

    expect(findRule(zip, "rels.dangling-target-in-use")).toMatchObject({ severity: "fatal", fixable: true });
    // Removing content is opt-in: a plain fix must leave it alone.
    expect(findRule(zip, "rels.dangling-target-in-use", { fix: true })!.fixed).toBe(false);

    inspectDocx(zip, { fix: true, aggressive: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml).not.toContain("<w:drawing");
    expect(xml).not.toContain("rId5");
    expect(xml).toContain("keep me"); // the rest of the document survives
    expect(zip.readAsText("word/_rels/document.xml.rels")).not.toContain("rId5");
    expect(rulesOf(zip)).not.toContain("rels.dangling-target-in-use");
  });

  it("unlinks a dead hyperlink but KEEPS the student's words", () => {
    const body = `<w:p><w:hyperlink r:id="rId5"><w:r><w:t>important text</w:t></w:r></w:hyperlink></w:p>${SECT}`;
    const zip = pkg(body, { "word/_rels/document.xml.rels": DEAD_IMAGE_RELS });
    inspectDocx(zip, { fix: true, aggressive: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml).not.toContain("w:hyperlink");
    expect(xml).toContain("<w:t>important text</w:t>");
  });

  it("refuses to guess when the reference sits in content it does not understand", () => {
    const body = `<w:p><w:customThing r:id="rId5"/></w:p>${SECT}`;
    const zip = pkg(body, { "word/_rels/document.xml.rels": DEAD_IMAGE_RELS });
    const f = findRule(zip, "rels.dangling-target-in-use", { fix: true, aggressive: true });
    expect(f).toMatchObject({ fixable: false, fixed: false });
    // Nothing was touched — the rel is still there rather than half-removed.
    expect(zip.readAsText("word/_rels/document.xml.rels")).toContain("rId5");
    expect(zip.readAsText("word/document.xml")).toContain("rId5");
  });

  it("flags an r:id the part references but the rels file never declares", () => {
    const body = `<w:p><w:hyperlink r:id="rIdNope"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:p>${SECT}`;
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
    const zip = pkg(body, { "word/_rels/document.xml.rels": rels });
    expect(findRule(zip, "rels.missing-relationship")).toMatchObject({ severity: "fatal", part: "word/document.xml" });
  });

  it("declares a part that [Content_Types].xml forgot", () => {
    const zip = pkg(P() + SECT, { "word/styles.xml": `<w:styles ${NS}/>` });
    expect(findRule(zip, "contenttypes.missing-override")).toMatchObject({ severity: "fatal", count: 1 });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("[Content_Types].xml")).toContain('PartName="/word/styles.xml"');
    expect(rulesOf(zip)).not.toContain("contenttypes.missing-override");
  });

  it("flags a torn XML part instead of trying to repair it", () => {
    const zip = pkg(P() + SECT);
    zip.addFile("word/document.xml", Buffer.from(`<w:document ${NS}><w:body><w:p></w:body></w:document>`, "utf8"));
    expect(findRule(zip, "xml.malformed")).toMatchObject({ severity: "fatal", fixable: false });
  });
});

describe("docx-doctor: references", () => {
  it("flags a paragraph style that styles.xml never defines", () => {
    const body = `<w:p><w:pPr><w:pStyle w:val="Ghost"/></w:pPr></w:p>${SECT}`;
    const zip = pkg(body, { "word/styles.xml": `<w:styles ${NS}><w:style w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>` });
    expect(findRule(zip, "ref.unknown-style")).toMatchObject({ severity: "warning", detail: "Ghost" });
  });

  describe("what Word refuses the whole file for", () => {
    const fatal = (z: DocxZip, rule: string) => inspectDocx(z).findings.find((f) => f.rule === rule && f.severity === "fatal");

    it("a list level naming a picture bullet that is not there", () => {
      const num = (inner: string) => `<w:numbering ${NS}>${inner}<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:lvlPicBulletId w:val="3"/></w:lvl></w:abstractNum></w:numbering>`;
      expect(fatal(pkg(P() + SECT, { "word/numbering.xml": num("") }), "ref.unknown-pic-bullet")).toMatchObject({ count: 1, fixable: false });
      expect(rulesOf(pkg(P() + SECT, { "word/numbering.xml": num('<w:numPicBullet w:numPicBulletId="3"><w:pict/></w:numPicBullet>') }))).not.toContain("ref.unknown-pic-bullet");
    });
  });

  it("treats numId 0 as legal (it means 'no numbering')", () => {
    const body = `<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>${SECT}`;
    const zip = pkg(body, { "word/numbering.xml": `<w:numbering ${NS}/>` });
    expect(rulesOf(zip)).not.toContain("ref.unknown-numid");
  });
});

describe("docx-doctor: table grid", () => {
  it("flags a row that spans more columns than the grid declares", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    expect(findRule(pkg(tbl + P() + SECT), "table.grid-mismatch")).toMatchObject({ severity: "warning", count: 1 });
  });

  it("accepts a legitimate gridSpan that fits", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    expect(rulesOf(pkg(tbl + P() + SECT))).not.toContain("table.grid-mismatch");
  });
});

describe("firstXmlError", () => {
  it("accepts well-formed markup, including self-closing and attribute '>'", () => {
    expect(firstXmlError(`<a><b x="1>2"/><c/></a>`)).toBeNull();
    expect(firstXmlError(`<?xml version="1.0"?><!-- <a> --><a><![CDATA[</a>]]></a>`)).toBeNull();
  });

  it("catches an unclosed and a mismatched tag", () => {
    expect(firstXmlError("<a><b></a>")).toBeTruthy();
    expect(firstXmlError("<a><b>")).toBeTruthy();
  });
});

describe("clearFalseDataDescriptors", () => {
  it("leaves an ordinary archive byte-identical", () => {
    const zip = new AdmZip();
    zip.addFile("a.xml", Buffer.from("<a/>", "utf8"));
    const buf = zip.toBuffer();
    expect(clearFalseDataDescriptors(buf).equals(buf)).toBe(true);
  });

  it("clears the flag on an entry that promises a descriptor but carries a real CRC", () => {
    const zip = new AdmZip();
    zip.addFile("a.xml", Buffer.from("<a/>", "utf8"));
    const buf = Buffer.from(zip.toBuffer());
    // Forge the streamed-entry flag in both the local header and the directory.
    const local = buf.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]));
    const central = buf.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
    buf.writeUInt16LE(buf.readUInt16LE(local + 6) | 0x0008, local + 6);
    buf.writeUInt16LE(buf.readUInt16LE(central + 8) | 0x0008, central + 8);

    const fixed = clearFalseDataDescriptors(buf);
    expect(fixed.readUInt16LE(local + 6) & 0x0008).toBe(0);
    expect(fixed.readUInt16LE(central + 8) & 0x0008).toBe(0);
    // …and the archive still reads.
    expect(new AdmZip(fixed).readAsText("a.xml")).toBe("<a/>");
  });

  it("returns the input untouched when it is not a zip at all", () => {
    const junk = Buffer.from("not a zip");
    expect(clearFalseDataDescriptors(junk).equals(junk)).toBe(true);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Zip-level flags. Every "Word" claim below was measured through real Word
// (Mac, 2026-11-02) on the seed document with exactly this forgery:
//   non-empty entry, bit 3 in the local header, no descriptor → refused (-10000);
//   the same on one single entry                              → refused;
//   the same on a directory or an empty file                  → opens;
//   bit 3 in the central directory only                       → opens;
//   a true streamed archive (real descriptors)                → opens;
//   CRC 0 on a non-empty entry, no flag                       → opens, unzip -t fails.
// ─────────────────────────────────────────────────────────────────────────────

/** Every entry's header offsets, read off the central directory. */
function zipOffsets(buf: Buffer): { name: string; central: number; local: number }[] {
  const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const out: { name: string; central: number; local: number }[] = [];
  let p = buf.readUInt32LE(eocd + 16);
  for (let n = buf.readUInt16LE(eocd + 10); n > 0; n--) {
    const nameLen = buf.readUInt16LE(p + 28);
    out.push({ name: buf.toString("utf8", p + 46, p + 46 + nameLen), central: p, local: buf.readUInt32LE(p + 42) });
    p += 46 + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32);
  }
  return out;
}

/** Set bit 3 on the entries `pick` chooses, the way adm-zip leaves a streamed upload. */
function forgeBit3(src: Buffer, pick: (name: string) => boolean = () => true, where: "both" | "central" | "local" = "both"): Buffer {
  const buf = Buffer.from(src);
  for (const e of zipOffsets(buf)) {
    if (!pick(e.name)) continue;
    if (where !== "local") buf.writeUInt16LE(buf.readUInt16LE(e.central + 8) | 0x0008, e.central + 8);
    if (where !== "central") buf.writeUInt16LE(buf.readUInt16LE(e.local + 6) | 0x0008, e.local + 6);
  }
  return buf;
}

const bit3Everywhere = (buf: Buffer) =>
  zipOffsets(buf).filter((e) => (buf.readUInt16LE(e.central + 8) | buf.readUInt16LE(e.local + 6)) & 0x0008).map((e) => e.name);

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (b: Buffer) => {
  let c = 0xffffffff;
  for (const x of b) c = CRC_TABLE[(c ^ x) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

/** A really streamed archive: bit 3, zeros in the local headers, the real values in a descriptor behind the data (stored). */
function streamedZip(files: Record<string, string>, opts: { signature?: boolean; directories?: string[] } = {}): Buffer {
  const signed = opts.signature !== false;
  const chunks: Buffer[] = [];
  const central: Buffer[] = [];
  let at = 0;
  const add = (name: string, data: Buffer) => {
    const nameBuf = Buffer.from(name, "utf8");
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0008, 6);
    local.writeUInt16LE(nameBuf.length, 26);
    const desc = Buffer.alloc(signed ? 16 : 12);
    let d = 0;
    if (signed) { desc.writeUInt32LE(0x08074b50, 0); d = 4; }
    desc.writeUInt32LE(crc, d);
    desc.writeUInt32LE(data.length, d + 4);
    desc.writeUInt32LE(data.length, d + 8);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0);
    c.writeUInt16LE(20, 4);
    c.writeUInt16LE(20, 6);
    c.writeUInt16LE(0x0008, 8);
    c.writeUInt32LE(crc, 16);
    c.writeUInt32LE(data.length, 20);
    c.writeUInt32LE(data.length, 24);
    c.writeUInt16LE(nameBuf.length, 28);
    c.writeUInt32LE(at, 42);
    central.push(c, nameBuf);
    chunks.push(local, nameBuf, data, desc);
    at += local.length + nameBuf.length + data.length + desc.length;
  };
  for (const dir of opts.directories ?? []) add(dir, Buffer.alloc(0));
  for (const [name, text] of Object.entries(files)) add(name, Buffer.from(text, "utf8"));
  const dirBytes = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(central.length / 2, 8);
  eocd.writeUInt16LE(central.length / 2, 10);
  eocd.writeUInt32LE(dirBytes.length, 12);
  eocd.writeUInt32LE(at, 16);
  return Buffer.concat([...chunks, dirBytes, eocd]);
}

const PKG_PARTS = (body = P() + SECT) => ({
  "[Content_Types].xml": CONTENT_TYPES(),
  "_rels/.rels": ROOT_RELS,
  "word/document.xml": doc(body),
});

/** A package carrying an empty entry too. */
function pkgWithEmpty(): Buffer {
  const zip = pkg(P() + SECT);
  zip.addFile("customXml/empty.bin", Buffer.alloc(0));
  zip.addFile("word/", Buffer.alloc(0));
  return zip.toBuffer();
}
const isEmptyEntry = (n: string) => n === "customXml/empty.bin" || n.endsWith("/");
const ruleOf = (r: { findings: { rule: string }[] }, rule: string) => r.findings.find((f) => f.rule === rule) as any;
const zipRules = (r: { findings: { rule: string }[] }) => r.findings.map((f) => f.rule).filter((x) => x.startsWith("zip."));

describe("inspectZipFlags", () => {
  it("is silent on an ordinary archive and on a truly streamed one", () => {
    const empty = { falseDescriptors: [], toleratedFalseDescriptors: [], zeroCrc: [] };
    expect(inspectZipFlags(pkg(P() + SECT).toBuffer())).toEqual(empty);
    expect(inspectZipFlags(pkgWithEmpty())).toEqual(empty);
    expect(inspectZipFlags(streamedZip(PKG_PARTS()))).toEqual(empty);
    expect(inspectZipFlags(streamedZip(PKG_PARTS(), { signature: false }))).toEqual(empty);
    expect(inspectZipFlags(streamedZip(PKG_PARTS(), { directories: ["word/"] }))).toEqual(empty);
    expect(inspectZipFlags(streamedZip(PKG_PARTS(), { directories: ["word/"], signature: false }))).toEqual(empty);
  });

  it("names a false flag on a non-empty entry, in both headers or the local one", () => {
    const buf = pkg(P() + SECT).toBuffer();
    expect(inspectZipFlags(forgeBit3(buf, (n) => n === "word/document.xml"))!.falseDescriptors).toEqual(["word/document.xml"]);
    expect(inspectZipFlags(forgeBit3(buf, (n) => n === "word/document.xml", "local"))!.falseDescriptors).toEqual(["word/document.xml"]);
    expect(inspectZipFlags(forgeBit3(buf))!.falseDescriptors).toHaveLength(3);
  });

  it("files what Word tolerates apart: empty entries and a central-only flag", () => {
    const buf = pkgWithEmpty();
    const empties = inspectZipFlags(forgeBit3(buf, isEmptyEntry))!;
    expect(empties.falseDescriptors).toEqual([]);
    expect(empties.toleratedFalseDescriptors.sort()).toEqual(["customXml/empty.bin", "word/"]);
    const central = inspectZipFlags(forgeBit3(buf, (n) => n === "word/document.xml", "central"))!;
    expect(central.falseDescriptors).toEqual([]);
    expect(central.toleratedFalseDescriptors).toEqual(["word/document.xml"]);
  });

  it("names a non-empty entry that records CRC 0, but never an empty one", () => {
    const buf = Buffer.from(pkgWithEmpty());
    const e = zipOffsets(buf).find((x) => x.name === "_rels/.rels")!;
    buf.writeUInt32LE(0, e.central + 16);
    buf.writeUInt32LE(0, e.local + 14);
    expect(inspectZipFlags(buf)).toEqual({ falseDescriptors: [], toleratedFalseDescriptors: [], zeroCrc: ["_rels/.rels"] });
  });

  it("returns null on what it cannot walk", () => {
    expect(inspectZipFlags(Buffer.from("not a zip"))).toBeNull();
    const buf = Buffer.from(pkg(P() + SECT).toBuffer());
    const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    const zip64 = Buffer.from(buf);
    zip64.writeUInt16LE(0xffff, eocd + 10);
    expect(inspectZipFlags(zip64)).toBeNull();
    const broken = Buffer.from(buf);
    broken.writeUInt32LE(0, broken.readUInt32LE(eocd + 16));
    expect(inspectZipFlags(broken)).toBeNull();
  });
});

describe("clearFalseDataDescriptors, the entries it used to skip", () => {
  it("clears a false flag on a CRC-0 directory and on an empty file", () => {
    const forged = forgeBit3(pkgWithEmpty(), isEmptyEntry);
    expect(bit3Everywhere(forged).sort()).toEqual(["customXml/empty.bin", "word/"]);
    expect(bit3Everywhere(clearFalseDataDescriptors(forged))).toEqual([]);
  });

  it("clears every flag adm-zip leaves on a streamed upload it re-wrote, directories included", () => {
    const rewritten = new AdmZip(streamedZip(PKG_PARTS(), { directories: ["word/"] })).toBuffer();
    expect(bit3Everywhere(rewritten).length).toBe(4);
    const fixed = clearFalseDataDescriptors(rewritten);
    expect(bit3Everywhere(fixed)).toEqual([]);
    expect(new AdmZip(fixed).readAsText("word/document.xml")).toContain("hello");
  });

  it("clears a flag standing in only one of the two headers", () => {
    const buf = pkg(P() + SECT).toBuffer();
    expect(bit3Everywhere(clearFalseDataDescriptors(forgeBit3(buf, undefined, "central")))).toEqual([]);
    expect(bit3Everywhere(clearFalseDataDescriptors(forgeBit3(buf, undefined, "local")))).toEqual([]);
  });

  it("leaves a truly streamed archive byte-identical, signed or not, directories included", () => {
    for (const b of [streamedZip(PKG_PARTS()), streamedZip(PKG_PARTS(), { signature: false, directories: ["word/"] })]) {
      expect(clearFalseDataDescriptors(b)).toBe(b);
    }
  });

  it("in a mixed archive clears the false flag and keeps the true one", () => {
    const streamed = streamedZip({ "a.xml": "<a/>", "b.xml": "<b/>" });
    // Turn a.xml into the adm-zip state: real values in the local header, no descriptor.
    const zip = new AdmZip();
    zip.addFile("first.xml", Buffer.from("<first/>", "utf8"));
    const normal = forgeBit3(zip.toBuffer());
    const merged = appendEntries(normal, streamed);
    const fixed = clearFalseDataDescriptors(merged);
    expect(bit3Everywhere(fixed).sort()).toEqual(["a.xml", "b.xml"]);
    const z = new AdmZip(fixed);
    for (const [n, t] of [["first.xml", "<first/>"], ["a.xml", "<a/>"], ["b.xml", "<b/>"]]) expect(z.readAsText(n)).toBe(t);
  });

  it("bails on a flagged entry whose local header disagrees and has no descriptor: it never guesses", () => {
    const buf = forgeBit3(pkg(P() + SECT).toBuffer());
    const e = zipOffsets(buf).find((x) => x.name === "word/document.xml")!;
    buf.writeUInt32LE(0, e.local + 14); // local CRC gone, nothing behind the data to recover it from
    expect(clearFalseDataDescriptors(buf)).toBe(buf);
    expect(inspectZipFlags(buf)).toBeNull();
  });
});

/** The two archives' entries in one zip: `head`'s entries first, then `tail`'s, offsets adjusted. */
function appendEntries(head: Buffer, tail: Buffer): Buffer {
  const eocdOf = (b: Buffer) => b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const h = eocdOf(head);
  const t = eocdOf(tail);
  const hDir = head.readUInt32LE(h + 16);
  const tDir = tail.readUInt32LE(t + 16);
  const tailCentral = Buffer.from(tail.subarray(tDir, t));
  let p = 0;
  for (let n = tail.readUInt16LE(t + 10); n > 0; n--) {
    tailCentral.writeUInt32LE(tailCentral.readUInt32LE(p + 42) + hDir, p + 42);
    p += 46 + tailCentral.readUInt16LE(p + 28) + tailCentral.readUInt16LE(p + 30) + tailCentral.readUInt16LE(p + 32);
  }
  const headCentral = head.subarray(hDir, h);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  const count = head.readUInt16LE(h + 10) + tail.readUInt16LE(t + 10);
  eocd.writeUInt16LE(count, 8);
  eocd.writeUInt16LE(count, 10);
  eocd.writeUInt32LE(headCentral.length + tailCentral.length, 12);
  eocd.writeUInt32LE(hDir + tDir, 16);
  return Buffer.concat([head.subarray(0, hDir), tail.subarray(0, tDir), headCentral, tailCentral, eocd]);
}

describe("docx-doctor: the package's own zip flags", () => {
  it("a healthy package reports no zip finding through either byte-level door", async () => {
    const buf = pkg(P() + SECT).toBuffer();
    expect(zipRules(await checkDocxBuffer(buf))).toEqual([]);
    const repaired = await repairDocxBuffer(buf);
    expect(zipRules(repaired)).toEqual([]);
    expect(repaired.changed).toBe(false);
    expect(repaired.buffer).toBe(buf);
  });

  it("a truly streamed package is healthy and stays byte-identical", async () => {
    const buf = streamedZip(PKG_PARTS(), { directories: ["word/"] });
    const checked = await checkDocxBuffer(buf);
    expect(zipRules(checked)).toEqual([]);
    expect(checked.ok).toBe(true);
    const repaired = await repairDocxBuffer(buf);
    expect(repaired.changed).toBe(false);
    expect(repaired.buffer).toBe(buf);
  });

  it("check: a false descriptor on a non-empty entry is FATAL, the doctor no longer says ok", async () => {
    const report = await checkDocxBuffer(forgeBit3(pkg(P() + SECT).toBuffer()));
    expect(report.ok).toBe(false);
    const f = ruleOf(report, "zip.false-data-descriptor");
    expect(f).toMatchObject({ severity: "fatal", part: "package", count: 3, fixable: true, fixed: false });
    expect(f.detail).toContain("word/document.xml");
    expect(report.findings[0].rule).toBe("zip.false-data-descriptor"); // fatals sort first
  });

  it("check: it still reads the parts behind a false flag (they used to read as nothing, so every check passed)", async () => {
    const zip = pkg(P() + SECT);
    zip.addFile("word/document.xml", Buffer.from("<w:document><w:body><w:p></w:body>", "utf8"));
    const report = await checkDocxBuffer(forgeBit3(zip.toBuffer()));
    expect(report.findings.map((x) => x.rule)).toEqual(expect.arrayContaining(["zip.false-data-descriptor", "xml.malformed"]));
  });

  it("check never modifies the buffer it is given", async () => {
    const forged = forgeBit3(pkg(P() + SECT).toBuffer());
    const copy = Buffer.from(forged);
    await checkDocxBuffer(forged);
    expect(forged.equals(copy)).toBe(true);
  });

  it("repair clears the flags, says so, and the bytes read back", async () => {
    const forged = forgeBit3(pkg(P() + SECT).toBuffer());
    const copy = Buffer.from(forged);
    const r = await repairDocxBuffer(forged);
    expect(forged.equals(copy)).toBe(true);
    expect(r.changed).toBe(true);
    expect(r.ok).toBe(true);
    expect(r.repairedParts).toEqual([]);
    expect(ruleOf(r, "zip.false-data-descriptor")).toMatchObject({ severity: "fatal", count: 3, fixed: true });
    expect(bit3Everywhere(r.buffer)).toEqual([]);
    expect(new AdmZip(r.buffer).readAsText("word/document.xml")).toContain("hello");
    // and it is now clean through the checking door
    expect(zipRules(await checkDocxBuffer(r.buffer))).toEqual([]);
  });

  it("repair reports the cleared flags alongside a part it also rewrote", async () => {
    // A table as the last body block: the doctor appends the spacer paragraph.
    const forged = forgeBit3(pkg(TBL() + SECT).toBuffer());
    const r = await repairDocxBuffer(forged);
    expect(r.repairedParts).toContain("word/document.xml");
    expect(r.changed).toBe(true);
    expect(ruleOf(r, "zip.false-data-descriptor")).toMatchObject({ count: 3, fixed: true });
    expect(bit3Everywhere(r.buffer)).toEqual([]);
    expect(r.ok).toBe(true);
  });

  it("an abandoned repair leaves the flags reported and unfixed, on the original bytes", async () => {
    const forged = forgeBit3(pkg(TBL() + SECT).toBuffer());
    // Sabotage only the first open: its rewrite cannot be serialised, so the repair gives up.
    const real = ZipManager.loadFromBuffer;
    const spy = vi.spyOn(ZipManager, "loadFromBuffer").mockImplementationOnce(async (b: Buffer) => {
      const z = await real.call(ZipManager, b);
      (z as any).toBuffer = () => { throw new Error("no bytes"); };
      return z;
    });
    try {
      const abandoned = await repairDocxBuffer(forged);
      expect(abandoned.buffer).toBe(forged);
      expect(abandoned.changed).toBe(false);
      expect(ruleOf(abandoned, "repair.abandoned")).toBeTruthy();
      expect(ruleOf(abandoned, "zip.false-data-descriptor")).toMatchObject({ fixed: false });
      expect(abandoned.ok).toBe(false);
    } finally {
      spy.mockRestore();
    }
  });

  it("flags on empty entries are a WARNING, and repair clears them too", async () => {
    const forged = forgeBit3(pkgWithEmpty(), isEmptyEntry);
    const checked = await checkDocxBuffer(forged);
    expect(checked.ok).toBe(true);
    expect(ruleOf(checked, "zip.false-data-descriptor")).toBeUndefined();
    expect(ruleOf(checked, "zip.false-data-descriptor-tolerated")).toMatchObject({ severity: "warning", count: 2, fixable: true, fixed: false });
    const r = await repairDocxBuffer(forged);
    expect(r.changed).toBe(true);
    expect(ruleOf(r, "zip.false-data-descriptor-tolerated")).toMatchObject({ count: 2, fixed: true });
    expect(bit3Everywhere(r.buffer)).toEqual([]);
  });

  it("a part nothing can read is fatal: a flag with no descriptor over a header that disagrees", async () => {
    const buf = forgeBit3(pkg(P() + SECT).toBuffer());
    const e = zipOffsets(buf).find((x) => x.name === "word/document.xml")!;
    buf.writeUInt32LE((buf.readUInt32LE(e.local + 14) ^ 1) >>> 0, e.local + 14);
    const report = await checkDocxBuffer(buf);
    expect(report.ok).toBe(false);
    expect(ruleOf(report, "zip.unreadable-part")).toMatchObject({ severity: "fatal", fixable: false });
    expect(ruleOf(report, "zip.unreadable-part").detail).toContain("word/document.xml");
    const r = await repairDocxBuffer(buf);
    expect(r.ok).toBe(false);
    expect(r.buffer).toBe(buf);
    expect(ruleOf(r, "zip.unreadable-part")).toMatchObject({ fixed: false });
  });

  it("damaged compressed data is the same fault", async () => {
    const zip = new AdmZip();
    for (const [name, xml] of Object.entries(PKG_PARTS())) zip.addFile(name, Buffer.from(xml.repeat(40), "utf8"));
    const buf = Buffer.from(zip.toBuffer());
    const e = zipOffsets(buf).find((x) => x.name === "word/document.xml")!;
    const dataAt = e.local + 30 + buf.readUInt16LE(e.local + 26) + buf.readUInt16LE(e.local + 28);
    for (let i = 8; i < 24; i++) buf[dataAt + i] ^= 0xff;
    const report = await checkDocxBuffer(buf);
    expect(report.ok).toBe(false);
    expect(ruleOf(report, "zip.unreadable-part")).toBeDefined();
  });

  it("a healthy package, a streamed one and a healed false flag have no unreadable part", async () => {
    for (const b of [pkg(P() + SECT).toBuffer(), streamedZip(PKG_PARTS()), forgeBit3(pkg(P() + SECT).toBuffer())]) {
      expect(ruleOf(await checkDocxBuffer(b), "zip.unreadable-part")).toBeUndefined();
    }
  });

  it("a CRC-0 part stays a warning: Word opens it, whatever this engine's reader says", async () => {
    const buf = Buffer.from(pkg(P() + SECT).toBuffer());
    const e = zipOffsets(buf).find((x) => x.name === "_rels/.rels")!;
    buf.writeUInt32LE(0, e.local + 14);
    buf.writeUInt32LE(0, e.central + 16);
    const report = await checkDocxBuffer(buf);
    expect(ruleOf(report, "zip.crc-zero")).toBeDefined();
    expect(ruleOf(report, "zip.unreadable-part")).toBeUndefined();
  });

  it("a CRC-0 entry is flagged, not fixable, and repair leaves it alone", async () => {
    const buf = Buffer.from(pkg(P() + SECT).toBuffer());
    const e = zipOffsets(buf).find((x) => x.name === "_rels/.rels")!;
    buf.writeUInt32LE(0, e.central + 16);
    buf.writeUInt32LE(0, e.local + 14);
    const checked = await checkDocxBuffer(buf);
    expect(ruleOf(checked, "zip.crc-zero")).toMatchObject({ severity: "warning", count: 1, fixable: false, fixed: false });
    expect(ruleOf(checked, "zip.crc-zero").detail).toContain("_rels/.rels");
    const r = await repairDocxBuffer(buf);
    expect(r.changed).toBe(false);
    expect(r.buffer).toBe(buf);
    expect(ruleOf(r, "zip.crc-zero")).toMatchObject({ fixed: false });
  });

  it("a CRC-0 entry with a false flag loses the flag, keeps the zero, and both are reported", async () => {
    const buf = forgeBit3(pkg(P() + SECT).toBuffer(), (n) => n === "_rels/.rels");
    const e = zipOffsets(buf).find((x) => x.name === "_rels/.rels")!;
    buf.writeUInt32LE(0, e.central + 16);
    buf.writeUInt32LE(0, e.local + 14);
    const checked = await checkDocxBuffer(buf);
    expect(ruleOf(checked, "zip.false-data-descriptor")).toMatchObject({ severity: "fatal", count: 1 });
    expect(ruleOf(checked, "zip.crc-zero")).toMatchObject({ count: 1 });
    const r = await repairDocxBuffer(buf);
    expect(r.changed).toBe(true);
    expect(bit3Everywhere(r.buffer)).toEqual([]);
    expect(ruleOf(r, "zip.false-data-descriptor")).toMatchObject({ fixed: true });
    expect(ruleOf(r, "zip.crc-zero")).toMatchObject({ fixed: false });
    expect(r.ok).toBe(true);
  });

  it("the real path: a streamed upload re-written by adm-zip is caught, then healed", async () => {
    const rewritten = new AdmZip(streamedZip(PKG_PARTS())).toBuffer();
    const checked = await checkDocxBuffer(rewritten);
    expect(ruleOf(checked, "zip.false-data-descriptor")).toMatchObject({ severity: "fatal", count: 3 });
    const r = await repairDocxBuffer(rewritten);
    expect(r.ok).toBe(true);
    expect(new AdmZip(r.buffer).readAsText("word/document.xml")).toContain("hello");
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Regressions from a live thesis Word refused to open AFTER the doctor had
// already declared it healthy. Both are xsd:sequence violations the first
// version of this file did not model, and both are silently fatal.
// ─────────────────────────────────────────────────────────────────────────────

describe("docx-doctor: the defects that still made Word refuse the file", () => {
  it("puts w:tblBorders back in schema order (top, left, bottom, right)", () => {
    // mdocxengine built this object as top/bottom/left/right — the order a human
    // thinks in — and the XML builder emitted the keys as given.
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>` +
      `<w:top w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:left w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:right w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:insideH w:val="single" w:sz="2" w:space="0" w:color="BFBFBF"/>` +
      `<w:insideV w:val="single" w:sz="2" w:space="0" w:color="BFBFBF"/>` +
      `</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);

    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({ severity: "fatal", detail: "w:tblBorders" });
    inspectDocx(zip, { fix: true });
    const sides = [...zip.readAsText("word/document.xml").matchAll(/<w:(top|left|bottom|right|insideH|insideV)\b/g)].map((m) => m[1]);
    expect(sides).toEqual(["top", "left", "bottom", "right", "insideH", "insideV"]);
    expect(rulesOf(zip)).not.toContain("sequence.out-of-order");
  });

  it("moves a w:trPr that got written after the row's first cell", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>` +
      `<w:trPr><w:tblHeader/></w:trPr></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);

    expect(findRule(zip, "sequence.out-of-order")).toMatchObject({ severity: "fatal", detail: "w:trPr" });
    inspectDocx(zip, { fix: true });
    const row = /<w:tr>[\s\S]*?<\/w:tr>/.exec(zip.readAsText("word/document.xml"))![0];
    expect(row.indexOf("<w:trPr>")).toBeLessThan(row.indexOf("<w:tc>"));
    expect(row).toContain("<w:t>cell</w:t>"); // the row's content is untouched
    expect(rulesOf(zip)).not.toContain("sequence.out-of-order");
  });

  it("does not disturb a row or table that is already correct", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>` +
      `<w:top w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:left w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `<w:right w:val="single" w:sz="4" w:space="0" w:color="808080"/>` +
      `</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);
    const before = zip.readAsText("word/document.xml");
    expect(inspectDocx(zip, { fix: true }).findings).toEqual([]);
    expect(zip.readAsText("word/document.xml")).toBe(before);
  });

  it("keeps w:pPr first in a paragraph whose properties drifted after a run", () => {
    const body = `<w:p><w:r><w:t>text</w:t></w:r><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>${SECT}`;
    const zip = pkg(body);
    inspectDocx(zip, { fix: true });
    const p = /<w:p>[\s\S]*?<\/w:p>/.exec(zip.readAsText("word/document.xml"))![0];
    expect(p.indexOf("<w:pPr>")).toBeLessThan(p.indexOf("<w:r>"));
    expect(p).toContain("<w:t>text</w:t>");
  });
});

describe("docx-doctor: a run loose in the body", () => {
  // THE defect that kept a real thesis unopenable after every other rule here
  // reported it healthy — found only by validating against the real OOXML schema.
  // mdocxengine's setParagraphText appended the new run AFTER a self-closing
  // <w:p/> instead of inside it.
  it("flags a bare <w:r> at body level as fatal", () => {
    const zip = pkg(`${P()}<w:r><w:t>loose</w:t></w:r>${SECT}`);
    expect(findRule(zip, "body.stray-run")).toMatchObject({ severity: "fatal", count: 1, fixable: true });
  });

  it("wraps a stray run that carries text, keeping the words", () => {
    const zip = pkg(`${P()}<w:r><w:t>loose</w:t></w:r>${SECT}`);
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml).toContain("<w:p><w:r><w:t>loose</w:t></w:r></w:p>");
    expect(rulesOf(zip)).not.toContain("body.stray-run");
  });

  it("drops a stray run that carries nothing, rather than adding a blank line", () => {
    const zip = pkg(`${P()}<w:r><w:t xml:space="preserve"></w:t></w:r>${SECT}`);
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    // The empty run is gone entirely — no empty paragraph left in its place.
    expect(xml).not.toContain('xml:space="preserve"');
    expect(xml).toContain("<w:p><w:r><w:t>hello</w:t></w:r></w:p><w:sectPr>");
    expect(rulesOf(zip)).not.toContain("body.stray-run");
  });

  it("keeps a stray run that holds a picture", () => {
    const zip = pkg(`${P()}<w:r><w:drawing><wp:inline/></w:drawing></w:r>${SECT}`);
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain("<w:p><w:r><w:drawing>");
  });

  it("reports — but does not guess at — some other illegal body child", () => {
    const zip = pkg(`${P()}<w:fldSimple w:instr="PAGE"/>${SECT}`);
    expect(findRule(zip, "body.illegal-child")).toMatchObject({ severity: "fatal", fixable: false });
  });

  // The same defect one layer up: a writer spliced paragraph PROPERTIES in after
  // a self-closing <w:p/>, so they landed beside the paragraph. Reported as an
  // unknown-block chip in the app, sitting exactly where a page break belonged.
  describe("an orphaned <w:pPr> in the body", () => {
    const ORPHAN = '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>';

    it("flags it fatal, and as fixable", () => {
      const zip = pkg(`${P()}<w:p w:rsidR="00A1"/>${ORPHAN}${SECT}`);
      expect(findRule(zip, "body.orphaned-ppr")).toMatchObject({ severity: "fatal", count: 1, fixable: true });
    });

    it("merges it back into the paragraph before it, keeping the section break", () => {
      const zip = pkg(`${P()}<w:p w:rsidR="00A1"/>${ORPHAN}${SECT}`);
      inspectDocx(zip, { fix: true });
      const xml = zip.readAsText("word/document.xml");
      // Inside the paragraph now — and the empty <w:p/> is paired so it can hold it.
      expect(xml).toContain('<w:p w:rsidR="00A1"><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>');
      expect(xml).not.toContain("/><w:pPr>");
      expect(rulesOf(zip)).not.toContain("body.orphaned-ppr");
      expect(rulesOf(zip)).not.toContain("body.illegal-child");
    });

    it("merges into a paragraph that already has properties, in CT_PPr order", () => {
      const zip = pkg(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>${ORPHAN}${SECT}`);
      inspectDocx(zip, { fix: true });
      const xml = zip.readAsText("word/document.xml");
      // w:jc precedes w:sectPr in CT_PPr, and both are inside the one pPr.
      expect(xml).toContain('<w:pPr><w:jc w:val="center"/><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>');
      expect(rulesOf(zip)).not.toContain("body.orphaned-ppr");
    });

    it("leaves one with no paragraph before it to the unfixable rule", () => {
      const zip = pkg(`${ORPHAN}${P()}${SECT}`);
      expect(rulesOf(zip)).not.toContain("body.orphaned-ppr");
      expect(findRule(zip, "body.illegal-child")).toMatchObject({ severity: "fatal", fixable: false });
    });
  });

  it("leaves the legal block-level elements alone", () => {
    const zip = pkg(`<w:bookmarkStart w:id="1" w:name="a"/>${P()}<w:bookmarkEnd w:id="1"/>${TBL()}${P()}${SECT}`);
    expect(rulesOf(zip)).not.toContain("body.stray-run");
    expect(rulesOf(zip)).not.toContain("body.illegal-child");
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Structural content models (src/doctor/structure.ts).
//
// Every rule here was mutation-tested: a healthy Word-authored thesis was broken
// in each of these ways, and the repaired result then re-validated against the
// real OOXML schema via modakerati-server/scripts/ooxml-validate. The corpus of
// real theses and templates produces ZERO of these findings, so none of them can
// fire on a healthy import.
// ─────────────────────────────────────────────────────────────────────────────

describe("docx-doctor: containers that may not be empty", () => {
  const CELL = (inner: string) =>
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
    `<w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;

  it("fills a table cell that has lost its paragraph", () => {
    const zip = pkg(CELL('<w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>') + P() + SECT);
    expect(findRule(zip, "table.empty-cell")).toMatchObject({ severity: "fatal", count: 1, fixable: true });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain("</w:tcPr><w:p/></w:tc>");
    expect(rulesOf(zip)).not.toContain("table.empty-cell");
  });

  it("keeps the cell's own properties when filling it", () => {
    const zip = pkg(CELL('<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>') + P() + SECT);
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain('<w:gridSpan w:val="2"/>');
  });

  it("leaves a cell that already has a paragraph alone", () => {
    const zip = pkg(CELL("<w:p><w:r><w:t>x</w:t></w:r></w:p>") + P() + SECT);
    expect(rulesOf(zip)).not.toContain("table.empty-cell");
  });

  it("uses a bidi spacer inside an RTL table", () => {
    const rtl = CELL('<w:tcPr/>').replace("<w:tblPr>", "<w:tblPr><w:bidiVisual/>");
    const zip = pkg(rtl + P() + SECT);
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain("<w:p><w:pPr><w:bidi/></w:pPr></w:p>");
  });
});

describe("docx-doctor: tables that cannot render", () => {
  it("rebuilds a missing tblGrid from the widest row", () => {
    const tbl =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>` +
      `<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);
    expect(findRule(zip, "table.missing-grid")).toMatchObject({ severity: "fatal", fixable: true });
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect((xml.match(/<w:gridCol\b/g) ?? []).length).toBe(3);
    expect(xml.indexOf("<w:tblGrid>")).toBeGreaterThan(xml.indexOf("</w:tblPr>"));
    expect(rulesOf(zip)).not.toContain("table.missing-grid");
  });

  it("counts gridSpan when sizing the rebuilt grid", () => {
    const tbl =
      `<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);
    inspectDocx(zip, { fix: true });
    expect((zip.readAsText("word/document.xml").match(/<w:gridCol\b/g) ?? []).length).toBe(3);
  });

  it("removes a row with no cells", () => {
    const tbl =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:tblHeader/></w:trPr></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const zip = pkg(tbl + P() + SECT);
    expect(findRule(zip, "table.empty-row")).toMatchObject({ severity: "fatal", count: 1 });
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect((xml.match(/<w:tr>/g) ?? []).length).toBe(1);
    expect(xml).toContain("<w:t>keep</w:t>");
  });

  it("removes a table with no rows", () => {
    const zip = pkg(`${P()}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid></w:tbl>${P("after")}${SECT}`);
    expect(findRule(zip, "table.empty")).toMatchObject({ severity: "fatal", count: 1 });
    inspectDocx(zip, { fix: true });
    const xml = zip.readAsText("word/document.xml");
    expect(xml).not.toContain("<w:tbl>");
    expect(xml).toContain("<w:t>after</w:t>"); // surrounding content survives
  });
});

describe("docx-doctor: content in the wrong container", () => {
  it("wraps a stray run in a HEADER story, not just the body", () => {
    // The setParagraphText bug could land a loose run in any story.
    const hdr = `<w:hdr ${NS}>${P()}<w:r><w:t>loose</w:t></w:r></w:hdr>`;
    const zip = pkg(P() + SECT, { "word/header1.xml": hdr },
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`);
    expect(findRule(zip, "story.stray-run")).toMatchObject({ severity: "fatal", part: "word/header1.xml" });
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/header1.xml")).toContain("<w:p><w:r><w:t>loose</w:t></w:r></w:p>");
    expect(rulesOf(zip)).not.toContain("story.stray-run");
  });

  it("reports a table nested inside a paragraph without guessing where to move it", () => {
    const body = `<w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:p>${SECT}`;
    const f = findRule(pkg(body), "paragraph.block-child");
    expect(f).toMatchObject({ severity: "fatal", fixable: false, detail: "w:tbl" });
  });

  it("accepts every legal run-level child of a paragraph", () => {
    const body =
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:bookmarkStart w:id="1" w:name="a"/>` +
      `<w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink>` +
      `<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>` +
      `<w:bookmarkEnd w:id="1"/></w:p>${SECT}`;
    expect(rulesOf(pkg(body))).not.toContain("paragraph.block-child");
  });
});

describe("docx-doctor: half-deleted fields", () => {
  it("notices a field begin with no end", () => {
    const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>x</w:t></w:r></w:p>${SECT}`;
    expect(findRule(pkg(body), "field.unbalanced")).toMatchObject({ severity: "warning", count: 1 });
  });

  it("notices a field end with no begin", () => {
    const body = `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>${SECT}`;
    expect(findRule(pkg(body), "field.unbalanced")).toMatchObject({ count: 1 });
  });

  it("accepts a well-formed caption field", () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> SEQ Table \\* ARABIC </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>${SECT}`;
    expect(rulesOf(pkg(body))).not.toContain("field.unbalanced");
  });
});

describe("docx-doctor: a reference resolving to the wrong KIND of part", () => {
  // The defect behind an unopenable combined thesis: a merge copied a paragraph
  // holding <c:chart r:id="rId7"> but not the chart part, so in the merged
  // package rId7 was an IMAGE relationship. The id resolves and image1.png
  // exists, so dangling-target, missing-relationship and full ECMA-376 schema
  // validation all pass — and Word refuses the file.
  const IMAGE_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const CHART_DRAWING =
    '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:docPr id="1" name="Chart 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId7"/>' +
    "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
  const DOC_RELS =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId7" Type="${IMAGE_TYPE}" Target="media/image1.png"/></Relationships>`;

  const chartPkg = () => {
    const zip = pkg(CHART_DRAWING + P() + SECT);
    zip.addFile("word/_rels/document.xml.rels", Buffer.from(DOC_RELS, "utf8"));
    zip.addFile("word/media/image1.png", Buffer.from([0x89, 0x50, 0x4e, 0x47]));
    return zip;
  };

  it("flags it as FATAL even though the id resolves and the part exists", () => {
    const f = findRule(chartPkg(), "rels.type-mismatch");
    expect(f).toMatchObject({ severity: "fatal", part: "word/document.xml", count: 1, fixable: true });
    expect(f!.message).toContain("chart");
  });

  it("no OTHER rule fires — this is exactly why it went unnoticed", () => {
    const rules = rulesOf(chartPkg());
    expect(rules).not.toContain("rels.dangling-target");
    expect(rules).not.toContain("rels.dangling-target-in-use");
    expect(rules).not.toContain("rels.missing-relationship");
  });

  it("leaves it alone on the automatic path (removing content is opt-in)", () => {
    const zip = chartPkg();
    inspectDocx(zip, { fix: true });
    expect(zip.readAsText("word/document.xml")).toContain("<c:chart");
  });

  it("removes the unusable drawing under the dead-links option", () => {
    const zip = chartPkg();
    const report = inspectDocx(zip, { fix: true, aggressive: true });
    expect(report.findings.find((f) => f.rule === "rels.type-mismatch")!.fixed).toBe(true);
    const out = zip.readAsText("word/document.xml");
    expect(out).not.toContain("<c:chart");
    expect(out).toContain("hello"); // the surrounding paragraph survives
  });

  // The repair keyed on the rId at first and deleted the student's picture along
  // with the chart, because the SAME rId7 is referenced correctly by a blip. Only
  // the wrongly-typed reference may be removed.
  it("removes the chart but keeps a picture sharing the same rId", () => {
    const PIC =
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
      '<a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const zip = pkg(CHART_DRAWING + PIC + P() + SECT);
    zip.addFile("word/_rels/document.xml.rels", Buffer.from(DOC_RELS, "utf8"));
    zip.addFile("word/media/image1.png", Buffer.from([0x89, 0x50, 0x4e, 0x47]));

    inspectDocx(zip, { fix: true, aggressive: true });
    const out = zip.readAsText("word/document.xml");
    expect(out).not.toContain("<c:chart");
    expect(out).toContain('<a:blip r:embed="rId7"/>');   // the picture SURVIVES
    expect(out).toContain("hello");
    // and the relationship it needs is still declared
    expect(zip.readAsText("word/_rels/document.xml.rels")).toContain('Id="rId7"');
  });

  it("does not flag a blip that correctly points at an image", () => {
    const zip = pkg(
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData>' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
      '<a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      P() + SECT,
    );
    zip.addFile("word/_rels/document.xml.rels", Buffer.from(DOC_RELS, "utf8"));
    zip.addFile("word/media/image1.png", Buffer.from([0x89, 0x50, 0x4e, 0x47]));
    expect(rulesOf(zip)).not.toContain("rels.type-mismatch");
  });
});

describe("docx-doctor: a namespace prefix nothing declares", () => {
  const ONLY_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const bare = (body: string) => {
    const zip = pkg(body);
    const xml = zip.readAsText("word/document.xml").replace(NS, ONLY_W);
    zip.updateFile("word/document.xml", Buffer.from(xml, "utf8"));
    return zip;
  };
  const withHeader = `<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`;

  it("is fatal: Word refuses the file though it is well-formed", () => {
    const zip = bare(withHeader);
    expect(findRule(zip, "xml.undeclared-prefix")).toMatchObject({ severity: "fatal", part: "word/document.xml", detail: "r" });
  });

  it("a known prefix is repaired by declaring it on the root, and nothing else moves", () => {
    const zip = bare(`<w:p><w:r><w:t>text</w:t></w:r></w:p>` + withHeader);
    const before = zip.readAsText("word/document.xml");
    inspectDocx(zip, { fix: true });
    const after = zip.readAsText("word/document.xml");
    expect(after).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    expect(after.slice(after.indexOf("<w:body>"))).toBe(before.slice(before.indexOf("<w:body>")));
    expect(rulesOf(zip)).not.toContain("xml.undeclared-prefix");
  });

  it("an unknown prefix cannot be repaired and stays fatal", () => {
    const zip = bare(`<w:p><foo:x/></w:p>${SECT}`);
    expect(findRule(zip, "xml.undeclared-prefix", { fix: true })).toMatchObject({ severity: "fatal", fixed: false, detail: "foo" });
  });

  it("reaches headers and notes, not only the document", () => {
    const zip = pkg(P() + SECT, { "word/header1.xml": `<?xml version="1.0"?><w:hdr ${ONLY_W}><w:p><w:r><w:drawing><a:x/></w:drawing></w:r></w:p></w:hdr>` },
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>');
    expect(findRule(zip, "xml.undeclared-prefix")).toMatchObject({ part: "word/header1.xml", detail: "a" });
  });

  it("a healthy package is untouched", () => {
    const zip = pkg(P() + SECT);
    expect(rulesOf(zip)).not.toContain("xml.undeclared-prefix");
  });
});

describe("docx-doctor: what a tag matcher passes and Word refuses", () => {
  it("a repeated attribute on one element is malformed", () => {
    expect(firstXmlError('<a x="1" x="2"/>')).toMatch(/repeats/);
    expect(firstXmlError('<a x="1" y="2" X="3"/>')).toBeNull();
    const zip = pkg(`<w:p w:rsidR="00A1" w:rsidR="00A2"><w:r><w:t>x</w:t></w:r></w:p>${SECT}`);
    expect(findRule(zip, "xml.malformed")).toMatchObject({ severity: "fatal", part: "word/document.xml" });
  });

  it("a value that spells an attribute is a value", () => {
    expect(firstXmlError(`<a x="y='1' y='2'"/>`)).toBeNull();
    expect(firstXmlError("<a>x='1' x='2'</a>")).toBeNull();
  });

  it("a second Default for one extension, in another case, is fatal and repaired to the first", () => {
    const zip = pkg(P() + SECT);
    const ct = zip.readAsText("[Content_Types].xml").replace("</Types>", '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Default Extension="ODTTF" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>');
    zip.updateFile("[Content_Types].xml", Buffer.from(ct, "utf8"));
    expect(findRule(zip, "contenttypes.duplicate")).toMatchObject({ severity: "fatal", detail: "odttf" });
    inspectDocx(zip, { fix: true });
    const after = zip.readAsText("[Content_Types].xml");
    expect(after).toContain('Extension="odttf"');
    expect(after).not.toContain('Extension="ODTTF"');
    expect(rulesOf(zip)).not.toContain("contenttypes.duplicate");
  });

  it("a second Override for one part is the same fault", () => {
    const zip = pkg(P() + SECT);
    const ct = zip.readAsText("[Content_Types].xml").replace("</Types>", '<Override PartName="/word/document.xml" ContentType="application/xml"/></Types>');
    zip.updateFile("[Content_Types].xml", Buffer.from(ct, "utf8"));
    expect(findRule(zip, "contenttypes.duplicate")).toMatchObject({ severity: "fatal", detail: "/word/document.xml" });
  });
});
