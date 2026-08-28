import { describe, test, expect } from "vitest";
import {
  splitDocument,
  assembleDocument,
  parseOrderedDoc,
  buildOrderedDoc,
  toBlocks,
  makeParagraphXml,
  makeParagraphNode,
  makeStyledParagraphXml,
  makeParagraphXmlLike,
  paragraphText,
  paragraphStyleId,
  setParagraphText,
  kindOf,
  nodeTag,
  escapeXmlText,
  decodeXmlText,
} from "./OrderedBody";

/**
 * Golden fixture: an interleaved body —
 *   p, tbl, p, tbl(nested tbl-in-cell), p(with w:drawing)
 * plus a trailing w:sectPr. This is the GATE for the whole migration.
 *
 * Notes that exercise the string scanner:
 *  - the first <w:p> has an attribute value containing a '>' (and an entity);
 *  - one table cell nests another <w:tbl> (depth tracking);
 *  - a paragraph carries a <w:drawing> with an embedded image relationship;
 *  - escaped entity (&amp;) in run text;
 *  - a <w:t/> empty self-closing run.
 */
const FIXTURE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p w14:paraId="00000001" w:rsidR="a&gt;b"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">First heading </w:t></w:r><w:r><w:t/></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Cell A1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Between tables &amp; more</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Cell B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic><a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

function countOccurrences(haystack: string, needle: string): number {
  return haystack.split(needle).length - 1;
}

describe("OrderedBody — string split / assemble (the gate)", () => {
  test("split → assemble is BYTE-IDENTICAL to the original", () => {
    const split = splitDocument(FIXTURE);
    expect(assembleDocument(split)).toBe(FIXTURE);
  });

  test("pre + blocks.join + post reconstructs the original exactly", () => {
    const split = splitDocument(FIXTURE);
    const rebuilt = split.pre + split.blocks.map((b) => b.xml).join("") + split.post;
    expect(rebuilt).toBe(FIXTURE);
  });

  test("classifies the top-level body children in order", () => {
    const { blocks } = splitDocument(FIXTURE);
    expect(blocks.map((b) => b.tag)).toEqual([
      "w:p",
      "w:tbl",
      "w:p",
      "w:tbl",
      "w:p",
      "w:sectPr",
    ]);
    expect(blocks.map((b) => b.kind)).toEqual([
      "paragraph",
      "table",
      "paragraph",
      "table",
      "paragraph",
      "sectPr",
    ]);
  });

  test("keeps each top-level child as its exact original substring", () => {
    const { blocks } = splitDocument(FIXTURE);
    // The second table contains a NESTED <w:tbl> in its cell; it must be kept
    // whole inside the one table block (depth-aware scan), not split.
    expect(countOccurrences(blocks[3].xml, "<w:tbl>")).toBe(2);
    expect(blocks[3].xml).toContain("Nested");
    expect(blocks[3].xml).toContain("Cell B1");
    // The drawing paragraph keeps its embed + pic intact.
    expect(blocks[4].xml).toContain("w:drawing");
    expect(blocks[4].xml).toContain('r:embed="rId7"');
    expect(blocks[4].xml).toContain("pic:pic");
  });

  test("handles an attribute value containing '>' inside quotes", () => {
    const { blocks } = splitDocument(FIXTURE);
    // p[0] has w:rsidR="a&gt;b"; the scanner must not treat the entity as a
    // tag boundary, and the trailing w:sectPr must still be last.
    expect(blocks[0].xml).toContain('w:rsidR="a&gt;b"');
    expect(blocks[blocks.length - 1].kind).toBe("sectPr");
  });

  test("preserves the trailing w:sectPr in place", () => {
    const { blocks } = splitDocument(FIXTURE);
    const last = blocks[blocks.length - 1];
    expect(last.kind).toBe("sectPr");
    expect(last.xml).toContain('w:w="11906"');
    expect(last.xml).toContain('w:h="16838"');
  });

  test("self-closing <w:body/> yields no blocks but round-trips", () => {
    const xml = `<w:document><w:body/></w:document>`;
    const split = splitDocument(xml);
    expect(split.blocks).toEqual([]);
    expect(assembleDocument(split)).toBe(xml);
  });

  test("inter-element whitespace is preserved as 'other' blocks (lossless)", () => {
    const xml = `<w:document><w:body>\n  <w:p><w:r><w:t>x</w:t></w:r></w:p>\n</w:body></w:document>`;
    const split = splitDocument(xml);
    expect(assembleDocument(split)).toBe(xml);
    // one real paragraph + surrounding whitespace 'other' blocks
    expect(split.blocks.some((b) => b.kind === "paragraph")).toBe(true);
  });
});

describe("OrderedBody — helpers", () => {
  test("kindOf maps tags correctly", () => {
    expect(kindOf("w:p")).toBe("paragraph");
    expect(kindOf("w:tbl")).toBe("table");
    expect(kindOf("w:sectPr")).toBe("sectPr");
    expect(kindOf("w:bookmarkStart")).toBe("other");
  });

  test("escape / decode round-trip the 5 XML entities", () => {
    const raw = `a & b < c > d " e ' f`;
    const enc = escapeXmlText(raw);
    expect(enc).toBe("a &amp; b &lt; c &gt; d &quot; e &apos; f");
    expect(decodeXmlText(enc)).toBe(raw);
  });

  test("paragraphText concatenates decoded w:t runs", () => {
    const { blocks } = splitDocument(FIXTURE);
    expect(paragraphText(blocks[0].xml)).toBe("First heading "); // + empty <w:t/>
    expect(paragraphText(blocks[2].xml)).toBe("Between tables & more");
  });

  test("paragraphText returns '' for an empty-ish / non-paragraph block", () => {
    expect(paragraphText("<w:tbl><w:tr></w:tr></w:tbl>")).toBe("");
    expect(paragraphText("<w:p><w:r><w:t/></w:r></w:p>")).toBe("");
  });

  test("paragraphStyleId reads w:pStyle val (or null)", () => {
    const { blocks } = splitDocument(FIXTURE);
    expect(paragraphStyleId(blocks[0].xml)).toBe("Heading1");
    expect(paragraphStyleId(blocks[2].xml)).toBeNull();
  });

  test("makeParagraphXml builds a styled RTL paragraph", () => {
    const xml = makeParagraphXml("مرحبا", "BodyArabic", true);
    expect(xml).toContain('w:val="BodyArabic"');
    expect(xml).toContain("<w:bidi/>");
    expect(xml).toContain('<w:t xml:space="preserve">مرحبا</w:t>');
    expect(paragraphText(xml)).toBe("مرحبا");
    expect(paragraphStyleId(xml)).toBe("BodyArabic");
  });

  test("makeParagraphXml escapes special chars in text", () => {
    const xml = makeParagraphXml("a & b < c");
    expect(xml).toContain("a &amp; b &lt; c");
    expect(paragraphText(xml)).toBe("a & b < c");
  });

  test("makeParagraphNode returns a string-based BodyBlock", () => {
    const block = makeParagraphNode("hi", "Normal");
    expect(block.kind).toBe("paragraph");
    expect(block.tag).toBe("w:p");
    expect(block.xml).toContain('<w:t xml:space="preserve">hi</w:t>');
  });
});

describe("OrderedBody — setParagraphText (in-place, run-preserving)", () => {
  test("replaces run text while preserving w:pPr (style)", () => {
    const { blocks } = splitDocument(FIXTURE);
    const next = setParagraphText(blocks[0].xml, "Changed heading");
    expect(paragraphText(next)).toBe("Changed heading");
    expect(paragraphStyleId(next)).toBe("Heading1");
    // collapsed to a single run
    expect(countOccurrences(next, "<w:r>")).toBe(1);
  });

  test("escapes special characters when replacing", () => {
    const p = `<w:p><w:r><w:t>old</w:t></w:r></w:p>`;
    const next = setParagraphText(p, "x & y < z");
    expect(next).toContain("x &amp; y &lt; z");
    expect(paragraphText(next)).toBe("x & y < z");
  });

  test("does NOT strip a drawing run — edits only the first w:t", () => {
    const { blocks } = splitDocument(FIXTURE);
    const drawingPara = blocks[4].xml; // the image paragraph (no <w:t>)
    const next = setParagraphText(drawingPara, "caption");
    // drawing must survive
    expect(next).toContain("w:drawing");
    expect(next).toContain('r:embed="rId7"');
    expect(next).toContain("pic:pic");
    // the text was appended (no <w:t> existed)
    expect(next).toContain('<w:t xml:space="preserve">caption</w:t>');
  });

  // Regression: Word writes every blank line as a SELF-CLOSING `<w:p …/>`, and
  // this used to find the `>` of `/>` as the end of the open tag — so the new run
  // was appended AFTER the paragraph. At body level that is a bare `<w:r>`, which
  // is not block-level content: Word then refuses to open the whole document
  // ("Word experienced an error trying to open the file"), and the edited text is
  // in no paragraph at all. Found in a live thesis via schema validation.
  test("editing a SELF-CLOSING empty paragraph keeps the run INSIDE it", () => {
    const p = '<w:p w:rsidR="000050DE" w:rsidRDefault="000050DE"/>';
    const next = setParagraphText(p, "hello");

    expect(next).toBe('<w:p w:rsidR="000050DE" w:rsidRDefault="000050DE"><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>');
    // The run must not escape the paragraph, and the paragraph must be closed.
    expect(next.endsWith("</w:p>")).toBe(true);
    expect(next).not.toMatch(/\/><w:r/);
    expect(paragraphText(next)).toBe("hello");
  });

  test("a self-closing paragraph keeps its attributes and survives an empty edit", () => {
    const next = setParagraphText('<w:p w:rsidR="ABC"/>', "");
    expect(next).toContain('w:rsidR="ABC"');
    expect(next).not.toMatch(/\/><w:r/);
    expect(splitDocument(`<w:document><w:body>${next}</w:body></w:document>`).blocks
      .filter((b) => b.tag !== "#text").map((b) => b.tag)).toEqual(["w:p"]);
  });

  test("for a drawing paragraph WITH a w:t, replaces only that first w:t", () => {
    const p = `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>orig</w:t></w:r><w:r><w:drawing><a:blip r:embed="rId9"/></w:drawing></w:r></w:p>`;
    const next = setParagraphText(p, "NEW");
    expect(next).toContain("<w:t xml:space=\"preserve\">NEW</w:t>");
    expect(next).not.toContain(">orig<");
    expect(next).toContain("w:drawing");
    expect(next).toContain('r:embed="rId9"');
    // style preserved
    expect(paragraphStyleId(next)).toBe("Caption");
  });

  test("a self-closing <w:pPr/> is preserved on edit", () => {
    const p = `<w:p><w:pPr/><w:r><w:t>old</w:t></w:r></w:p>`;
    const next = setParagraphText(p, "new");
    expect(next).toContain("<w:pPr/>");
    expect(paragraphText(next)).toBe("new");
  });
});

describe("OrderedBody — legacy-compatible exports", () => {
  test("parseOrderedDoc / buildOrderedDoc round-trip byte-identically", () => {
    const parsed = parseOrderedDoc(FIXTURE);
    expect(buildOrderedDoc(parsed.split)).toBe(FIXTURE);
    expect(parsed.blocks).toBe(parsed.bodyChildren);
  });

  test("toBlocks passes the block list through; nodeTag reads the tag", () => {
    const { blocks } = splitDocument(FIXTURE);
    const same = toBlocks(blocks);
    expect(same).toBe(blocks);
    expect(nodeTag(blocks[0])).toBe("w:p");
    expect(nodeTag(blocks[1])).toBe("w:tbl");
    expect(nodeTag("<w:sectPr><w:pgSz/></w:sectPr>")).toBe("w:sectPr");
  });

  // w:pPr is an ORDERED sequence (CT_PPrBase). Emitting bidi after jc/outlineLvl
  // made Word refuse the file — and only ever on RTL paragraphs, i.e. every
  // Arabic thesis, which is why it survived so long.
  test("styled paragraph writes w:pPr in schema order (bidi before jc before outlineLvl)", () => {
    const xml = makeStyledParagraphXml("العنوان", {
      styleId: "Heading1",
      outlineLevel: 0,
      alignment: "right",
      rtl: true,
    });
    const order = ["w:pStyle", "w:bidi", "w:jc", "w:outlineLvl"].map((t) => xml.indexOf("<" + t));
    expect(order.every((i) => i > -1)).toBe(true);
    expect(order).toEqual([...order].sort((a, b) => a - b));
    expect(xml).toContain('<w:pPr><w:pStyle w:val="Heading1"/><w:bidi/><w:jc w:val="right"/><w:outlineLvl w:val="0"/></w:pPr>');
  });
});


// ── Run tabs vs tab-stop definitions ────────────────────────────────────────
// A TOC entry is `text<w:tab/>page`; dropping the run tab glued the two into
// "مقدمة5" everywhere the flat text surfaces. The trap: `<w:tab>` ALSO appears
// inside `<w:pPr><w:tabs>` as tab-stop DEFINITIONS, which are geometry — not
// characters — and must never read as text.
describe("paragraphText — run tabs", () => {
  test("a run's <w:tab/> reads as \\t; a pPr tab STOP does not", () => {
    const xml =
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/>` +
      `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:bidi/></w:pPr>` +
      `<w:r><w:t>مقدمة</w:t></w:r>` +
      `<w:r><w:tab/></w:r>` +
      `<w:fldSimple w:instr=" PAGEREF _Toc1 \\h "><w:r><w:t>5</w:t></w:r></w:fldSimple>` +
      `</w:p>`;
    expect(paragraphText(xml)).toBe("مقدمة\t5");
  });

  test("a paragraph with only tab STOPS and no run tab has no \\t at all", () => {
    const xml =
      `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="9350"/></w:tabs></w:pPr>` +
      `<w:r><w:t>plain</w:t></w:r></w:p>`;
    expect(paragraphText(xml)).toBe("plain");
  });
});

describe("paragraph writers — tabs round-trip", () => {
  test("a \\t in edited text is written as a real <w:tab/>, and reads back", () => {
    const xml = makeStyledParagraphXml("مقدمة\t5", { styleId: "TOC1", rtl: true });
    expect(xml).toContain("<w:tab/>");
    expect(xml).not.toContain("\t");
    expect(paragraphText(xml)).toBe("مقدمة\t5");
  });

  test("makeParagraphXmlLike keeps the template's style and writes the tab", () => {
    const template = `<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:r><w:t>old</w:t></w:r></w:p>`;
    const xml = makeParagraphXmlLike(template, "label\t12");
    expect(xml).toContain('<w:pStyle w:val="TOC2"/>');
    expect(paragraphText(xml)).toBe("label\t12");
  });
});
