import { describe, test, expect } from "vitest";
import {
  upsertSectPrReference,
  upsertSectPrPageNumbering,
  upsertParagraphSectPr,
  removeParagraphSectPr,
  editBodySectPr,
  removeSectPrReferenceFromDocument,
  upsertSectPrVAlign,
  upsertSectPrPageBorders,
  SECTPR_ORDER,
} from "./sectPr";

describe("upsertSectPrPageNumbering", () => {
  test("inserts pgNumType AFTER pgSz/pgMar and BEFORE cols (CT_SectPr order)", () => {
    const start = '<w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:top="1440"/><w:cols w:space="708"/></w:sectPr>';
    const out = upsertSectPrPageNumbering(start, { format: "lowerRoman", start: 1 });
    expect(out).toBe(
      '<w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:top="1440"/>' +
        '<w:pgNumType w:fmt="lowerRoman" w:start="1"/>' +
        '<w:cols w:space="708"/></w:sectPr>',
    );
  });

  test("appends when no later-ordered sibling is present", () => {
    const out = upsertSectPrPageNumbering('<w:sectPr><w:pgSz w:w="1"/></w:sectPr>', { format: "decimal" });
    expect(out).toBe('<w:sectPr><w:pgSz w:w="1"/><w:pgNumType w:fmt="decimal"/></w:sectPr>');
  });

  test("replaces an existing pgNumType rather than duplicating it", () => {
    const start = '<w:sectPr><w:pgSz/><w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:cols/></w:sectPr>';
    const out = upsertSectPrPageNumbering(start, { format: "decimal", start: 1 });
    expect((out.match(/w:pgNumType/g) ?? []).length).toBe(1);
    expect(out).toContain('<w:pgNumType w:fmt="decimal" w:start="1"/>');
    expect(out.indexOf("w:pgNumType")).toBeLessThan(out.indexOf("w:cols"));
  });

  test("omitting start leaves the section continuing the previous sequence", () => {
    const out = upsertSectPrPageNumbering("<w:sectPr><w:cols/></w:sectPr>", { format: "decimal" });
    expect(out).toBe('<w:sectPr><w:pgNumType w:fmt="decimal"/><w:cols/></w:sectPr>');
    expect(out).not.toContain("w:start");
  });

  test("expands a self-closing sectPr, and no-ops with nothing to set", () => {
    expect(upsertSectPrPageNumbering("<w:sectPr/>", { start: 5 })).toBe(
      '<w:sectPr><w:pgNumType w:start="5"/></w:sectPr>',
    );
    expect(upsertSectPrPageNumbering("<w:sectPr/>", {})).toBe("<w:sectPr/>");
  });
});

describe("sectPr string helpers", () => {
  test("upsertSectPrReference inserts a header ref at the start of the sectPr", () => {
    const out = upsertSectPrReference("<w:sectPr><w:pgSz w:w=\"100\"/></w:sectPr>", "header", "default", "rId9");
    expect(out).toBe('<w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:pgSz w:w="100"/></w:sectPr>');
  });

  test("upsertSectPrReference replaces an existing ref of the same type", () => {
    const start = '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz/></w:sectPr>';
    const out = upsertSectPrReference(start, "header", "default", "rId2");
    expect(out).toContain('r:id="rId2"');
    expect(out).not.toContain('r:id="rId1"');
    expect((out.match(/w:headerReference/g) ?? []).length).toBe(1);
  });

  test("upsertSectPrReference expands a self-closing sectPr", () => {
    const out = upsertSectPrReference("<w:sectPr/>", "footer", "default", "rId5");
    expect(out).toBe('<w:sectPr><w:footerReference w:type="default" r:id="rId5"/></w:sectPr>');
  });

  test("upsertParagraphSectPr adds a pPr+sectPr to a bare paragraph", () => {
    const out = upsertParagraphSectPr("<w:p><w:r><w:t>x</w:t></w:r></w:p>", "<w:sectPr><w:type w:val=\"nextPage\"/></w:sectPr>");
    expect(out).toBe('<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>');
  });

  test("upsertParagraphSectPr appends sectPr inside an existing pPr (after other props)", () => {
    const out = upsertParagraphSectPr('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r/></w:p>', "<w:sectPr/>");
    expect(out).toBe('<w:p><w:pPr><w:jc w:val="center"/><w:sectPr/></w:pPr><w:r/></w:p>');
  });

  // A section break is usually applied to the blank line that ends a page, and
  // Word writes an empty paragraph SELF-CLOSING. Splicing after the first `>`
  // put the pPr after a tag that had already closed, so the break landed as a
  // bare <w:pPr> child of <w:body>: Word refuses the file and the app drew an
  // unknown-block chip where the page break belonged.
  test("upsertParagraphSectPr puts the sectPr INSIDE a self-closing <w:p/>", () => {
    const out = upsertParagraphSectPr('<w:p w:rsidR="00A1" w:rsidRDefault="00A1"/>', '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>');
    expect(out).toBe(
      '<w:p w:rsidR="00A1" w:rsidRDefault="00A1"><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>',
    );
    // The regression itself: nothing may follow the paragraph's close tag.
    expect(out.endsWith("</w:p>")).toBe(true);
    expect(out).not.toContain("/><w:pPr>");
  });

  test("upsertParagraphSectPr handles the bare self-closing <w:p/>", () => {
    expect(upsertParagraphSectPr("<w:p/>", "<w:sectPr/>")).toBe("<w:p><w:pPr><w:sectPr/></w:pPr></w:p>");
  });

  test("removeParagraphSectPr strips the sectPr only", () => {
    const out = removeParagraphSectPr('<w:p><w:pPr><w:jc w:val="left"/><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>');
    expect(out).toBe('<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>');
  });

  test("editBodySectPr targets the body sectPr and leaves tables before it untouched", () => {
    const doc = '<w:document><w:body><w:p/><w:tbl><w:tr/></w:tbl><w:sectPr><w:pgSz/></w:sectPr></w:body></w:document>';
    const out = editBodySectPr(doc, (s) => upsertSectPrReference(s, "header", "default", "rIdH"));
    expect(out.indexOf("<w:tbl>")).toBeLessThan(out.indexOf("<w:sectPr>"));
    expect(out).toContain('<w:headerReference w:type="default" r:id="rIdH"/>');
  });

  test("removeSectPrReferenceFromDocument deletes only the matching ref tag", () => {
    const doc = '<w:document><w:body><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>';
    const out = removeSectPrReferenceFromDocument(doc, "header", "rId1");
    expect(out).not.toContain('r:id="rId1"');
    expect(out).toContain('r:id="rId2"'); // footer ref untouched
  });
});

describe("SECTPR_ORDER", () => {
  test("headerReference comes first and pgBorders precedes vAlign", () => {
    expect(SECTPR_ORDER[0]).toBe("w:headerReference");
    const borders = SECTPR_ORDER.indexOf("w:pgBorders");
    const valign = SECTPR_ORDER.indexOf("w:vAlign");
    expect(borders).toBeGreaterThan(-1);
    expect(borders).toBeLessThan(valign);
  });
});

describe("upsertSectPrVAlign", () => {
  test("inserts vAlign into an empty sectPr", () => {
    expect(upsertSectPrVAlign("<w:sectPr></w:sectPr>", "center")).toBe(
      '<w:sectPr><w:vAlign w:val="center"/></w:sectPr>',
    );
  });

  test("expands a self-closing sectPr", () => {
    expect(upsertSectPrVAlign("<w:sectPr/>", "center")).toBe(
      '<w:sectPr><w:vAlign w:val="center"/></w:sectPr>',
    );
  });

  test("replaces an existing vAlign instead of duplicating", () => {
    const out = upsertSectPrVAlign('<w:sectPr><w:vAlign w:val="top"/></w:sectPr>', "center");
    expect(out).toContain('w:val="center"');
    expect(out.match(/<w:vAlign/g)).toHaveLength(1);
  });

  test("also replaces a paired-tag vAlign", () => {
    const out = upsertSectPrVAlign('<w:sectPr><w:vAlign w:val="top"></w:vAlign></w:sectPr>', "center");
    expect(out).toContain('w:val="center"');
    expect(out.match(/<w:vAlign/g)).toHaveLength(1);
  });

  test("lands after pgMar/cols and before titlePg — CT_SectPr is an ordered sequence", () => {
    const sect =
      '<w:sectPr><w:pgSz w:w="11906"/><w:pgMar w:top="1440"/><w:cols w:space="708"/><w:titlePg/></w:sectPr>';
    const out = upsertSectPrVAlign(sect, "center");
    expect(out.indexOf("<w:vAlign")).toBeGreaterThan(out.indexOf("<w:cols"));
    expect(out.indexOf("<w:vAlign")).toBeLessThan(out.indexOf("<w:titlePg"));
  });

  test("every other byte is preserved", () => {
    const sect =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440"/></w:sectPr>';
    const out = upsertSectPrVAlign(sect, "bottom");
    expect(out.replace(/<w:vAlign[^>]*\/>/, "")).toBe(sect);
  });
});

describe("upsertSectPrPageBorders", () => {
  test("draws all four edges with offsetFrom=page", () => {
    const out = upsertSectPrPageBorders("<w:sectPr></w:sectPr>", {
      style: "double", color: "6B4F2A", widthPt: 3,
    });
    expect(out).toContain('<w:pgBorders w:offsetFrom="page">');
    for (const edge of ["top", "left", "bottom", "right"]) {
      expect(out).toMatch(new RegExp(`<w:${edge} w:val="double" w:sz="24" w:space="24" w:color="6B4F2A"/>`));
    }
  });

  test("width is stored in EIGHTHS of a point, clamped to a minimum of 2", () => {
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "000000", widthPt: 1.5 }))
      .toContain('w:sz="12"');
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "000000", widthPt: 0.1 }))
      .toContain('w:sz="2"');
  });

  test("colour is normalised: leading # stripped, uppercased", () => {
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "#6b4f2a", widthPt: 1 }))
      .toContain('w:color="6B4F2A"');
  });

  test("offsetPt overrides the default space of 24", () => {
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "000000", widthPt: 1, offsetPt: 12 }))
      .toContain('w:space="12"');
  });

  test("replaces an existing pgBorders (paired form) instead of duplicating", () => {
    const existing =
      '<w:sectPr><w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/></w:pgBorders></w:sectPr>';
    const out = upsertSectPrPageBorders(existing, { style: "double", color: "6B4F2A", widthPt: 3 });
    expect(out.match(/<w:pgBorders/g)).toHaveLength(1);
    expect(out).not.toContain("FF0000");
  });

  test("lands after pgMar and BEFORE pgNumType, cols and vAlign — schema order", () => {
    const sect =
      '<w:sectPr><w:pgMar w:top="1440"/><w:pgNumType w:fmt="decimal"/><w:cols w:space="708"/><w:vAlign w:val="center"/></w:sectPr>';
    const out = upsertSectPrPageBorders(sect, { style: "single", color: "000000", widthPt: 1 });
    const b = out.indexOf("<w:pgBorders");
    expect(b).toBeGreaterThan(out.indexOf("<w:pgMar"));
    expect(b).toBeLessThan(out.indexOf("<w:pgNumType"));
    expect(b).toBeLessThan(out.indexOf("<w:cols"));
    expect(b).toBeLessThan(out.indexOf("<w:vAlign"));
  });

  test("handles attributes on the sectPr open tag (both forms)", () => {
    const paired = upsertSectPrPageBorders('<w:sectPr w:rsidR="00AB12CD"></w:sectPr>', {
      style: "single", color: "000000", widthPt: 1,
    });
    expect(paired).toContain('<w:sectPr w:rsidR="00AB12CD">');
    expect(paired).toContain("<w:pgBorders");
    const selfClosing = upsertSectPrVAlign('<w:sectPr w:rsidR="00AB12CD"/>', "center");
    expect(selfClosing).toBe('<w:sectPr w:rsidR="00AB12CD"><w:vAlign w:val="center"/></w:sectPr>');
  });

  test("every other byte is preserved", () => {
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/></w:sectPr>';
    const out = upsertSectPrPageBorders(sect, { style: "single", color: "000000", widthPt: 1 });
    expect(out.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/, "")).toBe(sect);
  });

  test("sz is clamped to the schema ceiling of 96 eighths", () => {
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "000000", widthPt: 20 }))
      .toContain('w:sz="96"');
  });

  test("offsetPt is rounded to an integer", () => {
    expect(upsertSectPrPageBorders("<w:sectPr/>", { style: "single", color: "000000", widthPt: 1, offsetPt: 12.5 }))
      .toContain('w:space="13"');
  });
});

// A tracked paragraph-property change holds the OLD properties in a nested w:pPr, so the first
// `</w:pPr>` in the paragraph is that nested one. The section break must land in the paragraph's
// own w:pPr, ahead of w:pPrChange (CT_PPr: rPr, sectPr, pPrChange), or Word reads no section at all.
describe("upsertParagraphSectPr with a tracked paragraph-property change", () => {
  const CHANGE = '<w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>';
  const SECT = '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>';

  test("the break goes into the paragraph's own pPr, before the pPrChange", () => {
    const out = upsertParagraphSectPr(`<w:p><w:pPr><w:jc w:val="center"/>${CHANGE}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`, SECT);
    expect(out).toBe(`<w:p><w:pPr><w:jc w:val="center"/>${SECT}${CHANGE}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`);
  });

  test("an existing break is replaced in place, and the old properties stay as they were", () => {
    const out = upsertParagraphSectPr(`<w:p><w:pPr><w:sectPr><w:cols/></w:sectPr>${CHANGE}</w:pPr></w:p>`, SECT);
    expect(out).toBe(`<w:p><w:pPr>${SECT}${CHANGE}</w:pPr></w:p>`);
  });

  test("a paragraph with no pPr of its own but a text box gets its own, not the text box's", () => {
    const box = '<w:r><w:txbxContent><w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p></w:txbxContent></w:r>';
    const out = upsertParagraphSectPr(`<w:p>${box}</w:p>`, SECT);
    expect(out).toBe(`<w:p><w:pPr>${SECT}</w:pPr>${box}</w:p>`);
  });
});
