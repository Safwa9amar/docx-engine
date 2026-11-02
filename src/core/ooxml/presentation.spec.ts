import { describe, test, expect } from "vitest";
import { wipeBlockPresentation, KEPT_RUN_PROPS, REMOVED_PARA_PROPS } from "./presentation";
import { Mdocxengine } from "../../index";

const text = (xml: string) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
const wipe = (x: string, o?: Parameters<typeof wipeBlockPresentation>[1]) => wipeBlockPresentation(x, o);

describe("wipeBlockPresentation — runs", () => {
  test("fonts, size and colour go; b/i/u/vertAlign/rtl/cs/lang/rStyle stay, in order", () => {
    const r = wipe(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="40"/><w:szCs w:val="40"/><w:color w:val="FF0000"/><w:b/><w:i/><w:u w:val="single"/><w:vertAlign w:val="superscript"/><w:rtl/><w:cs/><w:lang w:val="ar-DZ"/><w:rStyle w:val="Strong"/><w:highlight w:val="yellow"/></w:rPr><w:t>hi</w:t></w:r></w:p>',
    );
    expect(r.xml).toBe(
      '<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/><w:i/><w:u w:val="single"/><w:vertAlign w:val="superscript"/><w:rtl/><w:cs/><w:lang w:val="ar-DZ"/></w:rPr><w:t>hi</w:t></w:r></w:p>',
    );
    expect(r.runsStripped).toBe(1);
  });

  test("an rPr left empty is dropped; an untouched run is byte-identical", () => {
    const r = wipe('<w:p><w:r><w:rPr><w:sz w:val="40"/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r><w:r><w:rPr/><w:t>c</w:t></w:r><w:r w:rsidR="1"/></w:p>');
    expect(r.xml).toBe('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r><w:r><w:t>c</w:t></w:r><w:r w:rsidR="1"/></w:p>');
    expect(r.runsStripped).toBe(2);
  });

  test("tracked-change marks in an rPr are revision data and survive", () => {
    const r = wipe('<w:p><w:r><w:rPr><w:ins w:id="1" w:author="a"/><w:sz w:val="2"/></w:rPr><w:t>x</w:t></w:r></w:p>');
    expect(r.xml).toContain('<w:ins w:id="1" w:author="a"/>');
    expect(r.xml).not.toContain("w:sz");
  });
});

describe("wipeBlockPresentation — paragraphs", () => {
  test("look goes; pStyle/numPr/bidi/outlineLvl/mark rPr stay, in canonical order", () => {
    const r = wipe(
      '<w:p><w:pPr><w:jc w:val="both"/><w:outlineLvl w:val="0"/><w:spacing w:after="0"/><w:ind w:left="1"/><w:keepNext/><w:bidi/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:pStyle w:val="H1"/><w:rPr><w:sz w:val="9"/><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    );
    expect(r.xml).toBe(
      '<w:p><w:pPr><w:pStyle w:val="H1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:bidi/><w:outlineLvl w:val="0"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
    );
    expect(r.paragraphsStripped).toBe(1);
  });

  test("a pPr with only look is dropped; a pPr-less, self-closing or already-clean paragraph is untouched", () => {
    const x = '<w:p w:rsidR="1"/>';
    expect(wipe(x).xml).toBe(x);
    expect(wipe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>q</w:t></w:r></w:p>').xml).toBe('<w:p><w:r><w:t>q</w:t></w:r></w:p>');
    const clean = '<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>q</w:t></w:r></w:p>';
    expect(wipe(clean)).toMatchObject({ xml: clean, paragraphsStripped: 0 });
    const empty = '<w:p><w:pPr/><w:r><w:t>q</w:t></w:r></w:p>';
    expect(wipe(empty).xml).toBe(empty);
  });

  test("paragraph sectPr is kept unless asked", () => {
    const x = '<w:p><w:pPr><w:jc w:val="left"/><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr></w:p>';
    expect(wipe(x).xml).toBe('<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr></w:p>');
    const r = wipe(x, { removeParagraphSectPr: true });
    expect(r.xml).toBe("<w:p></w:p>");
    expect(r.sectPrRemoved).toBe(1);
  });

  test("a self-closing mark rPr is dropped", () => {
    expect(wipe('<w:p><w:pPr><w:pStyle w:val="A"/><w:rPr/></w:pPr></w:p>').xml).toBe('<w:p><w:pPr><w:pStyle w:val="A"/></w:pPr></w:p>');
  });
});

describe("wipeBlockPresentation — breaks", () => {
  test("page breaks and lastRenderedPageBreak go; line and column breaks stay; emptied run removed", () => {
    const r = wipe(
      '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:lastRenderedPageBreak/><w:t>a</w:t><w:br/><w:br w:type="column"/><w:br w:type="textWrapping"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"></w:br></w:r></w:p>',
    );
    expect(r.xml).toBe(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t><w:br/><w:br w:type="column"/><w:br w:type="textWrapping"/></w:r></w:p>',
    );
    expect(r.pageBreaksRemoved).toBe(3);
  });

  test("removePageBreaks:false keeps them", () => {
    const x = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    expect(wipe(x, { removePageBreaks: false })).toMatchObject({ xml: x, pageBreaksRemoved: 0 });
  });
});

describe("wipeBlockPresentation — never edits protected spans", () => {
  const dirty = '<w:rPr><w:sz w:val="50"/><w:color w:val="00FF00"/></w:rPr>';
  const pdirty = '<w:pPr><w:jc w:val="center"/><w:spacing w:after="9"/></w:pPr>';

  test("text boxes (also inside AlternateContent and drawing) are byte-identical", () => {
    const tx = `<w:txbxContent><w:p>${pdirty}<w:r>${dirty}<w:t>in box</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p></w:txbxContent>`;
    const x = `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx>${tx}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:textbox>${tx}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
    expect(wipe(x).xml).toBe(x);
    const pict = `<w:p><w:r><w:pict><v:textbox>${tx}</v:textbox></w:pict></w:r></w:p>`;
    expect(wipe(pict).xml).toBe(pict);
  });

  test("a run that carries a drawing still has its OWN rPr wiped, the drawing is untouched", () => {
    const d = `<w:drawing><wp:inline><a:graphic><a:r>${dirty}</a:r></a:graphic></wp:inline></w:drawing>`;
    const r = wipe(`<w:p><w:r>${dirty}${d}</w:r></w:p>`);
    expect(r.xml).toBe(`<w:p><w:r>${d}</w:r></w:p>`);
  });

  test("field instructions, fldSimple instr and bookmarks are unchanged", () => {
    const x =
      `<w:p><w:bookmarkStart w:id="1" w:name="b"/><w:r>${dirty}<w:instrText xml:space="preserve"> PAGEREF b \\h </w:instrText></w:r>` +
      `<w:fldSimple w:instr=" PAGE "><w:r>${dirty}<w:t>1</w:t></w:r></w:fldSimple><w:bookmarkEnd w:id="1"/></w:p>`;
    const r = wipe(x);
    expect(r.xml).toContain('<w:instrText xml:space="preserve"> PAGEREF b \\h </w:instrText>');
    expect(r.xml).toContain('<w:fldSimple w:instr=" PAGE ">');
    expect(r.xml).toContain('<w:bookmarkStart w:id="1" w:name="b"/>');
    expect(r.xml).toContain('<w:bookmarkEnd w:id="1"/>');
    expect(r.xml).not.toContain("w:sz");
  });

  test("math keeps its own rPr", () => {
    const m = `<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/><w:sz w:val="30"/></w:rPr><m:t>x</m:t></m:r></m:oMath>`;
    const x = `<w:p>${m}<m:oMathPara>${m}</m:oMathPara></w:p>`;
    expect(wipe(x).xml).toBe(x);
  });

  test("a self-closing protected element is carried through", () => {
    const x = '<w:p><w:r><w:rPr><w:sz w:val="9"/></w:rPr><w:instrText/><w:drawing/><w:t>a</w:t></w:r></w:p>';
    expect(wipe(x).xml).toBe('<w:p><w:r><w:instrText/><w:drawing/><w:t>a</w:t></w:r></w:p>');
  });

  test("an unclosed protected element is left to the normal rules instead of eating the block", () => {
    const x = '<w:p><w:r><w:rPr><w:sz w:val="9"/></w:rPr><w:drawing></w:r></w:p>';
    expect(wipe(x).xml).toBe('<w:p><w:r><w:drawing></w:r></w:p>');
  });
});

describe("wipeBlockPresentation — tables", () => {
  test("cell paragraphs are cleaned; tblPr/trPr/tcPr are untouched", () => {
    const x =
      '<w:tbl><w:tblPr><w:tblStyle w:val="T"/><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="9"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:jc w:val="left"/></w:trPr><w:tc><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="9"/><w:b/></w:rPr><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const r = wipe(x);
    expect(r.xml).toBe(
      x.replace('<w:pPr><w:jc w:val="center"/></w:pPr>', "").replace('<w:sz w:val="9"/>', ""),
    );
    expect(r.paragraphsStripped).toBe(1);
    expect(r.runsStripped).toBe(1);
  });
});

describe("constants", () => {
  test("exact lists", () => {
    expect(KEPT_RUN_PROPS).toEqual(["b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vertAlign", "rtl", "cs", "lang", "rStyle", "vanish", "specVanish", "noProof"]);
    expect(REMOVED_PARA_PROPS).toContain("textAlignment");
  });
});

describe("samples/2.docx", () => {
  test("text is identical before and after, on every block, and the result still loads", async () => {
    const e = await Mdocxengine.loadFromFile("samples/2.docx");
    const blocks = await e.document.getBlocks();
    expect(blocks.length).toBeGreaterThan(10);
    let changed = 0;
    for (const b of blocks) {
      const r = wipeBlockPresentation(b.xml, { removeParagraphSectPr: false });
      expect(text(r.xml), `block ${b.tag}`).toBe(text(b.xml));
      if (r.xml !== b.xml) changed++;
    }
    expect(changed).toBeGreaterThan(0);
    await e.document.saveBlocks(blocks.map((b) => ({ ...b, xml: wipeBlockPresentation(b.xml).xml })));
    const re = await Mdocxengine.loadFromBuffer(e.zip.toBuffer());
    expect((await re.document.getBlocks()).length).toBe(blocks.length);
  });
});
