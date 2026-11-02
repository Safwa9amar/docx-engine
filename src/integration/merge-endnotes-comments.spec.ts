import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, inspectDocx, type BodyBlock, type DocxZip } from "../index";

/**
 * Endnotes and comments across a merge.
 *
 * Endnotes travel the way footnotes do: the note is copied verbatim under a fresh
 * id, and its images and links through its own part's relationships.
 *
 * Comments do not travel at all: a thesis going to deposit never carries the
 * supervisor's balloons, and a `w:commentReference` left behind with no comments
 * part makes real Word refuse the file. So their anchors are stripped and counted.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = (t: string) => `${R_NS}/${t}`;

const para = (xml: string): BodyBlock => ({ kind: "paragraph", tag: "w:p", xml });
const text = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

const noteRun = (kind: "endnote" | "footnote", id: number | string, extra = "") =>
  `<w:r><w:rPr><w:rStyle w:val="${kind === "endnote" ? "Endnote" : "Footnote"}Reference"/></w:rPr><w:${kind}Reference ${extra}w:id="${id}"/></w:r>`;
const refPara = (kind: "endnote" | "footnote", id: number | string) => `<w:p><w:r><w:t>body </w:t></w:r>${noteRun(kind, id)}</w:p>`;

const notesPart = (kind: "endnote" | "footnote", notes: Record<string, string>, rootAttrs = "") =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${kind}s xmlns:w="${W_NS}" xmlns:r="${R_NS}"${rootAttrs ? " " + rootAttrs : ""}>` +
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
  Object.entries(notes).map(([id, inner]) => `<w:${kind} w:id="${id}">${inner}</w:${kind}>`).join("") +
  `</w:${kind}s>`;

const noteBody = (t: string, ppr = "") =>
  `<w:p>${ppr}<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${t}</w:t></w:r></w:p>`;

interface DocOpts {
  endnotes?: Record<string, string>;
  footnotes?: Record<string, string>;
  rootAttrs?: string;
}

async function doc(blocks: string[], o: DocOpts = {}): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  if (o.endnotes) {
    await e.endnotes.addEndnote("seed"); // registers the part, its content type and relationship
    e.zip.addFile("word/endnotes.xml", Buffer.from(notesPart("endnote", o.endnotes, o.rootAttrs), "utf8"));
  }
  if (o.footnotes) {
    e.zip.addFile("word/footnotes.xml", Buffer.from(notesPart("footnote", o.footnotes, o.rootAttrs), "utf8"));
  }
  await e.document.saveBlocks(blocks.map(para));
  return e;
}
const buf = async (blocks: string[], o: DocOpts = {}) => (await doc(blocks, o)).zip.toBuffer();

const bodyOf = async (e: Mdocxengine) => (await e.document.getBlocks()).map((b) => b.xml).join("");
const part = (e: Mdocxengine, name: string) => e.zip.readAsText(name) ?? "";
const refIds = (xml: string, kind: string) => [...xml.matchAll(new RegExp(`<w:${kind}Reference\\b[^>]*\\bw:id="([^"]+)"`, "g"))].map((m) => m[1]);
/** The text a real note holds, by id. */
const noteTexts = (xml: string, kind: string): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const m of xml.matchAll(new RegExp(`<w:${kind}\\b([^>]*)>([\\s\\S]*?)</w:${kind}>`, "g"))) {
    if (/w:type=/.test(m[1])) continue;
    out[/\bw:id="([^"]+)"/.exec(m[1])![1]] = [...m[2].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((x) => x[1]).join("").trim();
  }
  return out;
};
const fatals = (e: Mdocxengine) => inspectDocx(e.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal").map((f) => f.rule);

describe("merge: endnotes survive the way footnotes do", () => {
  test("a target with no endnotes part gets one, registered, with its separators", async () => {
    const t = await doc([text("target")]);
    expect(t.zip.getEntry("word/endnotes.xml")).toBeNull();
    await t.merge.appendDocument(await buf([refPara("endnote", 7)], { endnotes: { 7: noteBody("first note") } }));

    const notes = part(t, "word/endnotes.xml");
    expect(notes).toMatch(/<w:endnote\b[^>]*w:type="separator"/);
    expect(notes).toMatch(/<w:endnote\b[^>]*w:type="continuationSeparator"/);
    expect(part(t, "[Content_Types].xml")).toContain('PartName="/word/endnotes.xml"');
    expect(part(t, "word/_rels/document.xml.rels")).toContain("relationships/endnotes");
    const refs = refIds(await bodyOf(t), "endnote");
    expect(refs).toHaveLength(1);
    expect(noteTexts(notes, "endnote")[refs[0]]).toBe("first note");
    expect(fatals(t)).toEqual([]);
  });

  test("two parts using the same id keep their own notes, and the target's stays put", async () => {
    const t = await doc([refPara("endnote", 3)], { endnotes: { 3: noteBody("target note") } });
    await t.merge.appendDocument(await buf([refPara("endnote", 1)], { endnotes: { 1: noteBody("part two note") } }));
    await t.merge.appendDocument(await buf([refPara("endnote", 1), refPara("endnote", 2)], { endnotes: { 1: noteBody("part three a"), 2: noteBody("part three b") } }));

    const refs = refIds(await bodyOf(t), "endnote");
    const texts = noteTexts(part(t, "word/endnotes.xml"), "endnote");
    expect(refs.map((r) => texts[r])).toEqual(["target note", "part two note", "part three a", "part three b"]);
    expect(refs[0]).toBe("3");
    expect(new Set(refs).size).toBe(4);
  });

  test("a note is copied verbatim, separators never, unreferenced notes not at all", async () => {
    const rich = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>BOLD-MARKER</w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>';
    const t = await doc([text("target")]);
    await t.merge.appendDocument(await buf([refPara("endnote", 5)], { endnotes: { 5: rich, 6: noteBody("orphan") } }));
    const notes = part(t, "word/endnotes.xml");
    expect(notes).toContain(rich);
    expect(notes).not.toContain("orphan");
    expect(notes.match(/w:type="separator"/g)).toHaveLength(1);
  });

  test.each(["footnote", "endnote"] as const)("a self-closing %s does not swallow the note after it", async (kind) => {
    const src = await doc([refPara(kind, 1), refPara(kind, 2)], { [`${kind}s`]: { 2: noteBody("second") } });
    const name = `word/${kind}s.xml`;
    const xml = part(src, name);
    src.zip.addFile(name, Buffer.from(xml.replace(`<w:${kind} w:id="2">`, `<w:${kind} w:id="1"/><w:${kind} w:id="2">`), "utf8"));
    const t = await doc([text("target")]);
    await t.merge.appendDocument(src.zip.toBuffer());
    const refs = refIds(await bodyOf(t), kind);
    expect(noteTexts(part(t, name).replace(new RegExp(`<w:${kind} w:id="\\d+"/>`, "g"), ""), kind)[refs[1]]).toBe("second");
    expect(part(t, name)).toContain(`<w:${kind} w:id="${refs[0]}"/>`);
  });

  test("the same id as a footnote is a different note: each kind remaps on its own", async () => {
    const t = await doc([text("target")], { footnotes: { 1: noteBody("target footnote") } });
    await t.merge.appendDocument(
      await buf([`<w:p>${noteRun("footnote", 1)}${noteRun("endnote", 1)}</w:p>`], {
        footnotes: { 1: noteBody("src footnote") },
        endnotes: { 1: noteBody("src endnote") },
      }),
    );
    const body = await bodyOf(t);
    const fn = noteTexts(part(t, "word/footnotes.xml"), "footnote");
    const en = noteTexts(part(t, "word/endnotes.xml"), "endnote");
    expect(fn[refIds(body, "footnote")[0]]).toBe("src footnote");
    expect(en[refIds(body, "endnote")[0]]).toBe("src endnote");
    expect(refIds(body, "footnote")[0]).toBe("2");
    expect(refIds(body, "endnote")[0]).toBe("1");
  });

  test("attributes beside the id survive and both shape branches remap", async () => {
    const shape =
      `<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:txbxContent><w:p>${noteRun("endnote", 9, 'w:customMarkFollows="1" ')}</w:p></w:txbxContent></mc:Choice>` +
      `<mc:Fallback><w:txbxContent><w:p>${noteRun("endnote", 9, 'w:customMarkFollows="1" ')}</w:p></w:txbxContent></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
    const t = await doc([refPara("endnote", 1)], { endnotes: { 1: noteBody("taken") } });
    await t.merge.appendDocument(await buf([shape], { endnotes: { 9: noteBody("in a box") } }));
    const refs = refIds(await bodyOf(t), "endnote");
    expect(refs.slice(1)).toEqual(["2", "2"]);
    expect(await bodyOf(t)).toContain('w:customMarkFollows="1" w:id="2"');
    expect(noteTexts(part(t, "word/endnotes.xml"), "endnote")["2"]).toBe("in a box");
  });

  test("a note's hyperlink and picture go through the endnotes part's own relationships", async () => {
    const src = await doc([refPara("endnote", 1)], {
      endnotes: {
        1: `<w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdPic"/></w:drawing></w:r></w:p>`,
      },
    });
    src.zip.addFile("word/media/note-pic.png", Buffer.from([0x89, 0x50, 0x4e, 0x47, 7, 7, 7]));
    src.zip.addFile(
      "word/_rels/endnotes.xml.rels",
      Buffer.from(
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdLink" Type="${REL("hyperlink")}" Target="https://a.example/?a=1&amp;b=2" TargetMode="External"/>` +
          `<Relationship Id="rIdPic" Type="${REL("image")}" Target="media/note-pic.png"/></Relationships>`,
      ),
    );
    const t = await doc([text("target")]);
    await t.merge.appendDocument(src.zip.toBuffer());

    const notes = part(t, "word/endnotes.xml");
    const rels = part(t, "word/_rels/endnotes.xml.rels");
    const linkId = /<w:hyperlink r:id="([^"]+)"/.exec(notes)![1];
    const picId = /r:embed="([^"]+)"/.exec(notes)![1];
    expect(linkId).not.toBe("rIdLink");
    expect(picId).not.toBe("rIdPic");
    expect(new RegExp(`<Relationship\\b[^>]*Id="${linkId}"[^>]*`).exec(rels)![0]).toContain('TargetMode="External"');
    expect(rels).toContain("https://a.example/?a=1&amp;b=2"); // escaped once, not twice
    const picTarget = /Target="([^"]+)"/.exec(new RegExp(`<Relationship\\b[^>]*Id="${picId}"[^>]*`).exec(rels)![0])![1];
    expect(t.zip.getEntry(`word/${picTarget}`)!.getData().equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 7, 7, 7]))).toBe(true);
    expect(fatals(t)).toEqual([]);
  });

  test.each(["footnote", "endnote"] as const)("an OLE object inside a %s keeps its own relationship type and part", async (kind) => {
    const NS_O = 'xmlns:o="urn:schemas-microsoft-com:office:office"';
    const src = await doc([refPara(kind, 1)], { [`${kind}s`]: { 1: `<w:p><w:r><w:object><o:OLEObject ${NS_O} Type="Embed" ProgID="Equation.3" r:id="rIdOle"/></w:object></w:r></w:p>` } });
    const relsPath = `word/_rels/${kind}s.xml.rels`;
    src.zip.addFile("word/embeddings/oleObject1.bin", Buffer.from([0xd0, 0xcf, 1, 2]));
    src.zip.addFile(relsPath, Buffer.from(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOle" Type="${REL("oleObject")}" Target="embeddings/oleObject1.bin"/></Relationships>`));
    const t = await doc([text("target")]);
    await t.merge.appendDocument(src.zip.toBuffer());
    const id = /<o:OLEObject\b[^>]*\br:id="([^"]+)"/.exec(part(t, `word/${kind}s.xml`))![1];
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*`).exec(part(t, relsPath))![0];
    expect(rel).toContain(REL("oleObject"));
    const target = /Target="([^"]+)"/.exec(rel)![1];
    expect(target.startsWith("embeddings/")).toBe(true);
    expect(t.zip.getEntry(`word/${target}`)!.getData().equals(Buffer.from([0xd0, 0xcf, 1, 2]))).toBe(true);
    expect(t.zip.getEntries().some((e) => e.entryName.startsWith("word/media/"))).toBe(false);
  });

  test("namespaces the source's endnotes root declares reach the target's root", async () => {
    const extra = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
    const t = await doc([refPara("endnote", 1)], { endnotes: { 1: noteBody("a") } });
    await t.merge.appendDocument(
      await buf([refPara("endnote", 1)], {
        endnotes: { 1: '<w:p w16du:x="1"><w:r><w:t>b</w:t></w:r></w:p>' },
        rootAttrs: extra,
      }),
    );
    expect(/<w:endnotes\b[^>]*>/.exec(part(t, "word/endnotes.xml"))![0]).toContain(extra);
  });

  test("a fresh endnotes part declares the standard prefixes a Word note can use", async () => {
    const t = await doc([text("target")]);
    await t.merge.appendDocument(await buf([refPara("endnote", 1)], { endnotes: { 1: noteBody("a") } }));
    const root = /<w:endnotes\b[^>]*>/.exec(part(t, "word/endnotes.xml"))![0];
    for (const px of ["w", "r", "mc", "w14", "wp", "a", "pic"]) expect(root).toContain(`xmlns:${px}=`);
  });

  test.each(["footnote", "endnote"] as const)("a %s reference with no note in the source throws: Word refuses such a file, and it could bind to a stranger's note", async (kind) => {
    const t = await doc([refPara(kind, 1)], { [`${kind}s`]: { 1: noteBody("target's note") } });
    const before = part(t, `word/${kind}s.xml`);
    const src = await buf([refPara(kind, 1)], { [`${kind}s`]: {} });
    await expect(t.merge.appendDocument(src)).rejects.toThrow(new RegExp(`${kind} reference.* 1 .*no ${kind}`));
    expect(part(t, `word/${kind}s.xml`)).toBe(before);
  });
});

describe("merge: a notes part that cannot take the notes", () => {
  test.each(["footnote", "endnote"] as const)("a %s part with a self-closed root refuses, rather than leave the reference dangling", async (kind) => {
    const t = await doc([text("target")], { [`${kind}s`]: {} });
    t.zip.addFile(`word/${kind}s.xml`, Buffer.from(`<w:${kind}s xmlns:w="${W_NS}"/>`, "utf8"));
    await expect(t.merge.appendDocument(await buf([refPara(kind, 1)], { [`${kind}s`]: { 1: noteBody("a") } }))).rejects.toThrow(`cannot take the copied ${kind}s`);
  });
});

describe("merge: a note's own styles and lists travel with it", () => {
  const style = (id: string, extra = "", attrs = 'w:type="paragraph"') => `<w:style ${attrs} w:styleId="${id}"><w:name w:val="${id}"/>${extra}</w:style>`;
  const styleIds = (xml: string) => [...xml.matchAll(/<w:style\b[^>]*\bw:styleId="([^"]+)"/g)].map((m) => m[1]);
  const withStyles = async (e: Mdocxengine, styles: string) => {
    const sx = e.zip.readAsText("word/styles.xml")!;
    e.zip.addFile("word/styles.xml", Buffer.from(sx.replace("</w:styles>", styles + "</w:styles>"), "utf8"));
    return e.zip.toBuffer();
  };
  const noted = (kind: "footnote" | "endnote", inner: string) =>
    doc([refPara(kind, 1)], { [`${kind}s`]: { 1: inner } });

  test.each(["footnote", "endnote"] as const)("a style only a %s uses is imported, with its chain", async (kind) => {
    const src = await withStyles(
      await noted(kind, '<w:p><w:pPr><w:pStyle w:val="MyNoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="MyNoteMark"/></w:rPr><w:t>x</w:t></w:r></w:p>'),
      style("MyNoteBase") + style("MyNoteText", '<w:basedOn w:val="MyNoteBase"/>') + style("MyNoteMark", "", 'w:type="character"'),
    );
    const t = await doc([text("target")]);
    await t.merge.appendDocument(src);
    const got = styleIds(part(t, "word/styles.xml"));
    for (const id of ["MyNoteText", "MyNoteBase", "MyNoteMark"]) expect(got, id).toContain(id);
  });

  test.each(["footnote", "endnote"] as const)("styleMap retargets a %s's paragraph style, which then is not imported", async (kind) => {
    const src = await withStyles(await noted(kind, '<w:p><w:pPr><w:pStyle w:val="Mapped"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'), style("Mapped"));
    const t = await doc([text("target")]);
    await t.merge.appendDocument(src, { styleMap: { Mapped: "Normal" } });
    expect(styleIds(part(t, "word/styles.xml"))).not.toContain("Mapped");
    expect(part(t, `word/${kind}s.xml`)).toContain('<w:pStyle w:val="Normal"/>');
  });

  test.each(["footnote", "endnote"] as const)("a list only a %s uses is copied and its numId remapped", async (kind) => {
    const probe = await Mdocxengine.loadFromFile("samples/example.docx");
    const numId = /<w:num\b[^>]*w:numId="(\d+)"/.exec(probe.zip.readAsText("word/numbering.xml")!)![1];
    const src = (await noted(kind, `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`)).zip.toBuffer();
    const t = await doc([text("target")]);
    const before = (part(t, "word/numbering.xml").match(/<w:num\b/g) ?? []).length;
    await t.merge.appendDocument(src);
    const newId = /<w:numId w:val="(\d+)"/.exec(part(t, `word/${kind}s.xml`))![1];
    expect(newId).not.toBe(numId);
    expect(part(t, "word/numbering.xml")).toContain(`w:numId="${newId}"`);
    expect((part(t, "word/numbering.xml").match(/<w:num\b/g) ?? []).length).toBe(before + 1);
  });
});

describe("merge: comments are stripped from what is copied, and counted", () => {
  const ref = (id: number) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
  const commented = (id: number, word: string) =>
    `<w:p><w:r><w:t>before </w:t></w:r><w:commentRangeStart w:id="${id}"/><w:r><w:t>${word}</w:t></w:r><w:commentRangeEnd w:id="${id}"/>${ref(id)}<w:r><w:t> after</w:t></w:r></w:p>`;
  const clean = (word: string) => `<w:p><w:r><w:t>before </w:t></w:r><w:r><w:t>${word}</w:t></w:r><w:r><w:t> after</w:t></w:r></w:p>`;

  test("the anchors go, every word stays, the count comes back", async () => {
    const t = await doc([text("target")]);
    const r = await t.merge.appendDocument(await buf([commented(0, "alpha"), commented(1, "beta"), text("untouched")]));
    expect(r.commentsStripped).toBe(2);
    const blocks = await t.document.getBlocks();
    expect(blocks.slice(-3).map((b) => b.xml)).toEqual([clean("alpha"), clean("beta"), text("untouched")]);
    expect(await bodyOf(t)).not.toMatch(/<w:comment/);
  });

  test("a source with no comments returns 0 and its blocks arrive byte for byte", async () => {
    const t = await doc([text("target")]);
    const blocks = [text("one"), refPara("footnote", 1), '<w:p w:rsidR="1"><w:r><w:t>two</w:t></w:r></w:p>'];
    const r = await t.merge.appendDocument(await buf(blocks, { footnotes: { 1: noteBody("note") } }));
    expect(r.commentsStripped).toBe(0);
    const got = (await t.document.getBlocks()).slice(1).map((b) => b.xml);
    const newId = refIds(got[1], "footnote")[0];
    expect(got).toEqual([blocks[0], refPara("footnote", newId), blocks[2]]);
  });

  test("a range across paragraphs, and range marks standing between blocks", async () => {
    const src = await doc([]);
    await src.document.saveBlocks([
      { kind: "other", tag: "w:commentRangeStart", xml: '<w:commentRangeStart w:id="4"/>' },
      para(text("alpha")),
      para(`<w:p><w:r><w:t>beta</w:t></w:r><w:commentRangeEnd w:id="4"/>${ref(4)}</w:p>`),
      { kind: "other", tag: "w:commentRangeEnd", xml: '<w:commentRangeEnd w:id="5"/>' },
    ]);
    const t = await doc([text("target")]);
    const r = await t.merge.appendDocument(src.zip.toBuffer());
    expect(r.commentsStripped).toBe(2);
    expect((await t.document.getBlocks()).map((b) => b.xml)).toEqual([text("target"), text("alpha"), "<w:p><w:r><w:t>beta</w:t></w:r></w:p>"]);
  });

  test("a run holding other content keeps it; an emptied run goes, attributes or no rPr", async () => {
    const t = await doc([text("target")]);
    const src = [
      '<w:p><w:r><w:t>kept</w:t><w:commentReference w:id="1"/></w:r></w:p>',
      '<w:p><w:r w:rsidR="00AA"><w:commentReference w:id="2"></w:commentReference></w:r><w:r><w:t>x</w:t></w:r></w:p>',
      '<w:p><w:commentRangeStart w:id="3" w:displacedByCustomXml="next"></w:commentRangeStart><w:r><w:t>y</w:t></w:r></w:p>',
    ];
    const r = await t.merge.appendDocument(await buf(src));
    expect(r.commentsStripped).toBe(3);
    expect((await t.document.getBlocks()).slice(1).map((b) => b.xml)).toEqual([
      "<w:p><w:r><w:t>kept</w:t></w:r></w:p>",
      "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
      "<w:p><w:r><w:t>y</w:t></w:r></w:p>",
    ]);
  });

  test("anchors inside a table and a text box go too", async () => {
    const t = await doc([text("target")]);
    const tbl = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:commentRangeStart w:id="1"/><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r>${ref(1)}</w:p></w:tc></w:tr><w:commentRangeEnd w:id="1"/></w:tbl>`;
    const box = `<w:p><w:r><w:drawing><wps:txbx xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r>${ref(2)}</w:p></w:txbxContent></wps:txbx></w:drawing></w:r></w:p>`;
    const src = await doc([box]);
    await src.document.saveBlocks([{ kind: "table", tag: "w:tbl", xml: tbl }, para(box)]);
    const r = await t.merge.appendDocument(src.zip.toBuffer());
    expect(r.commentsStripped).toBe(2);
    const body = await bodyOf(t);
    expect(body).not.toMatch(/<w:comment/);
    expect(body).toContain("cell");
    expect(body).toContain("boxed");
  });

  test("anchors inside a copied note go too, and count with the body's", async () => {
    const t = await doc([text("target")]);
    const r = await t.merge.appendDocument(
      await buf([commented(0, "alpha"), refPara("endnote", 1), refPara("footnote", 1)], {
        endnotes: { 1: `<w:p><w:r><w:t>en</w:t></w:r>${ref(0)}${ref(7)}</w:p>` },
        footnotes: { 1: `<w:p><w:commentRangeStart w:id="8"/><w:r><w:t>fn</w:t></w:r><w:commentRangeEnd w:id="8"/></w:p>` },
      }),
    );
    expect(r.commentsStripped).toBe(3); // ids 0 (twice), 7 and 8
    expect(part(t, "word/endnotes.xml")).not.toMatch(/<w:comment/);
    expect(part(t, "word/footnotes.xml")).not.toMatch(/<w:comment/);
    expect(part(t, "word/endnotes.xml")).toContain(">en<");
  });

  test("a comment the source defines but never anchors still counts: its part stays behind", async () => {
    const src = await doc([commented(0, "alpha")]);
    src.zip.addFile("word/comments.xml", Buffer.from(`<?xml version="1.0"?><w:comments xmlns:w="${W_NS}"><w:comment w:id="0"><w:p/></w:comment><w:comment w:id="1"><w:p/></w:comment></w:comments>`));
    const t = await doc([text("target")]);
    const r = await t.merge.appendDocument(src.zip.toBuffer());
    expect(r.commentsStripped).toBe(2);
    expect(t.zip.getEntry("word/comments.xml")).toBeNull();
  });

  test("the comment-only character style is not imported for a reference that is gone", async () => {
    const e = await doc([commented(0, "alpha")]);
    const sx = e.zip.readAsText("word/styles.xml")!;
    e.zip.addFile("word/styles.xml", Buffer.from(sx.replace("</w:styles>", '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/></w:style></w:styles>'), "utf8"));
    const t = await doc([text("target")]);
    await t.merge.appendDocument(e.zip.toBuffer());
    expect(part(t, "word/styles.xml")).not.toContain('w:styleId="CommentReference"');
  });

  test("the target's own comments are not touched by an append", async () => {
    const t = await doc([commented(0, "theirs")]);
    t.zip.addFile("word/comments.xml", Buffer.from(`<w:comments xmlns:w="${W_NS}"><w:comment w:id="0"><w:p/></w:comment></w:comments>`));
    const r = await t.merge.appendDocument(await buf([commented(0, "alpha")]));
    expect(r.commentsStripped).toBe(1);
    expect((await t.document.getBlocks()).map((b) => b.xml)).toEqual([commented(0, "theirs"), clean("alpha")]);
    expect(t.zip.getEntry("word/comments.xml")).not.toBeNull();
  });
});

describe("merge.stripComments: the target's own comments", () => {
  const ref = (id: number) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
  const commented = (id: number) => `<w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>word</w:t></w:r><w:commentRangeEnd w:id="${id}"/>${ref(id)}</w:p>`;

  /** A document carrying Word's full comment family, one comment with its own picture and link. */
  async function withCommentParts(blocks: string[]): Promise<Mdocxengine> {
    const e = await doc(blocks);
    const add = async (name: string, relType: string, ct: string, xml: string) => {
      e.zip.addFile(`word/${name}`, Buffer.from(xml, "utf8"));
      await e.rels.addRelationship(await e.rels.genId(), relType, name);
      await e.contentTypes.addOverride(`/word/${name}`, ct);
    };
    const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    await add("comments.xml", REL("comments"), `${CT}.comments+xml`, `<?xml version="1.0"?><w:comments xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>a remark</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>another</w:t></w:r></w:p></w:comment></w:comments>`);
    await add("commentsExtended.xml", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended", `${CT}.commentsExtended+xml`, "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"/>");
    await add("commentsIds.xml", "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds", `${CT}.commentsIds+xml`, "<w16cid:commentsIds xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\"/>");
    await add("commentsExtensible.xml", "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible", `${CT}.commentsExtensible+xml`, "<w16cex:commentsExtensible xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\"/>");
    e.zip.addFile("word/media/comment-only.png", Buffer.from([0x89, 0x50, 1]));
    e.zip.addFile("word/media/shared.png", Buffer.from([0x89, 0x50, 2]));
    await e.rels.addRelationship("rIdShared", REL("image"), "media/shared.png");
    e.zip.addFile(
      "word/_rels/comments.xml.rels",
      Buffer.from(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="${REL("image")}" Target="media/comment-only.png"/>` +
          `<Relationship Id="rId2" Type="${REL("image")}" Target="media/shared.png"/>` +
          `<Relationship Id="rId3" Type="${REL("hyperlink")}" Target="https://a.example/" TargetMode="External"/></Relationships>`,
      ),
    );
    return e;
  }

  test("anchors, the comment parts, their relationships, content types and own relationships all go", async () => {
    const e = await withCommentParts([commented(0), commented(1), text("clean")]);
    const n = await e.merge.stripComments();
    expect(n).toBe(2);
    expect(await bodyOf(e)).toBe('<w:p><w:r><w:t>word</w:t></w:r></w:p><w:p><w:r><w:t>word</w:t></w:r></w:p>' + text("clean"));
    for (const name of ["comments", "commentsExtended", "commentsIds", "commentsExtensible"]) {
      expect(e.zip.getEntry(`word/${name}.xml`), name).toBeNull();
      expect(part(e, "[Content_Types].xml"), name).not.toContain(`/word/${name}.xml`);
    }
    expect(part(e, "word/_rels/document.xml.rels")).not.toMatch(/comments/i);
    expect(e.zip.getEntry("word/_rels/comments.xml.rels")).toBeNull();
    expect(fatals(e)).toEqual([]);
  });

  test("a picture only a comment used goes with it, one the body shares stays", async () => {
    const e = await withCommentParts([commented(0)]);
    await e.merge.stripComments();
    expect(e.zip.getEntry("word/media/comment-only.png")).toBeNull();
    expect(e.zip.getEntry("word/media/shared.png")).not.toBeNull();
  });

  test("a shared picture stays when the two rels spell its name differently", async () => {
    const e = await withCommentParts([commented(0)]);
    e.zip.addFile("word/media/my pic.png", Buffer.from([0x89, 0x50, 3]));
    await e.rels.addRelationship("rIdEnc", REL("image"), "media/my%20pic.png");
    const rels = part(e, "word/_rels/comments.xml.rels");
    e.zip.addFile("word/_rels/comments.xml.rels", Buffer.from(rels.replace("</Relationships>", `<Relationship Id="rId9" Type="${REL("image")}" Target="media/my pic.png"/></Relationships>`), "utf8"));
    await e.merge.stripComments();
    expect(e.zip.getEntry("word/media/my pic.png")).not.toBeNull();
  });

  test("a comment part no relationship names goes too", async () => {
    const e = await doc([commented(0)]);
    e.zip.addFile("word/comments.xml", Buffer.from(`<w:comments xmlns:w="${W_NS}"><w:comment w:id="0"><w:p/></w:comment></w:comments>`));
    expect(await e.merge.stripComments()).toBe(1);
    expect(e.zip.getEntry("word/comments.xml")).toBeNull();
  });

  test("an Override naming a picture that goes, goes with it", async () => {
    const e = await withCommentParts([commented(0)]);
    await e.contentTypes.addOverride("/word/media/comment-only.png", "image/png");
    await e.merge.stripComments();
    expect(part(e, "[Content_Types].xml")).not.toContain("comment-only.png");
    expect(part(e, "[Content_Types].xml")).toContain('PartName="/word/document.xml"');
  });

  test("an anchor in a header is stripped with the rest", async () => {
    const e = await doc([text("x")]);
    e.zip.addFile("word/header7.xml", Buffer.from(`<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>h</w:t></w:r>${ref(6)}</w:p></w:hdr>`, "utf8"));
    expect(await e.merge.stripComments()).toBe(1);
    expect(part(e, "word/header7.xml")).toBe(`<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>h</w:t></w:r></w:p></w:hdr>`);
  });

  test("it counts comments the document defines but no longer anchors", async () => {
    const e = await withCommentParts([commented(0)]);
    expect(await e.merge.stripComments()).toBe(2);
  });

  test("anchors in the document's own notes go too", async () => {
    const e = await doc([refPara("endnote", 1)], { endnotes: { 1: `<w:p><w:r><w:t>en</w:t></w:r>${ref(3)}</w:p>` }, footnotes: { 1: `<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>fn</w:t></w:r></w:p>` } });
    expect(await e.merge.stripComments()).toBe(2);
    expect(part(e, "word/endnotes.xml")).toContain("<w:p><w:r><w:t>en</w:t></w:r></w:p>");
    expect(part(e, "word/footnotes.xml")).toContain("<w:p><w:r><w:t>fn</w:t></w:r></w:p>");
  });

  test("it is idempotent, and a document with no comments is not rewritten at all", async () => {
    const e = await withCommentParts([commented(0)]);
    expect(await e.merge.stripComments()).toBe(2);
    expect(await e.merge.stripComments()).toBe(0);

    const plain = await Mdocxengine.loadFromFile("samples/example.docx");
    const names = new AdmZip(plain.zip.toBuffer()).getEntries().map((x) => x.entryName);
    const before = names.map((n) => part(plain, n));
    expect(await plain.merge.stripComments()).toBe(0);
    expect(names.map((n) => part(plain, n))).toEqual(before);
  });

  test("the words of a thesis with comments survive a strip byte for byte, outside the anchors", async () => {
    const e = await withCommentParts([commented(0), '<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>']);
    await e.merge.stripComments();
    const xml = await bodyOf(e);
    expect(xml).toContain("<w:r><w:t>word</w:t></w:r>");
    expect(xml).toContain('<w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
  });
});
