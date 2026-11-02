import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine } from "../index";
import { fixture, notesPart, P, texts, XML_HEAD, W_NS } from "./docx-fixture";

/**
 * Removing a comment anchor must remove nothing else. `<w:rPr />` is the form .NET's
 * XmlWriter gives an empty element, and read as an opening tag it ran the removal on to
 * the next `</w:rPr>` standing before a comment reference: the student's words between
 * them were gone, and nothing downstream looks at a single part's words.
 */

const COMMENTS = `${XML_HEAD}<w:comments xmlns:w="${W_NS}"><w:comment w:id="0" w:author="A" w:initials="A"><w:p><w:r><w:t>remark</w:t></w:r></w:p></w:comment></w:comments>`;
const REF = `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>`;
const SPACED_REF = `<w:r><w:rPr><w:rStyle w:val="CommentReference" /></w:rPr><w:commentReference w:id="0" /></w:r>`;

const stripped = async (doc: Buffer) => {
  const e = await Mdocxengine.loadFromBuffer(doc);
  const count = await e.merge.stripComments();
  return { count, body: e.zip.readAsText("word/document.xml")!, footnotes: e.zip.readAsText("word/footnotes.xml") ?? "" };
};

describe("a comment anchor leaves nothing else behind it", () => {
  const BODIES: Record<string, string> = {
    "<w:rPr /> runs before the reference":
      `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:rPr /><w:t xml:space="preserve">ONE </w:t></w:r><w:r><w:rPr /><w:t xml:space="preserve">TWO </w:t></w:r><w:commentRangeEnd w:id="0"/>${SPACED_REF}<w:r><w:t>after</w:t></w:r></w:p>`,
    "an rPr with an attribute":
      `<w:p><w:r><w:rPr w:rsidRPr="00A1"/><w:t xml:space="preserve">ONE </w:t></w:r>${REF}<w:r><w:t>after</w:t></w:r></w:p>`,
    "across table cells":
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:rPr /><w:t>ONE</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>TWO</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>THREE</w:t></w:r>${REF}</w:p></w:tc><w:tc><w:p><w:r><w:t>FOUR</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`,
    "starting inside a hyperlink":
      `<w:p><w:hyperlink w:anchor="x"><w:r><w:rPr /><w:t>LINK</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> MORE </w:t></w:r>${REF}</w:p>`,
    "a self-closed run":
      `<w:p><w:r /><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">ONE </w:t></w:r>${REF}</w:p>`,
  };

  test.each(Object.keys(BODIES))("stripComments: %s", async (name) => {
    const doc = fixture({ body: BODIES[name], comments: COMMENTS });
    const before = texts(new AdmZip(doc).readAsText("word/document.xml"));
    const { count, body } = await stripped(doc);
    expect(count).toBe(1);
    expect(texts(body)).toEqual(before);
    expect(body).not.toMatch(/<w:comment/);
    // Only the anchors went: every element that opened still closes.
    for (const tag of ["w:hyperlink", "w:tc", "w:tr", "w:p"]) {
      expect(body.match(new RegExp(`<${tag}[\\s>]`, "g"))?.length ?? 0, tag).toBe(body.match(new RegExp(`</${tag}>`, "g"))?.length ?? 0);
    }
  });

  test.each(Object.keys(BODIES))("appendDocument: %s", async (name) => {
    const src = fixture({ body: BODIES[name], comments: COMMENTS });
    const t = await Mdocxengine.loadFromBuffer(fixture({ body: P("target") }));
    const r = await t.merge.appendDocument(src);
    expect(r.commentsStripped).toBe(1);
    expect(texts(t.zip.readAsText("word/document.xml")!)).toEqual(["target", ...texts(new AdmZip(src).readAsText("word/document.xml"))]);
  });

  test("a notes part is one string: the removal cannot cross from note to note", async () => {
    const notes = notesPart(
      "footnote",
      `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:rPr /><w:t>NOTE ONE</w:t></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>NOTE TWO</w:t></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>NOTE THREE</w:t></w:r>${SPACED_REF}</w:p></w:footnote>`,
    );
    const body = `<w:p><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>`;
    const { footnotes } = await stripped(fixture({ body, footnotes: notes, comments: COMMENTS }));
    expect(texts(footnotes)).toEqual(["NOTE ONE", "NOTE TWO", "NOTE THREE"]);
    expect([...footnotes.matchAll(/<w:footnote\b[^>]*\bw:id="(-?\d+)"/g)].map((m) => m[1])).toEqual(["-1", "0", "1", "2", "3"]);
    expect(footnotes).not.toMatch(/<w:comment/);
  });

  test("a reference run with an empty rPr still goes whole", async () => {
    const doc = fixture({ body: `<w:p><w:r><w:t>kept</w:t></w:r><w:r w:rsidR="00B"><w:rPr /><w:commentReference w:id="0" /></w:r></w:p>`, comments: COMMENTS });
    const { body } = await stripped(doc);
    expect(body).toContain("<w:p><w:r><w:t>kept</w:t></w:r></w:p>");
  });

  test("a run holding other content keeps it, and loses only the anchor", async () => {
    const doc = fixture({ body: `<w:p><w:r><w:rPr /><w:commentReference w:id="0"/><w:t>kept</w:t></w:r></w:p>`, comments: COMMENTS });
    const { body } = await stripped(doc);
    expect(body).toContain("<w:p><w:r><w:rPr /><w:t>kept</w:t></w:r></w:p>");
  });
});
