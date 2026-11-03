import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { Mdocxengine, canonicalMainPart, checkDocxBuffer } from "../index";
import { fixture, P, SECT, noteText, note, notesPart, refP, png, relsPart, stylesDoc, NORMAL, style, texts, XML_HEAD, W_NS } from "./docx-fixture";

/**
 * The main document is whatever `_rels/.rels` calls the officeDocument. Word prints a package whose
 * main part is word/document2.xml; a reader of the fixed path sees an empty file and says nothing.
 */

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Move parts of a package the way another writer would have named them, rels and content types too. */
function renamed(docx: Buffer, moves: Record<string, string>): Buffer {
  const z = new AdmZip(docx);
  const out = new AdmZip();
  const relsOf = (p: string) => `${p.slice(0, p.lastIndexOf("/") + 1)}_rels/${p.slice(p.lastIndexOf("/") + 1)}.rels`;
  const allMoves = { ...moves };
  for (const [from, to] of Object.entries(moves)) allMoves[relsOf(from)] = relsOf(to);
  for (const e of z.getEntries()) {
    let bytes = e.getData();
    if (e.entryName === "[Content_Types].xml") {
      let t = bytes.toString("utf8");
      for (const [from, to] of Object.entries(moves)) t = t.replace(`PartName="/${from}"`, `PartName="/${to}"`);
      bytes = Buffer.from(t, "utf8");
    }
    if (e.entryName === "_rels/.rels") {
      let t = bytes.toString("utf8");
      for (const [from, to] of Object.entries(moves)) t = t.replace(`Target="${from}"`, `Target="${to}"`);
      bytes = Buffer.from(t, "utf8");
    }
    if (e.entryName === "word/_rels/document.xml.rels") {
      let t = bytes.toString("utf8");
      for (const [from, to] of Object.entries(moves)) if (from.startsWith("word/") && to.startsWith("word/")) t = t.replace(`Target="${from.slice(5)}"`, `Target="${to.slice(5)}"`);
      bytes = Buffer.from(t, "utf8");
    }
    out.addFile(allMoves[e.entryName] ?? e.entryName, bytes);
  }
  return out.toBuffer();
}

const withRootRels = (docx: Buffer, rels: string | null): Buffer => {
  const z = new AdmZip(docx);
  if (rels === null) z.deleteFile("_rels/.rels");
  else z.addFile("_rels/.rels", Buffer.from(relsPart(rels), "utf8"));
  return z.toBuffer();
};

const words = async (docx: Buffer) => texts(new AdmZip(docx).readAsText("word/document.xml")).join(" ");
const merge = async (a: Buffer, b: Buffer) => {
  const e = await Mdocxengine.loadFromBuffer(a);
  await e.merge.appendDocument(b, {});
  return e.zip.toBuffer();
};
const first = () => fixture({ body: P("Un.") });
const OFFICE = (target: string) => `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="${target}"/>`;

describe("a part's main document is the one its package names", () => {
  test("a later part at word/document2.xml brings every word", async () => {
    const later = renamed(fixture({ body: P("Mots du second fichier.") }), { "word/document.xml": "word/document2.xml" });
    expect(await words(await merge(first(), later))).toBe("Un. Mots du second fichier.");
  });

  test("its picture, its notes and its styles come through its own relationships", async () => {
    const later = renamed(
      fixture({
        body: `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/><a:graphic><a:graphicData><a:blip r:embed="rIdPic"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` + refP("footnote", 1, "Corps") + P("Style propre", `<w:pStyle w:val="Propre"/>`),
        footnotes: notesPart("footnote", note("footnote", 1, noteText("footnote", "Note du second fichier"))),
        styles: stylesDoc(NORMAL + style("paragraph", "Propre", "Propre Style", `<w:rPr><w:sz w:val="40"/></w:rPr>`)),
        parts: { "word/media/image1.png": png(3) },
        docRels: `<Relationship Id="rIdPic" Type="${R_NS}/image" Target="media/image1.png"/>`,
      }),
      { "word/document.xml": "word/document2.xml", "word/footnotes.xml": "word/footnotes2.xml", "word/styles.xml": "word/styles2.xml" },
    );
    const merged = new AdmZip(await merge(first(), later));
    expect(merged.readAsText("word/footnotes.xml")).toContain("Note du second fichier");
    expect(merged.readAsText("word/styles.xml")).toContain("Propre Style");
    const blip = /r:embed="([^"]+)"/.exec(merged.readAsText("word/document.xml"))![1];
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${blip}"[^>]*>`).exec(merged.readAsText("word/_rels/document.xml.rels"))![0];
    expect(merged.getEntry(`word/${/Target="([^"]+)"/.exec(rel)![1]}`)?.getData().equals(png(3))).toBe(true);
  });

  test("a main part in another folder keeps what its relationships reach", async () => {
    const z = new AdmZip(fixture({ body: `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/><a:graphic><a:graphicData><a:blip r:embed="rIdPic"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` + P("Ailleurs."), styles: null, docRels: `<Relationship Id="rIdPic" Type="${R_NS}/image" Target="media/image1.png"/>` }));
    const doc = z.readAsText("word/document.xml");
    const rels = z.readAsText("word/_rels/document.xml.rels");
    z.deleteFile("word/document.xml");
    z.deleteFile("word/_rels/document.xml.rels");
    z.addFile("contenu/principal.xml", Buffer.from(doc, "utf8"));
    z.addFile("contenu/_rels/principal.xml.rels", Buffer.from(rels, "utf8"));
    z.addFile("contenu/media/image1.png", png(5));
    z.addFile("_rels/.rels", Buffer.from(relsPart(OFFICE("contenu/principal.xml")), "utf8"));
    z.addFile("[Content_Types].xml", Buffer.from(z.readAsText("[Content_Types].xml").replace('PartName="/word/document.xml"', 'PartName="/contenu/principal.xml"'), "utf8"));

    const canon = new AdmZip(canonicalMainPart(z.toBuffer()));
    expect(canon.getEntry("contenu/principal.xml")).toBeNull();
    expect(canon.readAsText("word/document.xml")).toContain("Ailleurs.");
    const target = /Target="([^"]+)"/.exec(canon.readAsText("word/_rels/document.xml.rels"))![1];
    expect(target).toBe("../contenu/media/image1.png");

    const merged = new AdmZip(await merge(first(), z.toBuffer()));
    expect(merged.readAsText("word/document.xml")).toContain("Ailleurs.");
    const blip = /r:embed="([^"]+)"/.exec(merged.readAsText("word/document.xml"))![1];
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${blip}"[^>]*>`).exec(merged.readAsText("word/_rels/document.xml.rels"))![0];
    expect(merged.getEntry(`word/${/Target="([^"]+)"/.exec(rel)![1]}`)?.getData().equals(png(5))).toBe(true);
  });

  test("a stale word/document.xml nothing points at is not the document", async () => {
    const later = new AdmZip(renamed(fixture({ body: P("Le vrai texte.") }), { "word/document.xml": "word/document2.xml" }));
    later.addFile("word/document.xml", Buffer.from(`${XML_HEAD}<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Vieux texte.</w:t></w:r></w:p>${SECT}</w:body></w:document>`, "utf8"));
    expect(await words(await merge(first(), later.toBuffer()))).toBe("Un. Le vrai texte.");
  });

  test("a word/document.xml something else reaches cannot be replaced, and the merge says so", async () => {
    const later = new AdmZip(renamed(fixture({ body: P("Le vrai texte.") }), { "word/document.xml": "word/document2.xml" }));
    const rels = "word/_rels/document2.xml.rels";
    later.addFile(rels, Buffer.from(later.readAsText(rels).replace("</Relationships>", `<Relationship Id="rIdOld" Type="${R_NS}/aFChunk" Target="document.xml"/></Relationships>`), "utf8"));
    later.addFile("word/document.xml", Buffer.from(`${XML_HEAD}<w:document xmlns:w="${W_NS}"><w:body>${SECT}</w:body></w:document>`, "utf8"));
    await expect(merge(first(), later.toBuffer())).rejects.toThrow(/main document/);
  });

  test("the first part may be the renamed one too", async () => {
    const a = renamed(fixture({ body: P("Mots du premier.") }), { "word/document.xml": "word/document2.xml" });
    const out = await merge(a, fixture({ body: P("Deux.") }));
    expect(await words(out)).toBe("Mots du premier. Deux.");
    const z = new AdmZip(out);
    expect(z.readAsText("_rels/.rels")).toContain('Target="word/document.xml"');
    expect(z.readAsText("[Content_Types].xml")).toContain('PartName="/word/document.xml"');
    expect(z.getEntry("word/document2.xml")).toBeNull();
    expect((await checkDocxBuffer(out)).findings.filter((f) => f.severity === "fatal")).toEqual([]);
  });

  test("names are matched without regard to case, and the percent signs are read", async () => {
    const z = new AdmZip(fixture({ body: P("Casse du nom.") }));
    const doc = z.readAsText("word/document.xml");
    const rels = z.readAsText("word/_rels/document.xml.rels");
    z.deleteFile("word/document.xml");
    z.deleteFile("word/_rels/document.xml.rels");
    z.addFile("Word/Document 2.xml", Buffer.from(doc, "utf8"));
    z.addFile("Word/_rels/Document 2.xml.rels", Buffer.from(rels, "utf8"));
    z.addFile("_rels/.rels", Buffer.from(relsPart(OFFICE("/word/document%202.xml")), "utf8"));
    z.addFile("[Content_Types].xml", Buffer.from(z.readAsText("[Content_Types].xml").replace('PartName="/word/document.xml"', 'PartName="/word/document%202.xml"'), "utf8"));
    expect(await words(await merge(first(), z.toBuffer()))).toBe("Un. Casse du nom.");
  });

  test("a package in the usual shape comes back as the same bytes", () => {
    const a = fixture({ body: P("Rien a faire.") });
    expect(canonicalMainPart(a).equals(a)).toBe(true);
  });

  test("canonical twice is canonical once", () => {
    const later = renamed(fixture({ body: P("Mots.") }), { "word/document.xml": "word/document2.xml" });
    const once = canonicalMainPart(later);
    expect(canonicalMainPart(once).equals(once)).toBe(true);
  });
});

describe("a part with no main document stops the merge", () => {
  test("no _rels/.rels", async () => {
    await expect(merge(first(), withRootRels(fixture({ body: P("x") }), null))).rejects.toThrow(/main document/);
  });
  test("no officeDocument relationship", async () => {
    await expect(merge(first(), withRootRels(fixture({ body: P("x") }), ""))).rejects.toThrow(/main document/);
  });
  test("a relationship to a part the package does not hold", async () => {
    await expect(merge(first(), withRootRels(fixture({ body: P("x") }), OFFICE("word/absent.xml")))).rejects.toThrow(/main document/);
  });
  test("an external target", async () => {
    await expect(merge(first(), withRootRels(fixture({ body: P("x") }), `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="https://exemple.invalid/a.xml" TargetMode="External"/>`))).rejects.toThrow(/main document/);
  });
  test("two of them, naming different parts", async () => {
    await expect(merge(first(), withRootRels(fixture({ body: P("x") }), OFFICE("word/document.xml") + OFFICE("word/styles.xml").replace("rId1", "rId2")))).rejects.toThrow(/main document/);
  });
  test("this document's own, before anything is written", async () => {
    const e = await Mdocxengine.loadFromBuffer(withRootRels(fixture({ body: P("x") }), null));
    await expect(e.merge.appendDocument(first(), {})).rejects.toThrow(/main document/);
    expect(e.zip.readAsText("word/document.xml")).toContain(">x<");
  });
});
