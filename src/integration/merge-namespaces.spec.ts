import { describe, test, expect } from "vitest";
import { Mdocxengine, makeParagraphNode, type BodyBlock } from "../index";
import { carryRootNamespaces } from "../utils/rootNamespaces";

/**
 * Root namespace declarations across a merge.
 *
 * Blocks are copied as exact XML substrings, so a prefix the SOURCE root
 * declared (`w14`, `mc`, `wps`, `r`…) is unbound in a target whose root did not.
 * Word refuses such a file ("-10000"); real split theses hit it.
 */

const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  w10: "urn:schemas-microsoft-com:office:word",
  wne: "http://schemas.microsoft.com/office/word/2006/wordml",
};
const decl = (ps: string[]) => ps.map((p) => `xmlns:${p}="${(NS as any)[p]}"`).join(" ");

const LEGACY_ROOT = `<w:document xmlns:ve="${NS.mc}" ${decl(["o", "r", "m", "v", "wp", "w10", "w", "wne"])}>`;
const MODERN_ROOT = `<w:document ${decl(["w", "r", "mc", "w14", "wps", "wp14", "wp"])} mc:Ignorable="w14 wp14">`;

/** Scope-aware check: every element/attribute prefix is bound where it is used. */
function unboundPrefixes(xml: string): string[] {
  const bad = new Set<string>();
  const stack: Set<string>[] = [new Set(["xml", "xmlns"])];
  const isBound = (p: string) => stack.some((s) => s.has(p));
  for (const m of xml.matchAll(/<(\/?)([A-Za-z_][\w.-]*(?::[\w.-]+)?)((?:\s+[^<>]*?)?)(\/?)>/g)) {
    const [, closing, name, attrs, selfClose] = m;
    if (closing) {
      stack.pop();
      continue;
    }
    const scope = new Set<string>();
    for (const a of attrs.matchAll(/\sxmlns:([\w.-]+)=/g)) scope.add(a[1]);
    stack.push(scope);
    const pre = name.includes(":") ? name.split(":")[0] : null;
    if (pre && !isBound(pre)) bad.add(pre);
    for (const a of attrs.matchAll(/\s([\w.-]+):[\w.-]+="/g)) if (!isBound(a[1])) bad.add(a[1]);
    if (selfClose) stack.pop();
  }
  return [...bad].sort();
}

async function withRoot(root: string, blocks: BodyBlock[]): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  await e.document.saveBlocks(blocks);
  const xml = e.zip.readAsText("word/document.xml")!;
  e.zip.addFile("word/document.xml", Buffer.from(xml.replace(/<w:document\b[^>]*>/, root), "utf8"));
  return e;
}

const modernBlocks = (): BodyBlock[] => [
  makeParagraphNode("modern"),
  {
    kind: "paragraph",
    tag: "w:p",
    xml:
      '<w:p w14:paraId="1A2B3C4D"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
      '<wp:inline wp14:anchorId="0A0B0C0D"><wps:wsp><wps:bodyPr/></wps:wsp></wp:inline></w:drawing></mc:Choice>' +
      "<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>",
  },
];

describe("root namespaces survive a merge", () => {
  test("legacy 9-prefix target + modern source: no unbound prefix, Ignorable carried", async () => {
    const target = await withRoot(LEGACY_ROOT, [makeParagraphNode("T")]);
    expect(unboundPrefixes(target.zip.readAsText("word/document.xml")!)).toEqual([]);

    const source = await withRoot(MODERN_ROOT, modernBlocks());
    await target.merge.appendDocument(source.zip.toBuffer(), {});

    const out = target.zip.readAsText("word/document.xml")!;
    expect(unboundPrefixes(out)).toEqual([]);
    // Ignorable rides on the target's own MC prefix (`ve`); `mc` is bound too
    // because the copied blocks use `mc:AlternateContent`.
    expect(out).toMatch(/ve:Ignorable="[^"]*w14[^"]*wp14[^"]*"/);
    expect(out).toContain(`xmlns:mc="${NS.mc}"`);
    const re = await Mdocxengine.loadFromBuffer(target.zip.toBuffer());
    expect((await re.document.getBlocks()).length).toBeGreaterThan(2);
  });

  test("the checker really detects an unbound prefix", () => {
    expect(unboundPrefixes('<w:p xmlns:w="x" w14:paraId="1"><mc:Choice/></w:p>')).toEqual(["mc", "w14"]);
  });

  test("a target without footnotes gets a footnotes part that binds r", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    t.zip.deleteFile("word/footnotes.xml");
    const s = await Mdocxengine.loadFromFile("samples/example.docx");
    const id = (await s.footnotes.addFootnote("n")).id;
    const fx = s.zip.readAsText("word/footnotes.xml")!;
    s.zip.addFile(
      "word/footnotes.xml",
      Buffer.from(
        fx.replace(
          new RegExp(`(<w:footnote\\b[^>]*w:id="${id}"[^>]*>[\\s\\S]*?)(</w:p>\\s*</w:footnote>)`),
          '$1<w:hyperlink r:id="rIdH"><w:r><w:t>x</w:t></w:r></w:hyperlink>$2',
        ),
        "utf8",
      ),
    );
    await s.document.saveBlocks([
      makeParagraphNode("a"),
      { kind: "paragraph", tag: "w:p", xml: `<w:p><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>` },
    ]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const out = t.zip.readAsText("word/footnotes.xml")!;
    expect(out).toContain("r:id=");
    expect(unboundPrefixes(out)).toEqual([]);
    expect(out).toContain(`xmlns:r="${NS.r}"`);
  });

  test("a target without numbering gets a numbering root with the standard prefixes", async () => {
    const t = await Mdocxengine.loadFromFile("samples/example.docx");
    t.zip.deleteFile("word/numbering.xml");
    const s = await Mdocxengine.loadFromFile("samples/example.docx");
    const numId = /<w:num\b[^>]*w:numId="(\d+)"/.exec(s.zip.readAsText("word/numbering.xml")!)![1];
    await s.document.saveBlocks([
      {
        kind: "paragraph",
        tag: "w:p",
        xml: `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`,
      },
    ]);
    await t.merge.appendDocument(s.zip.toBuffer(), {});
    const out = t.zip.readAsText("word/numbering.xml")!;
    expect(out).toContain("<w:abstractNum");
    expect(out).toContain(`xmlns:r="${NS.r}"`);
    expect(out).toContain(`xmlns:w16cid=`);
    expect(unboundPrefixes(out)).toEqual([]);
  });

  test("a prefix bound to a different namespace throws", async () => {
    const target = await withRoot(LEGACY_ROOT, [makeParagraphNode("T")]);
    const bad = MODERN_ROOT.replace(NS.w14, "urn:something:else");
    const source = await withRoot(bad.replace(/xmlns:wp="[^"]*"/, `xmlns:wp="urn:other"`), modernBlocks());
    await expect(target.merge.appendDocument(source.zip.toBuffer(), {})).rejects.toThrow(/wp/);
  });
});

describe("carryRootNamespaces (unit)", () => {
  test("no-ops without a source or root, adds mc when the target has none", () => {
    expect(carryRootNamespaces("<a/>", null).xml).toBe("<a/>");
    expect(carryRootNamespaces("not xml", "<a/>").xml).toBe("not xml");
    const r = carryRootNamespaces(
      `<w:document xmlns:w="${NS.w}"><w:body/></w:document>`,
      `<w:document xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14 gone">`,
    );
    expect(r.added).toEqual(["mc", "w14"]);
    expect(r.xml).toContain('mc:Ignorable="w14"');
    expect(r.xml).toContain("<w:body/>");
  });

  test("unions into an existing Ignorable and handles a self-closing root", () => {
    const r = carryRootNamespaces(
      `<x xmlns:mc="${NS.mc}" xmlns:w14="${NS.w14}" mc:Ignorable="w14"/>`,
      `<x xmlns:mc="${NS.mc}" xmlns:wp14="${NS.wp14}" mc:Ignorable="wp14"/>`,
    );
    expect(r.xml).toContain('mc:Ignorable="w14 wp14"');
    expect(r.xml.endsWith("/>")).toBe(true);
  });
});
