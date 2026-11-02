import { describe, test, expect } from "vitest";
import { Mdocxengine, makeParagraphNode, type BodyBlock } from "../index";
import { ContentTypesManager } from "../core/PartsManagers/ContentTypesManager";

/**
 * VML pictures and OLE objects across a merge.
 *
 * `<v:imagedata r:id>` (a VML picture, or the preview of an OLE object) and
 * `<o:OLEObject r:id>` (the embedded object itself) address parts by relationship
 * id. Before they were remapped, every one came out of a merge with a dangling
 * relationship: a split thesis with 26 equation-editor objects lost all of them.
 */

const IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const OLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const NS_VO =
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const png = (tag: number) =>
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, tag, tag, tag, 1, 2, 3]);
const bin = (tag: number) => Buffer.from([0xd0, 0xcf, 0x11, 0xe0, tag, tag, 9, 9, 9]);

const vmlPara = (rid: string) =>
  `<w:p><w:r><w:pict ${NS_VO}><v:shape id="s1" style="width:20pt;height:10pt"><v:imagedata r:id="${rid}" o:title=""/></v:shape></w:pict></w:r></w:p>`;
const oleParaX = (preview: string, obj: string) =>
  `<w:p><w:r><w:object ${NS_VO} w:dxaOrig="400" w:dyaOrig="300"><v:shape id="s2" style="width:20pt;height:15pt"><v:imagedata r:id="${preview}" o:title=""/></v:shape>` +
  `<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="s2" DrawAspect="Content" ObjectID="_1234" r:id="${obj}"/></w:object></w:r></w:p>`;

interface SrcOpts {
  tag: number;
  missingOle?: boolean;
  footnoteImage?: boolean;
}

/** A source .docx carrying one VML picture and one OLE object (+ optional footnote image). */
async function source(o: SrcOpts): Promise<Buffer> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  const rels = e.merge["rels"];
  await rels.addRelationship("rIdV", IMG, "media/vml1.png");
  await rels.addRelationship("rIdP", IMG, "media/ole1.png");
  await rels.addRelationship("rIdO", OLE, "embeddings/oleObject1.bin");
  e.zip.addFile("word/media/vml1.png", png(o.tag));
  e.zip.addFile("word/media/ole1.png", png(o.tag + 1));
  if (!o.missingOle) e.zip.addFile("word/embeddings/oleObject1.bin", bin(o.tag));
  const ct = new ContentTypesManager(e.zip);
  await ct.addDefault("png", "image/png");
  await ct.addDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");

  const blocks: BodyBlock[] = [
    makeParagraphNode(`src${o.tag} start`),
    { kind: "paragraph", tag: "w:p", xml: vmlPara("rIdV") },
    { kind: "paragraph", tag: "w:p", xml: oleParaX("rIdP", "rIdO") },
  ];

  if (o.footnoteImage) {
    const id = (await e.footnotes.addFootnote("note")).id;
    const fx = e.zip.readAsText("word/footnotes.xml")!;
    const drawing =
      '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="100" cy="100"/><wp:docPr id="9" name="fn"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdFn"/>' +
      "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    const patched = fx.replace(
      new RegExp(`(<w:footnote\\b[^>]*w:id="${id}"[^>]*>[\\s\\S]*?)(</w:p>\\s*</w:footnote>)`),
      `$1${drawing}$2`,
    );
    expect(patched).not.toBe(fx);
    e.zip.addFile("word/footnotes.xml", Buffer.from(patched, "utf8"));
    e.zip.addFile(
      "word/_rels/footnotes.xml.rels",
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rIdFn" Type="${IMG}" Target="media/fn1.png"/></Relationships>`,
      ),
    );
    e.zip.addFile("word/media/fn1.png", png(o.tag + 7));
    blocks.push({
      kind: "paragraph",
      tag: "w:p",
      xml: `<w:p><w:r><w:t>with note</w:t></w:r><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`,
    });
  }
  await e.document.saveBlocks(blocks);
  return e.zip.toBuffer();
}

async function target(): Promise<Mdocxengine> {
  const e = await Mdocxengine.loadFromFile("samples/example.docx");
  await e.document.saveBlocks([makeParagraphNode("TARGET")]);
  return e;
}

function relsOf(e: Mdocxengine, path = "word/_rels/document.xml.rels") {
  const xml = e.zip.readAsText(path) ?? "";
  return [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
    id: /Id="([^"]+)"/.exec(m[0])?.[1] ?? "",
    type: /Type="([^"]+)"/.exec(m[0])?.[1] ?? "",
    target: /Target="([^"]+)"/.exec(m[0])?.[1] ?? "",
  }));
}

async function bodyXml(e: Mdocxengine) {
  return (await e.document.getBlocks()).map((b) => b.xml).join("");
}

/** Resolve every imagedata/OLEObject r:id to { type, bytes }. */
async function resolved(e: Mdocxengine) {
  const body = await bodyXml(e);
  const rels = relsOf(e);
  const pick = (re: RegExp) =>
    [...body.matchAll(re)].map((m) => {
      const rel = rels.find((r) => r.id === m[1]);
      expect(rel, `rId ${m[1]} must resolve`).toBeDefined();
      const entry = e.zip.getEntry(`word/${rel!.target}`);
      expect(entry, `${rel!.target} must exist`).toBeTruthy();
      return { type: rel!.type, target: rel!.target, bytes: entry!.getData() };
    });
  return {
    pics: pick(/<v:imagedata\b[^>]*\br:id="([^"]+)"/g),
    ole: pick(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/g),
  };
}

describe("MergeManager — VML pictures and OLE objects survive a merge", () => {
  test("every imagedata / OLEObject r:id resolves to the right kind with the same bytes", async () => {
    const t = await target();
    await t.merge.appendDocument(await source({ tag: 10 }));
    const { pics, ole } = await resolved(t);
    expect(pics).toHaveLength(2); // VML picture + OLE preview
    expect(pics.every((p) => p.type === IMG)).toBe(true);
    expect(pics[0].bytes.equals(png(10))).toBe(true);
    expect(pics[1].bytes.equals(png(11))).toBe(true);
    expect(ole).toHaveLength(1);
    expect(ole[0].type).toBe(OLE);
    expect(ole[0].bytes.equals(bin(10))).toBe(true);
  });

  test("the bin content type is declared", async () => {
    const t = await target();
    await t.merge.appendDocument(await source({ tag: 10 }));
    expect(t.zip.readAsText("[Content_Types].xml")).toContain(
      'Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"',
    );
  });

  test("two sources with the same part names keep both, byte for byte", async () => {
    const t = await target();
    await t.merge.appendDocument(await source({ tag: 10 }));
    await t.merge.appendDocument(await source({ tag: 40 }));
    const { pics, ole } = await resolved(t);
    expect(pics.map((p) => p.bytes[8])).toEqual([10, 11, 40, 41]);
    expect(ole.map((o) => o.bytes[4])).toEqual([10, 40]);
    expect(new Set(ole.map((o) => o.target)).size).toBe(2);
  });

  test("the merged buffer reloads with target + source blocks", async () => {
    const t = await target();
    const src = await Mdocxengine.loadFromBuffer(await source({ tag: 10 }));
    const srcCount = (await src.document.getBlocks()).length;
    await t.merge.appendDocument(await source({ tag: 10 }));
    const re = await Mdocxengine.loadFromBuffer(t.zip.toBuffer());
    expect((await re.document.getBlocks()).length).toBe(1 + srcCount);
  });

  test("a missing OLE part throws, naming the attribute and the path", async () => {
    const t = await target();
    await expect(t.merge.appendDocument(await source({ tag: 10, missingOle: true }))).rejects.toThrow(
      /r:id.*embeddings\/oleObject1\.bin/,
    );
  });

  test("a missing VML image part throws", async () => {
    const t = await target();
    const buf = await source({ tag: 10 });
    const s = await Mdocxengine.loadFromBuffer(buf);
    s.zip.deleteFile("word/media/vml1.png");
    await expect(t.merge.appendDocument(s.zip.toBuffer())).rejects.toThrow(/imagedata.*vml1\.png/);
  });

  test("an imagedata whose rId is not in the rels throws", async () => {
    const t = await target();
    const s = await Mdocxengine.loadFromBuffer(await source({ tag: 10 }));
    await s.document.saveBlocks([{ kind: "paragraph", tag: "w:p", xml: vmlPara("rIdNope") }]);
    await expect(t.merge.appendDocument(s.zip.toBuffer())).rejects.toThrow(/rIdNope/);
  });

  test("images inside a copied footnote resolve in the merged footnotes rels", async () => {
    const t = await target();
    await t.merge.appendDocument(await source({ tag: 10, footnoteImage: true }));
    const fx = t.zip.readAsText("word/footnotes.xml")!;
    const ids = [...fx.matchAll(/<a:blip\b[^>]*\br:embed="([^"]+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(1);
    const rel = relsOf(t, "word/_rels/footnotes.xml.rels").find((r) => r.id === ids[0]);
    expect(rel, "footnote image rId must resolve").toBeDefined();
    expect(rel!.type).toBe(IMG);
    expect(t.zip.getEntry(`word/${rel!.target}`)!.getData().equals(png(17))).toBe(true);
  });

  test("an external imagedata link is carried as an external relationship; a repeated rId maps once", async () => {
    const t = await target();
    const s = await Mdocxengine.loadFromBuffer(await source({ tag: 10 }));
    await s.merge["rels"].addRelationship("rIdX", IMG, "https://example.com/p.png", "External");
    const para = (attr: string) =>
      `<w:p><w:r><w:pict ${NS_VO}><v:shape id="s9"><v:imagedata ${attr}="rIdX"/></v:shape></w:pict></w:r></w:p>`;
    await s.document.saveBlocks([
      { kind: "paragraph", tag: "w:p", xml: para("r:href") },
      { kind: "paragraph", tag: "w:p", xml: para("r:href") },
    ]);
    await t.merge.appendDocument(s.zip.toBuffer());
    const body = await bodyXml(t);
    const ids = [...body.matchAll(/<v:imagedata\b[^>]*\br:href="([^"]+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe(ids[1]);
    const rel = relsOf(t).find((r) => r.id === ids[0])!;
    expect(rel.target).toBe("https://example.com/p.png");
    expect(t.zip.readAsText("word/_rels/document.xml.rels")).toContain('TargetMode="External"');
  });

  test("a footnote image that is dangling or repeated does not break the copy", async () => {
    const t = await target();
    const s = await Mdocxengine.loadFromBuffer(await source({ tag: 10, footnoteImage: true }));
    const fx = s.zip.readAsText("word/footnotes.xml")!;
    const blip = /<w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/.exec(fx)![0];
    const dangling = blip.replace("rIdFn", "rIdGone");
    s.zip.addFile(
      "word/footnotes.xml",
      Buffer.from(fx.replace(blip, blip + blip + dangling), "utf8"),
    );
    await t.merge.appendDocument(s.zip.toBuffer());
    const out = t.zip.readAsText("word/footnotes.xml")!;
    const embeds = [...out.matchAll(/r:embed="([^"]+)"/g)].map((m) => m[1]);
    expect(embeds).toHaveLength(3);
    expect(embeds[0]).toBe(embeds[1]);
    expect(embeds[2]).toBe("rIdGone");
  });

  test("hyperlinks inside a footnote are carried; anchors are untouched", async () => {
    const t = await target();
    const s = await Mdocxengine.loadFromBuffer(await source({ tag: 10, footnoteImage: true }));
    const fx = s.zip.readAsText("word/footnotes.xml")!;
    const links =
      '<w:hyperlink r:id="rIdH"><w:r><w:t>site</w:t></w:r></w:hyperlink>' +
      '<w:hyperlink w:anchor="bm1"><w:r><w:t>here</w:t></w:r></w:hyperlink>';
    const patched = fx.replace(/(<w:footnote\b[^>]*w:id="1"[^>]*>[\s\S]*?)(<\/w:p>\s*<\/w:footnote>)/, `$1${links}$2`);
    expect(patched).not.toBe(fx);
    s.zip.addFile("word/footnotes.xml", Buffer.from(patched, "utf8"));
    const rels = s.zip.readAsText("word/_rels/footnotes.xml.rels")!.replace(
      "</Relationships>",
      `<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/x" TargetMode="External"/></Relationships>`,
    );
    s.zip.addFile("word/_rels/footnotes.xml.rels", Buffer.from(rels, "utf8"));
    await t.merge.appendDocument(s.zip.toBuffer());

    const out = t.zip.readAsText("word/footnotes.xml")!;
    const ids = [...out.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(1);
    const xml = t.zip.readAsText("word/_rels/footnotes.xml.rels")!;
    const tag = [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((x) => x.includes(`Id="${ids[0]}"`))!;
    expect(tag).toContain('Target="https://example.org/x"');
    expect(tag).toContain('TargetMode="External"');
    expect(tag).toMatch(/relationships\/hyperlink"/);
    expect(out).toContain('<w:hyperlink w:anchor="bm1">');
    // the footnote image still resolves too
    expect(out).toMatch(/r:embed="rId\d+"/);
  });

  describe("footnote relationship ids never collide with the target's", () => {
    const HL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    const hlRel = (id: string, url: string) =>
      `<Relationship Id="${id}" Type="${HL}" Target="${url}" TargetMode="External"/>`;

    /** A docx with ONE footnote whose hyperlinks are `links` (rId -> url), rels in footnotes.xml.rels. */
    async function fnDoc(links: Record<string, string>, buf?: Mdocxengine): Promise<Mdocxengine> {
      const e = buf ?? (await Mdocxengine.loadFromFile("samples/example.docx"));
      const id = (await e.footnotes.addFootnote("n")).id;
      const fx = e.zip.readAsText("word/footnotes.xml")!;
      const xml = Object.keys(links)
        .map((r) => `<w:hyperlink r:id="${r}"><w:r><w:t>${links[r]}</w:t></w:r></w:hyperlink>`)
        .join("");
      const patched = fx.replace(
        new RegExp(`(<w:footnote\\b[^>]*w:id="${id}"[^>]*>[\\s\\S]*?)(</w:p>\\s*</w:footnote>)`),
        `$1${xml}$2`,
      );
      if (xml) expect(patched).not.toBe(fx);
      e.zip.addFile("word/footnotes.xml", Buffer.from(patched, "utf8"));
      const prev = e.zip.readAsText("word/_rels/footnotes.xml.rels");
      const relXml = Object.entries(links).map(([r, u]) => hlRel(r, u)).join("");
      const rels = prev
        ? prev.includes("</Relationships>")
          ? prev.replace("</Relationships>", `${relXml}</Relationships>`)
          : prev.replace(/\/>\s*$/, `>${relXml}</Relationships>`)
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relXml}</Relationships>`;
      e.zip.addFile("word/_rels/footnotes.xml.rels", Buffer.from(rels, "utf8"));
      await e.document.saveBlocks([
        makeParagraphNode("body"),
        { kind: "paragraph", tag: "w:p", xml: `<w:p><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>` },
      ]);
      return e;
    }

    /** Every footnote hyperlink's text (its URL) vs. what its r:id resolves to. */
    function resolveAll(e: Mdocxengine): { text: string; url: string | undefined }[] {
      const fx = e.zip.readAsText("word/footnotes.xml")!;
      const rels = e.zip.readAsText("word/_rels/footnotes.xml.rels")!;
      const url = new Map<string, string>();
      for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        url.set(/\bId="([^"]+)"/.exec(m[0])![1], /\bTarget="([^"]+)"/.exec(m[0])![1]);
      }
      return [...fx.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]+)"[^>]*><w:r><w:t>([^<]+)</g)].map((m) => ({
        text: m[2],
        url: url.get(m[1]),
      }));
    }

    test("a target that already owns rId1 keeps it; the source's links get fresh ids", async () => {
      const t = await fnDoc({ rId1: "https://a.example/A" });
      const src = await fnDoc({ rId1: "https://b.example/B", rId2: "https://c.example/C" });
      await t.merge.appendDocument(src.zip.toBuffer(), {});
      const all = resolveAll(t);
      expect(all.map((x) => x.text).sort()).toEqual([
        "https://a.example/A",
        "https://b.example/B",
        "https://c.example/C",
      ]);
      for (const x of all) expect(x.url).toBe(x.text);
    });

    test("a sequence of appends, each reusing rId1/rId2, never crosses wires", async () => {
      const t = await fnDoc({ rId1: "https://a.example/A" });
      await t.merge.appendDocument((await fnDoc({ rId1: "https://b.example/B" })).zip.toBuffer(), {});
      await t.merge.appendDocument((await fnDoc({ rId1: "https://c.example/C", rId2: "https://d.example/D" })).zip.toBuffer(), {});
      await t.merge.appendDocument((await fnDoc({ rId2: "https://e.example/E", rId1: "https://f.example/F" })).zip.toBuffer(), {});
      const all = resolveAll(t);
      expect(all).toHaveLength(6);
      for (const x of all) expect(x.url).toBe(x.text);
      const ids = [...t.zip.readAsText("word/footnotes.xml")!.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]+)"/g)].map((m) => m[1]);
      expect(new Set(ids).size).toBe(6);
    });

    test("an XML-escaped URL (&amp;) is carried as the same URL, not double-escaped", async () => {
      const t = await fnDoc({});
      const src = await fnDoc({ rId1: "https://b.example/?a=1&amp;b=2", rId2: "https://c.example/C" });
      await t.merge.appendDocument(src.zip.toBuffer(), {});
      const rels = t.zip.readAsText("word/_rels/footnotes.xml.rels")!;
      expect(rels).toContain("https://b.example/?a=1&amp;b=2");
      expect(rels).not.toContain("&amp;amp;");
    });
  });

  describe("every copied image gets its own part", () => {
    /** Source with N OLE objects whose previews are `ext` files, plus 2 png VML pictures. */
    async function previews(ext: string, n: number, base: number): Promise<Buffer> {
      const e = await Mdocxengine.loadFromFile("samples/example.docx");
      const ct = new ContentTypesManager(e.zip);
      await ct.addDefault("png", "image/png");
      await ct.addDefault(ext, `image/x-${ext}`);
      await ct.addDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");
      const blocks: BodyBlock[] = [];
      for (let i = 0; i < n; i++) {
        await e.merge["rels"].addRelationship(`rIdP${i}`, IMG, `media/prev${i}.${ext}`);
        await e.merge["rels"].addRelationship(`rIdO${i}`, OLE, `embeddings/oleObject${i}.bin`);
        e.zip.addFile(`word/media/prev${i}.${ext}`, Buffer.from([base, i, 1, 2, 3]));
        e.zip.addFile(`word/embeddings/oleObject${i}.bin`, Buffer.from([base + 100, i]));
        blocks.push({ kind: "paragraph", tag: "w:p", xml: oleParaX(`rIdP${i}`, `rIdO${i}`) });
      }
      for (let i = 0; i < 2; i++) {
        await e.merge["rels"].addRelationship(`rIdN${i}`, IMG, `media/pic${i}.png`);
        e.zip.addFile(`word/media/pic${i}.png`, Buffer.from([base, 50 + i, 9]));
        blocks.push({ kind: "paragraph", tag: "w:p", xml: vmlPara(`rIdN${i}`) });
      }
      await e.document.saveBlocks(blocks);
      return e.zip.toBuffer();
    }

    for (const ext of ["wmf", "emf", "tif", "gif"]) {
      test(`3 distinct .${ext} previews + 2 png pictures stay 5 distinct parts`, async () => {
        const t = await target();
        await t.merge.appendDocument(await previews(ext, 3, 7), {});
        const { pics } = await resolved(t);
        expect(pics.map((p) => [...p.bytes])).toEqual([
          [7, 0, 1, 2, 3], [7, 1, 1, 2, 3], [7, 2, 1, 2, 3], [7, 50, 9], [7, 51, 9],
        ]);
        expect(new Set(pics.map((p) => p.target)).size).toBe(5);
        expect(t.zip.readAsText("[Content_Types].xml")).toContain(`Extension="${ext}"`);
      });
    }

    test("a target that already holds media/image1.wmf keeps it", async () => {
      const t = await target();
      t.zip.addFile("word/media/image1.wmf", Buffer.from("EXISTING"));
      await t.merge.appendDocument(await previews("wmf", 3, 7), {});
      const { pics } = await resolved(t);
      expect(new Set(pics.map((p) => p.target)).size).toBe(5);
      expect(pics[0].bytes[0]).toBe(7);
      expect(t.zip.getEntry("word/media/image1.wmf")!.getData().toString()).toBe("EXISTING");
    });

    test("DrawingML images with wmf/emf/tif extensions do not collapse either", async () => {
      const e = await Mdocxengine.loadFromFile("samples/example.docx");
      const exts = ["wmf", "emf", "tif", "wmf"];
      const blocks: BodyBlock[] = [];
      for (let i = 0; i < exts.length; i++) {
        await e.merge["rels"].addRelationship(`rIdD${i}`, IMG, `media/d${i}.${exts[i]}`);
        e.zip.addFile(`word/media/d${i}.${exts[i]}`, Buffer.from([1, i]));
        blocks.push({
          kind: "paragraph",
          tag: "w:p",
          xml: `<w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdD${i}"/></w:drawing></w:r></w:p>`,
        });
      }
      await e.document.saveBlocks(blocks);
      const t = await target();
      await t.merge.appendDocument(e.zip.toBuffer(), {});
      const body = await bodyXml(t);
      const rels = relsOf(t);
      const got = [...body.matchAll(/r:embed="([^"]+)"/g)].map((m) => {
        const r = rels.find((x) => x.id === m[1])!;
        return [r.target, [...t.zip.getEntry(`word/${r.target}`)!.getData()]] as const;
      });
      expect(got.map((g) => g[1])).toEqual([[1, 0], [1, 1], [1, 2], [1, 3]]);
      expect(new Set(got.map((g) => g[0])).size).toBe(4);
    });
  });
});
