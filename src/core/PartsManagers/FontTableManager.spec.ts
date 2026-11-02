// src/core/PartsManagers/FontTableManager.spec.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import AdmZip from "adm-zip";
import { Mdocxengine } from "../../index";
import { checkDocxBuffer } from "../../doctor";
import { obfuscateFont } from "../fonts/obfuscate";

const SAMPLE = readFileSync(new URL("../../../samples/example.docx", import.meta.url));

/** sfnt with an OS/2 table carrying `fsType`, padded so obfuscation has 32+ bytes. */
function fakeFont(fsType: number, fill = 7): Buffer {
  const header = Buffer.alloc(12);
  header.writeUInt32BE(0x00010000, 0);
  header.writeUInt16BE(1, 4);
  const rec = Buffer.alloc(16);
  rec.write("OS/2", 0, "latin1");
  rec.writeUInt32BE(28, 8);
  rec.writeUInt32BE(10, 12);
  const table = Buffer.alloc(10);
  table.writeUInt16BE(fsType, 8);
  return Buffer.concat([header, rec, table, Buffer.alloc(40, fill)]);
}
let n = 0;
const fixedKey = () => `{00000000-0000-0000-0000-${String(++n).padStart(12, "0")}}`;

async function load(mutate?: (z: AdmZip) => void) {
  const z = new AdmZip(SAMPLE);
  mutate?.(z);
  const engine = await Mdocxengine.loadFromBuffer(z.toBuffer());
  engine.fontTable.newKey = fixedKey;
  return engine;
}
const text = (e: Mdocxengine, p: string) => e.zip.readAsText(p) ?? "";

describe("FontTableManager.embedFont", () => {
  it("embeds an editable font: odttf part, rel, fontTable entry, content type, settings flag", async () => {
    const e = await load();
    const font = fakeFont(0x0008);
    const res = await e.fontTable.embedFont({ name: "Amiri", regular: font });
    expect(res).toEqual({ name: "Amiri", embedded: true, variants: ["regular"], settingsUpdated: true });

    const ft = text(e, "word/fontTable.xml");
    const m = ft.match(/<w:font\b[^>]*w:name="Amiri"[^>]*>([\s\S]*?)<\/w:font>/);
    expect(m).not.toBeNull();
    const embed = m![1].match(/<w:embedRegular r:id="(rId\d+)" w:fontKey="(\{[^"]+\})"\/>/)!;
    const rels = text(e, "word/_rels/fontTable.xml.rels");
    expect(rels).toContain(`Id="${embed[1]}"`);
    expect(rels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"');
    const target = rels.match(new RegExp(`Id="${embed[1]}"[^>]*Target="([^"]+)"`))![1];
    const stored = e.zip.getEntry(`word/${target}`)!.getData();
    expect(obfuscateFont(stored, embed[2])).toEqual(font);
    expect(text(e, "[Content_Types].xml")).toMatch(/Extension="odttf"[^>]*obfuscatedFont/);
    expect(text(e, "word/settings.xml")).toContain("<w:embedTrueTypeFonts/>");

    const report = await checkDocxBuffer(Buffer.from(e.zip.toBuffer()));
    expect(report.findings.filter((f) => f.severity === "fatal")).toEqual([]);
  });

  it("refuses a restricted font and writes nothing", async () => {
    const e = await load();
    const before = text(e, "word/fontTable.xml");
    const res = await e.fontTable.embedFont({ name: "Simplified Arabic", regular: fakeFont(0x0002) });
    expect(res).toEqual({ name: "Simplified Arabic", embedded: false, reason: "restricted", variants: [], settingsUpdated: false });
    expect(text(e, "word/fontTable.xml")).toBe(before);
    expect(e.zip.getEntries().some((x) => x.entryName.startsWith("word/fonts/"))).toBe(false);
  });

  it("refuses all variants when one variant is not embeddable", async () => {
    const e = await load();
    const res = await e.fontTable.embedFont({ name: "Mixed", regular: fakeFont(0), bold: fakeFont(0x0004) });
    expect(res.embedded).toBe(false);
    expect(res.reason).toBe("preview-print");
  });

  it("refuses a font whose licence cannot be read", async () => {
    const e = await load();
    const res = await e.fontTable.embedFont({ name: "Broken", regular: Buffer.alloc(8) });
    expect(res).toMatchObject({ embedded: false, reason: "unknown" });
  });

  it("embeds every variant in schema order and numbers parts and rels upward", async () => {
    const e = await load();
    await e.fontTable.embedFont({ name: "A", regular: fakeFont(0) });
    const res = await e.fontTable.embedFont({
      name: "B", regular: fakeFont(0), bold: fakeFont(0), italic: fakeFont(0), boldItalic: fakeFont(0),
    });
    expect(res.variants).toEqual(["regular", "bold", "italic", "boldItalic"]);
    const b = text(e, "word/fontTable.xml").match(/<w:font\b[^>]*w:name="B"[^>]*>([\s\S]*?)<\/w:font>/)![1];
    const order = [...b.matchAll(/<w:(embed\w+)\b/g)].map((x) => x[1]);
    expect(order).toEqual(["embedRegular", "embedBold", "embedItalic", "embedBoldItalic"]);
    expect(e.zip.getEntry("word/fonts/font5.odttf")).toBeTruthy();
    expect(e.fontTable.embeddedFonts()).toEqual(["A", "B"]);
  });

  it("replaces an existing entry's embeds and keeps its other children", async () => {
    const e = await load((z) => z.addFile("word/fontTable.xml", Buffer.from(
      '<?xml version="1.0"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:font w:name="Amiri"><w:charset w:val="00"/><w:embedRegular r:id="rId9" w:fontKey="{X}"/></w:font>' +
      '<w:font w:name="Tahoma"/></w:fonts>')));
    await e.fontTable.embedFont({ name: "Amiri", regular: fakeFont(0) });
    await e.fontTable.embedFont({ name: "Tahoma", regular: fakeFont(0) });
    const ft = text(e, "word/fontTable.xml");
    expect(ft).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    expect(ft).toMatch(/<w:font w:name="Amiri"><w:charset w:val="00"\/><w:embedRegular r:id="rId\d+" w:fontKey="\{0/);
    expect(ft).not.toContain('r:id="rId9"');
    expect(ft).toMatch(/<w:font w:name="Tahoma"><w:embedRegular /);
    expect(e.fontTable.listFonts()).toEqual(["Amiri", "Tahoma"]);
  });

  it("creates fontTable.xml, its override and document rel when the file has none", async () => {
    const e = await load((z) => {
      z.deleteFile("word/fontTable.xml");
      const rels = z.readAsText("word/_rels/document.xml.rels").replace(/<Relationship\b[^>]*fontTable[^>]*\/>/g, "");
      z.addFile("word/_rels/document.xml.rels", Buffer.from(rels));
    });
    expect(e.fontTable.listFonts()).toEqual([]);
    await e.fontTable.embedFont({ name: "Amiri", regular: fakeFont(0) });
    expect(text(e, "word/fontTable.xml")).toContain('w:name="Amiri"');
    expect(text(e, "[Content_Types].xml")).toContain('PartName="/word/fontTable.xml"');
    expect(text(e, "word/_rels/document.xml.rels")).toMatch(/relationships\/fontTable"[^>]*Target="fontTable.xml"|Target="fontTable.xml"[^>]*relationships\/fontTable"/);
  });

  it("does not duplicate an existing document rel to fontTable.xml", async () => {
    const e = await load((z) => z.deleteFile("word/fontTable.xml"));
    await e.fontTable.embedFont({ name: "Amiri", regular: fakeFont(0) });
    expect(text(e, "word/_rels/document.xml.rels").match(/Target="fontTable\.xml"/g)).toHaveLength(1);
  });

  it("places embedTrueTypeFonts after the settings it must follow, once", async () => {
    const e = await load((z) => z.addFile("word/settings.xml", Buffer.from(
      '<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/></w:settings>')));
    await e.fontTable.embedFont({ name: "A", regular: fakeFont(0) });
    await e.fontTable.embedFont({ name: "B", regular: fakeFont(0) });
    const s = text(e, "word/settings.xml");
    expect(s).toContain('<w:zoom w:percent="100"/><w:embedTrueTypeFonts/><w:defaultTabStop');
    expect(s.match(/embedTrueTypeFonts/g)).toHaveLength(1);
  });

  it("puts embedTrueTypeFonts first when no earlier setting exists", async () => {
    const e = await load((z) => z.addFile("word/settings.xml", Buffer.from(
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/></w:settings>')));
    await e.fontTable.embedFont({ name: "A", regular: fakeFont(0) });
    expect(text(e, "word/settings.xml")).toMatch(/<w:settings [^>]*><w:embedTrueTypeFonts\/><w:defaultTabStop/);
  });

  it("reports settingsUpdated:false when the document has no settings part", async () => {
    const e = await load((z) => z.deleteFile("word/settings.xml"));
    const res = await e.fontTable.embedFont({ name: "A", regular: fakeFont(0) });
    expect(res.settingsUpdated).toBe(false);
    expect(res.embedded).toBe(true);
  });

  it("escapes the font name in XML", async () => {
    const e = await load();
    await e.fontTable.embedFont({ name: 'A&B "C"', regular: fakeFont(0) });
    expect(text(e, "word/fontTable.xml")).toContain('w:name="A&amp;B &quot;C&quot;"');
    expect(e.fontTable.listFonts()).toContain('A&B "C"');
  });

  it("generates a fresh GUID font key by default", async () => {
    const e = await load();
    e.fontTable = new (e.fontTable.constructor as any)(e.zip, e.contentTypes);
    const font = fakeFont(0);
    await e.fontTable.embedFont({ name: "Amiri", regular: font });
    const key = text(e, "word/fontTable.xml").match(/w:fontKey="(\{[0-9A-F-]{36}\})"/)![1];
    expect(key).not.toMatch(/^\{00000000-0000-0000-0000-/);
    expect(obfuscateFont(e.zip.getEntry("word/fonts/font1.odttf")!.getData(), key)).toEqual(font);
  });
});
