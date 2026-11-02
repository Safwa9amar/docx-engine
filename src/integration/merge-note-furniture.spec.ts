import { describe, test, expect } from "vitest";
import { Mdocxengine, inspectDocx, type DocxZip } from "../index";
import {
  fixture, note, notesPart, noteText, P, PAGE, refP, separators, settingsPart, type FixtureOptions,
} from "./docx-fixture";

/**
 * A note is more than its words: how it is numbered and where it stands live in
 * `word/settings.xml` and in the final `w:sectPr`, its separator in the notes part.
 * A merge has one final section, so the first part with notes of a kind decides for that
 * kind: it brings its properties and separators, and what a later part wanted differently
 * is reported, never silently mixed.
 */

type Kind = "footnote" | "endnote";
const merge = async (target: FixtureOptions, ...sources: FixtureOptions[]) => {
  const t = await Mdocxengine.loadFromBuffer(fixture(target));
  const conflicts = [];
  for (const s of sources) conflicts.push(...(await t.merge.appendDocument(fixture(s))).noteConflicts);
  return {
    t,
    conflicts,
    settings: t.zip.readAsText("word/settings.xml") ?? "",
    doc: t.zip.readAsText("word/document.xml")!,
    part: (kind: Kind) => t.zip.readAsText(`word/${kind}s.xml`)!,
    fatals: inspectDocx(t.zip as unknown as DocxZip).findings.filter((f) => f.severity === "fatal").map((f) => f.rule),
  };
};
const finalSect = (doc: string) => doc.slice(doc.lastIndexOf("<w:sectPr"), doc.lastIndexOf("</w:sectPr>"));
const propsIn = (xml: string, kind: Kind) => new RegExp(`<w:${kind}Pr>[\\s\\S]*?</w:${kind}Pr>`).exec(xml)?.[0];

const withNotes = (kind: Kind, o: { settings?: string; sect?: string; separators?: string; text?: string } = {}): FixtureOptions => ({
  body: refP(kind, 1, o.text ?? "body"),
  [`${kind}s`]: notesPart(kind, note(kind, 1, noteText(kind, o.text ?? "words")), o.separators ?? separators(kind)),
  ...(o.settings ? { settings: settingsPart(o.settings) } : {}),
  ...(o.sect ? { sect: o.sect } : {}),
});
const seqs = (kind: Kind, props: string, ids: number[] = [-1, 0]) =>
  `<w:${kind}Pr>${props}${ids.map((i) => `<w:${kind} w:id="${i}"/>`).join("")}</w:${kind}Pr>`;

describe.each(["footnote", "endnote"] as const)("%s properties", (kind) => {
  const FMT = `<w:numFmt w:val="decimalEnclosedParen"/>`;
  const RESTART = `<w:numRestart w:val="eachPage"/>`;

  test("a target with no such notes takes the part's, in settings and in the final section", async () => {
    const { settings, doc, fatals } = await merge(
      { body: P("cover") },
      withNotes(kind, { settings: seqs(kind, FMT + RESTART), sect: `<w:sectPr><w:${kind}Pr>${FMT}</w:${kind}Pr>${PAGE}</w:sectPr>` }),
    );
    expect(propsIn(settings, kind)).toBe(seqs(kind, FMT + RESTART));
    expect(propsIn(finalSect(doc), kind)).toBe(`<w:${kind}Pr>${FMT}</w:${kind}Pr>`);
    expect(fatals).toEqual([]);
  });

  test("written where the schemas want them", async () => {
    const { settings, doc } = await merge(
      { body: P("cover"), settings: settingsPart(`<w:zoom w:percent="100"/><w:compat/>`), sect: `<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>${PAGE}</w:sectPr>`, docRels: `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`, parts: { "word/header1.xml": `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>` }, contentTypes: `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` },
      withNotes(kind, { settings: seqs(kind, FMT), sect: `<w:sectPr><w:${kind}Pr>${FMT}</w:${kind}Pr>${PAGE}</w:sectPr>` }),
    );
    expect(settings.indexOf("<w:zoom")).toBeLessThan(settings.indexOf(`<w:${kind}Pr>`));
    expect(settings.indexOf(`<w:${kind}Pr>`)).toBeLessThan(settings.indexOf("<w:compat"));
    const sect = finalSect(doc);
    expect(sect.indexOf("<w:headerReference")).toBeLessThan(sect.indexOf(`<w:${kind}Pr>`));
    expect(sect.indexOf(`<w:${kind}Pr>`)).toBeLessThan(sect.indexOf("<w:pgSz"));
  });

  test("a target's stale setting gives way when it has none of those notes, and is cleared if the part has none", async () => {
    const { settings, doc } = await merge(
      { body: P("cover"), settings: settingsPart(seqs(kind, RESTART)), sect: `<w:sectPr><w:${kind}Pr>${RESTART}</w:${kind}Pr>${PAGE}</w:sectPr>` },
      withNotes(kind),
    );
    expect(propsIn(settings, kind)).toBeUndefined();
    expect(propsIn(finalSect(doc), kind)).toBeUndefined();
  });

  test("no settings part: one is made, registered, when the part has something to say", async () => {
    const { t, settings } = await merge({ body: P("cover") }, withNotes(kind, { settings: seqs(kind, FMT) }));
    expect(propsIn(settings, kind)).toBe(seqs(kind, FMT));
    expect(t.zip.readAsText("word/_rels/document.xml.rels")).toContain('Target="settings.xml"');
  });

  test("a target with those notes keeps its own, and the difference is reported, property by property", async () => {
    const { settings, conflicts } = await merge(
      withNotes(kind, { settings: seqs(kind, RESTART) }),
      withNotes(kind, { settings: seqs(kind, FMT) }),
    );
    expect(propsIn(settings, kind)).toBe(seqs(kind, RESTART));
    expect(conflicts.map((c) => c.property).sort()).toEqual(["numFmt", "numRestart"]);
    expect(conflicts.find((c) => c.property === "numRestart")).toMatchObject({ kind, kept: "eachPage", dropped: "continuous" });
  });

  test("the same effective properties are no conflict: a default written out, or a section saying what settings says", async () => {
    const defaults = kind === "footnote" ? `<w:numFmt w:val="decimal"/>` : `<w:numFmt w:val="lowerRoman"/>`;
    const { conflicts } = await merge(
      withNotes(kind),
      withNotes(kind, { settings: seqs(kind, defaults) }),
      withNotes(kind, { sect: `<w:sectPr><w:${kind}Pr>${defaults}</w:${kind}Pr>${PAGE}</w:sectPr>` }),
    );
    expect(conflicts).toEqual([]);
  });

  test("a start value is no conflict: a part set to carry on from the last one", async () => {
    const { conflicts } = await merge(withNotes(kind), withNotes(kind, { settings: seqs(kind, `<w:numStart w:val="31"/>`) }));
    expect(conflicts).toEqual([]);
  });

  test("a part with no notes of this kind changes nothing, whatever it says", async () => {
    const { settings, conflicts } = await merge(
      { body: P("cover"), settings: settingsPart(seqs(kind, RESTART)) },
      { body: P("no notes"), settings: settingsPart(seqs(kind, FMT)) },
    );
    expect(propsIn(settings, kind)).toBe(seqs(kind, RESTART));
    expect(conflicts).toEqual([]);
  });

  test("a third part is compared against what the second part brought", async () => {
    const { settings, conflicts } = await merge(
      { body: P("cover") },
      withNotes(kind, { settings: seqs(kind, FMT) }),
      withNotes(kind, { settings: seqs(kind, FMT) }),
      withNotes(kind, { settings: seqs(kind, RESTART) }),
    );
    expect(propsIn(settings, kind)).toBe(seqs(kind, FMT));
    expect(conflicts.map((c) => c.property).sort()).toEqual(["numFmt", "numRestart"]);
  });

  test("a mid-document section of the part keeps what it said", async () => {
    const mid = `<w:p><w:pPr><w:sectPr><w:${kind}Pr>${FMT}</w:${kind}Pr>${PAGE}</w:sectPr></w:pPr></w:p>`;
    const source = withNotes(kind, { settings: seqs(kind, FMT) });
    const { doc } = await merge({ body: P("cover") }, { ...source, body: mid + source.body });
    expect(doc).toContain(mid);
  });
});

describe.each(["footnote", "endnote"] as const)("%s separators", (kind) => {
  const BIDI = "<w:bidi/>";

  test("no notes part in the target: it gets the part's own separators", async () => {
    const { part, fatals } = await merge({ body: P("cover") }, withNotes(kind, { separators: separators(kind, BIDI) }));
    expect(part(kind)).toContain(separators(kind, BIDI));
    expect(part(kind).match(new RegExp(`w:type="separator"`, "g"))).toHaveLength(1);
    expect(fatals).toEqual([]);
  });

  test("a notes part with no notes in use: the part's separators replace the unused ones", async () => {
    const { part } = await merge({ body: P("cover"), [`${kind}s`]: notesPart(kind, "") }, withNotes(kind, { separators: separators(kind, BIDI) }));
    expect(part(kind)).toContain(separators(kind, BIDI));
    expect(part(kind).match(new RegExp(`w:type="separator"`, "g"))).toHaveLength(1);
  });

  test("a target with notes in use keeps its separators", async () => {
    const { part } = await merge(withNotes(kind), withNotes(kind, { separators: separators(kind, BIDI) }));
    expect(part(kind)).toContain(separators(kind));
    expect(part(kind)).not.toContain(BIDI);
  });

  test("older Word's ids, 0 and 1: the settings follow them, and the first real note clears them", async () => {
    const { part, settings } = await merge(
      { body: P("cover"), settings: settingsPart(seqs(kind, "")) },
      withNotes(kind, { separators: separators(kind, BIDI, [0, 1]), settings: seqs(kind, "", [0, 1]) }),
    );
    expect(part(kind)).toContain(separators(kind, BIDI, [0, 1]));
    expect(propsIn(settings, kind)).toBe(seqs(kind, "", [0, 1]));
    expect(new RegExp(`<w:${kind} w:id="2">`).test(part(kind))).toBe(true);
  });

  test("a target's settings keep their separator list when the part brings none", async () => {
    const { settings, part } = await merge(
      { body: P("cover"), [`${kind}s`]: notesPart(kind, ""), settings: settingsPart(seqs(kind, "")) },
      withNotes(kind, { separators: "" }),
    );
    expect(part(kind)).toContain(separators(kind));
    expect(propsIn(settings, kind)).toBe(seqs(kind, ""));
  });

  test("a note left unreferenced in the target keeps its id clear of the part's separators", async () => {
    const { part } = await merge(
      { body: P("cover"), [`${kind}s`]: notesPart(kind, note(kind, 1, noteText(kind, "orphan"))) },
      withNotes(kind, { separators: separators(kind, BIDI, [0, 1]) }),
    );
    const ids = [...part(kind).matchAll(new RegExp(`<w:${kind}\\b[^>]*\\bw:id="(-?\\d+)"`, "g"))].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(part(kind)).toContain("orphan");
  });
});

describe("a note marked w:type=\"normal\" is a note", () => {
  test.each(["footnote", "endnote"] as const)("%s: it travels, and its reference reaches it", async (kind) => {
    const { part, doc } = await merge(
      { body: P("cover") },
      { body: refP(kind, 1), [`${kind}s`]: notesPart(kind, note(kind, 1, noteText(kind, "typed normal"), 'w:type="normal"')) },
    );
    const id = new RegExp(`<w:${kind}Reference w:id="(\\d+)"`).exec(doc)![1];
    expect(part(kind)).toMatch(new RegExp(`w:type="normal" w:id="${id}">[\\s\\S]*typed normal`));
  });
});
