import { describe, test, expect } from "vitest";
import { Table } from ".";
import { makeTableXml } from "../body/OrderedBody";

describe("Table.fromXml", () => {
  test("parses a <w:tbl> string and reads its cells", async () => {
    const xml = makeTableXml([["Name", "Role"], ["Ada", "Author"]], { headerRow: true });
    const t = await Table.fromXml(xml);
    expect(t.getAllCellText()).toEqual([["Name", "Role"], ["Ada", "Author"]]);
    expect(t.getRowCount()).toBe(2);
  });

  test("round-trips fromGrid → toObject XML → fromXml", async () => {
    const grid = [["A", "B", "C"], ["1", "2", "3"]];
    const built = Table.fromGrid(grid);
    const xml = makeTableXml(grid);
    const reparsed = await Table.fromXml(xml);
    expect(reparsed.getAllCellText()).toEqual(built.getAllCellText());
  });
});

// ── Run-edge whitespace: the space-deletion data-loss bug ───────────────────
// Word splits a sentence across many runs, carrying spaces at run edges or in
// space-only `xml:space="preserve"` runs. parseXml's default trim deleted them,
// and because mutateTable round-trips fromXml → buildXml → SAVE, every table op
// wrote the glued text back into the document. Arabic theses (one run per word
// is routine there) glued into "اسمالجدول" for "اسم الجدول".
describe("Table.fromXml — whitespace preservation", () => {
  const CELL = (runs: string) =>
    `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:tr><w:tc><w:p>${runs}</w:p></w:tc></w:tr></w:tbl>`;

  test("a space-only xml:space='preserve' run survives", async () => {
    const t = await Table.fromXml(
      CELL(
        `<w:r><w:t>اسم</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve"> </w:t></w:r>` +
          `<w:r><w:t>الجدول</w:t></w:r>`,
      ),
    );
    expect(t.getAllCellText()).toEqual([["اسم الجدول"]]);
  });

  test("a trailing space INSIDE a run survives", async () => {
    const t = await Table.fromXml(
      CELL(`<w:r><w:t xml:space="preserve">Nom </w:t></w:r><w:r><w:t>complet</w:t></w:r>`),
    );
    expect(t.getAllCellText()).toEqual([["Nom complet"]]);
  });

  test("the mutate round-trip no longer destroys spaces (the SAVE path)", async () => {
    const { buildXml } = await import("@/utils/xmlUtils");
    const t = await Table.fromXml(
      CELL(
        `<w:r><w:t>تُرسل</w:t></w:r>` +
          `<w:r><w:t xml:space="preserve"> </w:t></w:r>` +
          `<w:r><w:t>البرقيات</w:t></w:r>`,
      ),
    );
    // The exact rebuild mutateTable performs before saveBlocks.
    const rebuilt = buildXml(t.toObject(), { rootName: "w:tbl", headless: true, pretty: false });
    const again = await Table.fromXml(rebuilt);
    expect(again.getAllCellText()).toEqual([["تُرسل البرقيات"]]);
    // And the serialized XML itself still carries the space-only run.
    expect(rebuilt).toContain('xml:space="preserve"');
  });
});
