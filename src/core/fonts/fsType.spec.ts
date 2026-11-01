import { describe, it, expect } from "vitest";
import { readFsType, embeddingPermission, canEmbed } from "./fsType";

/** A minimal sfnt: header + one table record + a 10-byte table. */
function fakeFont(tag: string, fsType = 0, tableLen = 10): Buffer {
  const header = Buffer.alloc(12);
  header.writeUInt32BE(0x00010000, 0);
  header.writeUInt16BE(1, 4);
  const rec = Buffer.alloc(16);
  rec.write(tag, 0, "latin1");
  rec.writeUInt32BE(28, 8);
  rec.writeUInt32BE(tableLen, 12);
  const table = Buffer.alloc(tableLen);
  if (tableLen >= 10) table.writeUInt16BE(fsType, 8);
  return Buffer.concat([header, rec, table]);
}

describe("readFsType", () => {
  it("reads fsType from the OS/2 table", () => {
    expect(readFsType(fakeFont("OS/2", 0x0008))).toBe(0x0008);
  });
  it("is null for a buffer shorter than an sfnt header", () => {
    expect(readFsType(Buffer.alloc(4))).toBeNull();
  });
  it("is null when the font has no OS/2 table", () => {
    expect(readFsType(fakeFont("head"))).toBeNull();
  });
  it("is null when a table record runs past the end", () => {
    const f = fakeFont("OS/2");
    expect(readFsType(f.subarray(0, 20))).toBeNull();
  });
  it("is null when the OS/2 table is truncated", () => {
    expect(readFsType(fakeFont("OS/2", 0, 6))).toBeNull();
  });
});

describe("embeddingPermission", () => {
  it.each([
    [null, "unknown"],
    [0x0000, "installable"],
    [0x0008, "editable"],
    [0x0004, "preview-print"],
    [0x0002, "restricted"],
    [0x000a, "editable"], // least restrictive bit wins
    [0x0208, "restricted"], // bitmap-only embedding cannot carry outlines
    [0x0001, "unknown"], // reserved bit only
  ] as const)("fsType %s → %s", (fs, want) => {
    expect(embeddingPermission(fs)).toBe(want);
  });
});

describe("canEmbed", () => {
  it("allows only installable and editable fonts", () => {
    expect(canEmbed("installable")).toBe(true);
    expect(canEmbed("editable")).toBe(true);
    expect(canEmbed("preview-print")).toBe(false);
    expect(canEmbed("restricted")).toBe(false);
    expect(canEmbed("unknown")).toBe(false);
  });
});
