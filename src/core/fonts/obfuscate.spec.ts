import { describe, it, expect } from "vitest";
import { fontKeyBytes, obfuscateFont, newFontKey } from "./obfuscate";

const GUID = "{00112233-4455-6677-8899-AABBCCDDEEFF}";

describe("fontKeyBytes", () => {
  it("reverses the GUID's bytes", () => {
    expect(fontKeyBytes(GUID).toString("hex")).toBe("ffeeddccbbaa99887766554433221100");
  });
  it("accepts a GUID without braces or dashes", () => {
    expect(fontKeyBytes("00112233445566778899AABBCCDDEEFF")).toEqual(fontKeyBytes(GUID));
  });
  it("rejects anything that is not 32 hex digits", () => {
    expect(() => fontKeyBytes("{not-a-guid}")).toThrow(/Invalid font key/);
  });
});

describe("obfuscateFont", () => {
  it("XORs the first 32 bytes with the key, twice over", () => {
    const out = obfuscateFont(Buffer.alloc(40), GUID);
    const key = fontKeyBytes(GUID);
    expect(out.subarray(0, 16)).toEqual(key);
    expect(out.subarray(16, 32)).toEqual(key);
    expect(out.subarray(32)).toEqual(Buffer.alloc(8));
  });
  it("is its own inverse and never mutates the input", () => {
    const font = Buffer.from(Array.from({ length: 64 }, (_, i) => i));
    const copy = Buffer.from(font);
    expect(obfuscateFont(obfuscateFont(font, GUID), GUID)).toEqual(copy);
    expect(font).toEqual(copy);
  });
  it("handles a font shorter than 32 bytes", () => {
    expect(obfuscateFont(Buffer.alloc(3), GUID).toString("hex")).toBe("ffeedd");
  });
});

describe("newFontKey", () => {
  it("formats 16 random bytes as a braced upper-case GUID", () => {
    const key = newFontKey(() => Buffer.from("00112233445566778899aabbccddeeff", "hex"));
    expect(key).toBe(GUID);
  });
  it("uses crypto randomness by default", () => {
    expect(newFontKey()).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
  });
});
