// ECMA-376 embedded-font obfuscation: XOR the first 32 bytes of the TTF with a
// 16-byte key taken from the w:fontKey GUID in reverse byte order.
import { randomBytes } from "node:crypto";

export function fontKeyBytes(guid: string): Buffer {
  const hex = guid.replace(/[{}-]/g, "");
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) throw new Error(`Invalid font key GUID: ${guid}`);
  const key = Buffer.alloc(16);
  for (let i = 0; i < 16; i++) key[i] = parseInt(hex.substr(30 - 2 * i, 2), 16);
  return key;
}

/** Returns a new buffer; XOR makes the same call de-obfuscate. */
export function obfuscateFont(font: Buffer, guid: string): Buffer {
  const out = Buffer.from(font);
  const key = fontKeyBytes(guid);
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) out[i] ^= key[i % 16];
  return out;
}

export function newFontKey(random: () => Buffer = () => randomBytes(16)): string {
  const h = random().toString("hex").toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20, 32)}}`;
}
