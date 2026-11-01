// A font's licence for embedding lives in its OS/2 table's fsType field
// (OpenType spec, "OS/2 — fsType"). Simplified Arabic and Traditional Arabic
// carry 0x0002 (restricted): Word refuses to embed them and so must we.

export type EmbeddingPermission = "installable" | "editable" | "preview-print" | "restricted" | "unknown";

/** fsType from the OS/2 table, or null when the buffer is not a usable sfnt. */
export function readFsType(font: Buffer): number | null {
  if (font.length < 12) return null;
  const numTables = font.readUInt16BE(4);
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + 16 * i;
    if (rec + 16 > font.length) return null;
    if (font.toString("latin1", rec, rec + 4) !== "OS/2") continue;
    const offset = font.readUInt32BE(rec + 8);
    if (offset + 10 > font.length) return null;
    return font.readUInt16BE(offset + 8);
  }
  return null;
}

/** Bits 1–3 are the usage permissions; when several are set the least restrictive applies. */
export function embeddingPermission(fsType: number | null): EmbeddingPermission {
  if (fsType === null) return "unknown";
  if (fsType & 0x0200) return "restricted";
  const usage = fsType & 0x000f;
  if (usage === 0) return "installable";
  if (usage & 0x0008) return "editable";
  if (usage & 0x0004) return "preview-print";
  if (usage & 0x0002) return "restricted";
  return "unknown";
}

/** We only embed fonts the student may keep editing with. */
export function canEmbed(p: EmbeddingPermission): boolean {
  return p === "installable" || p === "editable";
}
