import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { ContentTypesManager } from "./ContentTypesManager";

const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="ODTTF" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Default Extension="xml" ContentType="application/xml"/></Types>`;

describe("addDefault", () => {
  test("an extension differing only in case is the same extension: Word refuses a second Default", async () => {
    const zip = new AdmZip();
    zip.addFile("[Content_Types].xml", Buffer.from(CT, "utf-8"));
    await new ContentTypesManager(zip as any).addDefault("odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont");
    const xml = zip.readAsText("[Content_Types].xml");
    expect((xml.match(/<Default\b[^>]*Extension="odttf"/gi) ?? []).length).toBe(1);
  });
});
