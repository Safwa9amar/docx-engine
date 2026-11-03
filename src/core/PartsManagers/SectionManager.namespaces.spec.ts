import { describe, test, expect } from "vitest";
import AdmZip from "adm-zip";
import { SectionManager } from "./SectionManager";

describe("a section reference on a root that declares only xmlns:w", () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const doc = `<?xml version="1.0"?><w:document ${W}><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

  test.each(["setSectionHeader", "setSectionFooter"] as const)("%s declares r on the root", async (method) => {
    const zip = new AdmZip();
    zip.addFile("word/document.xml", Buffer.from(doc, "utf-8"));
    await new SectionManager(zip as any)[method](0, "rId5");
    const xml = zip.readAsText("word/document.xml");
    expect(xml).toContain('r:id="rId5"');
    expect(xml).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  });
});
