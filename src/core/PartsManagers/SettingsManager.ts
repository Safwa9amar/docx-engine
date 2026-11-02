/**
 * word/settings.xml writer — string surgery only.
 *
 * `CT_Settings` is an `xsd:sequence` of ~100 optional elements; Word refuses a
 * settings part whose children are out of order, and the xml2js round trip used
 * by `setEvenAndOddHeaders` rebuilds the whole part. This inserts or removes ONE
 * element at its schema position and leaves every other byte alone.
 */

import AdmZip from "adm-zip";
import { ContentTypesManager } from "./ContentTypesManager";
import { RelManager } from "./RelManager";

const SETTINGS_PATH = "word/settings.xml";
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const SETTINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

/** `CT_Settings` child sequence, ECMA-376 Part 1 17.15.1.78 (without the `w:` prefix). */
export const CT_SETTINGS_ORDER: readonly string[] = [
  "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
  "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
  "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
  "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
  "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
  "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate",
  "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
  "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection",
  "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
  "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope", "summaryLength",
  "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting",
  "bookFoldPrinting", "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
  "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
  "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
  "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne",
  "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
  "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
  "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
  "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
  "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
  "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions",
  "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
  "decimalSymbol", "listSeparator",
];

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const escRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Remove the first `<w:name …/>` / `<w:name …>…</w:name>`; returns its [start,end) or null. */
function findElement(xml: string, name: string): { start: number; end: number } | null {
  const re = new RegExp(`<w:${escRe(name)}(?=[\\s/>])[^>]*?(?:/>|>[\\s\\S]*?</w:${escRe(name)}>)`);
  const m = re.exec(xml);
  return m ? { start: m.index, end: m.index + m[0].length } : null;
}

/**
 * Pure transform. `attrs === null` removes the element; an object writes
 * `<w:name a="b"/>` (keys without a prefix get `w:`). An existing element is
 * replaced IN PLACE, a new one goes before the first later-in-sequence sibling
 * present (else before `</w:settings>`).
 *
 * @throws on a name that is not a `CT_Settings` child, and on a part with no
 * `<w:settings>` root.
 */
export function applySettingToXml(
  settingsXml: string,
  name: string,
  attrs: Record<string, string> | null,
): { xml: string; changed: boolean } {
  const rank = CT_SETTINGS_ORDER.indexOf(name);
  if (rank === -1) throw new Error(`SettingsManager.setSetting: "${name}" is not a CT_Settings element`);

  // `<w:settings …/>` (no children) → expand so there is somewhere to insert.
  let xml = settingsXml.replace(/<w:settings\b([^>]*?)\/>/, (_m, a) => `<w:settings${a}></w:settings>`);
  if (!/<\/w:settings>/.test(xml)) throw new Error("SettingsManager.setSetting: no <w:settings> root");

  const existing = findElement(xml, name);
  if (attrs === null) {
    if (!existing) return { xml: settingsXml, changed: false };
    return { xml: xml.slice(0, existing.start) + xml.slice(existing.end), changed: true };
  }

  const attrText = Object.entries(attrs)
    .map(([k, v]) => ` ${k.includes(":") ? k : `w:${k}`}="${escAttr(v)}"`)
    .join("");
  const el = `<w:${name}${attrText}/>`;

  if (existing) {
    if (xml.slice(existing.start, existing.end) === el) return { xml: settingsXml, changed: false };
    return { xml: xml.slice(0, existing.start) + el + xml.slice(existing.end), changed: true };
  }

  let at = xml.lastIndexOf("</w:settings>");
  for (const later of CT_SETTINGS_ORDER.slice(rank + 1)) {
    const f = findElement(xml, later);
    if (f && f.start < at) at = f.start;
  }
  xml = xml.slice(0, at) + el + xml.slice(at);
  return { xml, changed: true };
}

export class SettingsManager {
  private zip: AdmZip;

  constructor(zip: AdmZip) {
    this.zip = zip;
  }

  /**
   * Add, change or remove (`attrs === null`) ONE `word/settings.xml` element —
   * e.g. `w:updateFields`, `w:mirrorMargins`, `w:evenAndOddHeaders`. A package
   * with no settings part gets a minimal one (content type + relationship).
   */
  public async setSetting(
    name: string,
    attrs: Record<string, string> | null,
  ): Promise<{ changed: boolean }> {
    let xml = this.zip.readAsText(SETTINGS_PATH);
    const fresh = !xml;
    if (fresh) {
      if (attrs === null) return { changed: false };
      xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W_NS}"></w:settings>`;
    }
    const r = applySettingToXml(xml, name, attrs);
    if (!r.changed) return { changed: false };
    this.zip.addFile(SETTINGS_PATH, Buffer.from(r.xml, "utf-8"));
    if (fresh) {
      await new ContentTypesManager(this.zip).addOverride(`/${SETTINGS_PATH}`, SETTINGS_CT);
      const rels = new RelManager(this.zip);
      await rels.addRelationship(await rels.genId(), SETTINGS_REL, "settings.xml");
    }
    return { changed: true };
  }
}
