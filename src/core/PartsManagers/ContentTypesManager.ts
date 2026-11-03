import { ZipManager } from "@/utils/ZipManager";
import { parseXml, buildXml } from "@/utils/xmlUtils";
import logger from "@/utils/Logger";
import AdmZip from "adm-zip";

/**
 * Represents a single Default entry in [Content_Types].xml
 */
interface ContentTypeDefault {
  $: {
    Extension: string;
    ContentType: string;
  };
}

/**
 * Represents a single Override entry in [Content_Types].xml
 */
interface ContentTypeOverride {
  $: {
    PartName: string;
    ContentType: string;
  };
}

/**
 * Root structure of [Content_Types].xml
 */
interface ContentTypesXml {
  Types: {
    $: { xmlns: string };
    Default: ContentTypeDefault[];
    Override: ContentTypeOverride[];
  };
}

export class ContentTypesManager {
  private readonly zip: AdmZip;
  private readonly filePath = "[Content_Types].xml";
  private readonly ns = "http://schemas.openxmlformats.org/package/2006/content-types";

  constructor(zip: AdmZip) {
    this.zip = zip;
  }

  /** Generate a GUID */
  private generateGuid(): string {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (char) => {
      const rand = (Math.random() * 16) | 0;
      const value = char === "x" ? rand : (rand & 0x3) | 0x8;
      return value.toString(16);
    });
  }

  /**
   * Reads current [Content_Types].xml or returns a default structure
   */
  private async readTypes(): Promise<ContentTypesXml> {
    const xml = this.zip.readAsText(this.filePath);

    if (!xml) {
      return {
        Types: {
          $: { xmlns: this.ns },
          Default: [],
          Override: [],
        },
      };
    }

    const parsed = await parseXml(xml);

    // Normalize arrays
    parsed.Types.Default = Array.isArray(parsed.Types.Default)
      ? parsed.Types.Default
      : [parsed.Types.Default].filter(Boolean);
    parsed.Types.Override = Array.isArray(parsed.Types.Override)
      ? parsed.Types.Override
      : [parsed.Types.Override].filter(Boolean);

    return parsed as ContentTypesXml;
  }

  /**
   * Writes back the [Content_Types].xml into the zip
   */
  private async writeTypes(obj: ContentTypesXml): Promise<void> {
    const xml = buildXml(obj.Types, { headless: false, pretty: true, rootName: "Types" });
    this.zip.addFile(this.filePath, Buffer.from(xml, "utf-8"));
  }

  /**
   * Adds a Default content type if not exists
   * Example: Extension="xml" ContentType="application/xml"
   */
  async addDefault(extension: string, contentType: string): Promise<void> {
    const typesObj = await this.readTypes();

    // OPC compares extensions without case, and Word refuses a second Default for one.
    const exists = typesObj.Types.Default.some((d) => d.$.Extension.toLowerCase() === extension.toLowerCase());
    if (!exists) {
      typesObj.Types.Default.push({
        $: { Extension: extension, ContentType: contentType },
      });
      await this.writeTypes(typesObj);
      logger.info(`Added Default content type for extension: ${extension}`);
    }
  }

  /**
   * Adds an Override element if it doesn't exist yet.
   * Example:
   * PartName="/word/header1.xml"
   * ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
   */
  async addOverride(partName: string, contentType: string): Promise<void> {
    const typesObj = await this.readTypes();
    const overrides = typesObj.Types.Override;

    const exists = overrides.some((o) => o.$.PartName === partName);
    if (!exists) {
      overrides.push({
        $: {
          PartName: partName,
          ContentType: contentType,
        },
      });

      await this.writeTypes(typesObj);
      logger.info(`Added Override for part: ${partName}`);
    }
  }

  /**
   * Removes an Override entry by PartName
   */
  async removeOverride(partName: string): Promise<void> {
    const typesObj = await this.readTypes();
    typesObj.Types.Override = typesObj.Types.Override.filter((o) => o.$.PartName !== partName);

    await this.writeTypes(typesObj);
    logger.info(`Removed Override for part: ${partName}`);
  }

  /**
   * Checks if an override exists
   */
  async hasOverride(partName: string): Promise<boolean> {
    const typesObj = await this.readTypes();
    return typesObj.Types.Override.some((o) => o.$.PartName === partName);
  }

  /**
   * Helper to create a new unique part name with GUID
   */
  generateUniquePartName(prefix: string, extension = "xml"): string {
    return `/${prefix}/${this.generateGuid()}.${extension}`;
  }
}
