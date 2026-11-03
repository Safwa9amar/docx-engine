/**
 * Which part of a package is its main document, and putting it where this engine reads it.
 *
 * Every reader here, and every reader in the products above, opens `word/document.xml` by name. A
 * package need not keep it there: `_rels/.rels` says which part is the officeDocument, Word follows
 * that, and a package whose main part is `word/document2.xml` prints perfectly in Word. A reader of
 * the fixed path sees no blocks and does not throw, so a merge loses every word of that file with
 * all its checks green.
 *
 * Rather than teach every reader, the package is first put in the usual shape: the main part moves
 * to `word/document.xml`, and the parts the engine also reads by name (styles, notes, ...) to their
 * usual names, each with its relationships and its content type. The words and the markup keep
 * their bytes. A package already in that shape is not touched. One whose main part cannot be found
 * throws: there is nothing a merge could honestly do with it.
 */
import AdmZip from "adm-zip";
import { clearFalseDataDescriptors } from "@/doctor";

const CANONICAL_MAIN = "word/document.xml";
const ROOT_RELS = "_rels/.rels";
const CONTENT_TYPES = "[Content_Types].xml";
const MAIN_TYPE = /\/officeDocument$/;

/** Parts read by a fixed path, by the last word of their relationship type. */
const FIXED_PARTS: Record<string, string> = {
  styles: "word/styles.xml",
  numbering: "word/numbering.xml",
  footnotes: "word/footnotes.xml",
  endnotes: "word/endnotes.xml",
  settings: "word/settings.xml",
  comments: "word/comments.xml",
  fontTable: "word/fontTable.xml",
  webSettings: "word/webSettings.xml",
  theme: "word/theme/theme1.xml",
};

/** Only worth moving onto a free name: nothing is lost if another part already has it. */
const MOVED_IF_FREE = new Set(["fontTable", "webSettings", "theme"]);

const RELATIONSHIP = /<(?:[\w.-]+:)?Relationship\b(?:"[^"]*"|'[^']*'|[^>"'])*>/g;
const OVERRIDE = /<(?:[\w.-]+:)?Override\b(?:"[^"]*"|'[^']*'|[^>"'])*>(?:\s*<\/(?:[\w.-]+:)?Override>)?/g;
const attr = (tag: string, name: string) => new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
const attrValue = (tag: string, name: string) => {
  const m = attr(tag, name);
  return m ? (m[1] ?? m[2]) : undefined;
};
const setAttr = (tag: string, name: string, value: string) => tag.replace(new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`), `$1"${value}"`);

const decode = (s: string) => {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
};
const encode = (path: string) => path.split("/").map((s) => encodeURIComponent(s).replace(/%(?:2C|3D|40|3A|24|26|2B|3B)/g, decodeURIComponent)).join("/");

const relsOf = (part: string) => {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
};
/** "word/_rels/document.xml.rels" describes "word/document.xml"; "_rels/.rels" the package, "". */
const ownerOf = (rels: string) => rels.replace(/(^|\/)_rels\/([^/]*)\.rels$/i, "$1$2");

/** A relationship's target as a package path, read against the part that declares it. */
function resolve(owner: string, target: string): string {
  const raw = decode(target).replace(/\\/g, "/");
  const from = raw.startsWith("/") ? [] : owner.split("/").slice(0, -1);
  const out = [...from];
  for (const seg of raw.split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

function relative(fromPart: string, toPart: string): string {
  const from = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
}

const noMain = (why: string) => new Error(`This package has no main document that can be read: ${why}.`);

/** The package's entries by lower-cased name: a part name is not case-sensitive. */
function entryNames(zip: AdmZip): Map<string, string> {
  const names = new Map<string, string>();
  for (const e of zip.getEntries()) if (!e.isDirectory) names.set(e.entryName.toLowerCase(), e.entryName);
  return names;
}

const read = (zip: AdmZip, name: string | undefined) => (name ? (zip.getEntry(name)?.getData().toString("utf8") ?? "") : "");

/** The entry `_rels/.rels` calls the officeDocument. Throws when there is not exactly one. */
export function mainPartOf(zip: AdmZip): string {
  const names = entryNames(zip);
  const rootRels = names.get(ROOT_RELS);
  if (!rootRels) throw noMain("it has no _rels/.rels");
  const mains = [...read(zip, rootRels).matchAll(RELATIONSHIP)].map((m) => m[0]).filter((r) => MAIN_TYPE.test(attrValue(r, "Type") ?? ""));
  if (mains.length === 0) throw noMain("_rels/.rels names no officeDocument");
  if (mains.some((r) => attrValue(r, "TargetMode") === "External")) throw noMain("its officeDocument lies outside the package");
  const parts = new Set(mains.map((r) => resolve("", attrValue(r, "Target") ?? "").toLowerCase()));
  if (parts.size > 1) throw noMain("_rels/.rels names more than one officeDocument");
  const [part] = parts;
  const entry = names.get(part);
  if (!entry) throw noMain("the officeDocument _rels/.rels names is not in the package");
  return entry;
}

/**
 * Put this package in the usual shape, in place. Returns how many entries moved, 0 when it already was.
 * Throws, with nothing written, when the main part cannot be found or something else in the package
 * relies on the part that stands where it has to go.
 */
export function canonicalizeMainPart(zip: AdmZip): number {
  const main = mainPartOf(zip);
  const names = entryNames(zip);
  const lookup = (path: string) => names.get(path.toLowerCase());

  const moves = new Map<string, string>();
  if (main !== CANONICAL_MAIN) moves.set(main, CANONICAL_MAIN);
  for (const m of read(zip, lookup(relsOf(main))).matchAll(RELATIONSHIP)) {
    const type = /[^/]*$/.exec(attrValue(m[0], "Type") ?? "")?.[0] ?? "";
    const dest = FIXED_PARTS[type];
    const target = attrValue(m[0], "Target");
    if (!dest || target === undefined || attrValue(m[0], "TargetMode") === "External") continue;
    if (MOVED_IF_FREE.has(type) && lookup(dest)) continue;
    const entry = lookup(resolve(main, target));
    if (entry && entry !== dest && !moves.has(entry) && ![...moves.values()].includes(dest)) moves.set(entry, dest);
  }
  const mainRels = lookup(relsOf(main));
  const relsNameOff = mainRels !== undefined && mainRels !== relsOf(CANONICAL_MAIN);
  if (moves.size === 0 && !relsNameOff) return 0;

  // Every relationship in the package, to see what stands in the way and what follows a move.
  const allRels = [...names.values()].filter((n) => /(^|\/)_rels\/[^/]*\.rels$/i.test(n));
  const reached = new Set<string>();
  for (const rels of allRels) {
    for (const m of read(zip, rels).matchAll(RELATIONSHIP)) {
      const target = attrValue(m[0], "Target");
      if (target !== undefined && attrValue(m[0], "TargetMode") !== "External") reached.add(resolve(ownerOf(rels), target).toLowerCase());
    }
  }
  const stale = new Set<string>();
  for (const [from, dest] of moves) {
    const standing = lookup(dest);
    if (!standing || standing === from || moves.has(standing)) continue;
    if (reached.has(standing.toLowerCase())) throw noMain(`${dest} is a part of its own, and ${from} would have to take its place`);
    stale.add(standing);
    const standingRels = lookup(relsOf(standing));
    if (standingRels) stale.add(standingRels);
  }

  const written = new Map<string, Buffer>();
  const put = (name: string, xml: string) => written.set(name, Buffer.from(xml, "utf8"));
  const movedRels = new Map<string, string>();
  for (const [from, dest] of moves) movedRels.set(lookup(relsOf(from)) ?? relsOf(from), relsOf(dest));
  if (relsNameOff) movedRels.set(mainRels!, relsOf(CANONICAL_MAIN));

  for (const rels of allRels) {
    if (stale.has(rels)) continue;
    const owner = ownerOf(rels);
    const ownerEntry = owner === "" ? "" : (lookup(owner) ?? owner);
    const newOwner = moves.get(ownerEntry) ?? ownerEntry;
    const xml = read(zip, rels);
    const next = xml.replace(RELATIONSHIP, (tag) => {
      const target = attrValue(tag, "Target");
      if (target === undefined || attrValue(tag, "TargetMode") === "External") return tag;
      const path = resolve(owner, target);
      const entry = lookup(path);
      const final = (entry && moves.get(entry)) ?? entry ?? path;
      if (resolve(newOwner, target).toLowerCase() === final.toLowerCase()) return tag;
      return setAttr(tag, "Target", encode(target.startsWith("/") ? `/${final}` : newOwner === "" ? final : relative(newOwner, final)));
    });
    const dest = movedRels.get(rels) ?? rels;
    if (dest !== rels || next !== xml) put(dest, next);
  }

  const types = lookup(CONTENT_TYPES);
  if (types) {
    const xml = read(zip, types);
    const next = xml.replace(OVERRIDE, (tag) => {
      const part = resolve("", attrValue(tag, "PartName") ?? "");
      const entry = lookup(part);
      if (entry && stale.has(entry)) return "";
      const dest = entry && moves.get(entry);
      return dest ? setAttr(tag, "PartName", `/${dest}`) : tag;
    });
    if (next !== xml) put(types, next);
  }

  // Read everything before anything is deleted: a part may move onto a name another one is leaving.
  for (const [from, dest] of moves) written.set(dest, zip.getEntry(from)!.getData());
  const leaving = new Set([...stale, ...moves.keys(), ...movedRels.keys()]);
  for (const name of leaving) if (!written.has(name) && zip.getEntry(name)) zip.deleteFile(name);
  for (const [name, bytes] of written) zip.addFile(name, bytes);
  return moves.size + (relsNameOff ? 1 : 0);
}

/** The same, on the bytes. A package already in the usual shape comes back as the same bytes. */
export function canonicalMainPart(docx: Buffer): Buffer {
  const clean = clearFalseDataDescriptors(docx);
  const zip = new AdmZip(clean);
  return canonicalizeMainPart(zip) === 0 ? clean : clearFalseDataDescriptors(zip.toBuffer());
}
