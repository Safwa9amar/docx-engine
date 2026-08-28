# docx-engine

**Edit existing Word (.docx) files without mangling them.**

[![npm](https://img.shields.io/npm/v/docx-engine)](https://www.npmjs.com/package/docx-engine)
[![license: ISC](https://img.shields.io/badge/license-ISC-green.svg)](./LICENSE)

`docx-engine` is a TypeScript library for reading and editing `.docx` files that
already exist — theses, contracts, reports — by operating directly on the Office
Open XML inside the package. Everything you don't touch round-trips byte-for-byte.
What you do touch is written as the real Word structure (a `TOC` field, a `SEQ`
caption, a footnote part, a section's own header), in schema order, and every
write passes through a validating *docx doctor* so the file still opens in Word.

It is the engine behind [Kwill](https://kwill.app), the Algerian thesis editor,
where it has formatted Arabic, French and English theses in production since 2025.
Formerly published as `mdocxengine` (1.0.0); this is the same engine, renamed.

- **No Word, no LibreOffice, no Python.** Three runtime dependencies (`adm-zip`, `xml2js`, `xmldom`). Node 18+.
- **Arabic / RTL is first-class.** `w:bidi`, `w:rtl`, complex-script fonts and sizes, RTL tables and TOCs.
- **Real structures, not text that looks like them.** Table of contents, captions and cross-references, footnotes and endnotes, citations and bibliography, numbering definitions, per-section headers and footers, section breaks and front-matter numbering (i, ii, iii → 1, 2, 3).
- **Safe by default.** A doctor validates and repairs every write; ordered properties (`w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`) are emitted in the sequence the schema demands.

## Install

```bash
npm i docx-engine
```

## Quick start

```ts
import { Doc } from "docx-engine";

const doc = await Doc.open("memoire.docx");          // path or Buffer

await doc.addHeading("الفصل الأول", 1);
await doc.engine.toc.insertTOC({ title: "الفهرس", rtl: true });
await doc.engine.captions.insertFigureCaption(14, "Results");
await doc.frontMatterNumbering(14);                  // roman front matter, arabic body
await doc.save("memoire.docx");                      // or doc.toBuffer()
```

## What it can do

`Doc` is the high-level facade; `doc.engine` exposes the part managers underneath.

| Area | Facade (`Doc`) | Manager (`doc.engine.*`) |
|---|---|---|
| Read | `text()`, `wordCount()`, `blocks()`, `outline()`, `inferOutline()`, `tables()`, `images()`, `sections()`, `describe()`, `toMarkdownMap()` | — |
| Paragraphs & headings | `addParagraph()`, `addHeading()`, `editParagraph()`, `setHeadingLevel()`, `deleteBlock()`, `replaceText()`, `addMarkdown()` | `FormattingManager` |
| Text style | `setTextStyle()`, `getTextStyle()` — fonts, sizes, bold/italic, colour, Arabic complex-script variants | `StylesManager` |
| Tables | `addTable()`, `editTableCell()`, add/remove rows and columns, `mergeTableCells()`, `splitTableCells()`, `sortTable()`, `shadeTable()`, `setTableWidths()`, `setTableLayout()` | `Table` |
| Images | `addImage()` — inline or anchored, with positioning | `MediaManager` |
| Page & sections | `setPageSize()`, `setMargins()`, `setOrientation()`, `addPageNumbers()`, `frontMatterNumbering()`, `startOnNewPage()`, `setSectionPageBorders()`, `applySectionChrome()` | `PageLayoutManager`, `SectionManager` |
| Headers & footers | `setHeader()`, `setFooter()`, `setSectionHeader()`, `setSectionFooter()` | `HeaderManager`, `FooterManager` |
| Table of contents | — | `TableOfContentsManager` — `insertTOC()`, `removeTOC()`, `findTypedTOC()` |
| Captions & cross-refs | — | `CaptionManager` — `insertFigureCaption()`, `insertTableCaption()`, `convertTextCaptions()`, `renumber()`; `CrossReferenceManager` |
| Notes & citations | — | `FootnoteManager`, `EndnoteManager`, `CitationManager` — `addSource()`, `insertBibliography()` |
| Lists | — | `NumberingManager` — real numbering definitions |
| Shapes | — | `ShapeManager` — text boxes and cover-page shapes |
| Comments & revisions | — | `CommentsManager`, `TrackedChangesManager` |
| Merge | — | `MergeManager` — append one document to another, styles reconciled |
| Metadata | — | `MetadataManager` — core and app properties |

All managers are also exported directly for lower-level use, alongside the
`Paragraph`, `Run`, `Table` builders and the OOXML helpers
(`parseOrderedDoc`, `buildOrderedDoc`, `makeParagraphNode`, …).

## What it does not do

- **Render or paginate.** No page count, no layout. Use Word, LibreOffice or OnlyOffice for that.
- **Author charts or equations.** Existing `c:chart` and `m:oMath` content is preserved, kept in schema order and validated, but there is no authoring API for them here.
- **Generate documents from scratch.** It can, but that is not what it is optimised for — it is built to edit files that already exist.
- **Deep tracked changes.** `TrackedChangesManager` reads and emits basic revisions; there is no document compare.

## How it stays safe

- The zip is read once; untouched parts are written back unchanged.
- Body order is preserved as an ordered block list — paragraphs, tables, drawings and section breaks never get regrouped by tag.
- Every property container is emitted in ECMA-376 sequence (`src/core/ooxml/canonicalOrder.ts`).
- The doctor (`src/doctor/`) validates relationships, content types and structure on every write, and repairs what it can.
- 860+ tests, including full-document round-trips of real theses in Arabic and French, and a Word oracle.

## Licence

ISC © Hamza Safwan
