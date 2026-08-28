import { TableObject, TableRow, TableCell, CellMargins } from "./types";
import { parseXml } from "@/utils/xmlUtils";

/** Undo Table.fromXml's sentinel encoding: every U+E000 in any text value
 *  becomes the space it stood for. Walks the whole parsed tree — the sentinel
 *  is only ever injected into `<w:t>` bodies, so a global walk is safe. */
function restoreSentinelSpaces(node: unknown): void {
  if (node === null || typeof node !== "object") return;
  const obj = node as Record<string, unknown>;
  for (const key of Object.keys(obj)) {
    const v = obj[key];
    if (typeof v === "string") {
      if (v.includes("\uE000")) obj[key] = v.replace(/\uE000/g, " ");
    } else {
      restoreSentinelSpaces(v);
    }
  }
}

export interface CellTextOptions {
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontSize?: number;      // half-points
  fontFamily?: string;
  alignment?: "left" | "center" | "right" | "justify";
}

export interface BorderSide {
  style?: string;   // "single"|"double"|"dashed"|"dotted"|"thick"|"none"
  size?: number;    // eighth-points (4 = 0.5pt, 8 = 1pt)
  color?: string;   // hex without '#'
}

export interface TableBorderOptions {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
  insideH?: BorderSide;
  insideV?: BorderSide;
}

export interface FromGridOptions {
  /** Mark row 0 as a repeating header row (`w:tblHeader`). */
  headerRow?: boolean;
  /** Bold the header row's text (default false; only relevant with `headerRow`). */
  boldHeader?: boolean;
  /** Shade the header row with this hex fill (e.g. "D9D9D9"); omit for none. */
  headerFill?: string;
  /** Right-to-left table direction (`w:bidiVisual`). */
  rtl?: boolean;
  /** Table width as a percentage of page width (default 100). */
  widthPct?: number;
  /** Border set; defaults to a light grey single-line grid. Pass to override. */
  borders?: TableBorderOptions;
}

/** Default light-grey single-line border set used by {@link Table.fromGrid}. */
const DEFAULT_GRID_BORDERS: TableBorderOptions = {
  top: { style: "single", size: 4, color: "808080" },
  bottom: { style: "single", size: 4, color: "808080" },
  left: { style: "single", size: 4, color: "808080" },
  right: { style: "single", size: 4, color: "808080" },
  insideH: { style: "single", size: 2, color: "BFBFBF" },
  insideV: { style: "single", size: 2, color: "BFBFBF" },
};

/**
 * Wraps a parsed <w:tbl> and provides a full Table Design + Layout API.
 */
export class Table {
  private table: TableObject;

  constructor(table: TableObject) {
    this.table = table;
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  private getRows(): TableRow[] {
    const raw = this.table["w:tr"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  private setRows(rows: TableRow[]): void {
    this.table["w:tr"] = rows;
  }

  private getCells(row: TableRow): TableCell[] {
    const raw = row["w:tc"];
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
  }

  private setCells(row: TableRow, cells: TableCell[]): void {
    row["w:tc"] = cells;
  }

  private extractCellText(cell: TableCell): string {
    const paragraphs = cell["w:p"];
    if (!paragraphs) return "";
    const arr = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
    return arr
      .map((p: any) => {
        const runs = p["w:r"];
        if (!runs) return "";
        const runArr = Array.isArray(runs) ? runs : [runs];
        return runArr
          .map((r: any) => {
            const t = r["w:t"];
            if (!t) return "";
            return typeof t === "string" ? t : (t._ ?? "");
          })
          .join("");
      })
      .join("\n");
  }

  private ensureCellProps(cell: TableCell): any {
    if (!cell["w:tcPr"]) cell["w:tcPr"] = {};
    return cell["w:tcPr"];
  }

  private ensureRowProps(row: TableRow): any {
    if (!row["w:trPr"]) row["w:trPr"] = {};
    return row["w:trPr"];
  }

  private ensureTableProps(): any {
    if (!this.table["w:tblPr"]) this.table["w:tblPr"] = {};
    return this.table["w:tblPr"];
  }

  private buildBorderEl(side: BorderSide): any {
    return {
      $: {
        "w:val":   side.style ?? "single",
        "w:sz":    String(side.size ?? 4),
        "w:space": "0",
        "w:color": (side.color ?? "000000").replace("#", ""),
      },
    };
  }

  private emptyCell(): TableCell {
    return { "w:p": { "w:r": { "w:t": "" } } };
  }

  // ── Read ──────────────────────────────────────────────────────────────────

  /** Number of rows. */
  public getRowCount(): number {
    return this.getRows().length;
  }

  /** Number of columns in the given row (default row 0). */
  public getColumnCount(rowIndex = 0): number {
    const rows = this.getRows();
    if (!rows[rowIndex]) return 0;
    return this.getCells(rows[rowIndex]).length;
  }

  /** Raw cell object at [row, col], or null. */
  public getCell(row: number, col: number): TableCell | null {
    const rows = this.getRows();
    if (!rows[row]) return null;
    const cells = this.getCells(rows[row]);
    return cells[col] ?? null;
  }

  /** Plain text of cell [row, col]. */
  public getCellText(row: number, col: number): string {
    const cell = this.getCell(row, col);
    if (!cell) return "";
    return this.extractCellText(cell);
  }

  /**
   * Plain text of every cell as a row-major grid (`string[][]`). Ragged rows are
   * kept as-is (each row reflects its own cell count). Cell text is the joined
   * run text; intra-cell paragraph breaks become "\n".
   */
  public getAllCellText(): string[][] {
    return this.getRows().map((row) =>
      this.getCells(row).map((cell) => this.extractCellText(cell)),
    );
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // TABLE PROPERTIES
  // ═══════════════════════════════════════════════════════════════════════════

  /** Apply a named Word table style, e.g. "TableGrid", "LightShading-Accent1". */
  public setTableStyle(styleId: string): this {
    this.ensureTableProps()["w:tblStyle"] = { $: { "w:val": styleId } };
    return this;
  }

  /**
   * Table alignment on the page.
   * Equivalent to Table Properties → Table → Alignment.
   */
  public setTableAlignment(alignment: "left" | "center" | "right"): this {
    this.ensureTableProps()["w:jc"] = { $: { "w:val": alignment } };
    return this;
  }

  /**
   * Table indent from the left margin in twips.
   * Equivalent to Table Properties → Table → Indent from Left.
   */
  public setTableIndent(twips: number): this {
    this.ensureTableProps()["w:tblInd"] = {
      $: { "w:w": String(twips), "w:type": "dxa" },
    };
    return this;
  }

  /**
   * Text wrapping around the table.
   * Equivalent to Table Properties → Table → Text Wrapping.
   */
  public setTextWrapping(type: "none" | "around"): this {
    const tblPr = this.ensureTableProps();
    if (type === "around") {
      tblPr["w:tblpPr"] = { $: {} }; // Word fills positioning attrs when placed
    } else {
      delete tblPr["w:tblpPr"];
    }
    return this;
  }

  /**
   * Right-to-left table direction.
   * Equivalent to Table Properties → Table → Table direction.
   */
  public setTableDirection(rtl: boolean): this {
    const tblPr = this.ensureTableProps();
    if (rtl) {
      tblPr["w:bidiVisual"] = {};
    } else {
      delete tblPr["w:bidiVisual"];
    }
    return this;
  }

  /**
   * Set accessible alt text (title + description).
   * Equivalent to Table Properties → Alt Text.
   */
  public setAltText(title: string, description = ""): this {
    const tblPr = this.ensureTableProps();
    tblPr["w:tblCaption"]     = { $: { "w:val": title } };
    tblPr["w:tblDescription"] = { $: { "w:val": description } };
    return this;
  }

  /**
   * Set table width.
   * type "pct"  → percentage of page width (pass 0–100)
   * type "dxa"  → fixed twips
   * type "auto" → auto
   */
  public setTableWidth(value: number, type: "pct" | "dxa" | "auto" = "pct"): this {
    this.ensureTableProps()["w:tblW"] = {
      $: { "w:w": String(type === "pct" ? value * 50 : value), "w:type": type },
    };
    return this;
  }

  /**
   * Layout mode.
   * "autofit" → AutoFit Contents / AutoFit Window
   * "fixed"   → Fixed Column Width
   */
  public setLayoutMode(mode: "autofit" | "fixed"): this {
    this.ensureTableProps()["w:tblLayout"] = { $: { "w:type": mode } };
    return this;
  }

  /** Shorthand: AutoFit to contents. */
  public autoFitContents(): this {
    return this.setLayoutMode("autofit").setTableWidth(0, "auto");
  }

  /** Shorthand: AutoFit to window (page width). */
  public autoFitWindow(): this {
    return this.setLayoutMode("autofit").setTableWidth(100, "pct");
  }

  /** Shorthand: Fixed column widths. */
  public fixedColumnWidth(): this {
    return this.setLayoutMode("fixed");
  }

  /**
   * Default cell margins for the entire table (twips).
   * Equivalent to Table Properties → Options → Default cell margins.
   */
  public setDefaultCellMargins(margins: CellMargins): this {
    const tblPr = this.ensureTableProps();
    // CT_TblCellMar is an xsd:SEQUENCE — top, left, bottom, right. Key insertion
    // order here is the element order the XML builder emits, so writing them in
    // the "natural" top/bottom/left/right order produces a file Word rejects.
    tblPr["w:tblCellMar"] = {
      ...(margins.top    !== undefined ? { "w:top":    { $: { "w:w": String(margins.top),    "w:type": "dxa" } } } : {}),
      ...(margins.left   !== undefined ? { "w:left":   { $: { "w:w": String(margins.left),   "w:type": "dxa" } } } : {}),
      ...(margins.bottom !== undefined ? { "w:bottom": { $: { "w:w": String(margins.bottom), "w:type": "dxa" } } } : {}),
      ...(margins.right  !== undefined ? { "w:right":  { $: { "w:w": String(margins.right),  "w:type": "dxa" } } } : {}),
    };
    return this;
  }

  /**
   * Table-level borders.
   */
  public setTableBorders(borders: TableBorderOptions): this {
    // CT_TblBorders is an xsd:SEQUENCE: top, left, bottom, right, insideH,
    // insideV. These keys are emitted in INSERTION order, so assigning them
    // top/bottom/left/right — the order a human thinks in — writes
    // `<w:bottom>` before `<w:left>` and Word refuses to open the document
    // ("Word experienced an error trying to open the file"). Well-formed but
    // schema-invalid; nothing warns. Do not reorder these lines.
    const tblBorders: any = {};
    if (borders.top)     tblBorders["w:top"]      = this.buildBorderEl(borders.top);
    if (borders.left)    tblBorders["w:left"]     = this.buildBorderEl(borders.left);
    if (borders.bottom)  tblBorders["w:bottom"]   = this.buildBorderEl(borders.bottom);
    if (borders.right)   tblBorders["w:right"]    = this.buildBorderEl(borders.right);
    if (borders.insideH) tblBorders["w:insideH"]  = this.buildBorderEl(borders.insideH);
    if (borders.insideV) tblBorders["w:insideV"]  = this.buildBorderEl(borders.insideV);
    this.ensureTableProps()["w:tblBorders"] = tblBorders;
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // ROW PROPERTIES
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Set row height.
   * rule: "atLeast" (At least) | "exact" (Exact)
   * Equivalent to Table Properties → Row → Specify height.
   */
  public setRowHeight(rowIndex: number, heightTwips: number, rule: "atLeast" | "exact" = "atLeast"): this {
    const rows = this.getRows();
    if (!rows[rowIndex]) throw new Error(`Row ${rowIndex} not found`);
    this.ensureRowProps(rows[rowIndex])["w:trHeight"] = {
      $: { "w:val": String(heightTwips), "w:hRule": rule },
    };
    this.setRows(rows);
    return this;
  }

  /**
   * Allow/prevent a row from breaking across pages.
   * Equivalent to Table Properties → Row → Allow row to break across pages.
   */
  public setRowAllowBreak(rowIndex: number, allow: boolean): this {
    const rows = this.getRows();
    if (!rows[rowIndex]) throw new Error(`Row ${rowIndex} not found`);
    const trPr = this.ensureRowProps(rows[rowIndex]);
    if (allow) {
      delete trPr["w:cantSplit"];
    } else {
      trPr["w:cantSplit"] = {};
    }
    this.setRows(rows);
    return this;
  }

  /**
   * Mark a row as a repeating header at the top of each page.
   * Equivalent to Table Properties → Row → Repeat as header row.
   * Optionally set a background fill colour.
   */
  public setHeaderRow(rowIndex: number, fillColor?: string): this {
    const rows = this.getRows();
    if (!rows[rowIndex]) throw new Error(`Row ${rowIndex} not found`);
    this.ensureRowProps(rows[rowIndex])["w:tblHeader"] = {};
    if (fillColor) this.setRowShading(rowIndex, fillColor);
    this.setRows(rows);
    return this;
  }

  /** Shade all cells in a row with one background colour. */
  public setRowShading(rowIndex: number, fillColor: string): this {
    const count = this.getColumnCount(rowIndex);
    for (let col = 0; col < count; col++) {
      this.setCellShading(rowIndex, col, fillColor);
    }
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN PROPERTIES
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Set preferred width (twips) for every cell in a column.
   * Equivalent to Table Properties → Column → Preferred width.
   */
  public setColumnWidth(colIndex: number, widthTwips: number): this {
    const rows = this.getRows();
    rows.forEach((row) => {
      const cells = this.getCells(row);
      if (cells[colIndex]) {
        const tcPr = this.ensureCellProps(cells[colIndex]);
        tcPr["w:tcW"] = { $: { "w:w": String(widthTwips), "w:type": "dxa" } };
      }
    });
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CELL PROPERTIES
  // ═══════════════════════════════════════════════════════════════════════════

  /** Set cell preferred width in twips. */
  public setCellWidth(row: number, col: number, widthTwips: number): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    this.ensureCellProps(cell)["w:tcW"] = {
      $: { "w:w": String(widthTwips), "w:type": "dxa" },
    };
    return this;
  }

  /**
   * Set cell vertical alignment.
   * Equivalent to Table Properties → Cell → Vertical Alignment.
   */
  public setCellVerticalAlignment(
    row: number,
    col: number,
    alignment: "top" | "center" | "bottom",
  ): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    this.ensureCellProps(cell)["w:vAlign"] = { $: { "w:val": alignment } };
    return this;
  }

  /**
   * Set individual cell margins (overrides table default).
   * Equivalent to Table Properties → Cell → Options.
   */
  public setCellMargins(row: number, col: number, margins: CellMargins): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    const tcPr = this.ensureCellProps(cell);
    // CT_TcMar is an xsd:SEQUENCE — top, left, bottom, right (see the note on
    // setTableBorders; key order here IS element order).
    tcPr["w:tcMar"] = {
      ...(margins.top    !== undefined ? { "w:top":    { $: { "w:w": String(margins.top),    "w:type": "dxa" } } } : {}),
      ...(margins.left   !== undefined ? { "w:left":   { $: { "w:w": String(margins.left),   "w:type": "dxa" } } } : {}),
      ...(margins.bottom !== undefined ? { "w:bottom": { $: { "w:w": String(margins.bottom), "w:type": "dxa" } } } : {}),
      ...(margins.right  !== undefined ? { "w:right":  { $: { "w:w": String(margins.right),  "w:type": "dxa" } } } : {}),
    };
    return this;
  }

  /** Set cell background shading colour. */
  public setCellShading(row: number, col: number, fillColor: string): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    this.ensureCellProps(cell)["w:shd"] = {
      $: { "w:val": "clear", "w:color": "auto", "w:fill": fillColor.replace("#", "") },
    };
    return this;
  }

  /**
   * Format a cell's EXISTING text (every run in every paragraph), preserving
   * the content — unlike setCellContent, which replaces it. Pass only the
   * properties you want to change. sizeHalfPoints: Word half-points (24 = 12pt).
   */
  public setCellTextFormat(
    row: number,
    col: number,
    opts: { bold?: boolean; italic?: boolean; sizeHalfPoints?: number; fontFamily?: string },
  ): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    const paragraphs = cell["w:p"];
    if (!paragraphs) return this;
    const arr = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
    for (const p of arr as any[]) {
      const runs = p?.["w:r"];
      if (!runs) continue;
      const runArr = Array.isArray(runs) ? runs : [runs];
      for (const r of runArr) {
        if (!r || typeof r !== "object") continue;
        const rPr = (r["w:rPr"] ??= {});
        if (opts.bold !== undefined) {
          if (opts.bold) rPr["w:b"] = {};
          else delete rPr["w:b"];
        }
        if (opts.italic !== undefined) {
          if (opts.italic) rPr["w:i"] = {};
          else delete rPr["w:i"];
        }
        if (opts.sizeHalfPoints !== undefined) {
          rPr["w:sz"] = { $: { "w:val": String(opts.sizeHalfPoints) } };
          rPr["w:szCs"] = { $: { "w:val": String(opts.sizeHalfPoints) } };
        }
        if (opts.fontFamily !== undefined) {
          rPr["w:rFonts"] = {
            $: { "w:ascii": opts.fontFamily, "w:hAnsi": opts.fontFamily, "w:cs": opts.fontFamily },
          };
        }
      }
    }
    return this;
  }

  /** Grid column of the tc at [row, col] (cumulative gridSpans of preceding cells). */
  private gridColOf(row: number, col: number): number {
    const cells = this.getCells(this.getRows()[row] ?? { "w:tc": [] });
    let gc = 0;
    for (let c = 0; c < col && c < cells.length; c++) {
      const cs = Number((this.ensureCellProps(cells[c])["w:gridSpan"] as any)?.$?.["w:val"] ?? "1") || 1;
      gc += cs;
    }
    return gc;
  }

  /**
   * SPLIT a horizontally merged cell back into single cells: removes its
   * gridSpan and inserts the missing empty cells after it. Reverse of
   * mergeCellsHorizontal (the merged text stays in the first cell).
   */
  public splitCellHorizontal(row: number, col: number): this {
    const rows = this.getRows();
    if (!rows[row]) throw new Error(`Row ${row} not found`);
    const cells = this.getCells(rows[row]);
    const cell = cells[col];
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    const tcPr = this.ensureCellProps(cell);
    const cs = Number((tcPr["w:gridSpan"] as any)?.$?.["w:val"] ?? "1") || 1;
    if (cs <= 1) return this; // nothing to split
    delete tcPr["w:gridSpan"];
    const empties = Array.from({ length: cs - 1 }, () => this.emptyCell());
    cells.splice(col + 1, 0, ...empties);
    this.setCells(rows[row], cells);
    return this;
  }

  /**
   * SPLIT a vertically merged cell: removes vMerge from the restart cell at
   * [row, col] and from its continuation cells below (matched by GRID column).
   * Reverse of mergeCellsVertical (the merged text stays in the top cell).
   */
  public splitCellVertical(row: number, col: number): this {
    const rows = this.getRows();
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    delete this.ensureCellProps(cell)["w:vMerge"];
    const gc = this.gridColOf(row, col);
    for (let r = row + 1; r < rows.length; r++) {
      const cells = this.getCells(rows[r]);
      // Find the tc at grid column gc in this row.
      let cgc = 0;
      let target: TableCell | null = null;
      for (const c of cells) {
        if (cgc === gc) { target = c; break; }
        const cs = Number((this.ensureCellProps(c)["w:gridSpan"] as any)?.$?.["w:val"] ?? "1") || 1;
        cgc += cs;
        if (cgc > gc) break;
      }
      if (!target) break;
      const tcPr = this.ensureCellProps(target);
      const vm = tcPr["w:vMerge"] as any;
      const isCont = vm && !(vm.$ && vm.$["w:val"] === "restart");
      if (!isCont) break; // chain ended
      delete tcPr["w:vMerge"];
    }
    return this;
  }

  /** Move row `from` to position `to` (both 0-based). */
  public moveRow(from: number, to: number): this {
    const rows = this.getRows();
    if (!rows[from]) throw new Error(`Row ${from} not found`);
    const clamped = Math.max(0, Math.min(to, rows.length - 1));
    const [row] = rows.splice(from, 1);
    rows.splice(clamped, 0, row);
    this.setRows(rows);
    return this;
  }

  /** Move column `from` to position `to` (both 0-based, tc-indexed). */
  public moveColumn(from: number, to: number): this {
    const rows = this.getRows();
    rows.forEach((row) => {
      const cells = this.getCells(row);
      if (!cells[from]) return; // ragged/merged rows: skip rows without that tc
      const clamped = Math.max(0, Math.min(to, cells.length - 1));
      const [cell] = cells.splice(from, 1);
      cells.splice(clamped, 0, cell);
      this.setCells(row, cells);
    });
    return this;
  }

  /**
   * Set the font colour of a cell's EXISTING text (every run in every
   * paragraph), preserving the content — unlike setCellContent, which replaces
   * it. 6-hex colour, with or without '#'.
   */
  public setCellTextColor(row: number, col: number, color: string): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    const val = color.replace("#", "");
    const paragraphs = cell["w:p"];
    if (!paragraphs) return this;
    const arr = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
    for (const p of arr as any[]) {
      const runs = p?.["w:r"];
      if (!runs) continue;
      const runArr = Array.isArray(runs) ? runs : [runs];
      for (const r of runArr) {
        if (!r || typeof r !== "object") continue;
        const rPr = (r["w:rPr"] ??= {});
        rPr["w:color"] = { $: { "w:val": val } };
      }
    }
    return this;
  }

  /** Set text with rich formatting in a cell. */
  public setCellContent(
    row: number,
    col: number,
    text: string,
    opts: CellTextOptions = {},
  ): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);

    const rPr: any = {};
    if (opts.bold)       rPr["w:b"] = {};
    if (opts.italic)     rPr["w:i"] = {};
    if (opts.color)      rPr["w:color"] = { $: { "w:val": opts.color.replace("#", "") } };
    if (opts.fontSize) {
      rPr["w:sz"]   = { $: { "w:val": String(opts.fontSize) } };
      rPr["w:szCs"] = { $: { "w:val": String(opts.fontSize) } };
    }
    if (opts.fontFamily) {
      rPr["w:rFonts"] = {
        $: { "w:ascii": opts.fontFamily, "w:hAnsi": opts.fontFamily, "w:cs": opts.fontFamily },
      };
    }

    const pPr: any = opts.alignment
      ? { "w:jc": { $: { "w:val": opts.alignment } } }
      : undefined;

    cell["w:p"] = {
      ...(pPr ? { "w:pPr": pPr } : {}),
      "w:r": {
        ...(Object.keys(rPr).length ? { "w:rPr": rPr } : {}),
        "w:t": { _: text, $: { "xml:space": "preserve" } },
      },
    };
    return this;
  }

  /** Set plain text in a cell. */
  public setCellText(row: number, col: number, text: string): this {
    const cell = this.getCell(row, col);
    if (!cell) throw new Error(`Cell [${row},${col}] not found`);
    cell["w:p"] = { "w:r": { "w:t": text } };
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // ROW / COLUMN INSERT & DELETE  (Layout tab)
  // ═══════════════════════════════════════════════════════════════════════════

  private makeEmptyRow(colCount: number): TableRow {
    return { "w:tc": Array.from({ length: colCount }, () => this.emptyCell()) };
  }

  /**
   * Insert a blank row above the given index.
   * Equivalent to Layout → Insert Above.
   */
  public insertRowAbove(rowIndex: number): this {
    const rows = this.getRows();
    const colCount = this.getColumnCount();
    rows.splice(rowIndex, 0, this.makeEmptyRow(colCount));
    this.setRows(rows);
    return this;
  }

  /**
   * Insert a blank row below the given index.
   * Equivalent to Layout → Insert Below.
   */
  public insertRowBelow(rowIndex: number): this {
    const rows = this.getRows();
    const colCount = this.getColumnCount();
    rows.splice(rowIndex + 1, 0, this.makeEmptyRow(colCount));
    this.setRows(rows);
    return this;
  }

  /**
   * Append a row, optionally pre-filled with text.
   */
  public addRow(cellTexts?: string[]): this {
    const colCount = this.getColumnCount();
    const texts = cellTexts ?? Array(colCount).fill("");
    const newRow: TableRow = {
      "w:tc": texts.map((text) => ({ "w:p": { "w:r": { "w:t": text } } })),
    };
    const rows = this.getRows();
    rows.push(newRow);
    this.setRows(rows);
    return this;
  }

  /**
   * Remove the row at the given index.
   * Equivalent to Layout → Delete → Delete Rows.
   */
  public removeRow(index: number): this {
    const rows = this.getRows();
    if (index < 0 || index >= rows.length) throw new Error(`Row index ${index} out of bounds`);
    rows.splice(index, 1);
    this.setRows(rows);
    return this;
  }

  /**
   * Insert a blank column to the left of colIndex.
   * Equivalent to Layout → Insert Left.
   */
  public insertColumnLeft(colIndex: number): this {
    const rows = this.getRows();
    rows.forEach((row) => {
      const cells = this.getCells(row);
      cells.splice(colIndex, 0, this.emptyCell());
      this.setCells(row, cells);
    });
    return this;
  }

  /**
   * Insert a blank column to the right of colIndex.
   * Equivalent to Layout → Insert Right.
   */
  public insertColumnRight(colIndex: number): this {
    const rows = this.getRows();
    rows.forEach((row) => {
      const cells = this.getCells(row);
      cells.splice(colIndex + 1, 0, this.emptyCell());
      this.setCells(row, cells);
    });
    return this;
  }

  /**
   * Delete an entire column.
   * Equivalent to Layout → Delete → Delete Columns.
   */
  public deleteColumn(colIndex: number): this {
    const rows = this.getRows();
    rows.forEach((row) => {
      const cells = this.getCells(row);
      cells.splice(colIndex, 1);
      this.setCells(row, cells);
    });
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // MERGE / SPLIT  (Layout tab)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Horizontally merge cells in a single row from startCol to endCol (inclusive).
   * Equivalent to Layout → Merge Cells.
   */
  public mergeCellsHorizontal(rowIndex: number, startCol: number, endCol: number): this {
    if (startCol >= endCol) throw new Error("startCol must be less than endCol");
    const rows = this.getRows();
    const cells = this.getCells(rows[rowIndex]);

    // Collect text from all merged cells
    const mergedText = cells
      .slice(startCol, endCol + 1)
      .map((c) => this.extractCellText(c))
      .filter(Boolean)
      .join(" ");

    // Set gridSpan on the first cell
    const firstCell = cells[startCol];
    const tcPr = this.ensureCellProps(firstCell);
    tcPr["w:gridSpan"] = { $: { "w:val": String(endCol - startCol + 1) } };
    firstCell["w:p"] = { "w:r": { "w:t": mergedText } };

    // Remove the consumed cells
    cells.splice(startCol + 1, endCol - startCol);
    this.setCells(rows[rowIndex], cells);
    this.setRows(rows);
    return this;
  }

  /**
   * Vertically merge cells in a column from startRow to endRow (inclusive).
   * Equivalent to Layout → Merge Cells (vertical selection).
   */
  public mergeCellsVertical(colIndex: number, startRow: number, endRow: number): this {
    if (startRow >= endRow) throw new Error("startRow must be less than endRow");
    const rows = this.getRows();

    // First cell: vMerge with val="restart"
    const firstCell = this.getCells(rows[startRow])[colIndex];
    if (!firstCell) throw new Error(`Cell [${startRow},${colIndex}] not found`);
    this.ensureCellProps(firstCell)["w:vMerge"] = { $: { "w:val": "restart" } };

    // Subsequent cells: vMerge (no val = continuation)
    for (let r = startRow + 1; r <= endRow; r++) {
      const cell = this.getCells(rows[r])[colIndex];
      if (!cell) throw new Error(`Cell [${r},${colIndex}] not found`);
      this.ensureCellProps(cell)["w:vMerge"] = {};
      cell["w:p"] = { "w:r": { "w:t": "" } }; // empty content in continuation rows
    }

    this.setRows(rows);
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DISTRIBUTE  (Layout tab)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Set all rows to the same height (based on tallest row or a fixed value).
   * Equivalent to Layout → Distribute Rows.
   */
  public distributeRows(heightTwips?: number): this {
    const rows = this.getRows();
    const h = heightTwips ?? 400;
    rows.forEach((_, i) => this.setRowHeight(i, h, "atLeast"));
    return this;
  }

  /**
   * Set all columns to equal width.
   * Equivalent to Layout → Distribute Columns.
   */
  public distributeColumns(): this {
    const colCount = this.getColumnCount();
    if (colCount === 0) return this;
    const tblPr = this.ensureTableProps();
    const totalW = parseInt(tblPr["w:tblW"]?.["$"]?.["w:w"] ?? "9360", 10);
    const perCol = Math.floor(totalW / colCount);
    for (let c = 0; c < colCount; c++) this.setColumnWidth(c, perCol);
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SORT  (Layout tab)
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Sort data rows (skipping the header row) by the value in colIndex.
   * Equivalent to Layout → Sort.
   */
  public sortByColumn(
    colIndex: number,
    direction: "asc" | "desc" = "asc",
    hasHeaderRow = true,
  ): this {
    const rows = this.getRows();
    const headerRows = hasHeaderRow ? rows.splice(0, 1) : [];

    rows.sort((a, b) => {
      const aText = this.extractCellText(this.getCells(a)[colIndex] ?? {});
      const bText = this.extractCellText(this.getCells(b)[colIndex] ?? {});
      const aNum = parseFloat(aText);
      const bNum = parseFloat(bText);
      const numeric = !isNaN(aNum) && !isNaN(bNum);
      const cmp = numeric ? aNum - bNum : aText.localeCompare(bText);
      return direction === "asc" ? cmp : -cmp;
    });

    this.setRows([...headerRows, ...rows]);
    return this;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // OUTPUT
  // ═══════════════════════════════════════════════════════════════════════════

  /** Raw table object for serialisation. */
  public toObject(): TableObject {
    return this.table;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // FACTORY
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Build a styled `Table` from a grid of cell texts. The widest row determines
   * the column count; short rows are padded with empty cells. By default the
   * table gets a light-grey single-line border grid and 100% width.
   *
   * Replaces the hand-rolled `buildTable(header, rows, rtl)` helpers used by
   * document builders, and backs the string-level `makeTableXml` in OrderedBody.
   */
  /**
   * Parse a `<w:tbl>…</w:tbl>` XML string into a `Table`. Useful for reading a
   * table out of an OrderedBody block (`block.xml`) without the xml2js
   * full-document path — e.g. `(await Table.fromXml(block.xml)).getAllCellText()`.
   */
  public static async fromXml(tableXml: string): Promise<Table> {
    // trim: false — NOT the parser default. A table's cell text lives in
    // `<w:t>` runs, and Word splits a sentence across many runs with the
    // spaces riding at run edges or in space-only `xml:space="preserve"` runs.
    // The trimming parse deleted those spaces; because mutateTable round-trips
    // fromXml → buildXml → save, every table EDIT then wrote the glued text
    // back into the document — permanent data loss, one table op at a time.
    // Untrimmed parsing is safe for this XML: real Word document.xml is
    // single-line, and mutateTable rebuilds with pretty:false, so
    // whitespace-only inter-element text nodes do not occur here.
    // ⚠️ Space-only runs need MORE than trim:false. xml2js drops whitespace-
    // only text from any node that has attributes — and Word's space between
    // two styled words is exactly `<w:t xml:space="preserve"> </w:t>`. The
    // space is unrecoverable after the parse, so it is protected before it:
    // each space in a space-only `<w:t>` body becomes a private-use sentinel
    // (non-whitespace, so xml2js keeps it), and the parsed object is walked
    // once to restore real spaces. The in-memory table is then correct for
    // every reader AND for mutateTable's buildXml→save round-trip — no decode
    // at any exit point.
    const protectedXml = tableXml.replace(
      /(<w:t(?:\s[^>]*)?>)( +)(<\/w:t>)/g,
      (_m, open: string, spaces: string, close: string) => open + "\uE000".repeat(spaces.length) + close,
    );
    const parsed = await parseXml(protectedXml, { trim: false });
    restoreSentinelSpaces(parsed);
    // parseXml wraps the element under its root key ("w:tbl"); unwrap to the
    // inner object the Table class expects ({ "w:tblPr", "w:tblGrid", "w:tr" }).
    const inner =
      parsed && typeof parsed === "object" && "w:tbl" in parsed
        ? (parsed as Record<string, unknown>)["w:tbl"]
        : parsed;
    return new Table(inner as TableObject);
  }

  public static fromGrid(rows: string[][], opts: FromGridOptions = {}): Table {
    const headerRow = opts.headerRow ?? false;
    const boldHeader = opts.boldHeader ?? false;
    const cols = Math.max(1, ...rows.map((r) => r.length));

    const mkCell = (text: string, bold: boolean): TableCell => ({
      "w:p": {
        "w:r": {
          ...(bold ? { "w:rPr": { "w:b": {} } } : {}),
          "w:t": { _: text ?? "", $: { "xml:space": "preserve" } },
        },
      },
    });

    const mkRow = (cells: string[], bold: boolean): TableRow => ({
      "w:tc": Array.from({ length: cols }, (_, c) => mkCell(cells[c] ?? "", bold)),
    });

    const tableObj: TableObject = {
      "w:tblPr": {},
      "w:tblGrid": { "w:gridCol": Array.from({ length: cols }, () => ({})) },
      "w:tr": rows.map((r, i) => mkRow(r, boldHeader && headerRow && i === 0)),
    } as unknown as TableObject;

    const t = new Table(tableObj);
    t.setTableBorders(opts.borders ?? DEFAULT_GRID_BORDERS);
    t.setTableWidth(opts.widthPct ?? 100, "pct");
    if (opts.rtl) t.setTableDirection(true);
    if (headerRow && rows.length > 0) t.setHeaderRow(0, opts.headerFill);
    return t;
  }
}
