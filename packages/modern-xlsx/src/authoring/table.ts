/**
 * Table Layout Engine — generate styled XLSX tables from declarative options.
 *
 * Composes existing primitives (Worksheet.cell, StyleBuilder, addMergeCell,
 * setColumnWidth, frozenPane) into a high-level API that eliminates manual
 * cell coordinate math.
 *
 * `drawTable` writes styled cell ranges; it does not create an Excel table
 * (ListObject). Use `Worksheet.addTable` for that.
 */

import { columnToLetter, decodeCellRef, encodeRange } from '../format/cell-ref.js';
import { dateToSerial } from '../format/dates.js';
import type {
  AlignmentData,
  BorderStyle,
  FontData,
  PatternType,
  StylesData,
} from '../model/types.js';
import type { Workbook } from '../model/workbook.js';
import type { Worksheet } from '../model/worksheet.js';
import { StyleBuilder } from './style-builder.js';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** Result returned by drawTable with layout metadata. */
export interface TableResult {
  /** A1-style range covering the entire table (e.g. "A1:D10"), or `''` when nothing was drawn (no headers). */
  range: string;
  /** Total number of rows (header + data); 0 when nothing was drawn. */
  rowCount: number;
  /** Number of columns. */
  colCount: number;
  /** 0-based row index of the first data row in the worksheet. */
  firstDataRow: number;
  /** 0-based row index of the last data row in the worksheet. */
  lastDataRow: number;
}

/**
 * A value `drawTable` accepts in a data row. A `Date` is written as the serial
 * of its UTC instant in the workbook's date system (an invalid `Date` leaves the
 * cell empty); a `bigint` is written as a number when it is a safe integer and
 * as a string otherwise.
 */
export type TableCellValue = string | number | boolean | Date | bigint | null;

/** Every {@link PatternType}, as a record so a missing or extra member fails to compile. */
const PATTERN_TYPES = {
  none: true,
  solid: true,
  gray125: true,
  darkGray: true,
  mediumGray: true,
  lightGray: true,
  darkHorizontal: true,
  darkVertical: true,
  darkDown: true,
  darkUp: true,
  darkGrid: true,
  darkTrellis: true,
  lightHorizontal: true,
  lightVertical: true,
  lightDown: true,
  lightUp: true,
  lightGrid: true,
  lightTrellis: true,
} as const satisfies Record<PatternType, true>;

const isPatternType = (v: string): v is PatternType => Object.hasOwn(PATTERN_TYPES, v);

/** Per-cell style override. */
export interface CellStyle {
  font?: Partial<FontData>;
  fill?: { pattern?: PatternType | (string & {}); fgColor?: string };
  border?: Partial<{
    left: { style: BorderStyle; color?: string | null };
    right: { style: BorderStyle; color?: string | null };
    top: { style: BorderStyle; color?: string | null };
    bottom: { style: BorderStyle; color?: string | null };
  }>;
  alignment?: Partial<AlignmentData>;
  numberFormat?: string;
}

/** Column definition for drawTable. */
export interface TableColumn {
  /** Header label override. */
  header?: string;
  /** Fixed width in Excel character units. */
  width?: number;
  /** Horizontal alignment for data cells in this column. */
  align?: 'left' | 'center' | 'right';
  /** Number format code for data cells (e.g. '#,##0.00'). */
  numberFormat?: string;
}

/** Options for drawTable. */
export interface DrawTableOptions {
  /** Header labels. */
  headers: string[];
  /** Data rows — each row is an array of cell values. */
  rows: TableCellValue[][];
  /** Column definitions for header, width, alignment, and number format. */
  columns?: TableColumn[];
  /**
   * Fixed column widths array. A width above 0 wins over `columns[c].width`
   * and over `autoWidth` for that column; 0 or less counts as unset.
   */
  columnWidths?: number[];
  /** Origin cell reference (default: "A1"). */
  origin?: string;

  // --- Styling ---
  /** Header row font. Default: bold white. */
  headerFont?: Partial<FontData>;
  /** Header row background color (hex, no #). Default: '4472C4'. */
  headerColor?: string;
  /** Body font override. */
  bodyFont?: Partial<FontData>;
  /** Border style for all cells. Default: 'thin'. Set to null for no borders. */
  borderStyle?: BorderStyle | null;
  /** Border color (hex). Default: '000000'. */
  borderColor?: string;

  // --- Zebra striping ---
  /** Alternating row background color (hex). Null/undefined = no striping. */
  alternateRowColor?: string | null;

  // --- Alignment ---
  /** Default horizontal alignment for header cells. Default: 'center'. */
  headerAlign?: 'left' | 'center' | 'right';
  /** Default horizontal alignment for body cells. */
  bodyAlign?: 'left' | 'center' | 'right';
  /** Default vertical alignment for all cells. */
  verticalAlign?: 'top' | 'center' | 'bottom';

  // --- Content ---
  /** Enable text wrapping in body cells. Default: false. */
  wrapText?: boolean;
  /** Auto-calculate column widths from content. Default: false. */
  autoWidth?: boolean;
  /** Freeze the rows down to and including the header row. Default: false. */
  freezeHeader?: boolean;
  /** Add auto-filter dropdowns to the headers, spanning the header row through the last data row. Default: false. */
  autoFilter?: boolean;

  // --- Merge cells ---
  /** Cells to merge. Row/col are 0-based relative to data area (row 0 = first data row). */
  merges?: { row: number; col: number; rowSpan?: number; colSpan?: number }[];

  // --- Per-cell styling ---
  /** Per-cell style overrides. Key is "row,col" (0-based, data area). */
  cellStyles?: Record<string, CellStyle>;
}

// ---------------------------------------------------------------------------
// Style palette — pre-built style indices for each table region
// ---------------------------------------------------------------------------

interface StylePalette {
  header: number;
  bodyEven: number;
  bodyOdd: number;
  colBodyEven: (number | null)[];
  colBodyOdd: (number | null)[];
}

function buildBorderObj(opts: DrawTableOptions) {
  const bs = opts.borderStyle === undefined ? ('thin' satisfies BorderStyle) : opts.borderStyle;
  const bc = opts.borderColor ?? '000000';
  if (!bs) return undefined;
  return {
    left: { style: bs, color: bc },
    right: { style: bs, color: bc },
    top: { style: bs, color: bc },
    bottom: { style: bs, color: bc },
  };
}

function buildPalette(wb: Workbook, opts: DrawTableOptions, colCount: number): StylePalette {
  const styles = wb.styles;
  const borderObj = buildBorderObj(opts);

  // Header style
  const headerBuilder = new StyleBuilder()
    .font({ bold: true, color: 'FFFFFF', ...(opts.headerFont ?? {}) })
    .fill({ pattern: 'solid', fgColor: opts.headerColor ?? '4472C4' });
  if (borderObj) headerBuilder.border(borderObj);
  headerBuilder.alignment({
    horizontal: opts.headerAlign ?? 'center',
    vertical: opts.verticalAlign ?? null,
  });
  const headerIdx = headerBuilder.build(styles);

  const bodyEvenIdx = buildBodyStyle(styles, opts, borderObj, false);
  const bodyOddIdx = opts.alternateRowColor
    ? buildBodyStyle(styles, opts, borderObj, true)
    : bodyEvenIdx;

  // Per-column styles
  const colBodyEven: (number | null)[] = new Array<number | null>(colCount).fill(null);
  const colBodyOdd: (number | null)[] = new Array<number | null>(colCount).fill(null);

  if (opts.columns) {
    for (let c = 0; c < colCount; c++) {
      const col = opts.columns[c];
      if (!col?.align && !col?.numberFormat) continue;
      colBodyEven[c] = buildBodyStyle(styles, opts, borderObj, false, col);
      colBodyOdd[c] = opts.alternateRowColor
        ? buildBodyStyle(styles, opts, borderObj, true, col)
        : (colBodyEven[c] ?? null);
    }
  }

  return { header: headerIdx, bodyEven: bodyEvenIdx, bodyOdd: bodyOddIdx, colBodyEven, colBodyOdd };
}

function buildBodyStyle(
  styles: StylesData,
  opts: DrawTableOptions,
  borderObj: ReturnType<typeof buildBorderObj>,
  isOdd: boolean,
  col?: TableColumn,
): number {
  const sb = new StyleBuilder();
  if (opts.bodyFont) sb.font(opts.bodyFont);
  if (isOdd && opts.alternateRowColor) {
    sb.fill({ pattern: 'solid', fgColor: opts.alternateRowColor });
  }
  if (borderObj) sb.border(borderObj);
  const align: Partial<AlignmentData> = {};
  const horizontal = col?.align || opts.bodyAlign;
  if (horizontal) align.horizontal = horizontal;
  if (opts.verticalAlign) align.vertical = opts.verticalAlign;
  if (opts.wrapText) align.wrapText = true;
  if (Object.keys(align).length > 0) sb.alignment(align);
  if (col?.numberFormat) sb.numberFormat(col.numberFormat);
  return sb.build(styles);
}

// ---------------------------------------------------------------------------
// Auto-width calculation
// ---------------------------------------------------------------------------

function estimateWidth(value: TableCellValue | undefined): number {
  if (value == null) return 8;
  const str = value instanceof Date ? 'yyyy-mm-dd' : String(value);
  let len = 0;
  for (const ch of str) {
    const code = ch.codePointAt(0) ?? 0;
    len += code > 0x2e7f ? 2 : 1;
  }
  return Math.max(8, Math.ceil(len * 1.2) + 2);
}

function computeAutoWidths(
  headers: readonly string[],
  rows: readonly TableCellValue[][],
): number[] {
  const widths = headers.map((h) => estimateWidth(h));
  for (const row of rows) {
    for (let c = 0; c < widths.length; c++) {
      const w = estimateWidth(row[c]);
      const cur = widths[c];
      if (cur !== undefined && w > cur) widths[c] = w;
    }
  }
  return widths;
}

// ---------------------------------------------------------------------------
// Per-cell style override
// ---------------------------------------------------------------------------

function buildCellOverrideStyle(wb: Workbook, baseIdx: number, override: CellStyle): number {
  const styles = wb.styles;
  const base = styles.cellXfs[baseIdx];
  const sb = new StyleBuilder();

  // Merge font
  const baseFont = styles.fonts[base?.fontId ?? 0];
  sb.font({ ...baseFont, ...(override.font ?? {}) });

  // Merge fill
  const baseFill = styles.fills[base?.fillId ?? 0];
  if (override.fill) {
    const requested = override.fill.pattern ?? baseFill?.patternType ?? 'none';
    sb.fill({
      pattern: isPatternType(requested) ? requested : 'none',
      fgColor: override.fill.fgColor ?? baseFill?.fgColor ?? null,
    });
  } else if (baseFill && baseFill.patternType !== 'none') {
    sb.fill({
      pattern: isPatternType(baseFill.patternType) ? baseFill.patternType : 'none',
      fgColor: baseFill.fgColor,
    });
  }

  // Merge border
  const baseBorder = styles.borders[base?.borderId ?? 0];
  if (override.border || baseBorder) {
    const left = override.border?.left ?? baseBorder?.left;
    const right = override.border?.right ?? baseBorder?.right;
    const top = override.border?.top ?? baseBorder?.top;
    const bottom = override.border?.bottom ?? baseBorder?.bottom;
    if (left || right || top || bottom) {
      const borderOpts: Parameters<StyleBuilder['border']>[0] = {};
      if (left) borderOpts.left = left;
      if (right) borderOpts.right = right;
      if (top) borderOpts.top = top;
      if (bottom) borderOpts.bottom = bottom;
      sb.border(borderOpts);
    }
  }

  // Merge alignment
  const baseAlign = base?.alignment ?? {};
  if (override.alignment || Object.keys(baseAlign).length > 0) {
    sb.alignment({ ...baseAlign, ...(override.alignment ?? {}) });
  }

  // Number format
  if (override.numberFormat) {
    sb.numberFormat(override.numberFormat);
  } else if (base?.numFmtId && base.numFmtId > 0) {
    const fmt = styles.numFmts.find((f) => f.id === base.numFmtId);
    if (fmt) sb.numberFormat(fmt.formatCode);
  }

  return sb.build(styles);
}

// ---------------------------------------------------------------------------
// drawTable — main entry point
// ---------------------------------------------------------------------------

/**
 * Draw a styled table on a worksheet.
 *
 * Creates header row, data rows, applies styles, borders, column widths,
 * merge cells, zebra striping, frozen panes, and auto-filter.
 * `columns[c].header` overrides `headers[c]`. A `Date` cell is written as the serial of its UTC
 * instant in the workbook's date system (an invalid `Date` leaves the cell empty), and a column
 * holding a `Date` gets the number format `yyyy-mm-dd` unless `columns[c].numberFormat` is set;
 * a `bigint` cell is written as a number when it is a safe integer and as a string otherwise.
 *
 * @param wb - The parent Workbook (needed for shared styles).
 * @param ws - The target Worksheet to draw the table on.
 * @param opts - Table configuration: headers, rows, styling, and layout options.
 * @returns A TableResult with range metadata and row/column counts.
 * @throws {RangeError} When `headers` is empty.
 *
 * @example
 * ```ts
 * const result = drawTable(wb, ws, {
 *   headers: ['Name', 'Score'],
 *   rows: [['Alice', 95], ['Bob', 87]],
 *   headerColor: '4472C4',
 *   freezeHeader: true,
 * });
 * console.log(result.range); // 'A1:B3'
 * ```
 */
export function drawTable(wb: Workbook, ws: Worksheet, opts: DrawTableOptions): TableResult {
  const { headers, rows } = opts;
  const colCount = headers.length;
  if (colCount === 0) throw new RangeError('drawTable: headers must not be empty');
  const origin = opts.origin ? decodeCellRef(opts.origin) : { row: 0, col: 0 };

  // A `columns[c].header` override wins over `headers[c]`.
  const headerLabels = Array.from(
    { length: colCount },
    (_, c) => opts.columns?.[c]?.header ?? headers[c] ?? '',
  );

  // Date columns default to an ISO date format unless the caller set one.
  const columns: TableColumn[] = Array.from({ length: colCount }, (_, c) => {
    const col = opts.columns?.[c] ?? {};
    const hasDate = rows.some((row) => row[c] instanceof Date);
    return hasDate && col.numberFormat === undefined ? { ...col, numberFormat: 'yyyy-mm-dd' } : col;
  });

  // Build style palette once
  const palette = buildPalette(wb, { ...opts, columns }, colCount);

  // --- Column widths: columnWidths[c] > columns[c].width > auto (0 or less is unset) ---
  const autoWidths = opts.autoWidth ? computeAutoWidths(headerLabels, rows) : [];
  for (let c = 0; c < colCount; c++) {
    const fixed = opts.columnWidths?.[c];
    const colWidth = columns[c]?.width;
    const w =
      fixed !== undefined && fixed > 0
        ? fixed
        : colWidth !== undefined && colWidth > 0
          ? colWidth
          : autoWidths[c];
    if (w !== undefined && w > 0) ws.setColumnWidth(origin.col + c + 1, w);
  }

  // --- Header row ---
  const headerRowIdx = origin.row + 1; // 1-based worksheet row
  for (let c = 0; c < colCount; c++) {
    const ref = `${columnToLetter(origin.col + c)}${headerRowIdx}`;
    const cell = ws.cell(ref);
    cell.value = headerLabels[c] ?? '';
    cell.styleIndex = palette.header;
  }

  // --- Data rows ---
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const wsRowIdx = origin.row + r + 2; // 1-based, after header

    const isOdd = r % 2 === 1;

    for (let c = 0; c < colCount; c++) {
      const ref = `${columnToLetter(origin.col + c)}${wsRowIdx}`;
      const cell = ws.cell(ref);
      const val = toCellValue(row[c], wb.dateSystem);
      if (val !== null) {
        cell.value = val;
      }

      // Style: per-column override > zebra stripe > base body
      const colStyle = isOdd ? palette.colBodyOdd[c] : palette.colBodyEven[c];
      let styleIdx = colStyle ?? (isOdd ? palette.bodyOdd : palette.bodyEven);

      // Per-cell override
      const overrideKey = `${r},${c}`;
      const cellOverride = opts.cellStyles?.[overrideKey];
      if (cellOverride) {
        styleIdx = buildCellOverrideStyle(wb, styleIdx, cellOverride);
      }

      cell.styleIndex = styleIdx;
    }
  }

  // --- Merge cells ---
  if (opts.merges) {
    for (const merge of opts.merges) {
      const startRow = origin.row + merge.row + 1; // 0-based origin + 0-based data row + 1 for header
      const startCol = origin.col + merge.col;
      const endRow = startRow + (merge.rowSpan ?? 1) - 1;
      const endCol = startCol + (merge.colSpan ?? 1) - 1;
      ws.addMergeCell(encodeRange({ row: startRow, col: startCol }, { row: endRow, col: endCol }));
    }
  }

  // --- Frozen header: every row down to the header stays put ---
  if (opts.freezeHeader) {
    ws.frozenPane = { rows: origin.row + 1, cols: 0 };
  }

  // --- Auto filter: header row through the last data row ---
  if (opts.autoFilter) {
    ws.autoFilter = encodeRange(
      { row: origin.row, col: origin.col },
      { row: origin.row + rows.length, col: origin.col + colCount - 1 },
    );
  }

  // --- Result ---
  const totalRows = 1 + rows.length;
  return {
    range: encodeRange(
      { row: origin.row, col: origin.col },
      { row: origin.row + totalRows - 1, col: origin.col + colCount - 1 },
    ),
    rowCount: totalRows,
    colCount,
    firstDataRow: origin.row + 1,
    lastDataRow: origin.row + rows.length,
  };
}

function toCellValue(
  v: TableCellValue | undefined,
  system: Workbook['dateSystem'],
): string | number | boolean | null {
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : dateToSerial(v, system);
  if (typeof v === 'bigint') {
    const n = Number(v);
    return Number.isSafeInteger(n) ? n : String(v);
  }
  return v;
}

// ---------------------------------------------------------------------------
// drawTableFromData — create table from JSON array
// ---------------------------------------------------------------------------

/** Options for drawTableFromData. */
export interface DrawTableFromDataOptions extends Omit<DrawTableOptions, 'headers' | 'rows'> {
  /** Explicit column order. If omitted, uses Object.keys of first item. */
  headers?: string[];
  /** Map object keys to display headers. */
  headerMap?: Record<string, string>;
}

/**
 * Create a styled table from a JSON array.
 *
 * Automatically extracts headers from object keys and maps values to rows.
 * With no data and no explicit `headers`, nothing is drawn and the result has
 * an empty `range`; with explicit `headers` only the header row is drawn.
 *
 * @param wb - The parent Workbook (needed for shared styles).
 * @param ws - The target Worksheet to draw the table on.
 * @param data - Array of JSON objects representing the table rows.
 * @param opts - Optional overrides for headers, header labels, styling, and layout.
 * @returns A TableResult with range metadata and row/column counts.
 *
 * @example
 * ```ts
 * const result = drawTableFromData(wb, ws, [
 *   { name: 'Alice', score: 95 },
 *   { name: 'Bob', score: 87 },
 * ], { autoWidth: true });
 * ```
 */
export function drawTableFromData(
  wb: Workbook,
  ws: Worksheet,
  data: Record<string, unknown>[],
  opts?: DrawTableFromDataOptions,
): TableResult {
  const keys = opts?.headers ?? Object.keys(data[0] ?? {});
  const headers = keys.map(
    (k) =>
      (opts?.headerMap && Object.hasOwn(opts.headerMap, k) ? opts.headerMap[k] : undefined) ?? k,
  );
  if (headers.length === 0) {
    // Nothing to draw: the worksheet stays untouched.
    const origin = opts?.origin ? decodeCellRef(opts.origin) : { row: 0, col: 0 };
    return {
      range: '',
      rowCount: 0,
      colCount: 0,
      firstDataRow: origin.row + 1,
      lastDataRow: origin.row,
    };
  }

  const rows: TableCellValue[][] = data.map((item) =>
    keys.map((key) => {
      // An inherited Object.prototype member has no value, but a getter on the record's own
      // class prototype still counts.
      const val = key in Object.prototype && !Object.hasOwn(item, key) ? undefined : item[key];
      if (val === null || val === undefined) return null;
      if (
        typeof val === 'string' ||
        typeof val === 'number' ||
        typeof val === 'boolean' ||
        typeof val === 'bigint' ||
        val instanceof Date
      )
        return val;
      return String(val);
    }),
  );

  return drawTable(wb, ws, { ...opts, headers, rows });
}
