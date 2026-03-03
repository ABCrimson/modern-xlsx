/**
 * Sheet conversion utilities — helpers for converting worksheet data
 * to/from common JavaScript data structures.
 */

import { columnToLetter, decodeCellRef } from './cell-ref.js';
import type { CellData, RowData, SheetData } from './types.js';
import { Worksheet } from './workbook.js';

// ---------------------------------------------------------------------------
// sheetToJson
// ---------------------------------------------------------------------------

/** Options for the {@link sheetToJson} function. */
export interface SheetToJsonOptions {
  /** Row to use as header. `'A'` = use column letters (default), number = 1-based row index, string[] = explicit header. */
  header?: 'A' | number | string[];
  /** Range to extract, e.g. "A1:D10". If omitted, uses full data range. */
  range?: string;
  /** Default value for empty cells. */
  defval?: unknown;
  /** Maximum number of data rows to return. */
  sheetRows?: number;
}

interface DataBounds {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}

/**
 * Convert a Worksheet to an array of JSON objects.
 *
 * By default, the first row is treated as the header row.
 * Pass `header: 'A'` to use column letters as keys.
 */
export function sheetToJson<T extends Record<string, unknown> = Record<string, unknown>>(
  ws: Worksheet,
  opts?: SheetToJsonOptions,
): T[] {
  const rows = ws.rows;
  if (rows.length === 0) return [];

  const bounds = computeBounds(rows, opts?.range);
  const { headers, dataStartRow } = resolveHeaders(rows, bounds, opts?.header);

  const result = collectRows<T>(rows, headers, bounds, dataStartRow, opts?.defval);
  if (opts?.sheetRows !== undefined && opts.sheetRows > 0) {
    return result.slice(0, opts.sheetRows);
  }
  return result;
}

function computeBounds(rows: readonly RowData[], range?: string): DataBounds {
  if (range) {
    const [sRef, eRef] = range.split(':');
    if (sRef && eRef) {
      const s = decodeCellRef(sRef);
      const e = decodeCellRef(eRef);
      return { startRow: s.row, endRow: e.row, startCol: s.col, endCol: e.col };
    }
  }

  // Only scan all cells if no range given.
  // Use actual row indices (1-based) not array indices — otherwise sparse
  // sheets (e.g. rows at 1, 100, 200) would be truncated.
  const startRow = 0;
  let endRow = 0;
  const startCol = 0;
  let endCol = -1;

  for (const row of rows) {
    if (row.index > endRow) endRow = row.index;
    for (const cell of row.cells) {
      const { col } = decodeCellRef(cell.reference);
      if (col > endCol) endCol = col;
    }
  }

  return { startRow, endRow, startCol, endCol };
}

function resolveHeaders(
  rows: readonly RowData[],
  bounds: DataBounds,
  headerOpt?: 'A' | number | string[],
): { headers: string[]; dataStartRow: number } {
  if (headerOpt === 'A') {
    const headers = Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, i) =>
      columnToLetter(bounds.startCol + i),
    );
    return { headers, dataStartRow: bounds.startRow };
  }

  if (Array.isArray(headerOpt)) {
    return { headers: headerOpt, dataStartRow: bounds.startRow };
  }

  if (typeof headerOpt === 'number') {
    const headerRow = rows.find((r) => r.index === headerOpt);
    return {
      headers: buildHeaderFromRow(headerRow, bounds.startCol, bounds.endCol),
      dataStartRow: headerOpt,
    };
  }

  // Default: first data row is the header
  const firstRow = rows.find((r) => r.index >= bounds.startRow + 1);
  return {
    headers: buildHeaderFromRow(firstRow, bounds.startCol, bounds.endCol),
    dataStartRow: firstRow?.index ?? bounds.startRow + 1,
  };
}

function buildRowObject(
  row: RowData,
  headers: string[],
  colLetters: string[],
  bounds: DataBounds,
  defval?: unknown,
): Record<string, unknown> {
  const cellMap = new Map(row.cells.map((c) => [c.reference, c]));
  const obj: Record<string, unknown> = {};
  for (let col = bounds.startCol; col <= bounds.endCol; col++) {
    const idx = col - bounds.startCol;
    const ref = `${colLetters[idx]}${row.index}`;
    const cellData = cellMap.get(ref);
    const key = headers[idx] ?? colLetters[idx];
    if (key === undefined) continue;
    if (cellData?.value != null) {
      obj[key] = parseCellValue(cellData);
    } else if (defval !== undefined) {
      obj[key] = defval;
    }
  }
  return obj;
}

function collectRows<T extends Record<string, unknown>>(
  rows: readonly RowData[],
  headers: string[],
  bounds: DataBounds,
  dataStartRow: number,
  defval?: unknown,
): T[] {
  const colLetters = Array.from({ length: bounds.endCol - bounds.startCol + 1 }, (_, i) =>
    columnToLetter(bounds.startCol + i),
  );
  const result: T[] = [];
  for (const row of rows) {
    if (row.index <= dataStartRow || row.index > bounds.endRow + 1) continue;
    result.push(buildRowObject(row, headers, colLetters, bounds, defval) as T);
  }
  return result;
}

function buildHeaderFromRow(row: RowData | undefined, startCol: number, endCol: number): string[] {
  const headers: string[] = [];
  const cellMap = row ? new Map(row.cells.map((c) => [c.reference, c])) : null;
  for (let col = startCol; col <= endCol; col++) {
    const ref = `${columnToLetter(col)}${row?.index ?? 1}`;
    const cell = cellMap?.get(ref);
    headers.push(cell?.value ?? columnToLetter(col));
  }
  return headers;
}

function parseCellValue(cell: CellData): string | number | boolean {
  if (cell.value == null) return '';
  switch (cell.cellType) {
    case 'number':
      return Number.parseFloat(cell.value);
    case 'boolean':
      return cell.value === '1';
    default:
      return cell.value;
  }
}

// ---------------------------------------------------------------------------
// jsonToSheet
// ---------------------------------------------------------------------------

/** Options for the {@link jsonToSheet} function. */
export interface JsonToSheetOptions {
  /** Explicit header order. If omitted, uses Object.keys of first row. */
  header?: string[];
  /** Skip writing header row. */
  skipHeader?: boolean;
}

/**
 * Create a Worksheet from an array of JSON objects.
 */
export function jsonToSheet(data: Record<string, unknown>[], opts?: JsonToSheetOptions): Worksheet {
  if (data.length === 0) {
    return createEmptyWorksheet('Sheet1');
  }

  const firstRow = data[0];
  const headers = opts?.header ?? (firstRow ? Object.keys(firstRow) : []);
  const ws = new Worksheet(createEmptySheetData('Sheet1'));
  let rowIdx = 1;

  if (!opts?.skipHeader) {
    rowIdx = writeHeaderRow(ws, headers, rowIdx);
  }

  for (const record of data) {
    writeDataRow(ws, headers, record, rowIdx);
    rowIdx++;
  }

  return ws;
}

function writeHeaderRow(ws: Worksheet, headers: string[], rowIdx: number, startCol = 0): number {
  for (let col = 0; col < headers.length; col++) {
    const hdr = headers[col];
    if (hdr) ws.cell(`${columnToLetter(startCol + col)}${rowIdx}`).value = hdr;
  }
  return rowIdx + 1;
}

function writeDataRow(
  ws: Worksheet,
  headers: string[],
  record: Record<string, unknown>,
  rowIdx: number,
  startCol = 0,
): void {
  for (let col = 0; col < headers.length; col++) {
    const key = headers[col];
    const val = key ? record[key] : undefined;
    if (val === undefined || val === null) continue;

    const cell = ws.cell(`${columnToLetter(startCol + col)}${rowIdx}`);
    if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
      cell.value = val;
    } else {
      cell.value = String(val);
    }
  }
}

// ---------------------------------------------------------------------------
// aoaToSheet
// ---------------------------------------------------------------------------

/** Options for the {@link aoaToSheet} function. */
export interface AoaToSheetOptions {
  /** Origin cell reference, e.g. "A1" (default). */
  origin?: string;
}

/**
 * Create a Worksheet from a 2D array (array of arrays).
 */
export function aoaToSheet(data: unknown[][], opts?: AoaToSheetOptions): Worksheet {
  const origin = opts?.origin ? decodeCellRef(opts.origin) : { row: 0, col: 0 };
  const ws = new Worksheet(createEmptySheetData('Sheet1'));

  for (let r = 0; r < data.length; r++) {
    const rowArr = data[r];
    if (!rowArr) continue;
    for (let c = 0; c < rowArr.length; c++) {
      const val = rowArr[c];
      if (val === undefined || val === null) continue;

      const ref = `${columnToLetter(origin.col + c)}${origin.row + r + 1}`;
      const cell = ws.cell(ref);
      if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
        cell.value = val;
      } else {
        cell.value = String(val);
      }
    }
  }

  return ws;
}

// ---------------------------------------------------------------------------
// sheetToCsv
// ---------------------------------------------------------------------------

/** Options for the {@link sheetToCsv} function. */
export interface SheetToCsvOptions {
  /** Field separator (default: ","). */
  separator?: string;
  /** Force-quote all fields (default: false). */
  forceQuote?: boolean;
  /** Maximum number of rows to include in the output. */
  sheetRows?: number;
}

/**
 * Convert a Worksheet to a CSV string.
 */
export function sheetToCsv(ws: Worksheet, opts?: SheetToCsvOptions): string {
  const sep = opts?.separator ?? ',';
  const force = opts?.forceQuote ?? false;
  const rows = ws.rows;

  if (rows.length === 0) return '';

  const { maxCol, rowMap, minRowIdx, maxRowIdx } = buildRowIndex(rows);

  const effectiveMaxRow =
    opts?.sheetRows !== undefined && opts.sheetRows > 0
      ? Math.min(maxRowIdx, minRowIdx + opts.sheetRows - 1)
      : maxRowIdx;
  const colLetters = Array.from({ length: maxCol + 1 }, (_, i) => columnToLetter(i));
  const lines: string[] = [];
  for (let r = minRowIdx; r <= effectiveMaxRow; r++) {
    const row = rowMap.get(r);
    const cellMap = row ? new Map(row.cells.map((c) => [c.reference, c])) : null;
    const fields: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      const ref = `${colLetters[c]}${r}`;
      const cell = cellMap?.get(ref);
      fields.push(csvQuote(cell?.value ?? '', sep, force));
    }
    lines.push(fields.join(sep));
  }

  return lines.join('\n');
}

function buildRowIndex(rows: readonly RowData[]) {
  let maxCol = 0;
  const rowMap = new Map<number, RowData>();
  let minRowIdx = Number.MAX_SAFE_INTEGER;
  let maxRowIdx = 0;

  for (const row of rows) {
    rowMap.set(row.index, row);
    if (row.index < minRowIdx) minRowIdx = row.index;
    if (row.index > maxRowIdx) maxRowIdx = row.index;
    for (const cell of row.cells) {
      const { col } = decodeCellRef(cell.reference);
      if (col > maxCol) maxCol = col;
    }
  }

  return { maxCol, rowMap, minRowIdx, maxRowIdx };
}

function csvQuote(value: string, sep: string, force: boolean): string {
  if (force || value.includes(sep) || value.includes('"') || value.includes('\n')) {
    return `"${value.replaceAll('"', '""')}"`;
  }
  return value;
}

// ---------------------------------------------------------------------------
// sheetToHtml
// ---------------------------------------------------------------------------

/** Options for the {@link sheetToHtml} function. */
export interface SheetToHtmlOptions {
  /** CSS class for the table element. */
  className?: string;
  /** Include inline styles derived from cell data types. */
  includeStyle?: boolean;
  /** Include a <thead> element for the first row. */
  header?: boolean;
}

/**
 * Convert a Worksheet to an HTML table string.
 */
export function sheetToHtml(ws: Worksheet, opts?: SheetToHtmlOptions): string {
  const rows = ws.rows;
  if (rows.length === 0) return '<table></table>';

  const { maxCol, rowMap, minRowIdx, maxRowIdx } = buildRowIndex(rows);
  const colLetters = Array.from({ length: maxCol + 1 }, (_, i) => columnToLetter(i));
  const classAttr = opts?.className ? ` class="${escapeHtml(opts.className)}"` : '';
  const lines: string[] = [`<table${classAttr}>`];

  for (let r = minRowIdx; r <= maxRowIdx; r++) {
    const isFirst = r === minRowIdx;
    renderHtmlRow(lines, rowMap.get(r), r, maxCol, colLetters, isFirst, opts);
  }

  if (opts?.header) lines.push('</tbody>');
  lines.push('</table>');
  return lines.join('');
}

function renderHtmlRow(
  lines: string[],
  row: RowData | undefined,
  rowIdx: number,
  maxCol: number,
  colLetters: string[],
  isFirst: boolean,
  opts?: SheetToHtmlOptions,
): void {
  const cellMap = row ? new Map(row.cells.map((c) => [c.reference, c])) : null;
  const tag = opts?.header && isFirst ? 'th' : 'td';

  if (opts?.header && isFirst) lines.push('<thead>');
  lines.push('<tr>');

  for (let c = 0; c <= maxCol; c++) {
    const cell = cellMap?.get(`${colLetters[c]}${rowIdx}`);
    const val = cell?.value ?? '';
    const style = opts?.includeStyle ? htmlCellStyle(cell) : '';
    const styleAttr = style ? ` style="${style}"` : '';
    lines.push(`<${tag}${styleAttr}>${escapeHtml(String(val))}</${tag}>`);
  }

  lines.push('</tr>');
  if (opts?.header && isFirst) {
    lines.push('</thead>');
    lines.push('<tbody>');
  }
}

function htmlCellStyle(cell: CellData | undefined): string {
  if (!cell) return '';
  const parts: string[] = [];
  if (cell.cellType === 'number') parts.push('text-align:right');
  if (cell.cellType === 'boolean') parts.push('text-align:center');
  return parts.join(';');
}

function escapeHtml(str: string): string {
  return str
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

// ---------------------------------------------------------------------------
// sheetAddAoa — append array-of-arrays data to existing worksheet
// ---------------------------------------------------------------------------

/** Options for the {@link sheetAddAoa} function. */
export interface SheetAddAoaOptions {
  /** Starting cell reference (default: next row after existing data). */
  origin?: string;
}

/**
 * Append array-of-arrays data to an existing Worksheet.
 */
export function sheetAddAoa(ws: Worksheet, data: unknown[][], opts?: SheetAddAoaOptions): void {
  const origin = opts?.origin ? decodeCellRef(opts.origin) : { row: nextEmptyRow(ws), col: 0 };

  for (let r = 0; r < data.length; r++) {
    const rowArr = data[r];
    if (!rowArr) continue;
    writeAoaRow(ws, rowArr, origin.row + r + 1, origin.col);
  }
}

function writeAoaRow(ws: Worksheet, rowArr: unknown[], rowIdx: number, startCol: number): void {
  for (let c = 0; c < rowArr.length; c++) {
    const val = rowArr[c];
    if (val === undefined || val === null) continue;
    setCellValue(ws.cell(`${columnToLetter(startCol + c)}${rowIdx}`), val);
  }
}

// ---------------------------------------------------------------------------
// sheetAddJson — append JSON objects to existing worksheet
// ---------------------------------------------------------------------------

/** Options for the {@link sheetAddJson} function. */
export interface SheetAddJsonOptions {
  /** Explicit header order. If omitted, uses Object.keys of first row. */
  header?: string[];
  /** Skip writing header row (e.g. when appending to a sheet that already has headers). */
  skipHeader?: boolean;
  /** Starting cell reference (default: next row after existing data). */
  origin?: string;
}

/**
 * Append JSON objects to an existing Worksheet.
 */
export function sheetAddJson(
  ws: Worksheet,
  data: Record<string, unknown>[],
  opts?: SheetAddJsonOptions,
): void {
  if (data.length === 0) return;

  const firstRow = data[0];
  const headers = opts?.header ?? (firstRow ? Object.keys(firstRow) : []);
  const { row: startRow, col: startCol } = resolveAddOrigin(ws, opts?.origin);

  let rowIdx = startRow;

  if (!opts?.skipHeader) {
    rowIdx = writeHeaderRow(ws, headers, rowIdx, startCol);
  }

  for (const record of data) {
    writeJsonRow(ws, headers, record, rowIdx, startCol);
    rowIdx++;
  }
}

function resolveAddOrigin(ws: Worksheet, origin?: string): { row: number; col: number } {
  if (origin) {
    const decoded = decodeCellRef(origin);
    return { row: decoded.row + 1, col: decoded.col };
  }
  return { row: nextEmptyRow(ws) + 1, col: 0 };
}

function writeJsonRow(
  ws: Worksheet,
  headers: string[],
  record: Record<string, unknown>,
  rowIdx: number,
  startCol: number,
): void {
  for (let col = 0; col < headers.length; col++) {
    const key = headers[col];
    const val = key ? record[key] : undefined;
    if (val === undefined || val === null) continue;
    setCellValue(ws.cell(`${columnToLetter(startCol + col)}${rowIdx}`), val);
  }
}

function nextEmptyRow(ws: Worksheet): number {
  return ws.rows.at(-1)?.index ?? 0;
}

function setCellValue(cell: ReturnType<Worksheet['cell']>, val: unknown): void {
  if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
    cell.value = val;
  } else {
    cell.value = String(val);
  }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function createEmptySheetData(name: string): SheetData {
  return {
    name,
    worksheet: {
      dimension: null,
      rows: [],
      mergeCells: [],
      autoFilter: null,
      frozenPane: null,
      columns: [],
    },
  };
}

function createEmptyWorksheet(name: string): Worksheet {
  return new Worksheet(createEmptySheetData(name));
}
