/**
 * Cell reference utilities for converting between A1-style references and numeric indices.
 * All row/column indices are 0-based unless otherwise noted. Every codec enforces the
 * worksheet bounds `A1:XFD1048576` and throws `INVALID_CELL_REF` outside them, except
 * {@link decodeRow}, which throws `INVALID_ARGUMENT`.
 */

import { INVALID_ARGUMENT, INVALID_CELL_REF, ModernXlsxError } from '../model/errors.js';

/** Largest valid 0-based column index (`XFD`, column 16384 in A1 notation). */
export const MAX_COL = 16383;
/** Largest valid 0-based row index (row 1048576 in A1 notation). */
export const MAX_ROW = 1048575;

const CHAR_CODE_A = 65; // 'A'.charCodeAt(0)
const CELL_REF_RE = /^\$?([A-Z]+)\$?(\d+)$/;
const COLUMN_LETTERS_RE = /^[A-Za-z]{1,3}$/;

/** Whether `n` is an integer in `0..max`. */
function inBounds(n: number, max: number): boolean {
  return Number.isInteger(n) && n >= 0 && n <= max;
}

/**
 * Convert a 0-based column index to a column letter (0 → "A", 25 → "Z", 26 → "AA").
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` unless `col` is an integer in `0..MAX_COL`.
 */
export function columnToLetter(col: number): string {
  if (!inBounds(col, MAX_COL)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid column index: ${col}`);
  }
  let result = '';
  let c = col;
  do {
    result = String.fromCharCode(CHAR_CODE_A + (c % 26)) + result;
    c = Math.floor(c / 26) - 1;
  } while (c >= 0);
  return result;
}

/**
 * Convert a column letter to a 0-based column index ("A" → 0, "Z" → 25, "AA" → 26).
 * Lowercase letters are accepted.
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` unless `letter` is 1 to 3 ASCII letters
 * naming a column up to `XFD`.
 */
export function letterToColumn(letter: string): number {
  if (!COLUMN_LETTERS_RE.test(letter)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid column letters: '${letter}'`);
  }
  const upper = letter.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - CHAR_CODE_A + 1);
  }
  if (col - 1 > MAX_COL) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid column letters: '${letter}'`);
  }
  return col - 1;
}

/** A 0-based cell address with row and column indices. */
export interface CellAddress {
  /** 0-based row index. */
  readonly row: number;
  /** 0-based column index. */
  readonly col: number;
}

/**
 * Encode a 0-based row/col into an A1-style cell reference.
 * `encodeCellRef(0, 0)` → `"A1"`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` unless `row` is an integer in `0..MAX_ROW`
 * and `col` is an integer in `0..MAX_COL`.
 */
export function encodeCellRef(row: number, col: number): string {
  if (!inBounds(row, MAX_ROW) || !inBounds(col, MAX_COL)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid cell reference: row ${row}, col ${col}`);
  }
  return `${columnToLetter(col)}${row + 1}`;
}

/**
 * Decode an A1-style cell reference into 0-based row/col.
 * `decodeCellRef("A1")` → `{ row: 0, col: 0 }`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` when `ref` is not an uppercase A1 reference
 * (optionally with `$` markers) or lies outside `A1:XFD1048576`.
 */
export function decodeCellRef(ref: string): CellAddress {
  const match = ref.match(CELL_REF_RE);
  if (!match?.[1] || !match[2]) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid cell reference: ${ref}`);
  }
  const row = Number.parseInt(match[2], 10) - 1;
  const col = columnIndexOrNull(match[1]);
  if (col === null || !inBounds(row, MAX_ROW)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid cell reference: ${ref}`);
  }
  return { col, row };
}

/** Like {@link letterToColumn} but returns `null` instead of throwing. */
export function columnIndexOrNull(letters: string): number | null {
  try {
    return letterToColumn(letters);
  } catch {
    return null;
  }
}

/** A rectangular cell range defined by start and end addresses. */
export interface CellRange {
  /** Top-left corner of the range. */
  readonly start: CellAddress;
  /** Bottom-right corner of the range. */
  readonly end: CellAddress;
}

/**
 * Encode a range from 0-based start/end addresses.
 * `encodeRange({ row: 0, col: 0 }, { row: 9, col: 2 })` → `"A1:C10"`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` when either address is out of bounds.
 */
export function encodeRange(start: CellAddress, end: CellAddress): string {
  return `${encodeCellRef(start.row, start.col)}:${encodeCellRef(end.row, end.col)}`;
}

/**
 * Decode a range string into 0-based start/end addresses.
 * `decodeRange("A1:C10")` → `{ start: { row: 0, col: 0 }, end: { row: 9, col: 2 } }`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` when the string is not `start:end` or either
 * endpoint is not a valid cell reference.
 */
export function decodeRange(range: string): CellRange {
  const parts = range.split(':');
  const [startRef, endRef] = parts;
  if (parts.length !== 2 || !startRef || !endRef) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid range: ${range}`);
  }
  return {
    start: decodeCellRef(startRef),
    end: decodeCellRef(endRef),
  };
}

/** Result of splitting a cell reference into its components. */
export interface SplitCellRef {
  /** Column letters (e.g., "A", "XFD"). */
  readonly col: string;
  /** Row number as string (1-based, e.g., "1", "1048576"). */
  readonly row: string;
  /** Whether the column is absolute ($A). */
  readonly absCol: boolean;
  /** Whether the row is absolute ($1). */
  readonly absRow: boolean;
}

/**
 * Convert a 0-based row index to a 1-based row string.
 * `encodeRow(0)` → `"1"`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` unless `row` is an integer in `0..MAX_ROW`.
 */
export function encodeRow(row: number): string {
  if (!inBounds(row, MAX_ROW)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid row index: ${row}`);
  }
  return String(row + 1);
}

/**
 * Convert a 1-based row string to a 0-based row index.
 * `decodeRow("1")` → `0`
 *
 * @throws {ModernXlsxError} `INVALID_ARGUMENT` unless `rowStr` parses to a row in `1..1048576`.
 */
export function decodeRow(rowStr: string): number {
  const n = Number.parseInt(rowStr, 10);
  if (!Number.isFinite(n) || n < 1 || n - 1 > MAX_ROW) {
    throw new ModernXlsxError(INVALID_ARGUMENT, `Invalid row string: ${rowStr}`);
  }
  return n - 1;
}

const SPLIT_RE = /^(\$?)([A-Z]+)(\$?)(\d+)$/;

/**
 * Split a cell reference into column/row parts with absolute flags.
 * The ref is upper-cased (`String.prototype.toUpperCase`) before matching, so lowercase input is accepted,
 * and so is any character whose uppercase form is ASCII letters (`ß1` reads as `SS1`); the returned `col` is uppercase.
 * `splitCellRef("$A$1")` → `{ col: "A", row: "1", absCol: true, absRow: true }`
 *
 * @throws {ModernXlsxError} `INVALID_CELL_REF` when the upper-cased `ref` is not a well-formed A1 reference
 * or lies outside `A1:XFD1048576`.
 */
export function splitCellRef(ref: string): SplitCellRef {
  const match = ref.toUpperCase().match(SPLIT_RE);
  if (!match?.[2] || !match[4]) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid cell reference: ${ref}`);
  }
  const row = Number.parseInt(match[4], 10) - 1;
  if (columnIndexOrNull(match[2]) === null || !inBounds(row, MAX_ROW)) {
    throw new ModernXlsxError(INVALID_CELL_REF, `Invalid cell reference: ${ref}`);
  }
  return {
    col: match[2],
    row: match[4],
    absCol: match[1] === '$',
    absRow: match[3] === '$',
  };
}
