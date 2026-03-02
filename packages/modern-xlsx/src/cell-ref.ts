/**
 * Cell reference utilities for converting between A1-style references and numeric indices.
 * All row/column indices are 0-based unless otherwise noted.
 */

const CHAR_CODE_A = 65; // 'A'.charCodeAt(0)
const CELL_REF_RE = /^\$?([A-Z]+)\$?(\d+)$/;

/**
 * Convert a 0-based column index to a column letter (0 → "A", 25 → "Z", 26 → "AA").
 */
export function columnToLetter(col: number): string {
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
 */
export function letterToColumn(letter: string): number {
  const upper = letter.toUpperCase();
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * 26 + (upper.charCodeAt(i) - CHAR_CODE_A + 1);
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
 */
export function encodeCellRef(row: number, col: number): string {
  return `${columnToLetter(col)}${row + 1}`;
}

/**
 * Decode an A1-style cell reference into 0-based row/col.
 * `decodeCellRef("A1")` → `{ row: 0, col: 0 }`
 */
export function decodeCellRef(ref: string): CellAddress {
  const match = ref.match(CELL_REF_RE);
  if (!match?.[1] || !match[2]) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  return {
    col: letterToColumn(match[1]),
    row: Number.parseInt(match[2], 10) - 1,
  };
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
 */
export function encodeRange(start: CellAddress, end: CellAddress): string {
  return `${encodeCellRef(start.row, start.col)}:${encodeCellRef(end.row, end.col)}`;
}

/**
 * Decode a range string into 0-based start/end addresses.
 * `decodeRange("A1:C10")` → `{ start: { row: 0, col: 0 }, end: { row: 9, col: 2 } }`
 */
export function decodeRange(range: string): CellRange {
  const [startRef, endRef] = range.split(':');
  if (!startRef || !endRef) {
    throw new Error(`Invalid range: ${range}`);
  }
  return {
    start: decodeCellRef(startRef),
    end: decodeCellRef(endRef),
  };
}
