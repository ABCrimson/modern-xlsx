/**
 * Formula reference rewriter for row/column insert and delete operations.
 *
 * Parses a formula to AST, walks and transforms its references according to
 * the action, then serializes back to a formula string. Rules, matching Excel:
 *
 * - A single cell ref shifts, or becomes `#REF!` when its row/column is
 *   deleted or an insert pushes it past row 1,048,576 or column XFD.
 * - A range (`A1:B5`, and the whole-column/row forms `A:C` and `1:3`) shrinks
 *   when a delete covers part of it and becomes one `#REF!` when a delete
 *   covers all of it. On an insert, an end pushed past the sheet edge is
 *   clamped to the last row/column; a start pushed past it gives `#REF!`.
 * - A range's sheet is its start's sheet, else its end's (`Sheet1!A1:B5` has
 *   the sheet on the start only); both endpoints move together.
 * - 3-D references (`Sheet1:Sheet3!A1`), structured references and names
 *   pass through unchanged: the workbook's sheet order is unknown here.
 * - Union (`(A1,B2)`) and intersection (`A1:B5 B2:C3`) operands are rewritten
 *   one by one.
 *
 * @module formula/rewriter
 */

import { columnIndexOrNull, columnToLetter } from '../format/cell-ref.js';
import { MAX_COL_INDEX, MAX_ROW } from './grid.js';
import type { ASTNode, CellRefNode, ColumnRangeNode, RangeNode, RowRangeNode } from './parser.js';
import { parseFormula } from './parser.js';
import { serializeFormula } from './serializer.js';
import { mapRefs } from './walk.js';

// ---------------------------------------------------------------------------
// Action types
// ---------------------------------------------------------------------------

/**
 * A row or column insert/delete that {@link rewriteFormula} applies to the
 * references in a formula.
 *
 * - `start` uses a different base per axis. For `insert_rows` and
 *   `delete_rows` it is a 1-based row number as written in A1 notation
 *   (`3` is row 3). For `insert_cols` and `delete_cols` it is a 0-based
 *   column index (`0` is column A, `2` is column C).
 * - An insert shifts every reference at index `>= start` by `count`.
 * - A delete removes `[start, start + count)`: a cell ref inside it becomes
 *   `#REF!`, a range shrinks (or becomes `#REF!` when wholly inside it), and
 *   later references shift back by `count`.
 * - `sheet` names the sheet whose qualified references are rewritten; the
 *   name matches ignoring case, as in Excel. When it is `undefined`, only
 *   unqualified references (no sheet prefix) are touched.
 *
 * @remarks The row/column base asymmetry is kept for 1.x compatibility and
 * will be unified in 2.0.
 */
export type RewriteAction =
  | { type: 'insert_rows'; sheet?: string; start: number; count: number }
  | { type: 'delete_rows'; sheet?: string; start: number; count: number }
  | { type: 'insert_cols'; sheet?: string; start: number; count: number }
  | { type: 'delete_cols'; sheet?: string; start: number; count: number };

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const REF_ERROR: ASTNode = { type: 'error', value: '#REF!' };

/**
 * Whether a reference on `sheet` (its effective sheet; `undefined` when
 * unqualified) is on the action's sheet. An action without `sheet` matches
 * unqualified references only; one with `sheet` matches references qualified
 * with that name, ignoring case, as in Excel (`sheet1!A5` is on `Sheet1`).
 */
function sheetMatches(sheet: string | undefined, action: RewriteAction): boolean {
  if (sheet === undefined || action.sheet === undefined) return sheet === action.sheet;
  return sheet.toLowerCase() === action.sheet.toLowerCase();
}

/** Whether the action moves rows (`true`) or columns (`false`). */
function isRowAction(action: RewriteAction): boolean {
  return action.type === 'insert_rows' || action.type === 'delete_rows';
}

/**
 * Move the span `a..b` (row numbers or 0-based column indexes; either may be
 * the larger) along the action's axis.
 *
 * Delete: a span wholly inside the deleted block is gone; otherwise each end
 * inside the block snaps to the block's edge and ends past it move back by
 * `count`. Insert: ends at or past `start` move forward by `count`; a high
 * end past `max` is clamped to `max`, a low end past `max` loses the span.
 *
 * @returns The new values in the order `[a, b]` (each value goes back to the
 *   endpoint that held it), or `null` when the span becomes `#REF!`. A single
 *   cell is the span `v..v`, so it becomes `#REF!` when deleted or pushed off.
 */
function shiftSpan(
  a: number,
  b: number,
  action: RewriteAction,
  max: number,
): [number, number] | null {
  const isDelete = action.type === 'delete_rows' || action.type === 'delete_cols';
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  const span = isDelete
    ? deleteFromSpan(lo, hi, action.start, action.count)
    : insertIntoSpan(lo, hi, action.start, action.count, max);
  if (span === null) return null;
  return a <= b ? span : [span[1], span[0]];
}

/** {@link shiftSpan} for a delete of `start..start + count - 1`, on an ordered span `lo <= hi`. */
function deleteFromSpan(
  lo: number,
  hi: number,
  start: number,
  count: number,
): [number, number] | null {
  const end = start + count - 1;
  if (lo >= start && hi <= end) return null;
  const newLo = lo < start ? lo : lo > end ? lo - count : start;
  const newHi = hi > end ? hi - count : hi < start ? hi : start - 1;
  return [newLo, newHi];
}

/** {@link shiftSpan} for an insert of `count` at `start`, on an ordered span `lo <= hi`. */
function insertIntoSpan(
  lo: number,
  hi: number,
  start: number,
  count: number,
  max: number,
): [number, number] | null {
  const newLo = lo >= start ? lo + count : lo;
  if (newLo > max) return null;
  return [newLo, Math.min(hi >= start ? hi + count : hi, max)];
}

// ---------------------------------------------------------------------------
// AST transformation
// ---------------------------------------------------------------------------

/** Rebuild an AST node with its references adjusted per the action (unchanged leaves are shared). */
function transformNode(node: ASTNode, action: RewriteAction): ASTNode {
  return mapRefs(node, {
    cellRef: (ref) => transformCellRef(ref, action),
    range: (range) => transformRange(range, action),
    columnRange: (range) => transformColumnRange(range, action),
    rowRange: (range) => transformRowRange(range, action),
  });
}

/** Transform a single cell reference; a 3-D reference passes through. */
function transformCellRef(ref: CellRefNode, action: RewriteAction): ASTNode {
  if (ref.endSheet !== undefined || !sheetMatches(ref.sheet, action)) {
    return ref;
  }
  if (isRowAction(action)) {
    const span = shiftSpan(ref.row, ref.row, action, MAX_ROW);
    if (span === null) return REF_ERROR;
    return span[0] === ref.row ? ref : { ...ref, row: span[0] };
  }
  const colIdx = columnIndexOrNull(ref.col);
  if (colIdx === null) return REF_ERROR;
  const span = shiftSpan(colIdx, colIdx, action, MAX_COL_INDEX);
  if (span === null) return REF_ERROR;
  return span[0] === colIdx ? ref : { ...ref, col: columnToLetter(span[0]) };
}

/**
 * Transform a cell range as one span: it shrinks, clamps or becomes a single
 * `#REF!`, never a range with an error endpoint. The effective sheet
 * (`start.sheet ?? end.sheet`) is matched once and never copied onto `end`,
 * so `Sheet1!A1:B10` serializes back without a second prefix. A 3-D range
 * (`start.endSheet` set) passes through.
 */
function transformRange(node: RangeNode, action: RewriteAction): ASTNode {
  const { start, end } = node;
  if (start.endSheet !== undefined || !sheetMatches(start.sheet ?? end.sheet, action)) {
    return node;
  }
  if (isRowAction(action)) {
    const span = shiftSpan(start.row, end.row, action, MAX_ROW);
    if (span === null) return REF_ERROR;
    return { type: 'range', start: { ...start, row: span[0] }, end: { ...end, row: span[1] } };
  }
  const startIdx = columnIndexOrNull(start.col);
  const endIdx = columnIndexOrNull(end.col);
  if (startIdx === null || endIdx === null) return REF_ERROR;
  const span = shiftSpan(startIdx, endIdx, action, MAX_COL_INDEX);
  if (span === null) return REF_ERROR;
  return {
    type: 'range',
    start: { ...start, col: columnToLetter(span[0]) },
    end: { ...end, col: columnToLetter(span[1]) },
  };
}

/** Transform a whole-column range (`A:C`) on a column action; a row action or 3-D span leaves it as is. */
function transformColumnRange(node: ColumnRangeNode, action: RewriteAction): ASTNode {
  if (isRowAction(action) || node.endSheet !== undefined || !sheetMatches(node.sheet, action)) {
    return node;
  }
  const startIdx = columnIndexOrNull(node.startCol);
  const endIdx = columnIndexOrNull(node.endCol);
  if (startIdx === null || endIdx === null) return REF_ERROR;
  const span = shiftSpan(startIdx, endIdx, action, MAX_COL_INDEX);
  if (span === null) return REF_ERROR;
  return { ...node, startCol: columnToLetter(span[0]), endCol: columnToLetter(span[1]) };
}

/** Transform a whole-row range (`1:3`) on a row action; a column action or 3-D span leaves it as is. */
function transformRowRange(node: RowRangeNode, action: RewriteAction): ASTNode {
  if (!isRowAction(action) || node.endSheet !== undefined || !sheetMatches(node.sheet, action)) {
    return node;
  }
  const span = shiftSpan(node.startRow, node.endRow, action, MAX_ROW);
  if (span === null) return REF_ERROR;
  return { ...node, startRow: span[0], endRow: span[1] };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Rewrite references in a formula after a row/column insert or delete.
 *
 * @param formula - Formula string (with or without leading `=`).
 * @param action - The insert/delete action to apply. `start` is a 1-based row
 *   number for `insert_rows`/`delete_rows` and a 0-based column index for
 *   `insert_cols`/`delete_cols` (unified in 2.0). An insert shifts references
 *   at index `>= start`; a delete removes `[start, start + count)`, turning a
 *   cell ref inside it into `#REF!`, shrinking a range that overlaps it, and
 *   shifting the rest back. An `undefined` `sheet` touches unqualified
 *   references only. 3-D and structured references are never changed.
 * @returns The rewritten formula string, or `formula` unchanged when it does
 *   not parse or rewriting it exhausts the call stack.
 */
export function rewriteFormula(formula: string, action: RewriteAction): string {
  const { ast, errors } = parseFormula(formula);
  if (ast === null || errors.length > 0) {
    return formula; // Return unchanged on parse error
  }
  try {
    return serializeFormula(transformNode(ast, action));
  } catch (err) {
    // Backstop: operator chains are walked with loops, but the remaining
    // recursion (parentheses, functions, arrays: up to the parser's 256
    // levels) can still exhaust a small stack. Return the input unchanged,
    // as for a parse error.
    if (err instanceof RangeError) return formula;
    throw err;
  }
}
