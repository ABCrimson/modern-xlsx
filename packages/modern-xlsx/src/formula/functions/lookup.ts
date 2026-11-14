/**
 * Lookup and reference formula functions.
 *
 * Registers the built-in Excel functions: VLOOKUP, HLOOKUP, INDEX, MATCH,
 * CHOOSE, ROW, COLUMN, ROWS, COLUMNS.
 *
 * Matching rules:
 *   - Exact lookups (VLOOKUP/HLOOKUP with range_lookup FALSE, MATCH 0) are
 *     type-strict and case-insensitive: a number matches only a number, a
 *     boolean only a boolean, text only text (folded one code unit at a
 *     time, like the criteria functions).  A blank cell never matches.  A
 *     text lookup value containing `*`, `?` or `~` is an Excel wildcard
 *     pattern (see ./glob.js).
 *   - Approximate lookups (range_lookup TRUE, MATCH 1 and -1) compare only
 *     values of the lookup value's type and skip all others; no coercion
 *     happens, so text "2" never matches the number 2.
 *   - MATCH's match_type is truncated and reduced to its sign.
 *   - Lookup tables and arrays may be references or array constants.
 *   - INDEX with a zero row or column selects a whole column or row; because
 *     a cell value is a scalar, that is reduced to a single cell: the one
 *     cell of a one-cell line, else the cell on `ctx.currentCell`'s row (or
 *     column), else `#VALUE!`.
 *   - ROW() and COLUMN() without arguments read `ctx.currentCell`.  Given a
 *     single-area reference (cell, range, whole column or row, intersection),
 *     ROW and COLUMN return its top-left and ROWS and COLUMNS its declared
 *     size, without reading a cell; a 3-D or structured reference is `#REF!`.
 *
 * @module formula/functions/lookup
 */

import { toNumber } from '../coerce.js';
import { compareText } from '../collation.js';
import { isError } from '../errors.js';
import type { ASTNode } from '../parser.js';
import type { CellValue, EvalContext, EvaluateFn, FormulaFunction } from '../resolver.js';
import { areaExtent, argToMatrix, exceedsReadCap, isReferenceNode } from '../resolver.js';
import { compileGlob, hasWildcard } from './glob.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Case-insensitive text equality, folding one code unit at a time. */
function textEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const x = a.charAt(i);
    const y = b.charAt(i);
    if (x !== y && x.toLowerCase() !== y.toLowerCase()) return false;
  }
  return true;
}

/** Build the exact-match predicate for a lookup value (blank cells never match). */
function exactMatcher(target: CellValue): (value: CellValue) => boolean {
  if (typeof target === 'string') {
    if (hasWildcard(target)) {
      const glob = compileGlob(target);
      return (value) => typeof value === 'string' && glob(value);
    }
    return (value) => typeof value === 'string' && textEqual(value, target);
  }
  if (typeof target === 'number' || typeof target === 'boolean') {
    return (value) => value === target;
  }
  return () => false;
}

/**
 * Compare two values of the same type for approximate matching: negative if
 * a < b, 0 if equal, positive if a > b.  Callers guarantee matching types.
 */
function compareSameType(a: CellValue, b: CellValue): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return compareText(String(a).toLowerCase(), String(b).toLowerCase());
}

/** `true` when `value` has the same type as the lookup value (blank has none). */
function sameType(value: CellValue, target: CellValue): boolean {
  return value !== null && target !== null && typeof value === typeof target;
}

/** Resolve the optional `approx` argument (defaults to true). */
function resolveApprox(
  args: ASTNode[],
  argIndex: number,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): boolean | string {
  if (args.length <= argIndex) return true;
  const arg = args[argIndex];
  if (!arg) return true;
  const aVal = evaluate(arg, ctx);
  if (isError(aVal)) return aVal;
  if (typeof aVal === 'boolean') return aVal;
  if (typeof aVal === 'number') return aVal !== 0;
  return true;
}

/**
 * Find the index of the best approximate match (largest value <= target)
 * in a sorted ascending array, looking only at values of the target's type.
 * Returns -1 if no match found.
 */
function approxSearchAsc(values: CellValue[], target: CellValue): number {
  let bestIdx = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i] ?? null;
    if (!sameType(v, target)) continue;
    if (compareSameType(v, target) <= 0) bestIdx = i;
    else break;
  }
  return bestIdx;
}

/**
 * Find the index of the smallest value >= target in an array sorted
 * descending, looking only at values of the target's type.  Returns -1 if
 * none.
 */
function approxSearchDesc(values: CellValue[], target: CellValue): number {
  let bestIdx = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i] ?? null;
    if (!sameType(v, target)) continue;
    if (compareSameType(v, target) >= 0) bestIdx = i;
    else break;
  }
  return bestIdx;
}

/** Find the index of an exact match in a flat array.  Returns -1 if not found. */
function exactSearch(values: CellValue[], target: CellValue): number {
  const matches = exactMatcher(target);
  for (let i = 0; i < values.length; i++) {
    if (matches(values[i] ?? null)) return i;
  }
  return -1;
}

/**
 * Flatten a one-row or one-column matrix into a vector, keeping positions.
 * Returns `undefined` for a two-dimensional matrix.
 */
function matrixVector(matrix: CellValue[][]): CellValue[] | undefined {
  if (matrix.length === 1) return matrix[0] ?? [];
  if ((matrix[0]?.length ?? 0) === 1) return matrix.map((row) => row[0] ?? null);
  return undefined;
}

/** A lookup table: its declared size and the values read from it. */
interface LookupTable {
  rows: number;
  cols: number;
  values: CellValue[][];
}

/**
 * Read a lookup table or array.  A single-area reference (cell, range, whole
 * column or row) is sized by its declaration, because its matrix is clamped
 * to the sheet's populated bounds, which must not turn an in-range index into
 * `#REF!`.  Its error (`#REF!`, `#NULL!`) is returned as a string, and so is
 * `#NUM!` when the context has no bounds for the sheet and the reference
 * spans more than 1,048,576 cells.  Any other argument (array constant,
 * union, scalar) is read as a matrix and sized by it.
 */
function readTable(node: ASTNode, ctx: EvalContext, evaluate: EvaluateFn): LookupTable | string {
  const ext = areaExtent(node, ctx);
  if (typeof ext === 'string') return ext;
  if (ext !== undefined && exceedsReadCap(ext, ctx)) return '#NUM!';
  const values = argToMatrix(node, ctx, evaluate);
  if (ext === undefined) return { rows: values.length, cols: values[0]?.length ?? 0, values };
  return { rows: ext.rows, cols: ext.cols, values };
}

// ---------------------------------------------------------------------------
// VLOOKUP / HLOOKUP implementations
// ---------------------------------------------------------------------------

function vlookupImpl(args: ASTNode[], ctx: EvalContext, evaluate: EvaluateFn): CellValue {
  if (args.length < 3) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  const arg2 = args[2];
  if (!arg0 || !arg1 || !arg2) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  const table = readTable(arg1, ctx, evaluate);
  if (typeof table === 'string') return table;
  const matrix = table.values;
  const colIdxVal = toNumber(evaluate(arg2, ctx));
  if (typeof colIdxVal === 'string') return colIdxVal;
  const colIdx = Math.floor(colIdxVal);
  if (colIdx < 1) return '#VALUE!';
  if (colIdx > table.cols) return '#REF!';

  const approx = resolveApprox(args, 3, ctx, evaluate);
  if (typeof approx === 'string') return approx;

  // Extract first column for searching
  const firstCol: CellValue[] = matrix.map((row) => row[0] ?? null);

  const idx = approx ? approxSearchAsc(firstCol, lookupVal) : exactSearch(firstCol, lookupVal);
  return idx === -1 ? '#N/A' : (matrix[idx]?.[colIdx - 1] ?? null);
}

function hlookupImpl(args: ASTNode[], ctx: EvalContext, evaluate: EvaluateFn): CellValue {
  if (args.length < 3) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  const arg2 = args[2];
  if (!arg0 || !arg1 || !arg2) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  const table = readTable(arg1, ctx, evaluate);
  if (typeof table === 'string') return table;
  const matrix = table.values;
  const rowIdxVal = toNumber(evaluate(arg2, ctx));
  if (typeof rowIdxVal === 'string') return rowIdxVal;
  const rowIdx = Math.floor(rowIdxVal);
  if (rowIdx < 1) return '#VALUE!';
  if (rowIdx > table.rows) return '#REF!';

  const approx = resolveApprox(args, 3, ctx, evaluate);
  if (typeof approx === 'string') return approx;

  const firstRow = matrix[0];
  if (!firstRow) return '#N/A';

  const idx = approx ? approxSearchAsc(firstRow, lookupVal) : exactSearch(firstRow, lookupVal);
  return idx === -1 ? '#N/A' : (matrix[rowIdx - 1]?.[idx] ?? null);
}

// ---------------------------------------------------------------------------
// INDEX / MATCH implementations
// ---------------------------------------------------------------------------

/** The area INDEX reads: its size and a cell reader (1-based row and column within it). */
interface IndexArea {
  rows: number;
  cols: number;
  /** 0-based column and 1-based row of the top-left cell; set for a reference only. */
  origin?: { col: number; row: number };
  at(row: number, col: number): CellValue;
}

function indexImpl(args: ASTNode[], ctx: EvalContext, evaluate: EvaluateFn): CellValue {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  // INDEX reads one cell, so a single-area reference is validated against its
  // declared size and read directly with getCell, never resolved or clamped.
  // A union reads its first area (Excel's default area_num 1).  Any other
  // first argument (array constant, scalar) is read as a matrix.
  let first: ASTNode = arg0;
  while (first.type === 'binary_op' && first.op === ',') first = first.left;
  const extent = areaExtent(first, ctx);
  if (typeof extent === 'string') return extent;
  let area: IndexArea;
  if (extent !== undefined) {
    area = {
      rows: extent.rows,
      cols: extent.cols,
      origin: { col: extent.col, row: extent.row },
      at: (row, col) => ctx.getCell(extent.sheet, extent.col + col - 1, extent.row + row - 1),
    };
  } else {
    const matrix = argToMatrix(first, ctx, evaluate);
    area = {
      rows: matrix.length,
      cols: matrix[0]?.length ?? 0,
      at: (row, col) => matrix[row - 1]?.[col - 1] ?? null,
    };
  }

  const pos = indexPosition(arg1, args[2], area, ctx, evaluate);
  if (typeof pos === 'string') return pos;
  const [rowIdx, colIdx] = pos;

  if (rowIdx > area.rows || colIdx > area.cols) return '#REF!';
  if (rowIdx === 0 || colIdx === 0) return indexWholeLine(area, rowIdx, colIdx, ctx);
  return area.at(rowIdx, colIdx);
}

/**
 * Evaluate INDEX's `row_num` and optional `column_num`.  With two arguments
 * on a one-row area the index is the column; with the column omitted on a
 * multi-row area it is 1 for a single column and 0 (the whole row) otherwise.
 * Returns an error string for an error argument or a negative index; 0 means
 * the whole row or column.
 */
function indexPosition(
  rowArg: ASTNode,
  colArg: ASTNode | undefined,
  area: IndexArea,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): [row: number, col: number] | string {
  const first = toNumber(evaluate(rowArg, ctx));
  if (typeof first === 'string') return first;
  let second: number | string = 0;
  if (colArg) second = toNumber(evaluate(colArg, ctx));
  else if (area.rows > 1 && area.cols === 1) second = 1;
  if (typeof second === 'string') return second;
  const a = Math.trunc(first);
  const b = Math.trunc(second);
  if (a < 0 || b < 0) return '#VALUE!';
  // Two arguments on a one-row area: the index is the column.
  if (!colArg && area.rows === 1) return [1, a];
  return [a, b];
}

/**
 * INDEX with a zero row or column: a whole column or row.  A cell value is a
 * scalar, so the line is reduced to one cell: the line's only cell, else the
 * cell on the formula's own row (whole column) or column (whole row) when
 * the area is a reference and `ctx.currentCell` is set, else `#VALUE!`.
 */
function indexWholeLine(
  area: IndexArea,
  rowIdx: number,
  colIdx: number,
  ctx: EvalContext,
): CellValue {
  if (rowIdx === 0 && colIdx === 0) {
    return area.rows === 1 && area.cols === 1 ? area.at(1, 1) : '#VALUE!';
  }
  const wholeColumn = rowIdx === 0;
  const length = wholeColumn ? area.rows : area.cols;
  if (length === 1) return wholeColumn ? area.at(1, colIdx) : area.at(rowIdx, 1);
  if (!area.origin || !ctx.currentCell) return '#VALUE!';
  // Position of the formula within the area, along the selected line.
  const at = wholeColumn
    ? ctx.currentCell.row - area.origin.row + 1
    : ctx.currentCell.col - area.origin.col + 1;
  if (at < 1 || at > length) return '#VALUE!';
  return wholeColumn ? area.at(at, colIdx) : area.at(rowIdx, at);
}

function matchImpl(args: ASTNode[], ctx: EvalContext, evaluate: EvaluateFn): CellValue {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  // Excel's lookup_array is one row or one column; a 2-D area gives #N/A.
  const table = readTable(arg1, ctx, evaluate);
  if (typeof table === 'string') return table;
  const values = matrixVector(table.values);
  if (!values) return '#N/A';

  let matchType = 1;
  if (args.length >= 3) {
    const arg2 = args[2];
    if (!arg2) return '#VALUE!';
    const mtVal = toNumber(evaluate(arg2, ctx));
    if (typeof mtVal === 'string') return mtVal;
    // Excel documents only -1, 0 and 1.  Any other number follows Gnumeric's
    // rule: truncate toward zero, then take the sign (0.5 is exact, 2 is 1).
    matchType = Math.sign(Math.trunc(mtVal)) || 0;
  }

  return matchByType(values, lookupVal, matchType);
}

/** Dispatch MATCH logic by match type (-1, 0 or 1). */
function matchByType(values: CellValue[], lookupVal: CellValue, matchType: number): CellValue {
  let idx: number;
  if (matchType === 0) idx = exactSearch(values, lookupVal);
  else if (matchType > 0) idx = approxSearchAsc(values, lookupVal);
  else idx = approxSearchDesc(values, lookupVal);
  return idx === -1 ? '#N/A' : idx + 1;
}

/**
 * ROW, COLUMN, ROWS or COLUMNS of an argument that is neither an array
 * constant nor a single-area reference (those are sized without reading a
 * cell): a union of references is `#VALUE!` and is never dereferenced, so its
 * cells' errors cannot leak through. Any other argument gives its own error,
 * as in Excel (`COLUMN(XFE1)` is `#NAME?`, since an out-of-grid `XFE1` is a
 * name), otherwise `#VALUE!`.
 */
function nonReferenceArg(arg: ASTNode, ctx: EvalContext, evaluate: EvaluateFn): CellValue {
  if (isReferenceNode(arg)) return '#VALUE!';
  const value = evaluate(arg, ctx);
  return isError(value) ? value : '#VALUE!';
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

/**
 * Register all lookup and reference functions into the given registry.
 */
export function registerLookupFunctions(registry: Map<string, FormulaFunction>): void {
  registry.set('VLOOKUP', vlookupImpl);
  registry.set('HLOOKUP', hlookupImpl);
  registry.set('INDEX', indexImpl);
  registry.set('MATCH', matchImpl);

  // ---- CHOOSE ------------------------------------------------------------
  registry.set('CHOOSE', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const idxVal = toNumber(evaluate(arg0, ctx));
    if (typeof idxVal === 'string') return idxVal;
    const idx = Math.floor(idxVal);
    if (idx < 1 || idx >= args.length) return '#VALUE!';
    const argN = args[idx];
    if (!argN) return '#VALUE!';
    // An empty argument (CHOOSE(2,1,,3)) selects 0, as in Excel.
    if (argN.type === 'missing') return 0;
    return evaluate(argN, ctx);
  });

  // ---- ROW ---------------------------------------------------------------
  registry.set('ROW', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return ctx.currentCell?.row ?? '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    const ext = areaExtent(arg, ctx);
    if (ext === undefined) return nonReferenceArg(arg, ctx, evaluate);
    if (typeof ext === 'string') return ext;
    return ext.row;
  });

  // ---- COLUMN ------------------------------------------------------------
  registry.set('COLUMN', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return ctx.currentCell ? ctx.currentCell.col + 1 : '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    const ext = areaExtent(arg, ctx);
    if (ext === undefined) return nonReferenceArg(arg, ctx, evaluate);
    if (typeof ext === 'string') return ext;
    return ext.col + 1;
  });

  // ---- ROWS --------------------------------------------------------------
  registry.set('ROWS', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    if (arg.type === 'array') return arg.rows.length === 0 ? '#VALUE!' : arg.rows.length;
    const ext = areaExtent(arg, ctx);
    if (ext === undefined) return nonReferenceArg(arg, ctx, evaluate);
    if (typeof ext === 'string') return ext;
    return ext.rows;
  });

  // ---- COLUMNS -----------------------------------------------------------
  registry.set('COLUMNS', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    if (arg.type === 'array') {
      return arg.rows.length === 0 ? '#VALUE!' : (arg.rows[0]?.length ?? 0);
    }
    const ext = areaExtent(arg, ctx);
    if (ext === undefined) return nonReferenceArg(arg, ctx, evaluate);
    if (typeof ext === 'string') return ext;
    return ext.cols;
  });
}
