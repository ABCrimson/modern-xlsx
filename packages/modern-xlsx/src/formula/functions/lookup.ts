/**
 * Lookup and reference formula functions.
 *
 * Registers 9 built-in Excel functions: VLOOKUP, HLOOKUP, INDEX, MATCH,
 * CHOOSE, ROW, COLUMN, ROWS, COLUMNS.
 *
 * @module formula/functions/lookup
 */

import type { ASTNode } from '../parser.js';
import type { CellValue, EvalContext, FormulaFunction } from '../resolver.js';
import { resolveRange } from '../resolver.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function isError(val: CellValue): val is string {
  return typeof val === 'string' && val.length > 0 && val.charAt(0) === '#';
}

function toNumber(val: CellValue): number | string {
  if (val === null) return 0;
  switch (typeof val) {
    case 'number':
      return val;
    case 'boolean':
      return val ? 1 : 0;
    case 'string': {
      if (isError(val)) return val;
      const n = Number(val);
      return Number.isNaN(n) ? '#VALUE!' : n;
    }
  }
}

/** Convert column letter(s) to 0-based index. A=0, B=1, Z=25, AA=26. */
function letterToCol(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result - 1;
}

/**
 * Resolve an argument to a 2D matrix.  Range nodes use resolveRange;
 * scalars wrap as 1x1.
 */
function resolveMatrix(
  arg: ASTNode,
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue[][] {
  if (arg.type === 'range') {
    return resolveRange(arg, ctx);
  }
  return [[evaluate(arg, ctx)]];
}

/**
 * Compare two values for exact matching (case-insensitive string comparison).
 */
function valuesEqual(a: CellValue, b: CellValue): boolean {
  if (typeof a === 'number' && typeof b === 'number') return a === b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b;
  if (a === null && b === null) return true;
  return String(a ?? '').toLowerCase() === String(b ?? '').toLowerCase();
}

/**
 * Compare two values for approximate matching.
 * Returns negative if a < b, 0 if equal, positive if a > b.
 */
function compareForMatch(a: CellValue, b: CellValue): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  }
  const na = toNumber(a);
  const nb = toNumber(b);
  if (typeof na === 'number' && typeof nb === 'number') return na - nb;
  return String(a ?? '')
    .toLowerCase()
    .localeCompare(String(b ?? '').toLowerCase());
}

/** Resolve the optional `approx` argument (defaults to true). */
function resolveApprox(
  args: ASTNode[],
  argIndex: number,
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
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
 * in a sorted ascending array.  Returns -1 if no match found.
 */
function approxSearchAsc(values: CellValue[], target: CellValue): number {
  let bestIdx = -1;
  for (let i = 0; i < values.length; i++) {
    const cmp = compareForMatch(values[i] ?? null, target);
    if (cmp <= 0) bestIdx = i;
    else break;
  }
  return bestIdx;
}

/** Find the index of exact match in a flat array.  Returns -1 if not found. */
function exactSearch(values: CellValue[], target: CellValue): number {
  for (let i = 0; i < values.length; i++) {
    if (valuesEqual(values[i] ?? null, target)) return i;
  }
  return -1;
}

/** Flatten a 2D matrix into a 1D array. */
function flattenMatrix(matrix: CellValue[][]): CellValue[] {
  return matrix.flat();
}

// ---------------------------------------------------------------------------
// VLOOKUP / HLOOKUP implementations
// ---------------------------------------------------------------------------

function vlookupImpl(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue {
  if (args.length < 3) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  const arg2 = args[2];
  if (!arg0 || !arg1 || !arg2) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  const matrix = resolveMatrix(arg1, ctx, evaluate);
  const colIdxVal = toNumber(evaluate(arg2, ctx));
  if (typeof colIdxVal === 'string') return colIdxVal;
  const colIdx = Math.floor(colIdxVal);
  if (colIdx < 1) return '#VALUE!';
  if (colIdx > (matrix[0]?.length ?? 0)) return '#REF!';

  const approx = resolveApprox(args, 3, ctx, evaluate);
  if (typeof approx === 'string') return approx;

  // Extract first column for searching
  const firstCol: CellValue[] = matrix.map((row) => row[0] ?? null);

  if (approx) {
    const idx = approxSearchAsc(firstCol, lookupVal);
    return idx === -1 ? '#N/A' : (matrix[idx]?.[colIdx - 1] ?? null);
  }
  const idx = exactSearch(firstCol, lookupVal);
  return idx === -1 ? '#N/A' : (matrix[idx]?.[colIdx - 1] ?? null);
}

function hlookupImpl(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue {
  if (args.length < 3) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  const arg2 = args[2];
  if (!arg0 || !arg1 || !arg2) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  const matrix = resolveMatrix(arg1, ctx, evaluate);
  const rowIdxVal = toNumber(evaluate(arg2, ctx));
  if (typeof rowIdxVal === 'string') return rowIdxVal;
  const rowIdx = Math.floor(rowIdxVal);
  if (rowIdx < 1) return '#VALUE!';
  if (rowIdx > matrix.length) return '#REF!';

  const approx = resolveApprox(args, 3, ctx, evaluate);
  if (typeof approx === 'string') return approx;

  const firstRow = matrix[0];
  if (!firstRow) return '#N/A';

  if (approx) {
    const idx = approxSearchAsc(firstRow, lookupVal);
    return idx === -1 ? '#N/A' : (matrix[rowIdx - 1]?.[idx] ?? null);
  }
  const idx = exactSearch(firstRow, lookupVal);
  return idx === -1 ? '#N/A' : (matrix[rowIdx - 1]?.[idx] ?? null);
}

// ---------------------------------------------------------------------------
// INDEX / MATCH implementations
// ---------------------------------------------------------------------------

function indexImpl(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const matrix = resolveMatrix(arg0, ctx, evaluate);
  const rowVal = toNumber(evaluate(arg1, ctx));
  if (typeof rowVal === 'string') return rowVal;
  const rowIdx = Math.floor(rowVal);

  let colIdx = 1;
  if (args.length >= 3) {
    const arg2 = args[2];
    if (!arg2) return '#VALUE!';
    const cVal = toNumber(evaluate(arg2, ctx));
    if (typeof cVal === 'string') return cVal;
    colIdx = Math.floor(cVal);
  }

  if (rowIdx < 1 || colIdx < 1) return '#VALUE!';
  if (rowIdx > matrix.length) return '#REF!';
  const row = matrix[rowIdx - 1];
  if (!row || colIdx > row.length) return '#REF!';
  return row[colIdx - 1] ?? null;
}

function matchImpl(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';

  const lookupVal = evaluate(arg0, ctx);
  if (isError(lookupVal)) return lookupVal;

  const values = flattenMatrix(resolveMatrix(arg1, ctx, evaluate));

  let matchType = 1;
  if (args.length >= 3) {
    const arg2 = args[2];
    if (!arg2) return '#VALUE!';
    const mtVal = toNumber(evaluate(arg2, ctx));
    if (typeof mtVal === 'string') return mtVal;
    matchType = mtVal;
  }

  return matchByType(values, lookupVal, matchType);
}

/** Dispatch MATCH logic by match type. */
function matchByType(values: CellValue[], lookupVal: CellValue, matchType: number): CellValue {
  switch (matchType) {
    case 0: {
      const idx = exactSearch(values, lookupVal);
      return idx === -1 ? '#N/A' : idx + 1;
    }
    case 1: {
      const idx = approxSearchAsc(values, lookupVal);
      return idx === -1 ? '#N/A' : idx + 1;
    }
    case -1: {
      // Smallest value >= lookupVal (data sorted descending)
      let bestIdx = -1;
      for (let i = 0; i < values.length; i++) {
        const v = values[i];
        const cmp = compareForMatch(v ?? null, lookupVal);
        if (cmp >= 0) bestIdx = i;
        else break;
      }
      return bestIdx === -1 ? '#N/A' : bestIdx + 1;
    }
    default:
      return '#N/A';
  }
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
    return evaluate(argN, ctx);
  });

  // ---- ROW ---------------------------------------------------------------
  registry.set('ROW', (args): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    switch (arg.type) {
      case 'cell_ref':
        return arg.row;
      case 'range':
        return arg.start.row;
      default:
        return '#VALUE!';
    }
  });

  // ---- COLUMN ------------------------------------------------------------
  registry.set('COLUMN', (args): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    switch (arg.type) {
      case 'cell_ref':
        return letterToCol(arg.col) + 1;
      case 'range':
        return letterToCol(arg.start.col) + 1;
      default:
        return '#VALUE!';
    }
  });

  // ---- ROWS --------------------------------------------------------------
  registry.set('ROWS', (args): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    switch (arg.type) {
      case 'range':
        return Math.abs(arg.end.row - arg.start.row) + 1;
      case 'cell_ref':
        return 1;
      default:
        return '#VALUE!';
    }
  });

  // ---- COLUMNS -----------------------------------------------------------
  registry.set('COLUMNS', (args): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg = args[0];
    if (!arg) return '#VALUE!';
    switch (arg.type) {
      case 'range':
        return Math.abs(letterToCol(arg.end.col) - letterToCol(arg.start.col)) + 1;
      case 'cell_ref':
        return 1;
      default:
        return '#VALUE!';
    }
  });
}
