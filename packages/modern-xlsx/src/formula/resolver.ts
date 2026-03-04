/**
 * Cell reference resolver -- looks up cell values from workbook data.
 *
 * Provides an abstraction layer between the formula AST and the actual
 * workbook storage.  The {@link EvalContext} interface decouples resolution
 * from the `Workbook` class so tests can supply lightweight mock data.
 *
 * @module formula/resolver
 */

import type { ASTNode, CellRefNode, RangeNode } from './parser.js';

// ---------------------------------------------------------------------------
// Shared types
// ---------------------------------------------------------------------------

/** A resolved cell value: number, string, boolean, or null (empty). */
export type CellValue = number | string | boolean | null;

/**
 * A formula function receives raw AST arguments (not pre-evaluated) so that
 * functions like IF can short-circuit.  The third parameter is the evaluator
 * callback for evaluating individual arguments on demand.
 */
export type FormulaFunction = (
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
) => CellValue;

/**
 * Evaluation context: provides cell value lookup without coupling to the
 * Workbook class.  This abstraction enables testing with mock data and
 * future extensibility (e.g. cross-workbook references).
 */
export interface EvalContext {
  /** Get cell value by sheet name, 0-based column, and 1-based row. */
  getCell(sheet: string, col: number, row: number): CellValue;
  /** The name of the current sheet (used when CellRefNode has no sheet prefix). */
  currentSheet: string;
  /** Optional function registry for formula evaluation. */
  functions?: Map<string, FormulaFunction>;
}

// ---------------------------------------------------------------------------
// Resolution helpers
// ---------------------------------------------------------------------------

/** Convert column letter(s) to 0-based index.  A=0, B=1, Z=25, AA=26. */
function letterToColumnIndex(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result - 1;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Resolve a single cell reference to its value.
 *
 * Uses the optional `sheet` qualifier on the node, falling back to
 * `ctx.currentSheet` when absent.
 */
export function resolveRef(ref: CellRefNode, ctx: EvalContext): CellValue {
  const sheet = ref.sheet ?? ctx.currentSheet;
  const col = letterToColumnIndex(ref.col);
  return ctx.getCell(sheet, col, ref.row);
}

/**
 * Resolve a range reference to a 2D array of values (rows x cols).
 *
 * The range is normalised so that `start` is always the top-left corner
 * regardless of the order the user specified.
 */
export function resolveRange(range: RangeNode, ctx: EvalContext): CellValue[][] {
  const sheet = range.start.sheet ?? ctx.currentSheet;
  const startCol = letterToColumnIndex(range.start.col);
  const endCol = letterToColumnIndex(range.end.col);
  const startRow = range.start.row;
  const endRow = range.end.row;

  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  const minRow = Math.min(startRow, endRow);
  const maxRow = Math.max(startRow, endRow);

  const result: CellValue[][] = [];
  for (let r = minRow; r <= maxRow; r++) {
    const row: CellValue[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      row.push(ctx.getCell(sheet, c, r));
    }
    result.push(row);
  }
  return result;
}
