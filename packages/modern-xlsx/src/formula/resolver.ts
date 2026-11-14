/**
 * Cell reference resolver -- looks up cell values from workbook data.
 *
 * Provides an abstraction layer between the formula AST and the actual
 * workbook storage.  The {@link EvalContext} interface decouples resolution
 * from the `Workbook` class so tests can supply lightweight mock data.
 *
 * Besides the public {@link resolveRef} and {@link resolveRange}, it holds the
 * internal area model the evaluator and the built-in functions share: every
 * reference node (cell, range, whole column, whole row, and the `' '`
 * intersection of references) resolves to one rectangle, and
 * {@link argToMatrix} / {@link argValues} read a function argument, unions
 * included, as a matrix or as a row-major stream of cells.
 *
 * @module formula/resolver
 */

import { columnIndexOrNull } from '../format/cell-ref.js';
import { INVALID_ARGUMENT, ModernXlsxError } from '../model/errors.js';
import type { DateSystem } from '../model/types.js';
import { MAX_COL_INDEX, MAX_ROW } from './grid.js';
import type {
  ASTNode,
  BinaryOpNode,
  CellRefNode,
  ColumnRangeNode,
  RangeNode,
  RowRangeNode,
} from './parser.js';

/**
 * Largest range resolved when the context supplies no {@link EvalContext.getBounds}.
 *
 * @internal
 */
export const MAX_UNBOUNDED_CELLS = 1_048_576;

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
  /**
   * Optional populated extent of a sheet, in {@link EvalContext.getCell}'s
   * coordinates: `maxRow` is the last populated 1-based row and `maxCol` the
   * last populated 0-based column.  Return `{ maxRow: 0, maxCol: -1 }` for an
   * empty sheet.
   *
   * Ranges are clamped to these bounds (the declared top-left is kept, only
   * the bottom-right shrinks), so whole-column or whole-sheet ranges such as
   * `SUM(A1:XFD1048576)` read only populated cells.  Without bounds (method
   * absent or returning `undefined`), a range of more than 1,048,576 cells
   * evaluates to `#NUM!`.
   */
  getBounds?(sheet: string): { maxRow: number; maxCol: number } | undefined;
  /**
   * Optional position of the cell that holds the formula, in
   * {@link EvalContext.getCell}'s coordinates: `col` is 0-based and `row`
   * 1-based.
   *
   * A multi-cell reference in a scalar position (`=A1:A10*2`) reads the cell
   * of a single-column reference on this row, or of a single-row reference in
   * this column (Excel's implicit intersection).  Without it, such a reference
   * evaluates to `#VALUE!`.
   */
  currentCell?: { col: number; row: number };
  /**
   * Optional date system of the workbook the formula belongs to; absent
   * means `'date1900'`.  TEXT reads it, so serial numbers render as dates in
   * this system.
   */
  dateSystem?: DateSystem;
}

/**
 * The evaluator callback handed to every {@link FormulaFunction}: evaluates
 * one AST node in scalar context.
 *
 * @internal
 */
export type EvaluateFn = (node: ASTNode, ctx: EvalContext) => CellValue;

/**
 * A range read by the resolver (see {@link argToMatrix}).
 *
 * `rows` × `cols` is the declared area; `values` holds only the part inside
 * the sheet's populated bounds, anchored at the declared top-left, so it may
 * have fewer (or zero) rows and columns.  Every declared cell outside
 * `values` is blank.
 *
 * @internal
 */
export interface ResolvedRange {
  /** Sheet the range reads from. */
  sheet: string;
  /** Declared top-left column (0-based). */
  col: number;
  /** Declared top-left row (1-based). */
  row: number;
  /** Declared height. */
  rows: number;
  /** Declared width. */
  cols: number;
  /** Values inside the populated bounds, row-major from the declared top-left. */
  values: CellValue[][];
}

// ---------------------------------------------------------------------------
// Resolution helpers
// ---------------------------------------------------------------------------

/**
 * `true` when a 0-based column and a 1-based integer row lie inside the Excel
 * grid.  `col` comes from `columnIndexOrNull`, so letters that name no column
 * up to XFD (either case accepted) arrive as `null` and fail the check.
 */
function inGrid(col: number | null, row: number): col is number {
  return (
    col !== null &&
    col >= 0 &&
    col <= MAX_COL_INDEX &&
    Number.isInteger(row) &&
    row >= 1 &&
    row <= MAX_ROW
  );
}

/**
 * The declared extent of a range: sheet, top-left and size.
 *
 * @internal
 */
export type RangeExtent = Omit<ResolvedRange, 'values'>;

/**
 * Normalise a range so the top-left comes first, whatever order the user
 * wrote the endpoints in.  An endpoint outside the grid, or a 3-D prefix
 * (`Sheet1:Sheet3!A1:B2`, which the evaluator does not resolve), gives `'#REF!'`.
 *
 * @internal
 */
export function rangeExtent(range: RangeNode, ctx: EvalContext): RangeExtent | string {
  if (range.start.endSheet !== undefined || range.end.endSheet !== undefined) return '#REF!';
  const sheet = range.start.sheet ?? range.end.sheet ?? ctx.currentSheet;
  const startCol = columnIndexOrNull(range.start.col);
  const endCol = columnIndexOrNull(range.end.col);
  const startRow = range.start.row;
  const endRow = range.end.row;
  if (!inGrid(startCol, startRow) || !inGrid(endCol, endRow)) return '#REF!';

  const col = Math.min(startCol, endCol);
  const row = Math.min(startRow, endRow);
  return {
    sheet,
    col,
    row,
    rows: Math.max(startRow, endRow) - row + 1,
    cols: Math.max(startCol, endCol) - col + 1,
  };
}

/**
 * The last row and column of an extent that are actually read: the declared
 * bottom-right clamped to the sheet's bounds, or `'#NUM!'` when the context
 * has no bounds and the extent spans more than 1,048,576 cells.
 */
function readLimits(
  extent: RangeExtent,
  ctx: EvalContext,
): { lastRow: number; lastCol: number } | '#NUM!' {
  const { sheet, col, row, rows, cols } = extent;
  const lastRow = row + rows - 1;
  const lastCol = col + cols - 1;
  const bounds = ctx.getBounds?.(sheet);
  if (bounds) {
    return { lastRow: Math.min(lastRow, bounds.maxRow), lastCol: Math.min(lastCol, bounds.maxCol) };
  }
  return rows * cols > MAX_UNBOUNDED_CELLS ? '#NUM!' : { lastRow, lastCol };
}

/**
 * The resolver's read cap: without {@link EvalContext.getBounds}, an area of
 * more than {@link MAX_UNBOUNDED_CELLS} cells is not read and resolves to
 * `#NUM!`.  `true` when `extent` is over that cap and the sheet has no bounds.
 *
 * @internal
 */
export function exceedsReadCap(extent: RangeExtent, ctx: EvalContext): boolean {
  return (
    ctx.getBounds?.(extent.sheet) === undefined && extent.rows * extent.cols > MAX_UNBOUNDED_CELLS
  );
}

/**
 * Read the cells of a validated extent, clamped to the sheet's bounds or capped without them.
 *
 * @internal
 */
export function readExtent(extent: RangeExtent, ctx: EvalContext): ResolvedRange | '#NUM!' {
  const limits = readLimits(extent, ctx);
  if (typeof limits === 'string') return limits;
  const { sheet, col, row } = extent;
  const { lastRow, lastCol } = limits;

  const values: CellValue[][] = [];
  if (lastCol >= col) {
    for (let r = row; r <= lastRow; r++) {
      const line: CellValue[] = [];
      for (let c = col; c <= lastCol; c++) {
        line.push(ctx.getCell(sheet, c, r));
      }
      values.push(line);
    }
  }
  return { ...extent, values };
}

/** Stream the cells {@link readExtent} would read, row-major, without building the matrix. */
function* streamExtent(
  extent: RangeExtent,
  ctx: EvalContext,
): Generator<CellValue, void, undefined> {
  const limits = readLimits(extent, ctx);
  if (typeof limits === 'string') {
    yield limits;
    return;
  }
  const { sheet, col, row } = extent;
  const { lastRow, lastCol } = limits;
  if (lastCol < col) return;
  for (let r = row; r <= lastRow; r++) {
    for (let c = col; c <= lastCol; c++) {
      yield ctx.getCell(sheet, c, r);
    }
  }
}

// ---------------------------------------------------------------------------
// Areas: the rectangle a reference node denotes
// ---------------------------------------------------------------------------

/** The 1x1 extent of a cell reference; `'#REF!'` outside the grid or for a 3-D reference. */
function cellExtent(ref: CellRefNode, ctx: EvalContext): RangeExtent | string {
  if (ref.endSheet !== undefined) return '#REF!';
  const col = columnIndexOrNull(ref.col);
  if (!inGrid(col, ref.row)) return '#REF!';
  return { sheet: ref.sheet ?? ctx.currentSheet, col, row: ref.row, rows: 1, cols: 1 };
}

/**
 * The extent of a whole-column range (`A:C`): every row of the grid, so
 * 1,048,576 rows per column.  `'#REF!'` past XFD or for a 3-D reference.
 */
function columnRangeExtent(node: ColumnRangeNode, ctx: EvalContext): RangeExtent | string {
  if (node.endSheet !== undefined) return '#REF!';
  const startCol = columnIndexOrNull(node.startCol);
  const endCol = columnIndexOrNull(node.endCol);
  if (!inGrid(startCol, 1) || !inGrid(endCol, 1)) return '#REF!';
  const col = Math.min(startCol, endCol);
  return {
    sheet: node.sheet ?? ctx.currentSheet,
    col,
    row: 1,
    rows: MAX_ROW,
    cols: Math.max(startCol, endCol) - col + 1,
  };
}

/**
 * The extent of a whole-row range (`1:3`): every column of the grid, so
 * 16,384 columns per row.  `'#REF!'` outside 1..1,048,576 or for a 3-D
 * reference.
 */
function rowRangeExtent(node: RowRangeNode, ctx: EvalContext): RangeExtent | string {
  if (node.endSheet !== undefined) return '#REF!';
  if (!inGrid(0, node.startRow) || !inGrid(0, node.endRow)) return '#REF!';
  const row = Math.min(node.startRow, node.endRow);
  return {
    sheet: node.sheet ?? ctx.currentSheet,
    col: 0,
    row,
    rows: Math.max(node.startRow, node.endRow) - row + 1,
    cols: MAX_COL_INDEX + 1,
  };
}

/**
 * The overlap of the operands of a `' '` intersection chain.  The parser
 * folds `A1:C3 B2:D4 C1:C9` into `((A1:C3 B2:D4) C1:C9)`, so a long chain is
 * thousands of levels deep; its left spine is walked with a loop and only the
 * right operands (and the innermost left one) recurse.  The first error from
 * the left wins, as in a recursive walk.
 */
function intersectChain(node: BinaryOpNode, ctx: EvalContext): RangeExtent | string {
  const outer: BinaryOpNode[] = [];
  let deepest = node;
  while (deepest.left.type === 'binary_op' && deepest.left.op === ' ') {
    outer.push(deepest);
    deepest = deepest.left;
  }
  let acc = intersectExtent(areaExtent(deepest.left, ctx), deepest.right, ctx);
  for (const link of outer.toReversed()) acc = intersectExtent(acc, link.right, ctx);
  return acc;
}

/**
 * The overlap of an already resolved left operand and the reference `right`
 * (the `' '` operator).  `'#NULL!'` when they do not overlap or sit on
 * different sheets; `'#VALUE!'` when an operand is not a single area (a
 * value, or a union).  An error in `a`, then in `right`, propagates.
 */
function intersectExtent(
  a: RangeExtent | string | undefined,
  right: ASTNode,
  ctx: EvalContext,
): RangeExtent | string {
  if (typeof a === 'string') return a;
  const b = areaExtent(right, ctx);
  if (typeof b === 'string') return b;
  if (a === undefined || b === undefined) return '#VALUE!';
  if (a.sheet.toUpperCase() !== b.sheet.toUpperCase()) return '#NULL!';

  const col = Math.max(a.col, b.col);
  const row = Math.max(a.row, b.row);
  const lastCol = Math.min(a.col + a.cols, b.col + b.cols) - 1;
  const lastRow = Math.min(a.row + a.rows, b.row + b.rows) - 1;
  if (lastCol < col || lastRow < row) return '#NULL!';
  return { sheet: a.sheet, col, row, rows: lastRow - row + 1, cols: lastCol - col + 1 };
}

/**
 * The rectangle a single-area reference denotes: a cell, a range, a whole
 * column or row, or the `' '` intersection of such references.  Returns an
 * error string when the reference cannot be resolved (`'#REF!'` outside the
 * grid, for 3-D and structured references, which the evaluator does not resolve;
 * `'#NULL!'` for an empty intersection) and `undefined` when `node` is not a
 * single-area reference (a value, a function call, a union).
 *
 * The extent is the declared one, never clamped to
 * {@link EvalContext.getBounds}.
 *
 * @internal
 */
export function areaExtent(node: ASTNode, ctx: EvalContext): RangeExtent | string | undefined {
  switch (node.type) {
    case 'cell_ref':
      return cellExtent(node, ctx);
    case 'range':
      return rangeExtent(node, ctx);
    case 'column_range':
      return columnRangeExtent(node, ctx);
    case 'row_range':
      return rowRangeExtent(node, ctx);
    case 'structured_ref':
      return '#REF!';
    case 'binary_op':
      return node.op === ' ' ? intersectChain(node, ctx) : undefined;
    default:
      return undefined;
  }
}

/**
 * `true` when `node` denotes a reference rather than a value: a cell, a
 * range, a whole column or row, a structured reference, or the `' '`
 * intersection or `','` union of references.
 *
 * @internal
 */
export function isReferenceNode(node: ASTNode): boolean {
  // Walk the left spine of a ' ' / ',' chain with a loop: a left-folded
  // chain is thousands of levels deep, and only right operands recurse.
  let current = node;
  while (current.type === 'binary_op') {
    if (current.op !== ' ' && current.op !== ',') return false;
    if (!isReferenceNode(current.right)) return false;
    current = current.left;
  }
  switch (current.type) {
    case 'cell_ref':
    case 'range':
    case 'column_range':
    case 'row_range':
    case 'structured_ref':
      return true;
    default:
      return false;
  }
}

/**
 * The areas of a `','` union of references, left to right, with the
 * left-folded chain `((A1,B2),C3)` flattened by a loop to `[A1, B2, C3]`;
 * `undefined` when `node` is not such a union.  An operand that is itself a
 * parenthesised union (the right one in `(A1,(B2,C3))`) is kept whole.
 *
 * @internal
 */
export function unionAreas(node: ASTNode): ASTNode[] | undefined {
  if (node.type !== 'binary_op' || node.op !== ',' || !isReferenceNode(node)) return undefined;
  const areas: ASTNode[] = [];
  let current: ASTNode = node;
  while (current.type === 'binary_op' && current.op === ',') {
    areas.push(current.right);
    current = current.left;
  }
  areas.push(current);
  return areas.reverse();
}

/**
 * Read a function argument as a matrix of values, the way array-aware
 * functions (aggregates, lookups, logical folds) see it:
 *
 *   - a range, whole column or whole row resolves to its cells, clamped to
 *     {@link EvalContext.getBounds} like {@link resolveRange}; without bounds
 *     an area of more than 1,048,576 cells gives `[['#NUM!']]`;
 *   - a `' '` intersection resolves to the overlap rectangle, or
 *     `[['#NULL!']]` when the areas do not overlap;
 *   - a `','` union of references appends the rows of every area, left to
 *     right (see {@link unionAreas});
 *   - an array constant evaluates every element (`{}` gives `[['#VALUE!']]`);
 *   - a 3-D or structured reference gives `[['#REF!']]`;
 *   - anything else, a single cell included, is `[[evaluate(arg, ctx)]]`.
 *
 * Use {@link argValues} when only the cells matter: it streams them without
 * allocating the matrix.
 *
 * @internal
 */
export function argToMatrix(arg: ASTNode, ctx: EvalContext, evaluate: EvaluateFn): CellValue[][] {
  if (arg.type === 'array') {
    if (arg.rows.length === 0) return [['#VALUE!']];
    return arg.rows.map((row) => row.map((el) => evaluate(el, ctx)));
  }
  const areas = unionAreas(arg);
  if (areas) return areas.flatMap((area) => argToMatrix(area, ctx, evaluate));
  const area = arg.type === 'cell_ref' ? undefined : areaExtent(arg, ctx);
  if (area === undefined) return [[evaluate(arg, ctx)]];
  if (typeof area === 'string') return [[area]];
  const read = readExtent(area, ctx);
  return typeof read === 'string' ? [[read]] : read.values;
}

/**
 * Stream the cells of {@link argToMatrix}`(arg, ctx, evaluate)` row-major
 * without building the matrix, so an aggregate over a whole column
 * (`SUM(A:A)`) need not allocate a 1,048,576-row array.
 *
 * Cost: an area reads one {@link EvalContext.getCell} per cell it covers.
 * With {@link EvalContext.getBounds} that is only the populated part; without
 * it, a whole column is 1,048,576 calls and a whole row 16,384, and an area
 * past 1,048,576 cells (such as `A:B`) yields a single `'#NUM!'` instead.
 *
 * @internal
 */
export function* argValues(
  arg: ASTNode,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): Generator<CellValue, void, undefined> {
  if (arg.type === 'array') {
    if (arg.rows.length === 0) yield '#VALUE!';
    for (const row of arg.rows) {
      for (const el of row) yield evaluate(el, ctx);
    }
    return;
  }
  const areas = unionAreas(arg);
  if (areas) {
    for (const area of areas) yield* argValues(area, ctx, evaluate);
    return;
  }
  const area = arg.type === 'cell_ref' ? undefined : areaExtent(arg, ctx);
  if (area === undefined) yield evaluate(arg, ctx);
  else if (typeof area === 'string') yield area;
  else yield* streamExtent(area, ctx);
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Resolve a single cell reference to its value.
 *
 * Uses the optional `sheet` qualifier on the node, falling back to
 * `ctx.currentSheet` when absent.  Column letters may be either case.  A
 * reference outside the grid `A1:XFD1048576` (column letters that are not 1-3
 * letters up to XFD, or a row that is not an integer from 1 to 1,048,576)
 * resolves to `'#REF!'` without calling `ctx.getCell`, and so does a 3-D
 * reference (`Sheet1:Sheet3!A1`), which the evaluator does not resolve.
 */
export function resolveRef(ref: CellRefNode, ctx: EvalContext): CellValue {
  const extent = cellExtent(ref, ctx);
  return typeof extent === 'string' ? extent : ctx.getCell(extent.sheet, extent.col, extent.row);
}

/**
 * Resolve a range reference to a 2D array of values (rows x cols).
 *
 * The range is normalised so that `start` is always the top-left corner
 * regardless of the order the user specified.  When the context supplies
 * {@link EvalContext.getBounds}, the matrix is clamped to the populated
 * bounds (top-left kept), so it may be smaller than the declared range or
 * empty.  Column letters may be either case.  An endpoint outside the grid
 * `A1:XFD1048576`, or a 3-D prefix, resolves to `[['#REF!']]` without calling
 * `ctx.getCell`.
 *
 * @throws {ModernXlsxError} `INVALID_ARGUMENT` when the context has no bounds
 *   for the sheet and the range spans more than 1,048,576 cells.
 */
export function resolveRange(range: RangeNode, ctx: EvalContext): CellValue[][] {
  const extent = rangeExtent(range, ctx);
  if (typeof extent === 'string') return [[extent]];
  const resolved = readExtent(extent, ctx);
  if (typeof resolved !== 'string') return resolved.values;
  const n = extent.rows * extent.cols;
  throw new ModernXlsxError(
    INVALID_ARGUMENT,
    `Range spans ${n} cells; provide EvalContext.getBounds to evaluate whole-column or whole-sheet ranges`,
  );
}
