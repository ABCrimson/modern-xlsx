/**
 * Math and statistical formula functions.
 *
 * Registers the built-in Excel functions covering arithmetic
 * (ROUND, ROUNDUP, ROUNDDOWN, ABS, SQRT, MOD, INT, CEILING, FLOOR,
 * POWER, LOG, LN, PI, RAND) and aggregation
 * (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK,
 * SUMIF, COUNTIF, AVERAGEIF, SUMPRODUCT).
 *
 * Aggregates follow Excel's argument rule: cells of a reference and elements
 * of an array constant count only when they are numbers, while direct
 * arguments are coerced (`SUM(TRUE,"5")` is 6, `SUM("x")` is `#VALUE!`).
 * Criteria (`">=0"`, `"<>*x*"`, `"a~*b"`) compare numbers numerically and
 * text case-insensitively, with `*`/`?` wildcards from `./glob.ts`.
 *
 * @module formula/functions/math-stats
 */

import { excelPower, toNumber } from '../coerce.js';
import { compareText } from '../collation.js';
import { ERROR_SET, isError } from '../errors.js';
import { MAX_COL_INDEX, MAX_ROW } from '../grid.js';
import type { ASTNode } from '../parser.js';
import type {
  CellValue,
  EvalContext,
  EvaluateFn,
  FormulaFunction,
  ResolvedRange,
} from '../resolver.js';
import {
  areaExtent,
  argToMatrix,
  argValues,
  exceedsReadCap,
  isReferenceNode,
  readExtent,
  unionAreas,
} from '../resolver.js';
import { compileGlob, hasWildcard } from './glob.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** A value from an aggregate's arguments, tagged with where it came from. */
interface ArgValue {
  value: CellValue;
  /**
   * `true` for a cell of a reference argument (a cell, range, whole column or
   * row, intersection or union) and for an element of an array constant;
   * `false` for a direct argument (a literal, an operator or a function
   * result).  Excel coerces direct arguments only.
   */
  fromRef: boolean;
}

/**
 * Stream every value of the arguments in order, flattening references
 * (clamped to EvalContext.getBounds), unions and array constants, each tagged
 * with {@link ArgValue.fromRef}.  An unresolvable reference yields its error
 * (`#REF!`, or `#NUM!` for an area too large to read without bounds).
 */
function* collectValues(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: EvaluateFn,
): Generator<ArgValue, void, undefined> {
  for (const arg of args) {
    const fromRef = arg.type === 'array' || isReferenceNode(arg);
    for (const value of argValues(arg, ctx, evaluate)) yield { value, fromRef };
  }
}

/**
 * `'#NUM!'` when an argument (or one area of a union) is too large to read,
 * else `undefined`.  COUNT and COUNTA never return error values, so they
 * check this first instead of counting the resolver's `#NUM!` as a value.
 */
function unreadableArea(args: ASTNode[], ctx: EvalContext): '#NUM!' | undefined {
  const unreadable = (arg: ASTNode): boolean => {
    // A left-folded union is flattened by a loop, not by recursing on its left operand.
    const areas = unionAreas(arg);
    if (areas) return areas.some(unreadable);
    const extent = areaExtent(arg, ctx);
    return typeof extent === 'object' && exceedsReadCap(extent, ctx);
  };
  return args.some(unreadable) ? '#NUM!' : undefined;
}

/**
 * Read an argument as a {@link ResolvedRange} through argToMatrix.  A
 * reference keeps its declared extent, so cells clamped away by
 * EvalContext.getBounds still count as blank; an array constant, a union or
 * a scalar becomes the matrix of its values.  Returns the error string when
 * the reference cannot be resolved or is too large to read.
 */
function resolveMatrix(
  arg: ASTNode,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): ResolvedRange | string {
  const extent = arg.type === 'cell_ref' ? undefined : areaExtent(arg, ctx);
  if (typeof extent === 'string') return extent;
  if (extent !== undefined && exceedsReadCap(extent, ctx)) return '#NUM!';
  const values = argToMatrix(arg, ctx, evaluate);
  if (extent !== undefined) return { ...extent, values };
  let cols = 0;
  for (const row of values) cols = Math.max(cols, row.length);
  return { sheet: ctx.currentSheet, col: 0, row: 1, rows: values.length, cols, values };
}

/**
 * Read SUMIF's sum_range or AVERAGEIF's average_range the way Excel does:
 * a single-area reference keeps only its top-left cell and takes the size and
 * shape of `range` from there (`SUMIF(A1:A3,">0",B1)` reads B1:B3), cut at
 * the last row and column of the grid.  Any other argument (an array
 * constant, a union, a scalar) is read as is through {@link resolveMatrix}.
 * Returns the error string when the reference cannot be resolved or is too
 * large to read.
 */
function sizedLike(
  arg: ASTNode,
  range: ResolvedRange,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): ResolvedRange | string {
  const extent = areaExtent(arg, ctx);
  if (typeof extent === 'string') return extent;
  if (extent === undefined) return resolveMatrix(arg, ctx, evaluate);
  const { sheet, col, row } = extent;
  return readExtent(
    {
      sheet,
      col,
      row,
      rows: Math.min(range.rows, MAX_ROW - row + 1),
      cols: Math.min(range.cols, MAX_COL_INDEX - col + 1),
    },
    ctx,
  );
}

/**
 * Blank cells in a range: its declared area minus its non-blank cells, so
 * cells clamped away by EvalContext.getBounds still count as blank.
 */
function countBlank(range: ResolvedRange): number {
  let nonBlank = 0;
  for (const row of range.values) {
    for (const v of row) {
      if (v !== null && v !== '') nonBlank++;
    }
  }
  return range.rows * range.cols - nonBlank;
}

/** Number of cells actually present in a (possibly clamped) matrix. */
function presentCells(values: CellValue[][]): number {
  let n = 0;
  for (const row of values) n += row.length;
  return n;
}

// ---------------------------------------------------------------------------
// Criteria matching (for SUMIF / COUNTIF / AVERAGEIF)
// ---------------------------------------------------------------------------

type Matcher = (val: CellValue) => boolean;

/** Order a number against a numeric operand with an ordering operator (>, <, >=, <=). */
function compareNumeric(op: string, n: number, operand: number): boolean {
  switch (op) {
    case '>':
      return n > operand;
    case '<':
      return n < operand;
    case '>=':
      return n >= operand;
    case '<=':
      return n <= operand;
    default:
      return false;
  }
}

/** Order two lower-cased strings with the given ordering operator (>, <, >=, <=) via compareText. */
function compareStrings(op: string, a: string, b: string): boolean {
  switch (op) {
    case '>':
      return compareText(a, b) > 0;
    case '<':
      return compareText(a, b) < 0;
    case '>=':
      return compareText(a, b) >= 0;
    case '<=':
      return compareText(a, b) <= 0;
    default:
      return false;
  }
}

/**
 * Build a matcher for a numeric operand.
 *
 * `=` matches number cells equal to the operand and text cells that
 * {@link toNumber} reads as it (`'5'`, `' 5.0 '`; not `''`); `<>` matches
 * every other cell, blanks, booleans and text included.  The ordering
 * operators match number cells only: blanks, booleans and text, numeric text
 * included, never match.
 */
function buildNumericMatcher(op: string, operand: number): Matcher {
  const eq: Matcher = (val) =>
    typeof val === 'number'
      ? val === operand
      : typeof val === 'string' && toNumber(val) === operand;
  if (op === '=') return eq;
  if (op === '<>') return (val) => !eq(val);
  return (val) => typeof val === 'number' && compareNumeric(op, val, operand);
}

/**
 * Build the `=` matcher for a text operand, compared case-insensitively
 * through {@link compileGlob}, so `~*`, `~?` and `~~` are literals.
 *
 * With an unescaped `*` or `?` the operand is a wildcard pattern and matches
 * text cells only.  Without one it matches text cells and booleans by their
 * text form (`"true"` matches TRUE).  The empty operand matches blank cells
 * and empty text.  Number cells never match a text operand.
 */
function buildTextEquals(operand: string): Matcher {
  if (operand === '') return (val) => val === null || val === '';
  const match = compileGlob(operand);
  if (hasWildcard(operand)) return (val) => typeof val === 'string' && match(val);
  return (val) => {
    if (typeof val === 'string') return match(val);
    return typeof val === 'boolean' && match(val ? 'TRUE' : 'FALSE');
  };
}

/**
 * Build a matcher for a text operand.
 *
 * `=` (and a bare criterion, which means `=`) uses {@link buildTextEquals};
 * `<>` negates it, so blanks, numbers and booleans match `"<>x"`.  The
 * ordering operators compare text cells only, in collation order, with no
 * wildcards.
 */
function buildTextMatcher(op: string, operand: string): Matcher {
  if (op === '=') return buildTextEquals(operand);
  if (op === '<>') {
    const eq = buildTextEquals(operand);
    return (val) => !eq(val);
  }
  const lowerOp = operand.toLowerCase();
  return (val) => typeof val === 'string' && compareStrings(op, val.toLowerCase(), lowerOp);
}

/**
 * Build a matcher function from an Excel criteria value.
 *
 * A number criterion means `=` that number.  A text criterion is an optional
 * operator (`=`, `<>`, `>`, `<`, `>=`, `<=`) and an operand; an operand that
 * {@link toNumber} reads as a number (`"5.0"`, not `""` or hex) compares
 * numerically, any other operand as text.
 */
function buildMatcher(criteria: CellValue): Matcher {
  if (criteria === null) {
    return (val) => val === null || val === '';
  }
  switch (typeof criteria) {
    case 'number':
      return buildNumericMatcher('=', criteria);
    case 'boolean':
      return (val) => val === criteria;
    case 'string': {
      // Operator prefix; bare criteria mean '='.
      const opMatch = /^(<>|>=|<=|>|<|=)([\s\S]*)$/.exec(criteria);
      const op = opMatch?.[1] ?? '=';
      const operand = opMatch ? (opMatch[2] ?? '') : criteria;
      const num = toNumber(operand);
      return typeof num === 'number' ? buildNumericMatcher(op, num) : buildTextMatcher(op, operand);
    }
  }
}

// ---------------------------------------------------------------------------
// Aggregate helpers
// ---------------------------------------------------------------------------

/**
 * What SUM, AVERAGE, MIN and MAX make of one value, by Excel's argument rule:
 *   - an error value is returned as is;
 *   - a number counts wherever it comes from;
 *   - in a reference or an array constant, booleans, text and blanks are
 *     skipped (`undefined`);
 *   - a direct boolean counts as 1/0, and direct text counts when
 *     {@link toNumber} reads it as a number (`"5"`), else it is `#VALUE!`;
 *     a direct blank (an empty argument) is skipped.
 */
function countedNumber({ value, fromRef }: ArgValue): number | string | undefined {
  switch (typeof value) {
    case 'number':
      return value;
    case 'boolean':
      if (fromRef) return undefined;
      return value ? 1 : 0;
    case 'string':
      // ERROR_SET.has, not the isError type guard: its false side would
      // narrow `value` to `never`.
      if (ERROR_SET.has(value)) return value;
      return fromRef ? undefined : toNumber(value);
    default:
      return undefined;
  }
}

/**
 * Visit the numbers SUM, AVERAGE, MIN and MAX count ({@link countedNumber}),
 * in argument order.  Returns the first error string, or `undefined` once
 * every value was visited.
 */
function forEachCountedNumber(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: EvaluateFn,
  visit: (n: number) => void,
): string | undefined {
  for (const item of collectValues(args, ctx, evaluate)) {
    const n = countedNumber(item);
    if (typeof n === 'string') return n;
    if (n !== undefined) visit(n);
  }
  return undefined;
}

/**
 * Whether COUNT counts a value: numbers anywhere, and from direct arguments
 * also booleans and text that {@link toNumber} reads as a number.  Error
 * values are never counted.
 */
function countsAsNumber({ value, fromRef }: ArgValue): boolean {
  if (typeof value === 'number') return true;
  if (fromRef) return false;
  return (
    typeof value === 'boolean' || (typeof value === 'string' && typeof toNumber(value) === 'number')
  );
}

/**
 * Walk the criteria range and the value range in parallel, applying a matcher
 * to each criteria cell and collecting numeric values from the value range.
 *
 * Either matrix may be clamped by EvalContext.getBounds, so the walk covers
 * the union of both present areas inside the criteria range's declared size;
 * a criteria cell clamped away is blank and can still match ("" or "<>x").
 */
function conditionalCollect(range: ResolvedRange, sum: ResolvedRange, matcher: Matcher): number[] {
  const results: number[] = [];
  const rows = Math.min(range.rows, Math.max(range.values.length, sum.values.length));
  for (let r = 0; r < rows; r++) {
    const rangeRow = range.values[r];
    const sumRow = sum.values[r];
    const cols = Math.min(range.cols, Math.max(rangeRow?.length ?? 0, sumRow?.length ?? 0));
    for (let c = 0; c < cols; c++) {
      if (matcher(rangeRow?.[c] ?? null)) {
        const v = sumRow?.[c];
        if (typeof v === 'number') results.push(v);
      }
    }
  }
  return results;
}

// ---------------------------------------------------------------------------
// Function implementations
// ---------------------------------------------------------------------------

const sumifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const range = resolveMatrix(arg0, ctx, evaluate);
  if (typeof range === 'string') return range;
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  const arg2 = args[2];
  const sumRange = arg2 ? sizedLike(arg2, range, ctx, evaluate) : range;
  if (typeof sumRange === 'string') return sumRange;
  const nums = conditionalCollect(range, sumRange, matcher);
  let total = 0;
  for (const n of nums) total += n;
  return total;
};

const countifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const range = resolveMatrix(arg0, ctx, evaluate);
  if (typeof range === 'string') return range;
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  let count = 0;
  for (const row of range.values) {
    for (const cell of row) {
      if (matcher(cell)) count++;
    }
  }
  // Cells clamped away by EvalContext.getBounds are blank.
  if (matcher(null)) count += range.rows * range.cols - presentCells(range.values);
  return count;
};

const averageifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const range = resolveMatrix(arg0, ctx, evaluate);
  if (typeof range === 'string') return range;
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  const arg2 = args[2];
  const avgRange = arg2 ? sizedLike(arg2, range, ctx, evaluate) : range;
  if (typeof avgRange === 'string') return avgRange;
  const nums = conditionalCollect(range, avgRange, matcher);
  if (nums.length === 0) return '#DIV/0!';
  let total = 0;
  for (const n of nums) total += n;
  return total / nums.length;
};

const sumproductImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 1) return '#VALUE!';
  const ranges: ResolvedRange[] = [];
  for (const arg of args) {
    const range = resolveMatrix(arg, ctx, evaluate);
    if (typeof range === 'string') return range;
    ranges.push(range);
  }
  const dimError = validateDimensions(ranges);
  if (dimError) return dimError;

  // Only cells present in some clamped matrix can be non-blank; a blank
  // factor makes the product 0, so iterating their union is enough.
  let rows = 0;
  let cols = 0;
  for (const { values } of ranges) {
    rows = Math.max(rows, values.length);
    for (const row of values) cols = Math.max(cols, row.length);
  }
  let total = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const prod = cellProduct(ranges, r, c);
      if (typeof prod === 'string') return prod;
      total += prod;
    }
  }
  return total;
};

/** Validate all ranges have the same declared dimensions. */
function validateDimensions(ranges: ResolvedRange[]): CellValue {
  const first = ranges[0];
  if (!first) return '#VALUE!';
  for (const m of ranges) {
    if (m.rows !== first.rows || m.cols !== first.cols) return '#VALUE!';
  }
  return null;
}

/**
 * Compute product of corresponding cells across all ranges at (r, c).  Every
 * non-numeric entry, booleans and numeric text included, counts as 0, as in
 * Excel; an error value is returned as is.
 */
function cellProduct(ranges: ResolvedRange[], r: number, c: number): number | string {
  let product = 1;
  for (const m of ranges) {
    const v = m.values[r]?.[c] ?? null;
    if (isError(v)) return v;
    product *= typeof v === 'number' ? v : 0;
  }
  return product;
}

// ---------------------------------------------------------------------------
// Rounding
// ---------------------------------------------------------------------------

/**
 * Round `n` to `digits` decimal places (negative `digits` round left of the
 * point) the way Excel does, decimal-safe:
 *   - `'half'` (ROUND) rounds half away from zero: ROUND(-2.5,0) is -3;
 *   - `'up'` (ROUNDUP) rounds away from zero, `'down'` (ROUNDDOWN) toward it.
 *
 * The scaled magnitude is normalised to 15 significant digits before it is
 * rounded, so binary representation error does not leak into the result:
 * ROUNDUP(0.07,2) is 0.07 although 0.07 * 100 is 7.000000000000001.
 * `digits` is truncated.  When it asks for more precision than 15
 * significant digits hold, `n` is returned at 15 significant digits; when it
 * rounds away every digit, the result is 0, or for `'up'` the next power of
 * ten.  A non-finite input or result is `#NUM!`.
 */
function roundDecimal(n: number, digits: number, mode: 'half' | 'up' | 'down'): number | string {
  if (!Number.isFinite(n)) return '#NUM!';
  const d = Math.trunc(digits);
  if (n === 0) return 0;
  const sign = Math.sign(n);
  const abs = Math.abs(n);
  const mag = Math.floor(Math.log10(abs));
  if (d > 14 - mag) return Number(n.toPrecision(15));
  if (d < -(mag + 1)) {
    if (mode !== 'up') return 0;
    const r = sign * 10 ** -d;
    return Number.isFinite(r) ? r : '#NUM!';
  }
  // 10 ** d overflows to Infinity past 308 (reachable only for |n| < 1e-294),
  // so the scale is applied in two finite steps there; `down` divides by the
  // small factor first, which keeps ROUND(1E-300,310) exactly 1E-300.
  const up = (x: number) => (d > 308 ? x * 1e308 * 10 ** (d - 308) : x * 10 ** d);
  const down = (x: number) => (d > 308 ? x / 10 ** (d - 308) / 1e308 : x / 10 ** d);
  const scaled = Number((d >= 0 ? up(abs) : abs / 10 ** -d).toPrecision(15));
  let k: number;
  if (mode === 'half') k = Math.floor(scaled + 0.5);
  else if (mode === 'up') k = Math.ceil(scaled);
  else k = Math.floor(scaled);
  const r = sign * (d >= 0 ? down(k) : k * 10 ** -d);
  return Number.isFinite(r) ? r : '#NUM!';
}

/** ROUND, ROUNDUP or ROUNDDOWN: `roundDecimal(number, num_digits, mode)`. */
function roundFunction(mode: 'half' | 'up' | 'down'): FormulaFunction {
  return (args, ctx, evaluate) => {
    const [arg0, arg1] = args;
    if (!arg0 || !arg1) return '#VALUE!';
    const nVal = toNumber(evaluate(arg0, ctx));
    if (typeof nVal === 'string') return nVal;
    const dVal = toNumber(evaluate(arg1, ctx));
    if (typeof dVal === 'string') return dVal;
    return roundDecimal(nVal, dVal, mode);
  };
}

/**
 * CEILING's and FLOOR's step: `n` rounded by `step` (Math.ceil or Math.floor)
 * to a multiple of `sig`.  The quotient and the result are normalised to 15
 * significant digits, as in {@link roundDecimal}, so binary representation
 * error does not leak into the result: FLOOR(0.3,0.1) is 0.3 although
 * 0.3 / 0.1 is 2.9999999999999996, and CEILING(0.6,0.2) is 0.6, not
 * 0.6000000000000001.  A non-finite result is `#NUM!`.
 */
function toMultiple(n: number, sig: number, step: (x: number) => number): number | string {
  const q = Number((n / sig).toPrecision(15));
  const r = Number((step(q) * sig).toPrecision(15));
  return Number.isFinite(r) ? r : '#NUM!';
}

/**
 * LOG's base argument: 10 when omitted, an error propagated, `#NUM!` for a
 * base that is not positive or is 1.
 */
function logBase(
  arg: ASTNode | undefined,
  ctx: EvalContext,
  evaluate: EvaluateFn,
): number | string {
  if (arg === undefined) return 10;
  const base = toNumber(evaluate(arg, ctx));
  if (typeof base === 'string') return base;
  return base <= 0 || base === 1 ? '#NUM!' : base;
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

/**
 * Register all math and statistical functions into the given registry.
 */
export function registerMathStatsFunctions(registry: Map<string, FormulaFunction>): void {
  // SUM, AVERAGE, MIN and MAX share forEachCountedNumber's argument rule.
  // ---- SUM ---------------------------------------------------------------
  registry.set('SUM', (args, ctx, evaluate): CellValue => {
    let total = 0;
    const err = forEachCountedNumber(args, ctx, evaluate, (n) => {
      total += n;
    });
    return err ?? total;
  });

  // ---- AVERAGE -----------------------------------------------------------
  // Direct booleans and numeric text count here as they do for SUM and COUNT
  // (AVERAGE(TRUE,"3") is 2).  Microsoft's AVERAGE page says typed logical
  // values and number text are "not counted"; an Excel-authored fixture
  // should confirm this rule.
  registry.set('AVERAGE', (args, ctx, evaluate): CellValue => {
    let total = 0;
    let count = 0;
    const err = forEachCountedNumber(args, ctx, evaluate, (n) => {
      total += n;
      count++;
    });
    if (err !== undefined) return err;
    return count === 0 ? '#DIV/0!' : total / count;
  });

  // ---- MIN ---------------------------------------------------------------
  registry.set('MIN', (args, ctx, evaluate): CellValue => {
    let result = Number.POSITIVE_INFINITY;
    let found = false;
    const err = forEachCountedNumber(args, ctx, evaluate, (n) => {
      if (n < result) result = n;
      found = true;
    });
    if (err !== undefined) return err;
    return found ? result : 0;
  });

  // ---- MAX ---------------------------------------------------------------
  registry.set('MAX', (args, ctx, evaluate): CellValue => {
    let result = Number.NEGATIVE_INFINITY;
    let found = false;
    const err = forEachCountedNumber(args, ctx, evaluate, (n) => {
      if (n > result) result = n;
      found = true;
    });
    if (err !== undefined) return err;
    return found ? result : 0;
  });

  // ---- COUNT -------------------------------------------------------------
  // Error values are never counted and never returned (see countsAsNumber).
  registry.set('COUNT', (args, ctx, evaluate): CellValue => {
    const unreadable = unreadableArea(args, ctx);
    if (unreadable) return unreadable;
    let count = 0;
    for (const item of collectValues(args, ctx, evaluate)) {
      if (countsAsNumber(item)) count++;
    }
    return count;
  });

  // ---- COUNTA ------------------------------------------------------------
  registry.set('COUNTA', (args, ctx, evaluate): CellValue => {
    const unreadable = unreadableArea(args, ctx);
    if (unreadable) return unreadable;
    let count = 0;
    for (const { value } of collectValues(args, ctx, evaluate)) {
      if (value !== null && value !== '') count++;
    }
    return count;
  });

  // ---- COUNTBLANK --------------------------------------------------------
  registry.set('COUNTBLANK', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    let count = 0;
    for (const arg of args) {
      const range = resolveMatrix(arg, ctx, evaluate);
      if (typeof range === 'string') return range;
      count += countBlank(range);
    }
    return count;
  });

  // ---- ROUND -------------------------------------------------------------
  registry.set('ROUND', roundFunction('half'));

  // ---- ROUNDUP -----------------------------------------------------------
  registry.set('ROUNDUP', roundFunction('up'));

  // ---- ROUNDDOWN ---------------------------------------------------------
  registry.set('ROUNDDOWN', roundFunction('down'));

  // ---- ABS ---------------------------------------------------------------
  registry.set('ABS', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    return Math.abs(n);
  });

  // ---- SQRT --------------------------------------------------------------
  registry.set('SQRT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    if (n < 0) return '#NUM!';
    return Math.sqrt(n);
  });

  // ---- MOD ---------------------------------------------------------------
  registry.set('MOD', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    const d = toNumber(evaluate(arg1, ctx));
    if (typeof d === 'string') return d;
    if (d === 0) return '#DIV/0!';
    return n - d * Math.floor(n / d);
  });

  // ---- INT ---------------------------------------------------------------
  registry.set('INT', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    return Math.floor(n);
  });

  // ---- CEILING -----------------------------------------------------------
  registry.set('CEILING', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    const sig = toNumber(evaluate(arg1, ctx));
    if (typeof sig === 'string') return sig;
    if (sig === 0) return 0;
    if (n > 0 && sig < 0) return '#NUM!';
    return toMultiple(n, sig, Math.ceil);
  });

  // ---- FLOOR -------------------------------------------------------------
  registry.set('FLOOR', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    const sig = toNumber(evaluate(arg1, ctx));
    if (typeof sig === 'string') return sig;
    if (sig === 0) return '#DIV/0!';
    if (n > 0 && sig < 0) return '#NUM!';
    return toMultiple(n, sig, Math.floor);
  });

  // ---- POWER -------------------------------------------------------------
  registry.set('POWER', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const base = toNumber(evaluate(arg0, ctx));
    if (typeof base === 'string') return base;
    const exp = toNumber(evaluate(arg1, ctx));
    if (typeof exp === 'string') return exp;
    return excelPower(base, exp);
  });

  // ---- LOG ---------------------------------------------------------------
  registry.set('LOG', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    if (n <= 0) return '#NUM!';
    const base = logBase(args[1], ctx, evaluate);
    if (typeof base === 'string') return base;
    // log10 and log2 are exact for exact powers (LOG(1000) is 3, not
    // 2.9999999999999996 as the natural-log ratio gives).
    if (base === 10) return Math.log10(n);
    if (base === 2) return Math.log2(n);
    return Math.log(n) / Math.log(base);
  });

  // ---- LN ----------------------------------------------------------------
  registry.set('LN', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    if (n <= 0) return '#NUM!';
    return Math.log(n);
  });

  // ---- PI ----------------------------------------------------------------
  registry.set('PI', (): CellValue => Math.PI);

  // ---- RAND --------------------------------------------------------------
  registry.set('RAND', (): CellValue => Math.random());

  // ---- Conditional & product functions -----------------------------------
  registry.set('SUMIF', sumifImpl);
  registry.set('COUNTIF', countifImpl);
  registry.set('AVERAGEIF', averageifImpl);
  registry.set('SUMPRODUCT', sumproductImpl);
}
