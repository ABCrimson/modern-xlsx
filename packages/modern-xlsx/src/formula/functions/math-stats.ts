/**
 * Math and statistical formula functions.
 *
 * Registers 25 built-in Excel functions covering arithmetic
 * (ROUND, ROUNDUP, ROUNDDOWN, ABS, SQRT, MOD, INT, CEILING, FLOOR,
 * POWER, LOG, LN, PI, RAND) and aggregation
 * (SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK,
 * SUMIF, COUNTIF, AVERAGEIF, SUMPRODUCT).
 *
 * @module formula/functions/math-stats
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
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val === null) return 0;
  if (isError(val)) return val;
  const n = Number(val);
  return Number.isNaN(n) ? '#VALUE!' : n;
}

/** Coerce a CellValue to a numeric contribution (0 for non-numeric). */
function numericValue(v: CellValue): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  return 0;
}

/**
 * Collect all values from arguments, flattening ranges into individual cells.
 */
function collectValues(
  args: ASTNode[],
  ctx: EvalContext,
  evaluate: (node: ASTNode, ctx: EvalContext) => CellValue,
): CellValue[] {
  const values: CellValue[] = [];
  for (const arg of args) {
    if (arg.type === 'range') {
      const matrix = resolveRange(arg, ctx);
      for (const row of matrix) {
        for (const cell of row) {
          values.push(cell);
        }
      }
    } else {
      values.push(evaluate(arg, ctx));
    }
  }
  return values;
}

/**
 * Resolve a range argument to a 2D matrix.
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

// ---------------------------------------------------------------------------
// Criteria matching (for SUMIF / COUNTIF / AVERAGEIF)
// ---------------------------------------------------------------------------

type Matcher = (val: CellValue) => boolean;

/** Compare a numeric value against a numeric operand using the given operator. */
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
    case '=':
      return n === operand;
    case '<>':
      return n !== operand;
    default:
      return false;
  }
}

/** Compare two strings using the given operator. */
function compareStrings(op: string, a: string, b: string): boolean {
  switch (op) {
    case '=':
      return a === b;
    case '<>':
      return a !== b;
    case '>':
      return a.localeCompare(b) > 0;
    case '<':
      return a.localeCompare(b) < 0;
    case '>=':
      return a.localeCompare(b) >= 0;
    case '<=':
      return a.localeCompare(b) <= 0;
    default:
      return false;
  }
}

/** Build a matcher for operator-prefixed criteria like ">5" or "<>hello". */
function buildOperatorMatcher(op: string, operand: string): Matcher {
  const numOperand = Number(operand);
  const isNum = operand !== '' && !Number.isNaN(numOperand);

  if (isNum) {
    return (val) => {
      const n = typeof val === 'number' ? val : Number(val);
      if (Number.isNaN(n)) {
        if (op === '<>') return true;
        if (op === '=') return operand.toLowerCase() === String(val).toLowerCase();
        return false;
      }
      return compareNumeric(op, n, numOperand);
    };
  }

  const lowerOp = operand.toLowerCase();
  return (val) => compareStrings(op, String(val ?? '').toLowerCase(), lowerOp);
}

/**
 * Build a matcher function from an Excel criteria value.
 */
function buildMatcher(criteria: CellValue): Matcher {
  if (typeof criteria === 'number') {
    return (val) => typeof val === 'number' && val === criteria;
  }
  if (typeof criteria === 'boolean') {
    return (val) => val === criteria;
  }
  if (criteria === null) {
    return (val) => val === null || val === '';
  }

  const s = String(criteria);

  // Operator prefix patterns
  const opMatch = s.match(/^(<>|>=|<=|>|<|=)(.*)$/);
  if (opMatch) {
    const op = opMatch[1] ?? '';
    const val = opMatch[2] ?? '';
    return buildOperatorMatcher(op, val);
  }

  // Wildcard support: * and ?
  if (s.includes('*') || s.includes('?')) {
    const escaped = s.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    const pattern = escaped.replace(/\*/g, '.*').replace(/\?/g, '.');
    const re = new RegExp(`^${pattern}$`, 'i');
    return (val) => re.test(String(val ?? ''));
  }

  // Plain string: case-insensitive exact match
  const lower = s.toLowerCase();
  return (val) => String(val ?? '').toLowerCase() === lower;
}

// ---------------------------------------------------------------------------
// Aggregate helpers (reduce complexity of inline lambdas)
// ---------------------------------------------------------------------------

/** Check collected values for first error; return it or null. */
function firstError(values: CellValue[]): string | null {
  for (const v of values) {
    if (isError(v)) return v;
  }
  return null;
}

/**
 * Sum numeric values from a collected list.
 * Booleans count as 1/0; strings and nulls are skipped.
 */
function sumValues(values: CellValue[]): CellValue {
  const err = firstError(values);
  if (err) return err;
  let total = 0;
  for (const v of values) {
    if (typeof v === 'number') total += v;
    else if (typeof v === 'boolean') total += v ? 1 : 0;
  }
  return total;
}

/** Average numeric values (booleans count). */
function averageValues(values: CellValue[]): CellValue {
  const err = firstError(values);
  if (err) return err;
  let total = 0;
  let count = 0;
  for (const v of values) {
    if (typeof v === 'number' || typeof v === 'boolean') {
      total += numericValue(v);
      count++;
    }
  }
  return count === 0 ? '#DIV/0!' : total / count;
}

/**
 * Iterate two same-shaped matrices in parallel, applying a matcher to the
 * criteria range and collecting numeric values from the value range.
 */
function conditionalCollect(
  rangeMatrix: CellValue[][],
  valueMatrix: CellValue[][],
  matcher: Matcher,
): number[] {
  const results: number[] = [];
  for (let r = 0; r < rangeMatrix.length; r++) {
    const rangeRow = rangeMatrix[r];
    if (!rangeRow) continue;
    const valRow = valueMatrix[r];
    if (!valRow) continue;
    for (let c = 0; c < rangeRow.length; c++) {
      if (matcher(rangeRow[c] ?? null)) {
        const v = valRow[c];
        if (typeof v === 'number') results.push(v);
      }
    }
  }
  return results;
}

// ---------------------------------------------------------------------------
// Extracted implementations (reduce cognitive complexity)
// ---------------------------------------------------------------------------

const sumifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const rangeMatrix = resolveMatrix(arg0, ctx, evaluate);
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  const arg2 = args[2];
  const sumMatrix = arg2 ? resolveMatrix(arg2, ctx, evaluate) : rangeMatrix;
  const nums = conditionalCollect(rangeMatrix, sumMatrix, matcher);
  let total = 0;
  for (const n of nums) total += n;
  return total;
};

const countifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const rangeMatrix = resolveMatrix(arg0, ctx, evaluate);
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  let count = 0;
  for (const row of rangeMatrix) {
    for (const cell of row) {
      if (matcher(cell)) count++;
    }
  }
  return count;
};

const averageifImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 2) return '#VALUE!';
  const arg0 = args[0];
  const arg1 = args[1];
  if (!arg0 || !arg1) return '#VALUE!';
  const rangeMatrix = resolveMatrix(arg0, ctx, evaluate);
  const criteriaVal = evaluate(arg1, ctx);
  if (isError(criteriaVal)) return criteriaVal;
  const matcher = buildMatcher(criteriaVal);
  const arg2 = args[2];
  const avgMatrix = arg2 ? resolveMatrix(arg2, ctx, evaluate) : rangeMatrix;
  const nums = conditionalCollect(rangeMatrix, avgMatrix, matcher);
  if (nums.length === 0) return '#DIV/0!';
  let total = 0;
  for (const n of nums) total += n;
  return total / nums.length;
};

const sumproductImpl: FormulaFunction = (args, ctx, evaluate): CellValue => {
  if (args.length < 1) return '#VALUE!';
  const matrices = args.map((a) => resolveMatrix(a, ctx, evaluate));
  const dimError = validateDimensions(matrices);
  if (dimError) return dimError;

  const m0 = matrices[0];
  if (!m0) return '#VALUE!';
  const rows = m0.length;
  const cols = m0[0]?.length ?? 0;
  let total = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const prod = cellProduct(matrices, r, c);
      if (typeof prod === 'string') return prod;
      total += prod;
    }
  }
  return total;
};

/** Validate all matrices have the same dimensions. */
function validateDimensions(matrices: CellValue[][][]): CellValue {
  const first = matrices[0];
  if (!first) return '#VALUE!';
  const rows = first.length;
  const cols = first[0]?.length ?? 0;
  for (const m of matrices) {
    if (m.length !== rows || (m[0]?.length ?? 0) !== cols) return '#VALUE!';
  }
  return null;
}

/** Compute product of corresponding cells across all matrices at (r, c). */
function cellProduct(matrices: CellValue[][][], r: number, c: number): number | string {
  let product = 1;
  for (const m of matrices) {
    const v = m[r]?.[c] ?? null;
    if (isError(v)) return v;
    product *= numericValue(v);
  }
  return product;
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

/**
 * Register all math and statistical functions into the given registry.
 */
export function registerMathStatsFunctions(registry: Map<string, FormulaFunction>): void {
  // ---- SUM ---------------------------------------------------------------
  registry.set('SUM', (args, ctx, evaluate): CellValue => {
    return sumValues(collectValues(args, ctx, evaluate));
  });

  // ---- AVERAGE -----------------------------------------------------------
  registry.set('AVERAGE', (args, ctx, evaluate): CellValue => {
    return averageValues(collectValues(args, ctx, evaluate));
  });

  // ---- MIN ---------------------------------------------------------------
  registry.set('MIN', (args, ctx, evaluate): CellValue => {
    const values = collectValues(args, ctx, evaluate);
    const err = firstError(values);
    if (err) return err;
    let result = Number.POSITIVE_INFINITY;
    let found = false;
    for (const v of values) {
      if (typeof v === 'number') {
        if (v < result) result = v;
        found = true;
      }
    }
    return found ? result : 0;
  });

  // ---- MAX ---------------------------------------------------------------
  registry.set('MAX', (args, ctx, evaluate): CellValue => {
    const values = collectValues(args, ctx, evaluate);
    const err = firstError(values);
    if (err) return err;
    let result = Number.NEGATIVE_INFINITY;
    let found = false;
    for (const v of values) {
      if (typeof v === 'number') {
        if (v > result) result = v;
        found = true;
      }
    }
    return found ? result : 0;
  });

  // ---- COUNT -------------------------------------------------------------
  registry.set('COUNT', (args, ctx, evaluate): CellValue => {
    const values = collectValues(args, ctx, evaluate);
    const err = firstError(values);
    if (err) return err;
    let count = 0;
    for (const v of values) {
      if (typeof v === 'number') count++;
    }
    return count;
  });

  // ---- COUNTA ------------------------------------------------------------
  registry.set('COUNTA', (args, ctx, evaluate): CellValue => {
    const values = collectValues(args, ctx, evaluate);
    let count = 0;
    for (const v of values) {
      if (v !== null && v !== '') count++;
    }
    return count;
  });

  // ---- COUNTBLANK --------------------------------------------------------
  registry.set('COUNTBLANK', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const values = collectValues(args, ctx, evaluate);
    let count = 0;
    for (const v of values) {
      if (v === null || v === '') count++;
    }
    return count;
  });

  // ---- ROUND -------------------------------------------------------------
  registry.set('ROUND', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const nVal = toNumber(evaluate(arg0, ctx));
    if (typeof nVal === 'string') return nVal;
    const dVal = toNumber(evaluate(arg1, ctx));
    if (typeof dVal === 'string') return dVal;
    const d = Math.trunc(dVal);
    const factor = 10 ** d;
    return Math.round((nVal + Number.EPSILON * Math.sign(nVal)) * factor) / factor;
  });

  // ---- ROUNDUP -----------------------------------------------------------
  registry.set('ROUNDUP', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const nVal = toNumber(evaluate(arg0, ctx));
    if (typeof nVal === 'string') return nVal;
    const dVal = toNumber(evaluate(arg1, ctx));
    if (typeof dVal === 'string') return dVal;
    const d = Math.trunc(dVal);
    const factor = 10 ** d;
    const shifted = nVal * factor;
    return (nVal >= 0 ? Math.ceil(shifted) : Math.floor(shifted)) / factor;
  });

  // ---- ROUNDDOWN ---------------------------------------------------------
  registry.set('ROUNDDOWN', (args, ctx, evaluate): CellValue => {
    if (args.length < 2) return '#VALUE!';
    const arg0 = args[0];
    const arg1 = args[1];
    if (!arg0 || !arg1) return '#VALUE!';
    const nVal = toNumber(evaluate(arg0, ctx));
    if (typeof nVal === 'string') return nVal;
    const dVal = toNumber(evaluate(arg1, ctx));
    if (typeof dVal === 'string') return dVal;
    const d = Math.trunc(dVal);
    const factor = 10 ** d;
    const shifted = nVal * factor;
    return Math.trunc(shifted) / factor;
  });

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
    return Math.ceil(n / sig) * sig;
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
    return Math.floor(n / sig) * sig;
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
    const result = base ** exp;
    return Number.isFinite(result) ? result : '#NUM!';
  });

  // ---- LOG ---------------------------------------------------------------
  registry.set('LOG', (args, ctx, evaluate): CellValue => {
    if (args.length < 1) return '#VALUE!';
    const arg0 = args[0];
    if (!arg0) return '#VALUE!';
    const n = toNumber(evaluate(arg0, ctx));
    if (typeof n === 'string') return n;
    if (n <= 0) return '#NUM!';
    let base = 10;
    if (args.length >= 2) {
      const arg1 = args[1];
      if (!arg1) return '#VALUE!';
      const bVal = toNumber(evaluate(arg1, ctx));
      if (typeof bVal === 'string') return bVal;
      if (bVal <= 0 || bVal === 1) return '#NUM!';
      base = bVal;
    }
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
