/**
 * Tree-walk formula evaluator.
 *
 * Parses a formula string into an AST (via {@link parseFormula}) then
 * recursively evaluates each node, resolving cell references through the
 * supplied {@link EvalContext}.
 *
 * Follows Excel semantics for type coercion, error propagation, and
 * operator behaviour.
 *
 * @module formula/evaluator
 */

import type { ASTNode } from './parser.js';
import { parseFormula } from './parser.js';
import type { CellValue, EvalContext } from './resolver.js';
import { resolveRange, resolveRef } from './resolver.js';

// Re-export shared types so consumers can import from a single module.
export type { CellValue, EvalContext, FormulaFunction } from './resolver.js';

// ---------------------------------------------------------------------------
// Type coercion helpers (Excel semantics)
// ---------------------------------------------------------------------------

/** Returns `true` when `val` is an Excel error string (starts with `#`). */
function isError(val: CellValue): val is string {
  return typeof val === 'string' && val.length > 0 && val.charAt(0) === '#';
}

/**
 * Coerce a cell value to a number following Excel rules:
 *   number  -> itself
 *   boolean -> 1 / 0
 *   null    -> 0
 *   string  -> parseFloat, or "#VALUE!" if NaN
 *   error   -> propagated as-is
 */
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

/** Coerce a cell value to a string. */
function coerceToString(val: CellValue): string {
  if (val === null) return '';
  return typeof val === 'boolean' ? (val ? 'TRUE' : 'FALSE') : String(val);
}

// ---------------------------------------------------------------------------
// Comparison helpers
// ---------------------------------------------------------------------------

/** Type rank used for mixed-type comparison (Excel: blank < number < string < boolean). */
function typeRank(val: CellValue): number {
  if (val === null) return 0;
  switch (typeof val) {
    case 'number':
      return 1;
    case 'string':
      return 2;
    case 'boolean':
      return 3;
  }
}

/**
 * Compare two cell values following Excel rules:
 *   - Same-type: natural comparison (strings case-insensitive)
 *   - Mixed-type: number < string < boolean
 *
 * Returns negative / zero / positive like `Array#sort` comparator.
 */
function compareValues(a: CellValue, b: CellValue): number {
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;

  // Same type
  if (a === null) return 0;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') {
    return (a ? 1 : 0) - (b ? 1 : 0);
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Node evaluator
// ---------------------------------------------------------------------------

/**
 * Evaluate a single AST node, returning a {@link CellValue}.
 *
 * This is the recursive core of the evaluator.  It dispatches on the node
 * type and delegates cell lookups to the {@link EvalContext}.
 */
export function evaluateNode(node: ASTNode, ctx: EvalContext): CellValue {
  switch (node.type) {
    // -- literals ----------------------------------------------------------
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return node.value;

    // -- references --------------------------------------------------------
    case 'cell_ref':
      return resolveRef(node, ctx);

    case 'range': {
      // For scalar context, flatten a 1x1 range to its single value.
      const matrix = resolveRange(node, ctx);
      const firstRow = matrix[0];
      if (matrix.length === 1 && firstRow && firstRow.length === 1) {
        return firstRow[0] ?? null;
      }
      // Multi-cell ranges in scalar context are not yet fully supported;
      // return the top-left value (Excel implicit intersection).
      return firstRow?.[0] ?? null;
    }

    // -- named ranges (not yet resolved) -----------------------------------
    case 'name':
      return '#NAME?';

    // -- percent -----------------------------------------------------------
    case 'percent': {
      const operand = evaluateNode(node.operand, ctx);
      if (isError(operand)) return operand;
      const n = toNumber(operand);
      if (typeof n === 'string') return n; // error propagation
      return n / 100;
    }

    // -- unary operators ---------------------------------------------------
    case 'unary_op':
      return evaluateUnary(node.op, node.operand, ctx);

    // -- binary operators --------------------------------------------------
    case 'binary_op':
      return evaluateBinary(node.op, node.left, node.right, ctx);

    // -- function calls ----------------------------------------------------
    case 'function':
      return evaluateFunction(node.name, node.args, ctx);

    // -- array constants ---------------------------------------------------
    case 'array': {
      // Evaluate every element; return single value if 1x1.
      const rows = node.rows.map((row) => row.map((el) => evaluateNode(el, ctx)));
      const firstArrayRow = rows[0];
      if (rows.length === 1 && firstArrayRow && firstArrayRow.length === 1) {
        return firstArrayRow[0] ?? null;
      }
      // Multi-cell array in scalar context -> top-left.
      return firstArrayRow?.[0] ?? null;
    }

    default:
      return '#VALUE!';
  }
}

// ---------------------------------------------------------------------------
// Unary operator evaluation
// ---------------------------------------------------------------------------

function evaluateUnary(op: string, operandNode: ASTNode, ctx: EvalContext): CellValue {
  const operand = evaluateNode(operandNode, ctx);
  if (isError(operand)) return operand;

  const n = toNumber(operand);
  if (typeof n === 'string') return n; // #VALUE! or other error

  switch (op) {
    case '+':
      return n;
    case '-':
      return -n;
    default:
      return '#VALUE!';
  }
}

// ---------------------------------------------------------------------------
// Binary operator evaluation
// ---------------------------------------------------------------------------

function evaluateBinary(
  op: string,
  leftNode: ASTNode,
  rightNode: ASTNode,
  ctx: EvalContext,
): CellValue {
  const left = evaluateNode(leftNode, ctx);
  const right = evaluateNode(rightNode, ctx);

  // Error propagation (left first, then right).
  if (isError(left)) return left;
  if (isError(right)) return right;

  // -- concatenation -------------------------------------------------------
  if (op === '&') {
    return coerceToString(left) + coerceToString(right);
  }

  // -- comparison operators ------------------------------------------------
  switch (op) {
    case '=':
    case '<>':
    case '<':
    case '>':
    case '<=':
    case '>=':
      return evaluateComparison(op, left, right);
  }

  // -- arithmetic operators ------------------------------------------------
  const a = toNumber(left);
  if (typeof a === 'string') return a;
  const b = toNumber(right);
  if (typeof b === 'string') return b;

  switch (op) {
    case '+':
      return a + b;
    case '-':
      return a - b;
    case '*':
      return a * b;
    case '/':
      return b === 0 ? '#DIV/0!' : a / b;
    case '^':
      return a ** b;
    default:
      return '#VALUE!';
  }
}

// ---------------------------------------------------------------------------
// Comparison evaluation
// ---------------------------------------------------------------------------

function evaluateComparison(op: string, left: CellValue, right: CellValue): boolean {
  const cmp = compareValues(left, right);
  switch (op) {
    case '=':
      return cmp === 0;
    case '<>':
      return cmp !== 0;
    case '<':
      return cmp < 0;
    case '>':
      return cmp > 0;
    case '<=':
      return cmp <= 0;
    case '>=':
      return cmp >= 0;
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// Function call evaluation
// ---------------------------------------------------------------------------

function evaluateFunction(name: string, args: ASTNode[], ctx: EvalContext): CellValue {
  const upperName = name.toUpperCase();
  const fn = ctx.functions?.get(upperName);
  if (fn) {
    return fn(args, ctx, evaluateNode);
  }
  return '#NAME?';
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Parse and evaluate an Excel formula string.
 *
 * @param formula - Formula string, optionally starting with `=`.
 * @param ctx     - Evaluation context providing cell lookups.
 * @returns The computed cell value.
 */
export function evaluateFormula(formula: string, ctx: EvalContext): CellValue {
  const { ast, errors } = parseFormula(formula);
  if (!ast) {
    return errors.length > 0 ? '#VALUE!' : null;
  }
  if (errors.length > 0) {
    // Parse errors -> treat as #VALUE! (could also return partial result)
    return '#VALUE!';
  }
  return evaluateNode(ast, ctx);
}
