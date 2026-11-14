/**
 * Tree-walk formula evaluator.
 *
 * Parses a formula string into an AST (via {@link parseFormula}) then
 * recursively evaluates each node, resolving cell references through the
 * supplied {@link EvalContext}.  Left-folded chains (`1+1+...`, `A1%%...`,
 * `A1:B1:C1`) are walked with a loop, so a chain thousands of operators long
 * within the 8192-character limit evaluates instead of exhausting the stack.
 *
 * Follows Excel semantics for type coercion (shared helpers in
 * `formula/coerce.ts`), error propagation, and operator behaviour (`&` keeps
 * the first 32,767 characters of its result, as Excel does).
 *
 * Every node is evaluated in scalar context.  A reference that covers more
 * than one cell (a range, a whole column or row, an intersection) reads one
 * cell by implicit intersection with {@link EvalContext.currentCell}, or is
 * `#VALUE!`; a union is `#VALUE!`; an array constant reads its top-left
 * element.  Functions that take whole areas resolve their range arguments
 * themselves (`rangeExtent` / `areaExtent`); the resolver's
 * `argToMatrix` / `argValues` are the shared helpers for that.  3-D and structured
 * references evaluate to `#REF!`; the evaluator does not resolve them.
 *
 * @module formula/evaluator
 */

import { coerceToString, excelPower, MAX_TEXT_LENGTH, toNumber } from './coerce.js';
import { compareText } from './collation.js';
import { isError } from './errors.js';
import type { ASTNode, BinaryOpNode, PercentNode } from './parser.js';
import { parseFormula } from './parser.js';
import type { CellValue, EvalContext, RangeExtent } from './resolver.js';
import { areaExtent, resolveRef } from './resolver.js';

// ---------------------------------------------------------------------------
// Comparison helpers
// ---------------------------------------------------------------------------

/** A non-blank cell value. */
type Value = Exclude<CellValue, null>;

/**
 * Type rank used for mixed-type comparison: number < text < boolean, as in
 * Excel.  Blanks never reach it (see {@link compareValues}).
 */
function typeRank(val: Value): number {
  switch (typeof val) {
    case 'number':
      return 1;
    case 'string':
      return 2;
    case 'boolean':
      return 3;
  }
}

/** The value a blank takes when compared with `other`: 0, `''` or FALSE, matching its type. */
function blankAs(other: Value): Value {
  switch (typeof other) {
    case 'number':
      return 0;
    case 'string':
      return '';
    case 'boolean':
      return false;
  }
}

/**
 * Compare two cell values following Excel rules:
 *   - A blank is coerced to the other operand's type first (0, `''` or
 *     FALSE), so a blank equals 0, `""` and FALSE; two blanks are equal.
 *   - Same type: natural comparison (text case-insensitive, en-US collation).
 *   - Mixed type: number < text < boolean.
 *
 * Returns negative / zero / positive like an `Array#sort` comparator.
 */
function compareValues(a: CellValue, b: CellValue): number {
  if (a === null) return b === null ? 0 : compareNonBlank(blankAs(b), b);
  if (b === null) return compareNonBlank(a, blankAs(a));
  return compareNonBlank(a, b);
}

/** {@link compareValues} for two non-blank values. */
function compareNonBlank(a: Value, b: Value): number {
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;

  // Same type
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') {
    return compareText(a.toLowerCase(), b.toLowerCase());
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
 * Dispatches on the node type and delegates cell lookups to the
 * {@link EvalContext}.  A `RangeError` raised while evaluating (call-stack
 * exhaustion on a very deep AST, or another engine limit) becomes `#NUM!`
 * instead of propagating.
 */
export function evaluateNode(node: ASTNode, ctx: EvalContext): CellValue {
  return stackSafe(() => evalNode(node, ctx));
}

/**
 * Run `evaluate`, turning a `RangeError` (call-stack exhaustion, or an
 * engine limit such as an over-long string) into `#NUM!`.  Every other error
 * propagates.
 */
function stackSafe(evaluate: () => CellValue): CellValue {
  try {
    return evaluate();
  } catch (err) {
    if (err instanceof RangeError) return '#NUM!';
    throw err;
  }
}

/** The recursive core of the evaluator; see {@link evaluateNode}. */
function evalNode(node: ASTNode, ctx: EvalContext): CellValue {
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

    case 'range':
    case 'column_range':
    case 'row_range':
    case 'structured_ref':
      return implicitIntersection(areaExtent(node, ctx), ctx);

    // -- defined names: the evaluator has no name table, so a name is #NAME? --
    case 'name':
      return '#NAME?';

    // -- an empty function argument is blank --------------------------------
    case 'missing':
      return null;

    // -- percent and value operators: one loop over the left-folded chain ----
    case 'percent':
      return evaluateChain(node, ctx);

    // -- unary operators ---------------------------------------------------
    case 'unary_op':
      return evaluateUnary(node.op, node.operand, ctx);

    // -- binary operators --------------------------------------------------
    // The reference operators work on areas, so they are handled before
    // their operands are evaluated as values.
    case 'binary_op':
      switch (node.op) {
        case ' ':
          return implicitIntersection(areaExtent(node, ctx), ctx);
        case ',': // a union has no single value
          return '#VALUE!';
        case ':':
          return evaluateRangeOperator(node, ctx);
        default:
          return evaluateChain(node, ctx);
      }

    // -- function calls ----------------------------------------------------
    case 'function':
      return evaluateFunction(node.name, node.args, ctx);

    // -- array constants ---------------------------------------------------
    case 'array': {
      // Scalar context reads the top-left element; `{}` has none.
      const first = node.rows[0]?.[0];
      return first === undefined ? '#VALUE!' : evalNode(first, ctx);
    }

    // A node type outside ASTNode (hand-built input).
    default:
      return '#VALUE!';
  }
}

/**
 * The value of a reference in scalar context.  A 1x1 area yields its cell.
 * A single-column area yields its cell on {@link EvalContext.currentCell}'s
 * row, and a single-row area its cell in that column (Excel's implicit
 * intersection).  Any other area, and any multi-cell area without
 * `currentCell`, is `#VALUE!`.  An error string (an unresolvable reference)
 * passes through, and a non-area (`undefined`) is `#VALUE!`.
 */
function implicitIntersection(area: RangeExtent | string | undefined, ctx: EvalContext): CellValue {
  if (area === undefined) return '#VALUE!';
  if (typeof area === 'string') return area;
  const { sheet, col, row, rows, cols } = area;
  if (rows === 1 && cols === 1) return ctx.getCell(sheet, col, row);
  const at = ctx.currentCell;
  if (at) {
    if (cols === 1 && at.row >= row && at.row < row + rows) return ctx.getCell(sheet, col, at.row);
    if (rows === 1 && at.col >= col && at.col < col + cols) return ctx.getCell(sheet, at.col, row);
  }
  return '#VALUE!';
}

/**
 * A `binary_op` `':'`.  The parser builds one for a range whose endpoints are
 * not two plain cell references: a name endpoint (`A1:XFE1`, `XFE1:XFE2`,
 * `Total:A1`), a `#REF!` endpoint (`#REF!:C1`), a function or parenthesised endpoint (`A2:INDEX(A:A,B1)`), or a
 * chain (`A1:B1:C1`, left-folded).  Names are not resolved, so the first name
 * endpoint's `#NAME?` (left to right, along the chain too) propagates, as in
 * Excel; an error endpoint (`#REF!:C1`, or an error a parenthesised endpoint
 * yields, `A1:(#N/A)`) propagates the same way, whichever of a name or an
 * error comes first.  With neither, the result is `#VALUE!`: function
 * endpoints are not resolved to references yet.
 */
function evaluateRangeOperator(node: BinaryOpNode, ctx: EvalContext): CellValue {
  const rights: ASTNode[] = [];
  let left: ASTNode = node;
  while (left.type === 'binary_op' && left.op === ':') {
    rights.push(left.right);
    left = left.left;
  }
  if (left.type === 'name' || left.type === 'error') return evalNode(left, ctx);
  for (const endpoint of rights.toReversed()) {
    if (endpoint.type === 'name' || endpoint.type === 'error') return evalNode(endpoint, ctx);
  }
  return '#VALUE!';
}

// ---------------------------------------------------------------------------
// Unary operator evaluation
// ---------------------------------------------------------------------------

/**
 * Prefix `+` returns its operand unchanged, text and booleans included
 * (OpenFormula 1.3 part 4, 6.4.15), so `+"abc"` is `"abc"`.  Prefix `-`
 * converts its operand to a number first.
 */
function evaluateUnary(op: string, operandNode: ASTNode, ctx: EvalContext): CellValue {
  const operand = evalNode(operandNode, ctx);
  if (isError(operand)) return operand;

  switch (op) {
    case '+':
      return operand;
    case '-': {
      const n = toNumber(operand);
      return typeof n === 'string' ? n : -n; // #VALUE! or other error
    }
    default:
      return '#VALUE!';
  }
}

// ---------------------------------------------------------------------------
// Binary operator evaluation
// ---------------------------------------------------------------------------

/**
 * `l & r` as Excel computes it: the result keeps only its first 32,767
 * characters (Excel's text-length cap), so an over-long string is never built.
 */
function concatText(l: string, r: string): string {
  if (l.length + r.length <= MAX_TEXT_LENGTH) return l + r;
  if (l.length >= MAX_TEXT_LENGTH) return l.slice(0, MAX_TEXT_LENGTH);
  return l + r.slice(0, MAX_TEXT_LENGTH - l.length);
}

/** `true` for a node {@link evaluateChain} folds: a `%` or a value (non-reference) binary operator. */
function isChainLink(node: ASTNode): node is BinaryOpNode | PercentNode {
  if (node.type === 'percent') return true;
  return node.type === 'binary_op' && node.op !== ' ' && node.op !== ',' && node.op !== ':';
}

/**
 * Evaluate a left-folded chain of value operators and `%` with a loop.  The
 * parser folds `1+2+3` into `((1+2)+3)` and `A1%%` into `(A1%)%`, so a long
 * chain is thousands of levels deep within the 8192-character limit; only
 * right operands (and the innermost left one) recurse.  Each operator still
 * evaluates its left operand, then its right, and reports the left error
 * first.
 */
function evaluateChain(node: BinaryOpNode | PercentNode, ctx: EvalContext): CellValue {
  const links: (BinaryOpNode | PercentNode)[] = [];
  let innermost: ASTNode = node;
  while (isChainLink(innermost)) {
    links.push(innermost);
    innermost = innermost.type === 'percent' ? innermost.operand : innermost.left;
  }
  let acc = evalNode(innermost, ctx);
  for (const link of links.toReversed()) {
    acc =
      link.type === 'percent'
        ? applyPercent(acc)
        : applyBinary(link.op, acc, evalNode(link.right, ctx));
  }
  return acc;
}

/** Postfix `%`: the operand as a number, divided by 100; an error propagates. */
function applyPercent(operand: CellValue): CellValue {
  const n = toNumber(operand);
  return typeof n === 'string' ? n : n / 100;
}

/** Apply a value operator to its evaluated operands. */
function applyBinary(op: string, left: CellValue, right: CellValue): CellValue {
  // Error propagation (left first, then right).
  if (isError(left)) return left;
  if (isError(right)) return right;

  // -- concatenation -------------------------------------------------------
  if (op === '&') return concatText(coerceToString(left), coerceToString(right));

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
      return excelPower(a, b);
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
    return fn(args, ctx, evalNode);
  }
  return '#NAME?';
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Parse and evaluate an Excel formula string.
 *
 * A formula with parse errors evaluates to `#VALUE!`; a `RangeError` raised
 * while evaluating becomes `#NUM!` instead of propagating.
 *
 * @param formula - Formula string, optionally starting with `=`.
 * @param ctx     - Evaluation context providing cell lookups.
 * @returns The computed cell value.
 */
export function evaluateFormula(formula: string, ctx: EvalContext): CellValue {
  return stackSafe(() => {
    const { ast, errors } = parseFormula(formula);
    // A formula with parse errors is #VALUE!; an empty formula is blank.
    if (errors.length > 0) return '#VALUE!';
    if (!ast) return null;
    return evalNode(ast, ctx);
  });
}
