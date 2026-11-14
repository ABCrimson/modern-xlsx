/**
 * Formula serializer: converts an AST back into an Excel formula string.
 *
 * Handles operator precedence to emit minimal parentheses while preserving
 * correct evaluation order. Every binary operator is left-associative, as in
 * Excel, so only an equal-precedence right operand needs parentheses
 * (`2^(3^4)`, `A1-(B1-C1)`). A union (`,`) is written bare at the root and in
 * one pair of parentheses anywhere else (`SUM((A1,B2,C3))`).
 *
 * @module formula/serializer
 */

import type { ASTNode, BinaryOpNode, PercentNode } from './parser.js';

// ---------------------------------------------------------------------------
// Operator precedence (higher = tighter binding)
// ---------------------------------------------------------------------------

/**
 * Binary operator precedence, lowest to highest. The union operator `,` is
 * not listed: it binds loosest and is parenthesised on its own terms.
 */
const PRECEDENCE = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
  ' ': 6, // intersection
  ':': 7, // range
} as const satisfies Record<string, number>;

function hasPrecedence(op: string): op is keyof typeof PRECEDENCE {
  return op in PRECEDENCE;
}

function precedence(op: string): number {
  return hasPrecedence(op) ? PRECEDENCE[op] : 0;
}

/** `true` for a `binary_op` node with the union operator `,`. */
function isUnion(node: ASTNode): boolean {
  return node.type === 'binary_op' && node.op === ',';
}

// ---------------------------------------------------------------------------
// Sheet name quoting
// ---------------------------------------------------------------------------

/** A character that cannot appear in an unquoted sheet prefix. */
const UNSAFE_SHEET_CHAR = /[^A-Za-z0-9_.]/;

/** A1-style cell reference shape: 1-3 letters then digits (`Q1`, `xfd1048576`). */
const A1_REF_SHAPE = /^[A-Za-z]{1,3}[0-9]+$/;

/** R1C1-style reference shape, relative parts included (`R`, `C3`, `RC`, `R2C`, `R1C1`). */
const R1C1_REF_SHAPE = /^(?=.)(?:R[0-9]*)?(?:C[0-9]*)?$/i;

/** A name that starts with an `R<n>`, `C<n>` or `RC<n>` reference (`R2D2`, `C3PO`, `RC1x`), which Excel requires quoted. */
const R1C1_REF_PREFIX = /^(?:RC?|C)[0-9]/i;

/** `true` when Excel would misread `name` as a bare sheet prefix. */
function needsQuoting(name: string): boolean {
  const upper = name.toUpperCase();
  return (
    UNSAFE_SHEET_CHAR.test(name) ||
    /^[0-9.]/.test(name) ||
    A1_REF_SHAPE.test(name) ||
    R1C1_REF_SHAPE.test(name) ||
    R1C1_REF_PREFIX.test(name) ||
    upper === 'TRUE' ||
    upper === 'FALSE'
  );
}

/** Wrap `text` in `'...'`, doubling every inner `'`. */
function quote(text: string): string {
  return `'${text.replaceAll("'", "''")}'`;
}

/**
 * Format a sheet name as a reference prefix (without the `!`), quoting it
 * when Excel would misread it bare.
 *
 * The name is wrapped in `'...'`, with every inner `'` doubled, when it has a
 * character outside `[A-Za-z0-9_.]` (non-ASCII letters included: the
 * tokenizer also reads them bare, but the quoted form is always valid),
 * starts with a digit or `.`, has the shape of an A1 or R1C1 reference
 * (`Q1`, `R1C1`, `C`), starts with an
 * `R<n>`, `C<n>` or `RC<n>` reference (`R2D2`, `C3PO`, `RC1x`), or is `TRUE` or `FALSE` in any case.
 * Otherwise it is returned unchanged.
 *
 * @param name - The sheet name, unescaped.
 * @returns The prefix text, e.g. `Sheet1` or `'My Data'`.
 *
 * @example
 * ```ts
 * quoteSheetName('Sheet1'); // 'Sheet1'
 * quoteSheetName('Q1'); // "'Q1'"
 * quoteSheetName("O'Brien"); // "'O''Brien'"
 * ```
 */
export function quoteSheetName(name: string): string {
  return needsQuoting(name) ? quote(name) : name;
}

/**
 * The `Sheet!` prefix of a reference node, or `''` without a sheet. A 3-D
 * span is quoted as one unit (`'Sheet 1:Sheet3'!`) when either name needs it.
 */
function sheetPrefix(node: { sheet?: string; endSheet?: string }): string {
  const { sheet, endSheet } = node;
  if (sheet === undefined) return '';
  if (endSheet === undefined) return `${quoteSheetName(sheet)}!`;
  const span = `${sheet}:${endSheet}`;
  return needsQuoting(sheet) || needsQuoting(endSheet) ? `${quote(span)}!` : `${span}!`;
}

/** `$` for an absolute part, `''` for a relative one. */
function abs(absolute: boolean): string {
  return absolute ? '$' : '';
}

// ---------------------------------------------------------------------------
// Core serializer
// ---------------------------------------------------------------------------

/**
 * Serialize an AST node into an Excel formula string.
 *
 * Roundtrip property: for well-formed formulas,
 * `serializeFormula(parseFormula(f).ast)` produces a semantically equivalent formula.
 * Whole-column, whole-row, structured and 3-D references are written as they
 * were parsed; a missing argument is written as nothing (`IF(A1,,2)`).
 */
export function serializeFormula(ast: ASTNode): string {
  return serializeNode(ast, true);
}

/** Serialize `ast`; `atRoot` is `true` only for the whole formula, where a union needs no parentheses. */
function serializeNode(ast: ASTNode, atRoot: boolean): string {
  switch (ast.type) {
    case 'number':
      return String(ast.value);

    case 'string':
      return `"${ast.value.replaceAll('"', '""')}"`;

    case 'boolean':
      return ast.value ? 'TRUE' : 'FALSE';

    case 'error':
      return ast.value;

    case 'cell_ref':
      return `${sheetPrefix(ast)}${abs(ast.absCol)}${ast.col}${abs(ast.absRow)}${ast.row}`;

    case 'range':
      return `${serializeNode(ast.start, false)}:${serializeNode(ast.end, false)}`;

    case 'column_range':
      return `${sheetPrefix(ast)}${abs(ast.absStart)}${ast.startCol}:${abs(ast.absEnd)}${ast.endCol}`;

    case 'row_range':
      return `${sheetPrefix(ast)}${abs(ast.absStart)}${ast.startRow}:${abs(ast.absEnd)}${ast.endRow}`;

    case 'structured_ref':
      return ast.text;

    case 'missing':
      return '';

    case 'name':
      return ast.name;

    case 'function':
      return `${ast.name}(${ast.args.map((arg) => serializeNode(arg, false)).join(',')})`;

    case 'binary_op': {
      if (ast.op === ',') {
        const list = serializeUnionList(ast);
        return atRoot ? list : `(${list})`;
      }
      return serializeBinaryChain(ast);
    }

    case 'unary_op': {
      const operand = ast.operand;
      if (operand.type === 'binary_op' || operand.type === 'unary_op') {
        return `${ast.op}${parenthesize(operand)}`;
      }
      return `${ast.op}${serializeNode(operand, false)}`;
    }

    case 'percent':
      return serializePercent(ast);

    case 'array':
      return `{${ast.rows
        .map((row) => row.map((el) => serializeNode(el, false)).join(','))
        .join(';')}}`;
  }
}

/**
 * A postfix `%`. A `%` chain (`A1%%`) nests one level per `%`, thousands deep
 * within 8192 characters, so it is counted with a loop; an operator operand
 * is parenthesised (`(A1+B1)%`).
 */
function serializePercent(node: PercentNode): string {
  let operand = node.operand;
  let percents = 1;
  while (operand.type === 'percent') {
    operand = operand.operand;
    percents++;
  }
  const suffix = '%'.repeat(percents);
  if (operand.type === 'binary_op' || operand.type === 'unary_op') {
    return `${parenthesize(operand)}${suffix}`;
  }
  return `${serializeNode(operand, false)}${suffix}`;
}

/** Serialize `node` in parentheses; a union brings its own pair. */
function parenthesize(node: ASTNode): string {
  const s = serializeNode(node, false);
  return isUnion(node) ? s : `(${s})`;
}

/**
 * The comma-separated operands of a union, without the enclosing
 * parentheses. A left-nested chain `((A1,B2),C3)` flattens to `A1,B2,C3`;
 * a union in the right operand keeps its own parentheses.
 */
function serializeUnionList(node: BinaryOpNode): string {
  const operands: ASTNode[] = [];
  let current: ASTNode = node;
  while (current.type === 'binary_op' && current.op === ',') {
    operands.push(current.right);
    current = current.left;
  }
  operands.push(current);
  return operands
    .reverse()
    .map((operand) => serializeNode(operand, false))
    .join(',');
}

/**
 * A non-union binary operation. The parser folds `A1+B1+C1` into
 * `((A1+B1)+C1)`, so a long chain is thousands of levels deep; the left
 * operands that need no parentheses are walked with a loop, and only right
 * operands (and the innermost left one) recurse.
 */
function serializeBinaryChain(node: BinaryOpNode): string {
  const chain: BinaryOpNode[] = [node];
  let deepest = node;
  while (continuesChain(deepest.left, deepest.op)) {
    deepest = deepest.left;
    chain.push(deepest);
  }
  let out = wrapBinaryChild(deepest.left, deepest.op, 'left');
  for (const link of chain.toReversed()) {
    out += `${link.op}${wrapBinaryChild(link.right, link.op, 'right')}`;
  }
  return out;
}

/** `true` when a left operand is a non-union binary_op written without parentheses. */
function continuesChain(left: ASTNode, parentOp: string): left is BinaryOpNode {
  return left.type === 'binary_op' && !isUnion(left) && !needsParens(left, parentOp, 'left');
}

/**
 * Whether a child of a binary operator needs parentheses.
 *
 * Every binary operator is left-associative, so a child needs parens if:
 * - It is a binary_op with lower precedence than the parent.
 * - It is the right child and a binary_op with **equal** precedence
 *   (`A1-(B1-C1)`, `2^(3^4)`).
 * - It is a unary_op or percent under a reference operator (`:` or ` `),
 *   which bind tighter than unary `-` and `%` (`(-A1) B1`).
 *
 * A union child never does here: {@link serializeNode} parenthesises it.
 */
function needsParens(child: ASTNode, parentOp: string, side: 'left' | 'right'): boolean {
  if (isUnion(child)) return false;

  const parentPrec = precedence(parentOp);

  if (child.type === 'unary_op' || child.type === 'percent') {
    return parentPrec > PRECEDENCE['^'];
  }
  if (child.type !== 'binary_op') return false;

  const childPrec = precedence(child.op);
  return childPrec < parentPrec || (childPrec === parentPrec && side === 'right');
}

/** Serialize a child of a binary operator, in parentheses when {@link needsParens} says so. */
function wrapBinaryChild(child: ASTNode, parentOp: string, side: 'left' | 'right'): string {
  const s = serializeNode(child, false);
  return needsParens(child, parentOp, side) ? `(${s})` : s;
}
