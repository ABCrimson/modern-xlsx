/**
 * Formula serializer: converts an AST back into an Excel formula string.
 *
 * Handles operator precedence to emit minimal parentheses while preserving
 * correct evaluation order.
 *
 * @module formula/serializer
 */

import type { ASTNode, BinaryOpNode } from './parser.js';

// ---------------------------------------------------------------------------
// Operator precedence (higher = tighter binding)
// ---------------------------------------------------------------------------

const PRECEDENCE: Record<string, number> = {
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
};

function precedence(op: string): number {
  return PRECEDENCE[op] ?? 0;
}

// ---------------------------------------------------------------------------
// Sheet name quoting
// ---------------------------------------------------------------------------

/** Returns true if a sheet name needs to be quoted (contains spaces or special chars). */
function needsQuoting(name: string): boolean {
  return /[^A-Za-z0-9_]/.test(name);
}

/** Quote a sheet name for use in a cell reference prefix. */
function quoteSheet(name: string): string {
  if (needsQuoting(name)) {
    return `'${name.replace(/'/g, "''")}'`;
  }
  return name;
}

// ---------------------------------------------------------------------------
// Core serializer
// ---------------------------------------------------------------------------

/**
 * Serialize an AST node into an Excel formula string.
 *
 * Roundtrip property: for well-formed formulas,
 * `serializeFormula(parseFormula(f).ast)` produces a semantically equivalent formula.
 */
export function serializeFormula(ast: ASTNode): string {
  switch (ast.type) {
    case 'number':
      return String(ast.value);

    case 'string':
      return `"${ast.value.replace(/"/g, '""')}"`;

    case 'boolean':
      return ast.value ? 'TRUE' : 'FALSE';

    case 'error':
      return ast.value;

    case 'cell_ref': {
      let result = '';
      if (ast.sheet !== undefined) {
        result += `${quoteSheet(ast.sheet)}!`;
      }
      if (ast.absCol) result += '$';
      result += ast.col;
      if (ast.absRow) result += '$';
      result += String(ast.row);
      return result;
    }

    case 'range':
      return `${serializeFormula(ast.start)}:${serializeFormula(ast.end)}`;

    case 'name':
      return ast.name;

    case 'function':
      return `${ast.name}(${ast.args.map(serializeFormula).join(',')})`;

    case 'binary_op': {
      const parentPrec = precedence(ast.op);
      const leftStr = wrapBinaryChild(ast.left, parentPrec, 'left', ast.op);
      const rightStr = wrapBinaryChild(ast.right, parentPrec, 'right', ast.op);
      return `${leftStr}${ast.op}${rightStr}`;
    }

    case 'unary_op': {
      const operand = ast.operand;
      if (operand.type === 'binary_op' || operand.type === 'unary_op') {
        return `${ast.op}(${serializeFormula(operand)})`;
      }
      return `${ast.op}${serializeFormula(operand)}`;
    }

    case 'percent': {
      const operand = ast.operand;
      if (operand.type === 'binary_op' || operand.type === 'unary_op') {
        return `(${serializeFormula(operand)})%`;
      }
      return `${serializeFormula(operand)}%`;
    }

    case 'array':
      return `{${ast.rows.map((row) => row.map(serializeFormula).join(',')).join(';')}}`;
  }
}

/**
 * Wrap a child of a binary operator in parentheses if needed.
 *
 * A child needs parens if:
 * - It is a binary_op with lower precedence than the parent.
 * - It is a right child binary_op with **equal** precedence for left-associative
 *   operators (to preserve evaluation order). Exception: `^` is right-associative.
 */
function wrapBinaryChild(
  child: ASTNode,
  parentPrec: number,
  side: 'left' | 'right',
  parentOp: string,
): string {
  const s = serializeFormula(child);
  if (child.type !== 'binary_op') return s;

  const childPrec = precedence((child as BinaryOpNode).op);

  if (childPrec < parentPrec) {
    return `(${s})`;
  }

  // For equal precedence on the right side of a left-associative operator, add parens.
  // `^` is right-associative, so no parens needed for equal precedence on the right.
  if (childPrec === parentPrec && side === 'right' && parentOp !== '^') {
    return `(${s})`;
  }

  return s;
}
