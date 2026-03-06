/**
 * Formula reference rewriter for row/column insert and delete operations.
 *
 * Parses a formula to AST, walks and transforms cell references according
 * to the specified action, then serializes back to a formula string.
 *
 * @module formula/rewriter
 */

import { columnToLetter, letterToColumn } from '../cell-ref.js';
import type { ASTNode, CellRefNode } from './parser.js';
import { parseFormula } from './parser.js';
import { serializeFormula } from './serializer.js';

// ---------------------------------------------------------------------------
// Action types
// ---------------------------------------------------------------------------

export type RewriteAction =
  | { type: 'insert_rows'; sheet?: string; start: number; count: number }
  | { type: 'delete_rows'; sheet?: string; start: number; count: number }
  | { type: 'insert_cols'; sheet?: string; start: number; count: number }
  | { type: 'delete_cols'; sheet?: string; start: number; count: number };

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Check whether a cell ref's sheet matches the action's target sheet. */
function sheetMatches(ref: CellRefNode, action: RewriteAction): boolean {
  // If action has no sheet, it applies to the current (unqualified) sheet.
  // Unqualified refs (no sheet) match when action has no sheet.
  // Sheet-qualified refs match when action sheet matches ref sheet.
  if (action.sheet === undefined) {
    return ref.sheet === undefined;
  }
  return ref.sheet === action.sheet;
}

// ---------------------------------------------------------------------------
// AST transformation
// ---------------------------------------------------------------------------

/** Deep-clone and transform an AST node, adjusting cell references per the action. */
function transformNode(node: ASTNode, action: RewriteAction): ASTNode {
  switch (node.type) {
    case 'cell_ref':
      return transformCellRef(node, action);

    case 'range': {
      const start = transformCellRef(node.start, action);
      const end = transformCellRef(node.end, action);
      // If both endpoints are still cell refs, return a proper range node
      if (start.type === 'cell_ref' && end.type === 'cell_ref') {
        return { type: 'range', start, end };
      }
      // If either became #REF!, build a synthetic binary ':' to preserve the shape
      return {
        type: 'binary_op',
        op: ':',
        left: start,
        right: end,
      };
    }

    case 'function':
      return {
        type: 'function',
        name: node.name,
        args: node.args.map((arg) => transformNode(arg, action)),
      };

    case 'binary_op':
      return {
        type: 'binary_op',
        op: node.op,
        left: transformNode(node.left, action),
        right: transformNode(node.right, action),
      };

    case 'unary_op':
      return {
        type: 'unary_op',
        op: node.op,
        operand: transformNode(node.operand, action),
      };

    case 'percent':
      return {
        type: 'percent',
        operand: transformNode(node.operand, action),
      };

    case 'array':
      return {
        type: 'array',
        rows: node.rows.map((row) => row.map((cell) => transformNode(cell, action))),
      };

    // Leaf nodes that don't contain references
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
    case 'name':
      return node;
  }
}

const REF_ERROR: ASTNode = { type: 'error', value: '#REF!' };

/** Adjust a row number for an insert or delete action. Returns null for #REF!. */
function adjustRow(row: number, start: number, count: number, isDelete: boolean): number | null {
  if (isDelete) {
    if (row >= start && row < start + count) return null;
    return row >= start + count ? row - count : row;
  }
  return row >= start ? row + count : row;
}

/** Adjust a column index for an insert or delete action. Returns null for #REF!. */
function adjustCol(colIdx: number, start: number, count: number, isDelete: boolean): number | null {
  if (isDelete) {
    if (colIdx >= start && colIdx < start + count) return null;
    return colIdx >= start + count ? colIdx - count : colIdx;
  }
  return colIdx >= start ? colIdx + count : colIdx;
}

/** Transform a single cell reference according to the action. */
function transformCellRef(ref: CellRefNode, action: RewriteAction): CellRefNode | ASTNode {
  if (!sheetMatches(ref, action)) {
    return ref;
  }

  switch (action.type) {
    case 'insert_rows':
    case 'delete_rows': {
      const newRow = adjustRow(ref.row, action.start, action.count, action.type === 'delete_rows');
      if (newRow === null) return REF_ERROR;
      return newRow === ref.row ? ref : { ...ref, row: newRow };
    }
    case 'insert_cols':
    case 'delete_cols': {
      const colIdx = letterToColumn(ref.col);
      const newCol = adjustCol(colIdx, action.start, action.count, action.type === 'delete_cols');
      if (newCol === null) return REF_ERROR;
      return newCol === colIdx ? ref : { ...ref, col: columnToLetter(newCol) };
    }
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Rewrite cell references in a formula after a row/column insert or delete.
 *
 * @param formula - Formula string (with or without leading `=`).
 * @param action - The insert/delete action to apply.
 * @returns The rewritten formula string.
 */
export function rewriteFormula(formula: string, action: RewriteAction): string {
  const { ast, errors } = parseFormula(formula);
  if (ast === null || errors.length > 0) {
    return formula; // Return unchanged on parse error
  }
  const transformed = transformNode(ast, action);
  return serializeFormula(transformed);
}
