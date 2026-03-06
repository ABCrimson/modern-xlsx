/**
 * Shared formula expansion: adjusts relative references from a master cell
 * to a child cell by applying row/column deltas.
 *
 * In OOXML, shared formulas store the formula once in a master cell and only
 * record the `si` (shared index) in child cells. This module reconstructs the
 * child formula by parsing the master, shifting relative references, and
 * serializing back.
 *
 * @module formula/shared
 */

import { columnToLetter, decodeCellRef, letterToColumn } from '../cell-ref.js';
import type { ASTNode, CellRefNode } from './parser.js';
import { parseFormula } from './parser.js';
import { serializeFormula } from './serializer.js';

// ---------------------------------------------------------------------------
// AST transformation
// ---------------------------------------------------------------------------

/** Shift a single cell reference by the given row/col deltas, respecting absolute flags. */
function shiftCellRef(ref: CellRefNode, rowDelta: number, colDelta: number): ASTNode {
  let newRow = ref.row;
  let newCol = ref.col;
  let invalid = false;

  if (!ref.absRow) {
    newRow = ref.row + rowDelta;
    if (newRow < 1) {
      invalid = true;
    }
  }

  if (!ref.absCol) {
    const colIdx = letterToColumn(ref.col) + colDelta;
    if (colIdx < 0) {
      invalid = true;
    } else {
      newCol = columnToLetter(colIdx);
    }
  }

  if (invalid) {
    return { type: 'error', value: '#REF!' };
  }

  const result: CellRefNode = {
    type: 'cell_ref',
    col: newCol,
    row: newRow,
    absCol: ref.absCol,
    absRow: ref.absRow,
  };
  if (ref.sheet !== undefined) {
    result.sheet = ref.sheet;
  }
  return result;
}

/** Deep-clone and shift all cell references in an AST node. */
function shiftNode(node: ASTNode, rowDelta: number, colDelta: number): ASTNode {
  switch (node.type) {
    case 'cell_ref':
      return shiftCellRef(node, rowDelta, colDelta);

    case 'range': {
      const start = shiftCellRef(node.start, rowDelta, colDelta);
      const end = shiftCellRef(node.end, rowDelta, colDelta);
      // If either endpoint became an error, the range becomes partial error
      if (start.type === 'error' || end.type === 'error') {
        // Serialize as error range (e.g., #REF!:A12 or A1:#REF!)
        // but we need to keep as a range node with the error substituted
        return {
          type: 'range',
          start:
            start.type === 'cell_ref'
              ? start
              : { type: 'cell_ref', col: '', row: 0, absCol: false, absRow: false },
          end:
            end.type === 'cell_ref'
              ? end
              : { type: 'cell_ref', col: '', row: 0, absCol: false, absRow: false },
        };
      }
      if (start.type !== 'cell_ref' || end.type !== 'cell_ref') {
        return { type: 'error', value: '#REF!' };
      }
      return { type: 'range', start, end };
    }

    case 'function':
      return {
        type: 'function',
        name: node.name,
        args: node.args.map((arg) => shiftNode(arg, rowDelta, colDelta)),
      };

    case 'binary_op':
      return {
        type: 'binary_op',
        op: node.op,
        left: shiftNode(node.left, rowDelta, colDelta),
        right: shiftNode(node.right, rowDelta, colDelta),
      };

    case 'unary_op':
      return {
        type: 'unary_op',
        op: node.op,
        operand: shiftNode(node.operand, rowDelta, colDelta),
      };

    case 'percent':
      return {
        type: 'percent',
        operand: shiftNode(node.operand, rowDelta, colDelta),
      };

    case 'array':
      return {
        type: 'array',
        rows: node.rows.map((row) => row.map((cell) => shiftNode(cell, rowDelta, colDelta))),
      };

    // Leaf nodes
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
    case 'name':
      return node;
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Expand a shared formula from its master cell to a child cell.
 *
 * Relative references are shifted by the row/column difference between
 * masterRef and childRef. Absolute references ($A$1) are preserved.
 *
 * @param masterFormula - The formula string in the master cell.
 * @param masterRef - The master cell reference (e.g., "B1").
 * @param childRef - The child cell reference (e.g., "B3").
 * @returns The expanded formula string for the child cell.
 */
export function expandSharedFormula(
  masterFormula: string,
  masterRef: string,
  childRef: string,
): string {
  const master = decodeCellRef(masterRef);
  const child = decodeCellRef(childRef);

  // decodeCellRef returns 0-based indices, but our AST uses 1-based rows
  // and letter columns. Deltas are the same regardless of base.
  const rowDelta = child.row - master.row;
  const colDelta = child.col - master.col;

  if (rowDelta === 0 && colDelta === 0) {
    return masterFormula;
  }

  const { ast, errors } = parseFormula(masterFormula);
  if (ast === null || errors.length > 0) {
    return masterFormula;
  }

  const shifted = shiftNode(ast, rowDelta, colDelta);
  return serializeFormula(shifted);
}
