/**
 * Shared formula expansion: adjusts relative references from a master cell
 * to a child cell by applying row/column deltas.
 *
 * In OOXML, shared formulas store the formula once in a master cell and only
 * record the `si` (shared index) in child cells. This module reconstructs the
 * child formula by parsing the master, shifting relative references, and
 * serializing back. A reference shifted off the sheet (above row 1, left of
 * column A, past row 1,048,576 or column XFD) becomes `#REF!`; a range with
 * either end off the sheet becomes one `#REF!`, as Excel writes it.
 *
 * @module formula/shared
 */

import { columnIndexOrNull, columnToLetter, decodeCellRef } from '../format/cell-ref.js';
import { MAX_COL_INDEX, MAX_ROW } from './grid.js';
import type { ASTNode, CellRefNode, ColumnRangeNode, RowRangeNode } from './parser.js';
import { parseFormula } from './parser.js';
import { serializeFormula } from './serializer.js';
import { mapRefs } from './walk.js';

// ---------------------------------------------------------------------------
// AST transformation
// ---------------------------------------------------------------------------

const REF_ERROR: ASTNode = { type: 'error', value: '#REF!' };

/**
 * Shift one part of a reference by `delta` unless it is `absolute`.
 *
 * @returns The new value, or `null` when it leaves `min..max`.
 */
function shiftPart(
  value: number,
  delta: number,
  absolute: boolean,
  min: number,
  max: number,
): number | null {
  if (absolute) return value;
  const next = value + delta;
  return next < min || next > max ? null : next;
}

/** Shift a column's letters by `delta` unless `absolute`; `null` when it leaves `A..XFD`. */
function shiftColumn(letters: string, delta: number, absolute: boolean): string | null {
  const idx = columnIndexOrNull(letters);
  if (idx === null) return null;
  const next = shiftPart(idx, delta, absolute, 0, MAX_COL_INDEX);
  return next === null ? null : columnToLetter(next);
}

/**
 * Shift a single cell reference by the given row/col deltas, respecting
 * absolute flags. Sheet and 3-D prefixes are kept.
 */
function shiftCellRef(ref: CellRefNode, rowDelta: number, colDelta: number): ASTNode {
  const row = shiftPart(ref.row, rowDelta, ref.absRow, 1, MAX_ROW);
  const col = shiftColumn(ref.col, colDelta, ref.absCol);
  return row === null || col === null ? REF_ERROR : { ...ref, row, col };
}

/** Shift the relative columns of a whole-column range (`A:C`). */
function shiftColumnRange(node: ColumnRangeNode, colDelta: number): ASTNode {
  const startCol = shiftColumn(node.startCol, colDelta, node.absStart);
  const endCol = shiftColumn(node.endCol, colDelta, node.absEnd);
  return startCol === null || endCol === null ? REF_ERROR : { ...node, startCol, endCol };
}

/** Shift the relative rows of a whole-row range (`1:3`). */
function shiftRowRange(node: RowRangeNode, rowDelta: number): ASTNode {
  const startRow = shiftPart(node.startRow, rowDelta, node.absStart, 1, MAX_ROW);
  const endRow = shiftPart(node.endRow, rowDelta, node.absEnd, 1, MAX_ROW);
  return startRow === null || endRow === null ? REF_ERROR : { ...node, startRow, endRow };
}

/** Rebuild an AST node with every cell reference shifted (unchanged leaves are shared). */
function shiftNode(node: ASTNode, rowDelta: number, colDelta: number): ASTNode {
  return mapRefs(node, {
    cellRef: (ref) => shiftCellRef(ref, rowDelta, colDelta),
    range: (range) => {
      const start = shiftCellRef(range.start, rowDelta, colDelta);
      const end = shiftCellRef(range.end, rowDelta, colDelta);
      // Either end off the sheet: Excel writes the whole range as #REF!.
      if (start.type !== 'cell_ref' || end.type !== 'cell_ref') {
        return REF_ERROR;
      }
      return { type: 'range', start, end };
    },
    columnRange: (range) => shiftColumnRange(range, colDelta),
    rowRange: (range) => shiftRowRange(range, rowDelta),
  });
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
 * @returns The expanded formula string for the child cell, or `masterFormula`
 *   unchanged when it does not parse or shifting it exhausts the call stack.
 * @throws {ModernXlsxError} INVALID_CELL_REF when masterRef or childRef is not an uppercase A1 reference inside A1:XFD1048576; shifts past the sheet edge never throw and give #REF!
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

  try {
    return serializeFormula(shiftNode(ast, rowDelta, colDelta));
  } catch (err) {
    // Backstop: operator chains are walked with loops, but the remaining
    // recursion (parentheses, functions, arrays: up to the parser's 256
    // levels) can still exhaust a small stack. Return the input unchanged,
    // as for a parse error.
    if (err instanceof RangeError) return masterFormula;
    throw err;
  }
}
