import { describe, expect, it, vi } from 'vitest';
import { argToMatrix, argValues, isReferenceNode } from '../src/formula/resolver.js';
import type { ASTNode, CellRefNode, CellValue, EvalContext, RangeNode } from '../src/index.js';
import { evaluateNode, parseFormula, resolveRange, resolveRef } from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build a mock EvalContext from a sheet -> row/col -> value map. */
function mockContext(
  data: Record<string, Record<string, CellValue>>,
  currentSheet = 'Sheet1',
): EvalContext {
  return {
    currentSheet,
    getCell(sheet: string, col: number, row: number): CellValue {
      const key = `${col},${row}`;
      return data[sheet]?.[key] ?? null;
    },
  };
}

/** Shorthand for building a CellRefNode. */
function cellRef(col: string, row: number, sheet?: string): CellRefNode {
  return { type: 'cell_ref', col, row, absCol: false, absRow: false, ...(sheet ? { sheet } : {}) };
}

/** Shorthand for building a RangeNode. */
function range(
  startCol: string,
  startRow: number,
  endCol: string,
  endRow: number,
  sheet?: string,
): RangeNode {
  return {
    type: 'range',
    start: cellRef(startCol, startRow, sheet),
    end: cellRef(endCol, endRow, sheet),
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('resolveRef', () => {
  it('resolves a cell on the current sheet', () => {
    const ctx = mockContext({ Sheet1: { '0,1': 42 } });
    expect(resolveRef(cellRef('A', 1), ctx)).toBe(42);
  });

  it('resolves a cell on an explicit sheet', () => {
    const ctx = mockContext({ Data: { '2,3': 'hello' } });
    expect(resolveRef(cellRef('C', 3, 'Data'), ctx)).toBe('hello');
  });

  it('returns null for a missing cell', () => {
    const ctx = mockContext({ Sheet1: {} });
    expect(resolveRef(cellRef('Z', 99), ctx)).toBe(null);
  });

  it('returns null for a missing sheet', () => {
    const ctx = mockContext({});
    expect(resolveRef(cellRef('A', 1, 'NoSheet'), ctx)).toBe(null);
  });

  it('resolves boolean values', () => {
    const ctx = mockContext({ Sheet1: { '0,1': true } });
    expect(resolveRef(cellRef('A', 1), ctx)).toBe(true);
  });

  it('resolves multi-letter column (AA = col 26)', () => {
    const ctx = mockContext({ Sheet1: { '26,1': 100 } });
    expect(resolveRef(cellRef('AA', 1), ctx)).toBe(100);
  });

  it('falls back to currentSheet when no sheet on ref', () => {
    const ctx = mockContext({ MySheet: { '1,2': 'yes' } }, 'MySheet');
    expect(resolveRef(cellRef('B', 2), ctx)).toBe('yes');
  });

  it('resolves lowercase columns in hand-built nodes', () => {
    const getCell = vi.fn<EvalContext['getCell']>(() => 7);
    const ctx: EvalContext = { currentSheet: 'Sheet1', getCell };
    expect(resolveRef(cellRef('a', 1), ctx)).toBe(7);
    expect(resolveRef(cellRef('xfd', 1048576), ctx)).toBe(7);
    expect(getCell.mock.calls).toEqual([
      ['Sheet1', 0, 1],
      ['Sheet1', 16383, 1048576],
    ]);
  });

  it.each([
    ['column past XFD', cellRef('XFE', 1)],
    ['row 0', cellRef('A', 0)],
    ['row 1048577', cellRef('A', 1048577)],
    ['empty column', cellRef('', 1)],
    ['non-letter column', cellRef('A1', 1)],
    ['fractional row', cellRef('A', 1.5)],
  ])('returns #REF! without calling getCell for a %s', (_label, node) => {
    const getCell = vi.fn<EvalContext['getCell']>(() => 1);
    expect(resolveRef(node, { currentSheet: 'Sheet1', getCell })).toBe('#REF!');
    expect(getCell).not.toHaveBeenCalled();
  });
});

describe('resolveRange', () => {
  const data: Record<string, Record<string, CellValue>> = {
    Sheet1: {
      '0,1': 1,
      '1,1': 2,
      '2,1': 3,
      '0,2': 4,
      '1,2': 5,
      '2,2': 6,
    },
  };

  it('resolves a simple 2x3 range', () => {
    const ctx = mockContext(data);
    const result = resolveRange(range('A', 1, 'C', 2), ctx);
    expect(result).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
  });

  it('resolves a single-cell range', () => {
    const ctx = mockContext(data);
    const result = resolveRange(range('B', 1, 'B', 1), ctx);
    expect(result).toEqual([[2]]);
  });

  it('normalises reversed ranges (end before start)', () => {
    const ctx = mockContext(data);
    const result = resolveRange(range('C', 2, 'A', 1), ctx);
    expect(result).toEqual([
      [1, 2, 3],
      [4, 5, 6],
    ]);
  });

  it('returns null for empty cells within the range', () => {
    const ctx = mockContext({ Sheet1: { '0,1': 10 } });
    const result = resolveRange(range('A', 1, 'B', 2), ctx);
    expect(result).toEqual([
      [10, null],
      [null, null],
    ]);
  });

  it('resolves a cross-sheet range', () => {
    const ctx = mockContext({ Other: { '0,1': 'x', '1,1': 'y' } });
    const result = resolveRange(range('A', 1, 'B', 1, 'Other'), ctx);
    expect(result).toEqual([['x', 'y']]);
  });

  it('resolves a single column range', () => {
    const ctx = mockContext(data);
    const result = resolveRange(range('A', 1, 'A', 2), ctx);
    expect(result).toEqual([[1], [4]]);
  });

  it('resolves a single row range', () => {
    const ctx = mockContext(data);
    const result = resolveRange(range('A', 1, 'C', 1), ctx);
    expect(result).toEqual([[1, 2, 3]]);
  });

  it('resolves lowercase columns in hand-built nodes', () => {
    const ctx = mockContext(data);
    expect(resolveRange(range('a', 1, 'c', 1), ctx)).toEqual([[1, 2, 3]]);
  });

  it.each([
    ['an end column past XFD', range('A', 1, 'XFE', 1)],
    ['a start column past XFD', range('XFE', 1, 'A', 1)],
    ['an end row past 1048576', range('A', 1, 'A', 1048577)],
    ['a start row of 0', range('A', 0, 'A', 1)],
  ])('returns [["#REF!"]] without calling getCell for %s', (_label, node) => {
    const getCell = vi.fn<EvalContext['getCell']>(() => 1);
    expect(resolveRange(node, { currentSheet: 'Sheet1', getCell })).toEqual([['#REF!']]);
    expect(getCell).not.toHaveBeenCalled();
  });
});

// ---------------------------------------------------------------------------
// Argument helpers: argToMatrix, argValues, isReferenceNode
// ---------------------------------------------------------------------------

/** Parse `formula`, failing the test on a parse error. */
function ast(formula: string): ASTNode {
  const { ast: node, errors } = parseFormula(formula);
  expect(errors).toEqual([]);
  if (!node) throw new Error(`no AST for ${formula}`);
  return node;
}

describe('argToMatrix', () => {
  const ctx = mockContext({
    Sheet1: { '0,1': 'A1', '1,1': 'B1', '1,2': 'B2', '2,2': 'C2', '1,3': 'B3' },
  });

  it('evaluates every element of an array constant', () => {
    expect(argToMatrix(ast('{1,2;3,4}'), ctx, evaluateNode)).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it('resolves a range', () => {
    expect(argToMatrix(ast('B1:B3'), ctx, evaluateNode)).toEqual([['B1'], ['B2'], ['B3']]);
  });

  it('gives the overlap rectangle of an intersection', () => {
    expect(argToMatrix(ast('B1:B3 B2:C2'), ctx, evaluateNode)).toEqual([['B2']]);
    expect(argToMatrix(ast('A1:C3 B:B'), ctx, evaluateNode)).toEqual([['B1'], ['B2'], ['B3']]);
  });

  it('gives #NULL! for disjoint references', () => {
    expect(argToMatrix(ast('A1 C3'), ctx, evaluateNode)).toEqual([['#NULL!']]);
  });

  it('appends the rows of both areas of a union', () => {
    expect(argToMatrix(ast('(A1,B2)'), ctx, evaluateNode)).toEqual([['A1'], ['B2']]);
  });

  it('gives #REF! for 3-D and structured references', () => {
    expect(argToMatrix(ast('Sheet1:Sheet3!A1'), ctx, evaluateNode)).toEqual([['#REF!']]);
    expect(argToMatrix(ast('Table1[Col]'), ctx, evaluateNode)).toEqual([['#REF!']]);
  });

  it('wraps any other argument as a 1x1 matrix', () => {
    expect(argToMatrix(ast('1+1'), ctx, evaluateNode)).toEqual([[2]]);
    expect(argToMatrix(ast('B2'), ctx, evaluateNode)).toEqual([['B2']]);
  });

  it('clamps a whole column to getBounds', () => {
    const bounded: EvalContext = { ...ctx, getBounds: () => ({ maxRow: 3, maxCol: 2 }) };
    expect(argToMatrix(ast('B:B'), bounded, evaluateNode)).toEqual([['B1'], ['B2'], ['B3']]);
  });
});

describe('argValues', () => {
  it('yields the cells of argToMatrix row-major', () => {
    const ctx = mockContext({ Sheet1: { '0,1': 1, '1,1': 2, '0,2': 3, '1,2': 4 } });
    for (const formula of ['A1:B2', '{1,2;3,4}', '(A1,B2)', 'A1:B2 B:B', 'A1 B2', '7']) {
      const node = ast(formula);
      expect([...argValues(node, ctx, evaluateNode)]).toEqual(
        argToMatrix(node, ctx, evaluateNode).flat(),
      );
    }
  });

  it('walks a whole column once per row without building a matrix', () => {
    let calls = 0;
    const ctx: EvalContext = {
      currentSheet: 'Sheet1',
      getCell(_sheet, col, row) {
        calls++;
        return col === 0 && row === 1_048_576 ? 1 : null;
      },
    };
    const values = [...argValues(ast('A:A'), ctx, evaluateNode)];
    expect(values).toHaveLength(1_048_576);
    expect(values.at(-1)).toBe(1);
    expect(calls).toBe(1_048_576);
  });

  it('walks a whole row across every column', () => {
    const ctx = mockContext({ Sheet1: { '16383,2': 'XFD2' } });
    const values = [...argValues(ast('2:2'), ctx, evaluateNode)];
    expect(values).toHaveLength(16_384);
    expect(values.at(-1)).toBe('XFD2');
  });
});

describe('isReferenceNode', () => {
  it('is true for cell, range, whole-column and whole-row references', () => {
    for (const formula of ['A1', 'A1:B2', 'A:A', '1:1']) {
      expect(isReferenceNode(ast(formula))).toBe(true);
    }
  });

  it('is true for intersections and unions of references', () => {
    expect(isReferenceNode(ast('A1:B2 B:B'))).toBe(true);
    expect(isReferenceNode(ast('(A1,B2)'))).toBe(true);
  });

  it('is false for values and function calls', () => {
    expect(isReferenceNode(ast('1'))).toBe(false);
    expect(isReferenceNode(ast('SUM(A1)'))).toBe(false);
    expect(isReferenceNode(ast('{1,2}'))).toBe(false);
  });
});
