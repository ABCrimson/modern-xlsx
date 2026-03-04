import { describe, expect, it } from 'vitest';
import type { CellRefNode, CellValue, EvalContext, RangeNode } from '../src/index.js';
import { resolveRange, resolveRef } from '../src/index.js';

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
});
