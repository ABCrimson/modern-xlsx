import { describe, expect, it } from 'vitest';
import { columnToLetter } from '../src/format/cell-ref.js';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import type { CellValue } from '../src/index.js';
import { evaluateFormula } from '../src/index.js';
import { ev } from './helpers/formula-context.js';

/** Like `ev`, but with `currentCell` set (the cell that holds the formula). */
function evAt(
  formula: string,
  currentCell: { col: number; row: number } | undefined,
  cells: Record<string, CellValue> = {},
): CellValue {
  return evaluateFormula(formula, {
    currentSheet: 'Sheet1',
    functions: createDefaultFunctions(),
    ...(currentCell === undefined ? {} : { currentCell }),
    getCell: (sheet, col, row) => cells[`${sheet}!${columnToLetter(col)}${row}`] ?? null,
  });
}

// ---------------------------------------------------------------------------
// VLOOKUP
// ---------------------------------------------------------------------------

describe('VLOOKUP type-strict and wildcard matching', () => {
  const c: Record<string, CellValue> = {
    'Sheet1!A1': 1,
    'Sheet1!A2': 2,
    'Sheet1!B1': 'a',
    'Sheet1!B2': 'b',
    'Sheet1!A7': 'apple',
    'Sheet1!A8': 'banana',
    'Sheet1!B7': 10,
    'Sheet1!B8': 20,
  };

  it('text never matches a number, exact or approximate', () => {
    expect(ev('VLOOKUP("1",A1:B2,2,FALSE)', c)).toBe('#N/A');
    expect(ev('VLOOKUP("1",A1:B2,2,TRUE)', c)).toBe('#N/A');
    expect(ev('VLOOKUP(1,A1:B2,2,FALSE)', c)).toBe('a');
  });

  it('empty text never matches a blank cell', () => {
    expect(ev('VLOOKUP("",A1:B3,2,FALSE)', c)).toBe('#N/A');
  });

  it('wildcards match text exactly', () => {
    expect(ev('VLOOKUP("ap*",A7:B8,2,FALSE)', c)).toBe(10);
    expect(ev('HLOOKUP("ap*",A7:B8,2,FALSE)', c)).toBe('banana');
  });

  it('reads array constants', () => {
    expect(ev('VLOOKUP(2,{1,"a";2,"b"},2,FALSE)')).toBe('b');
    expect(ev('HLOOKUP(2,{1,2;"a","b"},2,FALSE)')).toBe('b');
  });
});

describe('VLOOKUP', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 1,
    'Sheet1!A2': 2,
    'Sheet1!A3': 3,
    'Sheet1!A4': 4,
    'Sheet1!B1': 'one',
    'Sheet1!B2': 'two',
    'Sheet1!B3': 'three',
    'Sheet1!B4': 'four',
  };

  it('exact match found', () => {
    expect(ev('VLOOKUP(3,A1:B4,2,FALSE)', cells)).toBe('three');
  });

  it('exact match not found returns #N/A', () => {
    expect(ev('VLOOKUP(5,A1:B4,2,FALSE)', cells)).toBe('#N/A');
  });

  it('approximate match (default)', () => {
    // Sorted ascending data: finds largest <= 3.5 which is 3
    const cells2: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!A3': 3,
      'Sheet1!A4': 5,
      'Sheet1!B1': 'one',
      'Sheet1!B2': 'two',
      'Sheet1!B3': 'three',
      'Sheet1!B4': 'five',
    };
    expect(ev('VLOOKUP(3.5,A1:B4,2)', cells2)).toBe('three');
  });

  it('approximate match: value smaller than all returns #N/A', () => {
    expect(ev('VLOOKUP(0,A1:B4,2,TRUE)', cells)).toBe('#N/A');
  });

  it('col_index out of range returns #REF!', () => {
    expect(ev('VLOOKUP(1,A1:B4,3,FALSE)', cells)).toBe('#REF!');
  });

  it('case-insensitive string lookup', () => {
    const strCells: Record<string, CellValue> = {
      'Sheet1!A1': 'apple',
      'Sheet1!A2': 'banana',
      'Sheet1!A3': 'cherry',
      'Sheet1!B1': 1,
      'Sheet1!B2': 2,
      'Sheet1!B3': 3,
    };
    expect(ev('VLOOKUP("BANANA",A1:B3,2,FALSE)', strCells)).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// HLOOKUP
// ---------------------------------------------------------------------------

describe('HLOOKUP', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 1,
    'Sheet1!B1': 2,
    'Sheet1!C1': 3,
    'Sheet1!A2': 'one',
    'Sheet1!B2': 'two',
    'Sheet1!C2': 'three',
  };

  it('exact match found', () => {
    expect(ev('HLOOKUP(2,A1:C2,2,FALSE)', cells)).toBe('two');
  });

  it('exact match not found returns #N/A', () => {
    expect(ev('HLOOKUP(5,A1:C2,2,FALSE)', cells)).toBe('#N/A');
  });

  it('approximate match', () => {
    const cells2: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!B1': 3,
      'Sheet1!C1': 5,
      'Sheet1!A2': 'one',
      'Sheet1!B2': 'three',
      'Sheet1!C2': 'five',
    };
    expect(ev('HLOOKUP(4,A1:C2,2)', cells2)).toBe('three');
  });

  it('row_index out of range returns #REF!', () => {
    expect(ev('HLOOKUP(1,A1:C2,3,FALSE)', cells)).toBe('#REF!');
  });
});

// ---------------------------------------------------------------------------
// INDEX
// ---------------------------------------------------------------------------

describe('INDEX', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 10,
    'Sheet1!B1': 20,
    'Sheet1!C1': 30,
    'Sheet1!A2': 40,
    'Sheet1!B2': 50,
    'Sheet1!C2': 60,
  };

  it('returns value at row,col', () => {
    expect(ev('INDEX(A1:C2,2,3)', cells)).toBe(60);
  });

  it('returns value at row 1, col 1', () => {
    expect(ev('INDEX(A1:C2,1,1)', cells)).toBe(10);
  });

  it('out of range returns #REF!', () => {
    expect(ev('INDEX(A1:C2,3,1)', cells)).toBe('#REF!');
  });

  it('column out of range returns #REF!', () => {
    expect(ev('INDEX(A1:C2,1,4)', cells)).toBe('#REF!');
  });

  it('negative row or col returns #VALUE!', () => {
    expect(ev('INDEX(A1:C2,-1,1)', cells)).toBe('#VALUE!');
    expect(ev('INDEX(A1:C2,1,-1)', cells)).toBe('#VALUE!');
  });

  it('zero index without a current cell returns #VALUE!', () => {
    expect(ev('INDEX(A1:C2,0,1)', cells)).toBe('#VALUE!');
  });

  it('two arguments on a one-row range index the column', () => {
    expect(ev('INDEX(A1:C1,2)', cells)).toBe(20);
    expect(ev('INDEX(A1:C1,4)', cells)).toBe('#REF!');
  });

  it('zero index on a one-row range selects the single cell', () => {
    expect(ev('INDEX(A1:C1,0,2)', cells)).toBe(20);
  });

  it('zero row intersects the whole column on the current row', () => {
    expect(evAt('INDEX(A1:C2,0,1)', { col: 5, row: 2 }, cells)).toBe(40);
    expect(evAt('INDEX(A1:C2,0,1)', { col: 5, row: 9 }, cells)).toBe('#VALUE!');
  });

  it('zero column intersects the whole row on the current column', () => {
    expect(evAt('INDEX(A1:C2,2,0)', { col: 1, row: 20 }, cells)).toBe(50);
  });

  it('reads array constants', () => {
    expect(ev('INDEX({1,2,3},2)')).toBe(2);
    expect(ev('INDEX({1,2;3,4},2,1)')).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// MATCH
// ---------------------------------------------------------------------------

describe('MATCH', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 10,
    'Sheet1!A2': 20,
    'Sheet1!A3': 30,
    'Sheet1!A4': 40,
  };

  it('exact match (type 0)', () => {
    expect(ev('MATCH(30,A1:A4,0)', cells)).toBe(3);
  });

  it('exact match not found returns #N/A', () => {
    expect(ev('MATCH(25,A1:A4,0)', cells)).toBe('#N/A');
  });

  it('less than or equal (type 1, default)', () => {
    expect(ev('MATCH(25,A1:A4,1)', cells)).toBe(2);
  });

  it('less than or equal: exact match returns position', () => {
    expect(ev('MATCH(30,A1:A4,1)', cells)).toBe(3);
  });

  it('greater than or equal (type -1, sorted descending)', () => {
    const descCells: Record<string, CellValue> = {
      'Sheet1!A1': 40,
      'Sheet1!A2': 30,
      'Sheet1!A3': 20,
      'Sheet1!A4': 10,
    };
    expect(ev('MATCH(25,A1:A4,-1)', descCells)).toBe(2);
  });

  it('value smaller than all in ascending, type 1 returns #N/A', () => {
    expect(ev('MATCH(5,A1:A4,1)', cells)).toBe('#N/A');
  });

  it('match_type is truncated, then reduced to its sign', () => {
    expect(ev('MATCH(25,A1:A4,2)', cells)).toBe(ev('MATCH(25,A1:A4,1)', cells));
    expect(ev('MATCH(30,A1:A4,0.5)', cells)).toBe(3);
    expect(ev('MATCH(25,A1:A4,0.5)', cells)).toBe('#N/A');
    const descCells: Record<string, CellValue> = {
      'Sheet1!A1': 40,
      'Sheet1!A2': 30,
      'Sheet1!A3': 20,
      'Sheet1!A4': 10,
    };
    expect(ev('MATCH(25,A1:A4,-2)', descCells)).toBe(ev('MATCH(25,A1:A4,-1)', descCells));
    expect(ev('MATCH(25,A1:A4,-2)', descCells)).toBe(2);
  });

  it('reads array constants', () => {
    expect(ev('MATCH(2,{1,2,3},0)')).toBe(2);
  });

  it('wildcards apply to exact text matches', () => {
    const c: Record<string, CellValue> = {
      'Sheet1!A5': 'apple',
      'Sheet1!A7': 'apple',
      'Sheet1!A8': 'banana',
    };
    expect(ev('MATCH("ap*",A1:A5,0)', c)).toBe(5);
    expect(ev('MATCH("b?nana",A7:A8,0)', c)).toBe(2);
    expect(ev('MATCH("ap*",A1:A5,1)', c)).toBe('#N/A');
  });

  it('exact matching is type-strict', () => {
    const c: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!C1': '1',
      'Sheet1!C2': '2',
    };
    expect(ev('MATCH("2",A1:A2,0)', c)).toBe('#N/A');
    expect(ev('MATCH(1,C1:C2,0)', c)).toBe('#N/A');
  });

  it('approximate matching skips values of another type', () => {
    const c: Record<string, CellValue> = { 'Sheet1!A1': 'x', 'Sheet1!A2': 1, 'Sheet1!A3': 5 };
    expect(ev('MATCH(3,A1:A3,1)', c)).toBe(2);
  });

  it('match_type 2 on 1..4 finds position 3 for 3', () => {
    const c: Record<string, CellValue> = {
      'Sheet1!B1': 1,
      'Sheet1!B2': 2,
      'Sheet1!B3': 3,
      'Sheet1!B4': 4,
    };
    expect(ev('MATCH(3,B1:B4,2)', c)).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// CHOOSE
// ---------------------------------------------------------------------------

describe('CHOOSE', () => {
  it('returns first value', () => {
    expect(ev('CHOOSE(1,"a","b","c")')).toBe('a');
  });

  it('returns second value', () => {
    expect(ev('CHOOSE(2,"a","b","c")')).toBe('b');
  });

  it('returns third value', () => {
    expect(ev('CHOOSE(3,"a","b","c")')).toBe('c');
  });

  it('index 0 returns #VALUE!', () => {
    expect(ev('CHOOSE(0,"a","b")')).toBe('#VALUE!');
  });

  it('index out of range returns #VALUE!', () => {
    expect(ev('CHOOSE(5,"a","b","c")')).toBe('#VALUE!');
  });

  it('with numeric values', () => {
    expect(ev('CHOOSE(2,10,20,30)')).toBe(20);
  });

  it('a missing selected argument is 0', () => {
    expect(ev('CHOOSE(2,1,,3)')).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// ROW / COLUMN
// ---------------------------------------------------------------------------

describe('ROW', () => {
  it('returns row number of cell ref', () => {
    expect(ev('ROW(A5)')).toBe(5);
  });

  it('returns row number of range start', () => {
    expect(ev('ROW(B3:D7)')).toBe(3);
  });

  it('no argument returns the current row', () => {
    expect(evAt('ROW()', { col: 2, row: 7 })).toBe(7);
  });

  it('no argument without a current cell returns #VALUE!', () => {
    expect(ev('ROW()')).toBe('#VALUE!');
  });
});

describe('COLUMN', () => {
  it('returns column number of A (1)', () => {
    expect(ev('COLUMN(A1)')).toBe(1);
  });

  it('returns column number of C (3)', () => {
    expect(ev('COLUMN(C5)')).toBe(3);
  });

  it('returns column number of range start', () => {
    expect(ev('COLUMN(B3:D7)')).toBe(2);
  });

  it('no argument returns the current column', () => {
    expect(evAt('COLUMN()', { col: 2, row: 7 })).toBe(3);
  });

  it('no argument without a current cell returns #VALUE!', () => {
    expect(ev('COLUMN()')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// ROWS / COLUMNS
// ---------------------------------------------------------------------------

describe('ROWS', () => {
  it('counts rows in range', () => {
    expect(ev('ROWS(A1:A5)')).toBe(5);
  });

  it('counts rows in multi-column range', () => {
    expect(ev('ROWS(A1:C3)')).toBe(3);
  });

  it('single cell = 1 row', () => {
    expect(ev('ROWS(A1)')).toBe(1);
  });

  it('counts rows in an array constant', () => {
    expect(ev('ROWS({1,2;3,4;5,6})')).toBe(3);
  });
});

describe('COLUMNS', () => {
  it('counts columns in range', () => {
    expect(ev('COLUMNS(A1:C1)')).toBe(3);
  });

  it('counts columns in multi-row range', () => {
    expect(ev('COLUMNS(A1:D5)')).toBe(4);
  });

  it('single cell = 1 column', () => {
    expect(ev('COLUMNS(A1)')).toBe(1);
  });

  it('counts columns in an array constant', () => {
    expect(ev('COLUMNS({1,2,3})')).toBe(3);
  });
});

// Expected values from Excel (Microsoft 365, COM probe on 2026-11-12), except
// the 3-D rows: Excel gives #VALUE! there, and 1.3.0 evaluates every 3-D
// reference to #REF! (D/deferred.json "TS-060 eval").
describe('ROW, COLUMN, ROWS and COLUMNS size every single-area reference', () => {
  it('whole columns and rows', () => {
    expect(ev('ROWS(A:A)')).toBe(1048576);
    expect(ev('COLUMNS(1:1)')).toBe(16384);
    expect(ev('ROW(A:A)')).toBe(1);
    expect(ev('COLUMN(A:A)')).toBe(1);
  });

  it('a reversed range reads its top row', () => {
    expect(ev('ROW(A5:A1)')).toBe(1);
    expect(ev('ROW(B5:A1)')).toBe(1);
    expect(ev('ROWS(B5:A1)')).toBe(5);
  });

  it('an intersection is sized by its overlap', () => {
    // A1:B5 ∩ B2:C3 is B2:B3.
    expect(ev('ROWS(A1:B5 B2:C3)')).toBe(2);
  });

  it('array constants are sized by their elements', () => {
    expect(ev('ROWS({1;2})')).toBe(2);
    expect(ev('COLUMNS({1,2,3})')).toBe(3);
  });

  it('a 3-D reference is #REF! for all four', () => {
    for (const f of ['ROW', 'COLUMN', 'ROWS', 'COLUMNS']) {
      expect(ev(`${f}(Sheet1:Sheet3!A1:B2)`)).toBe('#REF!');
      expect(ev(`${f}(Sheet1:Sheet3!B2)`)).toBe('#REF!');
    }
  });
});

describe('ev helper', () => {
  it('keys cells past column Z with columnToLetter', () => {
    expect(ev('COLUMN(AA1)')).toBe(27);
    expect(ev('AA1', { 'Sheet1!AA1': 7 })).toBe(7);
  });
});

describe('reference tables, extents and size limits', () => {
  /** Evaluate with `getBounds` limiting the sheet to its populated cells. */
  function evBounded(
    formula: string,
    cells: Record<string, CellValue>,
    bounds: { maxRow: number; maxCol: number } | undefined,
    currentCell?: { col: number; row: number },
  ): CellValue {
    return evaluateFormula(formula, {
      currentSheet: 'Sheet1',
      functions: createDefaultFunctions(),
      ...(currentCell === undefined ? {} : { currentCell }),
      getBounds: () => bounds,
      getCell: (sheet, col, row) => cells[`${sheet}!${columnToLetter(col)}${row}`] ?? null,
    });
  }

  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 'k1',
    'Sheet1!B1': 1,
    'Sheet1!A2': 'k2',
    'Sheet1!B2': 2,
    'Sheet1!A3': 'k3',
    'Sheet1!B3': 3,
  };
  const narrow = { maxRow: 3, maxCol: 1 };

  it('COLUMN and COLUMNS of an out-of-grid column give #NAME?, never throw', () => {
    expect(ev('COLUMN(XFE1)')).toBe('#NAME?');
    expect(ev('COLUMNS(XFD1:XFE1)')).toBe('#NAME?');
    expect(ev('COLUMN(XFD1)')).toBe(16384);
  });

  it('VLOOKUP over a whole-column table uses its declared width', () => {
    expect(evBounded('VLOOKUP("k2",A:C,3,FALSE)', cells, narrow)).toBe(null);
    expect(evBounded('VLOOKUP("k2",A:C,4,FALSE)', cells, narrow)).toBe('#REF!');
    expect(evBounded('VLOOKUP("k2",A:B,2,FALSE)', cells, narrow)).toBe(2);
  });

  it('INDEX over a whole column or row validates against the declared size', () => {
    expect(evBounded('INDEX(A:A,5)', cells, narrow)).toBe(null);
    expect(evBounded('INDEX(A:B,0,2)', cells, narrow, { col: 5, row: 2 })).toBe(2);
    expect(evBounded('INDEX(A:A,1048577)', cells, narrow)).toBe('#REF!');
  });

  it('an oversized table without bounds gives #NUM!, not #N/A', () => {
    expect(ev('VLOOKUP("k2",A1:B1048576,2,FALSE)', cells)).toBe('#NUM!');
    expect(ev('VLOOKUP("k2",A:B,2,FALSE)', cells)).toBe('#NUM!');
    expect(ev('MATCH("k2",A1:B1048576,0)', cells)).toBe('#NUM!');
  });

  it('INDEX of a union reads the first area', () => {
    const c: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!C1': 7,
      'Sheet1!C2': 8,
      'Sheet1!C3': 9,
    };
    expect(ev('INDEX((A1:A2,C1:C3),2)', c)).toBe(2);
    expect(ev('INDEX((A1:A2,C1:C3),3)', c)).toBe('#REF!');
  });
});
