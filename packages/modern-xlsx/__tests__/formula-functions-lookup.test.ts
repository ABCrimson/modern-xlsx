import { describe, expect, it } from 'vitest';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import type { CellValue } from '../src/index.js';
import { evaluateFormula } from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const functions = createDefaultFunctions();

function ev(formula: string, cells?: Record<string, CellValue>): CellValue {
  return evaluateFormula(formula, {
    currentSheet: 'Sheet1',
    functions,
    getCell: (sheet: string, col: number, row: number): CellValue => {
      const letter = String.fromCharCode(65 + col);
      const key = `${sheet}!${letter}${row}`;
      return cells?.[key] ?? null;
    },
  });
}

// ---------------------------------------------------------------------------
// VLOOKUP
// ---------------------------------------------------------------------------

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

  it('row or col < 1 returns #VALUE!', () => {
    expect(ev('INDEX(A1:C2,0,1)', cells)).toBe('#VALUE!');
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

  it('no argument returns #VALUE!', () => {
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

  it('no argument returns #VALUE!', () => {
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
});
