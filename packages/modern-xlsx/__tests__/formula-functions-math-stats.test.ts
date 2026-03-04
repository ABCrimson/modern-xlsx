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
// SUM
// ---------------------------------------------------------------------------

describe('SUM', () => {
  it('sums literal numbers', () => {
    expect(ev('SUM(1,2,3)')).toBe(6);
  });

  it('sums a range of cells', () => {
    expect(
      ev('SUM(A1:A3)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 20,
        'Sheet1!A3': 30,
      }),
    ).toBe(60);
  });

  it('skips non-numeric values in ranges', () => {
    expect(
      ev('SUM(A1:A3)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 'text',
        'Sheet1!A3': 30,
      }),
    ).toBe(40);
  });

  it('treats booleans as 1/0', () => {
    expect(ev('SUM(TRUE,FALSE,1)')).toBe(2);
  });

  it('propagates errors', () => {
    expect(ev('SUM(1,#REF!,3)')).toBe('#REF!');
  });

  it('returns 0 for no numeric values', () => {
    expect(ev('SUM()')).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// AVERAGE
// ---------------------------------------------------------------------------

describe('AVERAGE', () => {
  it('calculates mean of numbers', () => {
    expect(ev('AVERAGE(2,4,6)')).toBe(4);
  });

  it('calculates mean of range', () => {
    expect(
      ev('AVERAGE(A1:A3)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 20,
        'Sheet1!A3': 30,
      }),
    ).toBe(20);
  });

  it('returns #DIV/0! for no numeric values', () => {
    expect(
      ev('AVERAGE(A1:A2)', {
        'Sheet1!A1': 'text',
        'Sheet1!A2': null,
      }),
    ).toBe('#DIV/0!');
  });

  it('propagates errors', () => {
    expect(ev('AVERAGE(1,#N/A)')).toBe('#N/A');
  });
});

// ---------------------------------------------------------------------------
// MIN / MAX
// ---------------------------------------------------------------------------

describe('MIN', () => {
  it('finds minimum', () => {
    expect(ev('MIN(5,2,8,1)')).toBe(1);
  });

  it('finds minimum in range', () => {
    expect(
      ev('MIN(A1:A3)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 5,
        'Sheet1!A3': 20,
      }),
    ).toBe(5);
  });

  it('returns 0 for no numeric values', () => {
    expect(ev('MIN()')).toBe(0);
  });

  it('handles negative numbers', () => {
    expect(ev('MIN(3,-2,1)')).toBe(-2);
  });
});

describe('MAX', () => {
  it('finds maximum', () => {
    expect(ev('MAX(5,2,8,1)')).toBe(8);
  });

  it('finds maximum in range', () => {
    expect(
      ev('MAX(A1:A3)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 50,
        'Sheet1!A3': 20,
      }),
    ).toBe(50);
  });

  it('returns 0 for no numeric values', () => {
    expect(ev('MAX()')).toBe(0);
  });

  it('handles negative numbers', () => {
    expect(ev('MAX(-3,-2,-1)')).toBe(-1);
  });
});

// ---------------------------------------------------------------------------
// COUNT / COUNTA / COUNTBLANK
// ---------------------------------------------------------------------------

describe('COUNT', () => {
  it('counts numeric values', () => {
    expect(ev('COUNT(1,2,3)')).toBe(3);
  });

  it('excludes non-numeric', () => {
    expect(
      ev('COUNT(A1:A4)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 'text',
        'Sheet1!A3': null,
        'Sheet1!A4': 20,
      }),
    ).toBe(2);
  });
});

describe('COUNTA', () => {
  it('counts non-empty values', () => {
    expect(
      ev('COUNTA(A1:A4)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': 'text',
        'Sheet1!A3': null,
        'Sheet1!A4': 0,
      }),
    ).toBe(3);
  });

  it('counts booleans', () => {
    expect(ev('COUNTA(TRUE,FALSE,0)')).toBe(3);
  });
});

describe('COUNTBLANK', () => {
  it('counts blank cells', () => {
    expect(
      ev('COUNTBLANK(A1:A4)', {
        'Sheet1!A1': 10,
        'Sheet1!A2': null,
        'Sheet1!A3': '',
        'Sheet1!A4': 20,
      }),
    ).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// ROUND / ROUNDUP / ROUNDDOWN
// ---------------------------------------------------------------------------

describe('ROUND', () => {
  it('rounds to 2 decimal places', () => {
    expect(ev('ROUND(3.14159,2)')).toBe(3.14);
  });

  it('rounds to 0 decimal places', () => {
    expect(ev('ROUND(3.5,0)')).toBe(4);
  });

  it('rounds with negative digits', () => {
    expect(ev('ROUND(1234,-2)')).toBe(1200);
  });

  it('propagates errors', () => {
    expect(ev('ROUND(#REF!,2)')).toBe('#REF!');
  });
});

describe('ROUNDUP', () => {
  it('rounds up positive number', () => {
    expect(ev('ROUNDUP(3.141,2)')).toBe(3.15);
  });

  it('rounds up negative number (away from zero)', () => {
    expect(ev('ROUNDUP(-3.141,2)')).toBe(-3.15);
  });

  it('rounds up to 0 decimal places', () => {
    expect(ev('ROUNDUP(3.1,0)')).toBe(4);
  });
});

describe('ROUNDDOWN', () => {
  it('rounds down positive number', () => {
    expect(ev('ROUNDDOWN(3.149,2)')).toBe(3.14);
  });

  it('rounds down negative number (toward zero)', () => {
    expect(ev('ROUNDDOWN(-3.149,2)')).toBe(-3.14);
  });

  it('rounds down to 0 decimal places', () => {
    expect(ev('ROUNDDOWN(3.9,0)')).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// ABS / SQRT / MOD / INT
// ---------------------------------------------------------------------------

describe('ABS', () => {
  it('absolute of positive', () => {
    expect(ev('ABS(5)')).toBe(5);
  });

  it('absolute of negative', () => {
    expect(ev('ABS(-5)')).toBe(5);
  });

  it('absolute of zero', () => {
    expect(ev('ABS(0)')).toBe(0);
  });
});

describe('SQRT', () => {
  it('square root of 9', () => {
    expect(ev('SQRT(9)')).toBe(3);
  });

  it('square root of 2', () => {
    expect(ev('SQRT(2)')).toBeCloseTo(Math.SQRT2, 4);
  });

  it('negative returns #NUM!', () => {
    expect(ev('SQRT(-1)')).toBe('#NUM!');
  });

  it('sqrt of 0', () => {
    expect(ev('SQRT(0)')).toBe(0);
  });
});

describe('MOD', () => {
  it('basic modulus', () => {
    expect(ev('MOD(7,3)')).toBe(1);
  });

  it('divisor 0 returns #DIV/0!', () => {
    expect(ev('MOD(7,0)')).toBe('#DIV/0!');
  });

  it('Excel MOD: result has same sign as divisor', () => {
    expect(ev('MOD(-7,3)')).toBe(2);
  });

  it('negative divisor', () => {
    expect(ev('MOD(7,-3)')).toBe(-2);
  });
});

describe('INT', () => {
  it('floors positive number', () => {
    expect(ev('INT(3.7)')).toBe(3);
  });

  it('floors negative number', () => {
    expect(ev('INT(-3.2)')).toBe(-4);
  });

  it('integer unchanged', () => {
    expect(ev('INT(5)')).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// CEILING / FLOOR
// ---------------------------------------------------------------------------

describe('CEILING', () => {
  it('rounds up to multiple', () => {
    expect(ev('CEILING(4.2,1)')).toBe(5);
  });

  it('rounds up to multiple of 0.5', () => {
    expect(ev('CEILING(4.2,0.5)')).toBe(4.5);
  });

  it('significance 0 returns 0', () => {
    expect(ev('CEILING(4.2,0)')).toBe(0);
  });

  it('positive number with negative significance returns #NUM!', () => {
    expect(ev('CEILING(4.2,-1)')).toBe('#NUM!');
  });
});

describe('FLOOR', () => {
  it('rounds down to multiple', () => {
    expect(ev('FLOOR(4.8,1)')).toBe(4);
  });

  it('rounds down to multiple of 0.5', () => {
    expect(ev('FLOOR(4.8,0.5)')).toBe(4.5);
  });

  it('significance 0 returns #DIV/0!', () => {
    expect(ev('FLOOR(4.8,0)')).toBe('#DIV/0!');
  });

  it('positive number with negative significance returns #NUM!', () => {
    expect(ev('FLOOR(4.8,-1)')).toBe('#NUM!');
  });
});

// ---------------------------------------------------------------------------
// POWER / LOG / LN / PI
// ---------------------------------------------------------------------------

describe('POWER', () => {
  it('2^3 = 8', () => {
    expect(ev('POWER(2,3)')).toBe(8);
  });

  it('10^0 = 1', () => {
    expect(ev('POWER(10,0)')).toBe(1);
  });

  it('negative exponent', () => {
    expect(ev('POWER(2,-1)')).toBe(0.5);
  });
});

describe('LOG', () => {
  it('log base 10: LOG(100)', () => {
    expect(ev('LOG(100)')).toBeCloseTo(2, 10);
  });

  it('log base 2: LOG(8,2)', () => {
    expect(ev('LOG(8,2)')).toBeCloseTo(3, 10);
  });

  it('negative returns #NUM!', () => {
    expect(ev('LOG(-1)')).toBe('#NUM!');
  });

  it('zero returns #NUM!', () => {
    expect(ev('LOG(0)')).toBe('#NUM!');
  });

  it('base 1 returns #NUM!', () => {
    expect(ev('LOG(10,1)')).toBe('#NUM!');
  });
});

describe('LN', () => {
  it('natural log of e', () => {
    expect(ev('LN(2.718281828)')).toBeCloseTo(1, 5);
  });

  it('natural log of 1', () => {
    expect(ev('LN(1)')).toBe(0);
  });

  it('negative returns #NUM!', () => {
    expect(ev('LN(-1)')).toBe('#NUM!');
  });
});

describe('PI', () => {
  it('returns pi', () => {
    expect(ev('PI()')).toBeCloseTo(Math.PI, 7);
  });
});

describe('RAND', () => {
  it('returns a number between 0 and 1', () => {
    const result = ev('RAND()');
    expect(typeof result).toBe('number');
    expect(result as number).toBeGreaterThanOrEqual(0);
    expect(result as number).toBeLessThan(1);
  });
});

// ---------------------------------------------------------------------------
// SUMIF / COUNTIF / AVERAGEIF
// ---------------------------------------------------------------------------

describe('SUMIF', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 10,
    'Sheet1!A2': 20,
    'Sheet1!A3': 30,
    'Sheet1!A4': 40,
    'Sheet1!B1': 100,
    'Sheet1!B2': 200,
    'Sheet1!B3': 300,
    'Sheet1!B4': 400,
  };

  it('sums values matching exact number', () => {
    expect(ev('SUMIF(A1:A4,20)', cells)).toBe(20);
  });

  it('sums with ">" criteria', () => {
    expect(ev('SUMIF(A1:A4,">20")', cells)).toBe(70);
  });

  it('sums with "<>" criteria', () => {
    expect(ev('SUMIF(A1:A4,"<>20")', cells)).toBe(80);
  });

  it('sums with separate sum_range', () => {
    expect(ev('SUMIF(A1:A4,">20",B1:B4)', cells)).toBe(700);
  });

  it('sums with "<=" criteria', () => {
    expect(ev('SUMIF(A1:A4,"<=20")', cells)).toBe(30);
  });
});

describe('COUNTIF', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 10,
    'Sheet1!A2': 20,
    'Sheet1!A3': 30,
    'Sheet1!A4': 20,
  };

  it('counts exact matches', () => {
    expect(ev('COUNTIF(A1:A4,20)', cells)).toBe(2);
  });

  it('counts with ">" criteria', () => {
    expect(ev('COUNTIF(A1:A4,">15")', cells)).toBe(3);
  });

  it('counts string matches (case-insensitive)', () => {
    const strCells: Record<string, CellValue> = {
      'Sheet1!A1': 'apple',
      'Sheet1!A2': 'APPLE',
      'Sheet1!A3': 'banana',
    };
    expect(ev('COUNTIF(A1:A3,"apple")', strCells)).toBe(2);
  });

  it('counts with wildcard *', () => {
    const strCells: Record<string, CellValue> = {
      'Sheet1!A1': 'apple',
      'Sheet1!A2': 'application',
      'Sheet1!A3': 'banana',
    };
    expect(ev('COUNTIF(A1:A3,"app*")', strCells)).toBe(2);
  });

  it('counts with wildcard ?', () => {
    const strCells: Record<string, CellValue> = {
      'Sheet1!A1': 'cat',
      'Sheet1!A2': 'bat',
      'Sheet1!A3': 'cart',
    };
    expect(ev('COUNTIF(A1:A3,"?at")', strCells)).toBe(2);
  });
});

describe('AVERAGEIF', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 10,
    'Sheet1!A2': 20,
    'Sheet1!A3': 30,
    'Sheet1!A4': 40,
    'Sheet1!B1': 100,
    'Sheet1!B2': 200,
    'Sheet1!B3': 300,
    'Sheet1!B4': 400,
  };

  it('averages matching values', () => {
    expect(ev('AVERAGEIF(A1:A4,">20")', cells)).toBe(35);
  });

  it('averages with separate range', () => {
    expect(ev('AVERAGEIF(A1:A4,">20",B1:B4)', cells)).toBe(350);
  });

  it('returns #DIV/0! when no matches', () => {
    expect(ev('AVERAGEIF(A1:A4,">100")', cells)).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// SUMPRODUCT
// ---------------------------------------------------------------------------

describe('SUMPRODUCT', () => {
  it('sum of products of two arrays', () => {
    const cells: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!A3': 3,
      'Sheet1!B1': 4,
      'Sheet1!B2': 5,
      'Sheet1!B3': 6,
    };
    // 1*4 + 2*5 + 3*6 = 4+10+18 = 32
    expect(ev('SUMPRODUCT(A1:A3,B1:B3)', cells)).toBe(32);
  });

  it('single array just sums', () => {
    const cells: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!A3': 3,
    };
    expect(ev('SUMPRODUCT(A1:A3)', cells)).toBe(6);
  });

  it('different-sized arrays return #VALUE!', () => {
    const cells: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 2,
      'Sheet1!B1': 3,
      'Sheet1!B2': 4,
      'Sheet1!B3': 5,
    };
    expect(ev('SUMPRODUCT(A1:A2,B1:B3)', cells)).toBe('#VALUE!');
  });

  it('treats non-numeric as 0', () => {
    const cells: Record<string, CellValue> = {
      'Sheet1!A1': 1,
      'Sheet1!A2': 'text',
      'Sheet1!B1': 10,
      'Sheet1!B2': 20,
    };
    // 1*10 + 0*20 = 10
    expect(ev('SUMPRODUCT(A1:A2,B1:B2)', cells)).toBe(10);
  });
});
