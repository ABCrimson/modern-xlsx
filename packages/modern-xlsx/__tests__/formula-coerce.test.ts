import { describe, expect, it } from 'vitest';
import XLSX from 'xlsx';
import { formatGeneralNumber, numberToExcelText } from '../src/format/number-text.js';
import { excelPower, toBool, toNumber } from '../src/formula/coerce.js';
import { EXCEL_ERROR_LITERALS, isError } from '../src/formula/errors.js';
import { evaluateFormula } from '../src/formula/evaluator.js';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import { MAX_COL_INDEX, MAX_ROW } from '../src/formula/grid.js';
import type { CellValue, EvalContext } from '../src/formula/resolver.js';

const functions = createDefaultFunctions();

/** Evaluate `formula` against a Map keyed by A1 address on the current sheet. */
function ev(formula: string, cells: ReadonlyMap<string, CellValue> = new Map()): CellValue {
  const ctx: EvalContext = {
    currentSheet: 'Sheet1',
    functions,
    getCell: (_sheet, col, row) => cells.get(`${String.fromCharCode(65 + col)}${row}`) ?? null,
  };
  return evaluateFormula(formula, ctx);
}

describe('isError: closed set of Excel error literals', () => {
  it.each(['#N/A', '#BUSY!', '#SPILL!', ...EXCEL_ERROR_LITERALS])('%s is an error', (lit) => {
    expect(isError(lit)).toBe(true);
  });

  it.each(['#1', '#hashtag', '#', '#n/a', '#VALUE', ''])('%j is not an error', (text) => {
    expect(isError(text)).toBe(false);
  });

  it('non-strings are never errors', () => {
    expect(isError(0)).toBe(false);
    expect(isError(true)).toBe(false);
    expect(isError(null)).toBe(false);
  });

  it('holds 17 distinct literals, none a prefix of another', () => {
    expect(new Set(EXCEL_ERROR_LITERALS).size).toBe(17);
    for (const a of EXCEL_ERROR_LITERALS) {
      for (const b of EXCEL_ERROR_LITERALS) {
        if (a !== b) expect(b.startsWith(a)).toBe(false);
      }
    }
  });
});

describe('grid limits', () => {
  it('match Excel 2007+ (1,048,576 rows; XFD = 0-based column 16,383)', () => {
    expect(MAX_ROW).toBe(1_048_576);
    expect(MAX_COL_INDEX).toBe(16_383);
  });
});

describe('toNumber: decimal literal, else #VALUE!', () => {
  it.each([
    '',
    ' ',
    '0x10',
    '0b11',
    '0o7',
    'Infinity',
    '-Infinity',
    '1,23',
    '$5',
    '5%',
    '1e400',
    '1,2345',
    'abc',
  ])('%j -> #VALUE!', (text) => {
    expect(toNumber(text)).toBe('#VALUE!');
  });

  it.each([
    [' 1 ', 1],
    ['1e3', 1000],
    ['1,234', 1234],
    ['1,234.5', 1234.5],
    ['-1,234,567', -1_234_567],
    ['+.5', 0.5],
    ['5.', 5],
    ['1E-2', 0.01],
  ] as const)('%j -> %d', (text, expected) => {
    expect(toNumber(text)).toBe(expected);
  });

  it('keeps the non-string rules', () => {
    expect(toNumber(null)).toBe(0);
    expect(toNumber(true)).toBe(1);
    expect(toNumber(false)).toBe(0);
    expect(toNumber(-2.5)).toBe(-2.5);
    expect(toNumber('#N/A')).toBe('#N/A');
  });
});

describe('toBool', () => {
  it('follows Excel rules', () => {
    expect(toBool(null)).toBe(false);
    expect(toBool(0)).toBe(false);
    expect(toBool(-1)).toBe(true);
    expect(toBool('true')).toBe(true);
    expect(toBool('FALSE')).toBe(false);
    expect(toBool('#DIV/0!')).toBe('#DIV/0!');
    expect(toBool('#1')).toBe('#VALUE!');
    expect(toBool('yes')).toBe('#VALUE!');
  });
});

describe('numberToExcelText: formula number -> text', () => {
  it.each([
    [0.1 + 0.2, '0.3'],
    [1 / 3, '0.333333333333333'],
    [1e21, '1E+21'],
    [-0, '0'],
    [123456789012345, '123456789012345'],
  ] as const)('%d -> %s', (n, text) => {
    expect(numberToExcelText(n)).toBe(text);
  });

  // Observed in Excel 16.0 (build 20527) via `=(x)&""` on 2026-11-10: the plain
  // form is kept while it fits in 20 characters (sign excluded).
  it.each([
    [1e-10, '0.0000000001'],
    [1e-18, '0.000000000000000001'],
    [1e-19, '1E-19'],
    [-1e-18, '-0.000000000000000001'],
    [1.5e-17, '0.000000000000000015'],
    [1.5e-18, '1.5E-18'],
    [1.23456789012345e-4, '0.000123456789012345'],
    [1.23456789012345e-5, '1.23456789012345E-05'],
    [1.23456789e-10, '0.000000000123456789'],
    [1.23456789e-11, '1.23456789E-11'],
    [1e-4 / 3, '3.33333333333333E-05'],
    [1e-3 / 3, '0.000333333333333333'],
    [1e19, '10000000000000000000'],
    [-1e19, '-10000000000000000000'],
    [1.23456789012345e19, '12345678901234500000'],
    [1e20, '1E+20'],
    [-1e21, '-1E+21'],
    [1e15, '1000000000000000'],
    [2 ** 53, '9007199254740990'],
    [2 ** 60, '1152921504606850000'],
    [1e16 / 3, '3333333333333330'],
    [1e14 / 3, '33333333333333.3'],
    [10 / 3, '3.33333333333333'],
    [-1 / 3, '-0.333333333333333'],
    [2 / 3, '0.666666666666667'],
    [-1234.5678, '-1234.5678'],
    [1e100, '1E+100'],
    [1.5e-100, '1.5E-100'],
    [1e-300, '1E-300'],
    [1.23456e-16, '1.23456E-16'],
    [150000, '150000'],
  ] as const)('Excel: %d -> %s', (n, text) => {
    expect(numberToExcelText(n)).toBe(text);
  });

  // Known divergences: Excel 16.0 (build 20527) values observed 2026-11-11.
  it.fails.each([
    [1.23456789012345e99, '1.2345678901235E+99'],
    [9.99999999999999e307, '1E+308'],
  ] as const)('known gap, Excel: %d -> %s', (n, text) => {
    expect(numberToExcelText(n)).toBe(text);
  });
});

describe('formatGeneralNumber: Excel General display (11 characters)', () => {
  it.each([
    0,
    -0,
    1 / 3,
    0.1 + 0.2,
    123456789012345,
    1e21,
    1e-10,
    -1234.5678,
    12345678901,
    0.000123456,
    100,
  ])('matches SheetJS SSF General for %d', (v) => {
    expect(formatGeneralNumber(v)).toBe(XLSX.SSF.format('General', v));
  });

  // Observed in Excel 16.0 (build 20527) CSV export of General cells on 2026-11-10.
  it.each([
    [1 / 3, '0.333333333'],
    [2 / 3, '0.666666667'],
    [-2 / 3, '-0.666666667'],
    [123456789012345, '1.23457E+14'],
    [1e21, '1E+21'],
    [1e-10, '1E-10'],
    [12345678901.5, '12345678902'],
    [-12345678901.9, '-12345678902'],
    [99999999999.7, '1E+11'],
    [9999999999.7, '10000000000'],
    [1234567890.5, '1234567891'],
    [1e-7, '0.0000001'],
    [1.23456e-5, '1.23456E-05'],
    [1.23e-5, '0.0000123'],
    [1.5e-9, '1.5E-09'],
    [1e-9, '0.000000001'],
    [1.2345e-9, '1.2345E-09'],
    [-0.000123456, '-0.000123456'],
    [123456.789012, '123456.789'],
    [1e10, '10000000000'],
    [1e11, '1E+11'],
    [999999999999, '1E+12'],
    [-2147483649, '-2147483649'],
    [1.234567e-5, '1.23457E-05'],
    [Math.PI, '3.141592654'],
    [0.9999999999, '1'],
    [1e-5, '0.00001'],
    [2 ** 53, '9.0072E+15'],
    [1e100, '1E+100'],
    [-1.5e-100, '-1.5E-100'],
    [1.23456789e100, '1.2346E+100'],
    [1.23456789e-100, '1.2346E-100'],
    [0.00012345678901, '0.000123457'],
    [0.0000999999999, '1E-04'],
    [0.000099999, '0.000099999'],
    [1.99999999999e-5, '2E-05'],
    [999999.99999999, '1000000'],
  ] as const)('Excel: %d -> %s', (n, text) => {
    expect(formatGeneralNumber(n)).toBe(text);
  });

  // Known divergence: Excel 16.0 (build 20527) value observed 2026-11-11.
  it.fails('known gap, Excel: 1.23456789E+99 -> 1.2346E+99', () => {
    expect(formatGeneralNumber(1.23456789e99)).toBe('1.2346E+99');
  });
});

describe('excelPower', () => {
  it.each([
    [0, 0, '#NUM!'],
    [0, -1, '#DIV/0!'],
    [-1, 0.5, '#NUM!'],
    [10, 400, '#NUM!'],
    [2, 10, 1024],
    [2, -1, 0.5],
    [-8, 3, -512],
    // Real odd roots, as in Excel 16.0: =POWER(-8,1/3) is -2.
    [-8, 1 / 3, -2],
    [-8, -1 / 3, -0.5],
  ] as const)('%d ^ %d -> %s', (a, b, expected) => {
    expect(excelPower(a, b)).toBe(expected);
  });
});

describe('evaluateFormula with the shared coercion', () => {
  const a1Hash = new Map<string, CellValue>([['A1', '#1']]);

  it.each([
    ['LEN("#1")', 2],
    ['"#tag"&"x"', '#tagx'],
    ['#N/A+1', '#N/A'],
    ['LEN(#N/A)', '#N/A'],
    ['""+1', '#VALUE!'],
    ['" "+1', '#VALUE!'],
    ['"0x10"+0', '#VALUE!'],
    ['"Infinity"+0', '#VALUE!'],
    ['" 1 "+1', 2],
    ['VALUE("")', '#VALUE!'],
    ['VALUE("1,234")', 1234],
    ['VALUE("0x10")', '#VALUE!'],
    ['0.1+0.2&""', '0.3'],
    ['LEN(0.1+0.2)', 3],
    ['1/3&""', '0.333333333333333'],
    ['1E21&""', '1E+21'],
    ['CONCATENATE(0.1+0.2)', '0.3'],
    ['LEFT(1E21,3)', '1E+'],
    ['(-1)^0.5', '#NUM!'],
    ['(-1)^(1/3)', -1],
    ['POWER(-8,1/3)', -2],
    ['10^400', '#NUM!'],
    ['0^-1', '#DIV/0!'],
    ['0^0', '#NUM!'],
    ['POWER(0,0)', '#NUM!'],
    ['POWER(0,-1)', '#DIV/0!'],
    ['POWER(2,10)', 1024],
  ] as const)('%s = %j', (formula, expected) => {
    expect(ev(formula)).toBe(expected);
  });

  it('LEN of a cell holding the text "#1" is 2', () => {
    expect(ev('LEN(A1)', a1Hash)).toBe(2);
    expect(ev('A1&"x"', a1Hash)).toBe('#1x');
  });
});

describe('toNumber: linear-time literal matching', () => {
  it('rejects a 32,767-digit run with a trailing junk character', () => {
    // No wall-clock bound: a quadratic regression would stay far under the 10 s testTimeout, so only a bench can catch it.
    expect(toNumber(`${'1'.repeat(32_767)}x`)).toBe('#VALUE!');
  });

  it('keeps the accept set of the decimal literal', () => {
    for (const [text, n] of [
      ['1', 1],
      ['1.', 1],
      ['1.5', 1.5],
      ['.5', 0.5],
      ['-1e5', -100_000],
      ['+1.5E-3', 0.0015],
      [' 12 ', 12],
    ] as const) {
      expect(toNumber(text)).toBe(n);
    }
    for (const text of ['1e', '.', '', '1.2.3', 'abc', '1 2']) {
      expect(toNumber(text)).toBe('#VALUE!');
    }
  });
});
