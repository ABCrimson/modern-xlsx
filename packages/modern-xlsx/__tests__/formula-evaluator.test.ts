import { assert, describe, expect, it, vi } from 'vitest';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import type { ASTNode, CellValue, EvalContext, FormulaFunction } from '../src/index.js';
import { evaluateFormula, evaluateNode } from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build a mock EvalContext from a flat cell map. Keys are "COL,ROW" (0-based col, 1-based row). */
function mockContext(
  data: Record<string, CellValue> = {},
  currentSheet = 'Sheet1',
  functions?: Map<string, FormulaFunction>,
): EvalContext {
  return {
    currentSheet,
    ...(functions && { functions }),
    getCell(sheet: string, col: number, row: number): CellValue {
      const key = `${sheet}!${col},${row}`;
      if (key in data) return data[key] ?? null;
      // Fallback: try without sheet prefix for convenience
      const shortKey = `${col},${row}`;
      return data[shortKey] ?? null;
    },
  };
}

/** Shorthand to evaluate a formula with default empty context. */
function ev(
  formula: string,
  data?: Record<string, CellValue>,
  fns?: Map<string, FormulaFunction>,
): CellValue {
  return evaluateFormula(formula, mockContext(data, 'Sheet1', fns));
}

// ---------------------------------------------------------------------------
// Arithmetic
// ---------------------------------------------------------------------------

describe('arithmetic', () => {
  it('addition: 1+2 = 3', () => {
    expect(ev('1+2')).toBe(3);
  });

  it('subtraction: 10-3 = 7', () => {
    expect(ev('10-3')).toBe(7);
  });

  it('multiplication: 2*3 = 6', () => {
    expect(ev('2*3')).toBe(6);
  });

  it('division: 10/4 = 2.5', () => {
    expect(ev('10/4')).toBe(2.5);
  });

  it('exponentiation: 2^3 = 8', () => {
    expect(ev('2^3')).toBe(8);
  });

  it('nested parentheses: (1+2)*3 = 9', () => {
    expect(ev('(1+2)*3')).toBe(9);
  });

  it('operator precedence: 2+3*4 = 14', () => {
    expect(ev('2+3*4')).toBe(14);
  });

  it('chained addition: 1+2+3 = 6', () => {
    expect(ev('1+2+3')).toBe(6);
  });

  it('division result is float: 10/3', () => {
    expect(ev('10/3')).toBeCloseTo(3.333333, 5);
  });

  it('exponentiation is left-associative, like Excel: 2^3^2 = 64', () => {
    // (2^3)^2 = 8^2 = 64
    expect(ev('2^3^2')).toBe(64);
  });

  it('100k nested parentheses evaluate to #VALUE! without throwing', () => {
    expect(ev(`${'('.repeat(100_000)}1${')'.repeat(100_000)}`)).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// Percent
// ---------------------------------------------------------------------------

describe('percent', () => {
  it('50% = 0.5', () => {
    expect(ev('50%')).toBe(0.5);
  });

  it('100% = 1', () => {
    expect(ev('100%')).toBe(1);
  });

  it('200%+1 = 3', () => {
    expect(ev('200%+1')).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// Concatenation
// ---------------------------------------------------------------------------

describe('concatenation', () => {
  it('"hello"&" "&"world" = "hello world"', () => {
    expect(ev('"hello"&" "&"world"')).toBe('hello world');
  });

  it('number & string: 1&"x" = "1x"', () => {
    expect(ev('1&"x"')).toBe('1x');
  });

  it('boolean & string: TRUE&"!" = "TRUE!"', () => {
    expect(ev('TRUE&"!"')).toBe('TRUE!');
  });
});

// ---------------------------------------------------------------------------
// Comparison
// ---------------------------------------------------------------------------

describe('comparison', () => {
  it('1=1 is true', () => {
    expect(ev('1=1')).toBe(true);
  });

  it('1=2 is false', () => {
    expect(ev('1=2')).toBe(false);
  });

  it('1<>2 is true', () => {
    expect(ev('1<>2')).toBe(true);
  });

  it('1<>1 is false', () => {
    expect(ev('1<>1')).toBe(false);
  });

  it('1<2 is true', () => {
    expect(ev('1<2')).toBe(true);
  });

  it('2<1 is false', () => {
    expect(ev('2<1')).toBe(false);
  });

  it('1>2 is false', () => {
    expect(ev('1>2')).toBe(false);
  });

  it('2>1 is true', () => {
    expect(ev('2>1')).toBe(true);
  });

  it('1<=1 is true', () => {
    expect(ev('1<=1')).toBe(true);
  });

  it('1>=1 is true', () => {
    expect(ev('1>=1')).toBe(true);
  });

  it('string comparison is case-insensitive: "abc"="ABC"', () => {
    expect(ev('"abc"="ABC"')).toBe(true);
  });

  it('string ordering: "a"<"b"', () => {
    expect(ev('"a"<"b"')).toBe(true);
  });

  it('mixed types order number < text < FALSE < TRUE', () => {
    expect(ev('2<"1"')).toBe(true);
    expect(ev('"a"<TRUE')).toBe(true);
    expect(ev('"z"<FALSE')).toBe(true);
    expect(ev('FALSE<TRUE')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Blank cells in comparisons
// ---------------------------------------------------------------------------

describe('blank comparison', () => {
  it('a blank equals 0, "" and FALSE (coerced to the other side)', () => {
    expect(ev('A1=0')).toBe(true);
    expect(ev('A1=""')).toBe(true);
    expect(ev('A1=FALSE')).toBe(true);
  });

  it('a blank orders as 0 against numbers and "" against text', () => {
    expect(ev('A1<1')).toBe(true);
    expect(ev('A1<"a"')).toBe(true);
    expect(ev('A1>-1')).toBe(true);
  });

  it('a blank is not 1 and not unequal to 0', () => {
    expect(ev('A1=1')).toBe(false);
    expect(ev('A1<>0')).toBe(false);
  });

  it('a blank is below TRUE (as FALSE) and two blanks are equal', () => {
    expect(ev('A1<TRUE')).toBe(true);
    expect(ev('A1=B1')).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Unary operators
// ---------------------------------------------------------------------------

describe('unary operators', () => {
  it('-5 = -5', () => {
    expect(ev('-5')).toBe(-5);
  });

  it('+3 = 3', () => {
    expect(ev('+3')).toBe(3);
  });

  it('--5 = 5 (double negation)', () => {
    expect(ev('--5')).toBe(5);
  });

  it('-TRUE = -1', () => {
    expect(ev('-TRUE')).toBe(-1);
  });

  it('unary + returns its operand unchanged (OpenFormula 6.4.15)', () => {
    expect(ev('+"abc"')).toBe('abc');
    expect(ev('+TRUE')).toBe(true);
    expect(ev('+"1"')).toBe('1');
  });

  it('unary - still coerces: -"1" = -1', () => {
    expect(ev('-"1"')).toBe(-1);
  });
});

// ---------------------------------------------------------------------------
// References in scalar context: implicit intersection and reference operators
// ---------------------------------------------------------------------------

describe('references in scalar context', () => {
  /** Evaluate with the formula sitting at `currentCell` (0-based col, 1-based row). */
  function evAt(
    formula: string,
    data: Record<string, CellValue>,
    currentCell?: { col: number; row: number },
  ): CellValue {
    return evaluateFormula(formula, { ...mockContext(data), ...(currentCell && { currentCell }) });
  }

  const column = { '0,1': 5, '0,2': 7, '0,3': 9 };
  const grid = { '0,1': 'A1', '1,1': 'B1', '1,2': 'B2', '2,2': 'C2', '2,3': 'C3' };

  it('a multi-cell range without currentCell is #VALUE!', () => {
    expect(evAt('A1:A3', column)).toBe('#VALUE!');
  });

  it('a single-column range intersects the formula row', () => {
    expect(evAt('A1:A3', column, { col: 4, row: 2 })).toBe(7);
  });

  it('a single-row range intersects the formula column', () => {
    expect(evAt('A2:C2', grid, { col: 1, row: 9 })).toBe('B2');
  });

  it('a range outside the formula row or column is #VALUE!', () => {
    expect(evAt('A1:A3', column, { col: 4, row: 9 })).toBe('#VALUE!');
    expect(evAt('A2:C2', grid, { col: 5, row: 9 })).toBe('#VALUE!');
  });

  it('a 2-D range is #VALUE! even with currentCell', () => {
    expect(evAt('A1:B2', grid, { col: 0, row: 1 })).toBe('#VALUE!');
  });

  it('a 1x1 range yields its cell', () => {
    expect(evAt('B2:B2', grid)).toBe('B2');
  });

  it('whole column and row ranges intersect the same way', () => {
    expect(evAt('A:A', column, { col: 3, row: 3 })).toBe(9);
    expect(evAt('2:2', grid, { col: 2, row: 7 })).toBe('C2');
    expect(evAt('A:A', column)).toBe('#VALUE!');
  });

  it('an array constant yields its top-left element', () => {
    expect(ev('{1,2;3,4}')).toBe(1);
  });

  it('an empty array constant is #VALUE!', () => {
    expect(evaluateNode({ type: 'array', rows: [] }, mockContext())).toBe('#VALUE!');
  });

  it('intersection yields the 1x1 overlap', () => {
    expect(evAt('B1:B3 B2:B2', grid)).toBe('B2');
    expect(evAt('B:B 2:2', grid)).toBe('B2');
  });

  it('a multi-cell intersection follows implicit intersection', () => {
    expect(evAt('A1:C3 B1:B3', grid, { col: 0, row: 2 })).toBe('B2');
    expect(evAt('A1:C3 B1:B3', grid)).toBe('#VALUE!');
  });

  it('disjoint references intersect to #NULL!', () => {
    expect(evAt('A1 C3', grid)).toBe('#NULL!');
  });

  it('a union in scalar context is #VALUE!', () => {
    expect(evAt('(A1,B2)', grid)).toBe('#VALUE!');
  });

  it('a range operator built outside the parser is #VALUE!', () => {
    const node: ASTNode = {
      type: 'binary_op',
      op: ':',
      left: { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: false },
      right: { type: 'cell_ref', col: 'A', row: 2, absCol: false, absRow: false },
    };
    expect(evaluateNode(node, mockContext(column))).toBe('#VALUE!');
  });

  it('structured and 3-D references are #REF!', () => {
    expect(ev('Table1[Col]')).toBe('#REF!');
    expect(ev('Sheet1:Sheet3!A1')).toBe('#REF!');
    expect(ev('Sheet1:Sheet3!A1:B2')).toBe('#REF!');
  });

  it('a missing argument evaluates to blank', () => {
    expect(evaluateNode({ type: 'missing' }, mockContext())).toBe(null);
  });
});

// ---------------------------------------------------------------------------
// Cell references
// ---------------------------------------------------------------------------

describe('cell references', () => {
  it('resolves a cell value from context', () => {
    // A1 -> col=0, row=1
    expect(ev('A1', { '0,1': 42 })).toBe(42);
  });

  it('uses cell value in arithmetic', () => {
    expect(ev('A1+10', { '0,1': 5 })).toBe(15);
  });

  it('empty cell is 0 in arithmetic: A1+5 where A1 is empty', () => {
    expect(ev('A1+5')).toBe(5);
  });

  it('empty cell is "" in concatenation: A1&"x" where A1 is empty', () => {
    expect(ev('A1&"x"')).toBe('x');
  });

  it('resolves cross-sheet reference: Sheet2!B3', () => {
    expect(ev('Sheet2!B3', { 'Sheet2!1,3': 99 })).toBe(99);
  });

  it('resolves multi-letter column: AA1', () => {
    expect(ev('AA1', { '26,1': 7 })).toBe(7);
  });

  it.each(['XFE1', 'A0', 'ZZZ1', 'A1048577'])(
    '%s is a name outside A1:XFD1048576 and gives #NAME? without calling getCell',
    (formula) => {
      const getCell = vi.fn<EvalContext['getCell']>(() => 1);
      expect(evaluateFormula(formula, { currentSheet: 'Sheet1', getCell })).toBe('#NAME?');
      expect(getCell).not.toHaveBeenCalled();
    },
  );

  it('never calls getCell outside the grid for XFE1+A0', () => {
    const getCell = vi.fn<EvalContext['getCell']>(() => 1);
    expect(evaluateFormula('XFE1+A0', { currentSheet: 'Sheet1', getCell })).toBe('#NAME?');
    for (const [, col, row] of getCell.mock.calls) {
      expect(col).toBeGreaterThanOrEqual(0);
      expect(col).toBeLessThanOrEqual(16383);
      expect(row).toBeGreaterThanOrEqual(1);
      expect(row).toBeLessThanOrEqual(1048576);
    }
  });

  // Excel (COM, 2026-11-11): each of these shows #NAME?, and IFERROR(XFE1,7) shows 7.
  it.each(['A1:XFE1', 'XFE1:XFE2', 'A1:A1048577', 'Total:A1'])(
    'a range with a name endpoint (%s) gives #NAME? without calling getCell',
    (formula) => {
      const getCell = vi.fn<EvalContext['getCell']>(() => 1);
      expect(evaluateFormula(formula, { currentSheet: 'Sheet1', getCell })).toBe('#NAME?');
      expect(getCell).not.toHaveBeenCalled();
    },
  );

  it.each([
    ['SUM(A1:A99999999999)', '#NAME?'],
    ['SUM(A1,XFE1)', '#NAME?'],
    ['COLUMN(XFE1)', '#NAME?'],
    ['ROW(A0)', '#NAME?'],
    ['ROWS(A1:A1048577)', '#NAME?'],
    ['COLUMNS(XFD1:XFE1)', '#NAME?'],
    ['COLUMNS(Total)', '#NAME?'],
    ['IFERROR(XFE1,7)', 7],
  ])('%s gives %s through the built-in functions', (formula, expected) => {
    const functions = createDefaultFunctions();
    const getCell = vi.fn<EvalContext['getCell']>(() => 1);
    expect(evaluateFormula(formula, { currentSheet: 'Sheet1', getCell, functions })).toBe(expected);
    for (const [, col, row] of getCell.mock.calls) {
      expect(col).toBeLessThanOrEqual(16383);
      expect(row).toBeLessThanOrEqual(1048576);
    }
  });

  it.each([
    ['ROWS(A:A)', 1048576],
    ['ROW(A:A)', 1],
    ['ROWS(1:1)', 1],
    ['COLUMNS(1:1)', 16384],
  ])('%s is %s and never passes on a cell error by implicit intersection', (formula, expected) => {
    const functions = createDefaultFunctions();
    const getCell = vi.fn<EvalContext['getCell']>(() => '#DIV/0!');
    const result = evaluateFormula(formula, {
      currentSheet: 'Sheet1',
      currentCell: { col: 1, row: 2 },
      getCell,
      functions,
    });
    expect(result).toBe(expected);
    expect(getCell).not.toHaveBeenCalled();
  });

  it('still reads the last cell XFD1048576', () => {
    const getCell = vi.fn<EvalContext['getCell']>(() => 5);
    expect(evaluateFormula('XFD1048576', { currentSheet: 'Sheet1', getCell })).toBe(5);
    expect(getCell.mock.calls).toEqual([['Sheet1', 16383, 1048576]]);
  });
});

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

describe('error propagation', () => {
  it('1+#REF! = "#REF!"', () => {
    expect(ev('1+#REF!')).toBe('#REF!');
  });

  it('#DIV/0!+1 = "#DIV/0!"', () => {
    expect(ev('#DIV/0!+1')).toBe('#DIV/0!');
  });

  it('#VALUE! propagates through concatenation', () => {
    expect(ev('#VALUE!&"x"')).toBe('#VALUE!');
  });

  it('#N/A in comparison propagates', () => {
    expect(ev('#N/A=1')).toBe('#N/A');
  });

  it('division by zero: 1/0 = "#DIV/0!"', () => {
    expect(ev('1/0')).toBe('#DIV/0!');
  });

  it('0/0 = "#DIV/0!"', () => {
    expect(ev('0/0')).toBe('#DIV/0!');
  });
});

// ---------------------------------------------------------------------------
// Type coercion
// ---------------------------------------------------------------------------

describe('type coercion', () => {
  it('"5"+3 = 8 (string to number)', () => {
    expect(ev('"5"+3')).toBe(8);
  });

  it('TRUE+1 = 2', () => {
    expect(ev('TRUE+1')).toBe(2);
  });

  it('FALSE+1 = 1', () => {
    expect(ev('FALSE+1')).toBe(1);
  });

  it('"abc"+1 = "#VALUE!" (non-numeric string)', () => {
    expect(ev('"abc"+1')).toBe('#VALUE!');
  });

  it('null (empty cell) + 5 = 5', () => {
    expect(ev('A1+5')).toBe(5);
  });

  it('-"3" = -3 (unary minus on numeric string)', () => {
    expect(ev('-"3"')).toBe(-3);
  });

  it('-"abc" = "#VALUE!" (unary minus on non-numeric string)', () => {
    expect(ev('-"abc"')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// Function calls
// ---------------------------------------------------------------------------

describe('function calls', () => {
  it('unknown function returns "#NAME?"', () => {
    expect(ev('FOO(1,2)')).toBe('#NAME?');
  });

  it('calls a registered function', () => {
    const fns = new Map<string, FormulaFunction>();
    fns.set('SUM', (args, ctx, evaluate) => {
      let total = 0;
      for (const arg of args) {
        const val = evaluate(arg, ctx);
        if (typeof val === 'number') total += val;
      }
      return total;
    });
    expect(ev('SUM(1,2,3)', {}, fns)).toBe(6);
  });

  it('function names are case-insensitive', () => {
    const fns = new Map<string, FormulaFunction>();
    fns.set('ADD', (args, ctx, evaluate) => {
      const [argA, argB] = args;
      assert.isDefined(argA);
      assert.isDefined(argB);
      const a = evaluate(argA, ctx);
      const b = evaluate(argB, ctx);
      if (typeof a === 'number' && typeof b === 'number') return a + b;
      return '#VALUE!';
    });
    expect(ev('add(10,20)', {}, fns)).toBe(30);
  });

  it('function receives raw AST nodes for lazy evaluation', () => {
    const fns = new Map<string, FormulaFunction>();
    fns.set('IF', (args, ctx, evaluate) => {
      const [condArg, thenArg] = args;
      assert.isDefined(condArg);
      assert.isDefined(thenArg);
      const cond = evaluate(condArg, ctx);
      if (cond === true || cond === 1) return evaluate(thenArg, ctx);
      return args[2] ? evaluate(args[2], ctx) : false;
    });
    // Only the true branch should be evaluated
    expect(ev('IF(TRUE,42,1/0)', {}, fns)).toBe(42);
  });
});

// ---------------------------------------------------------------------------
// Named ranges
// ---------------------------------------------------------------------------

describe('named ranges', () => {
  it('returns "#NAME?" for unresolved named range', () => {
    // tokenize('MyRange') yields a single 'name' token; the NameNode is built directly
    // to exercise evaluateNode, which returns #NAME? for an unresolved name.
    const node: ASTNode = { type: 'name', name: 'MyRange' };
    const ctx = mockContext();
    expect(evaluateNode(node, ctx)).toBe('#NAME?');
  });
});

// ---------------------------------------------------------------------------
// Literal values
// ---------------------------------------------------------------------------

describe('literals', () => {
  it('number literal', () => {
    expect(ev('42')).toBe(42);
  });

  it('string literal', () => {
    expect(ev('"hello"')).toBe('hello');
  });

  it('boolean TRUE', () => {
    expect(ev('TRUE')).toBe(true);
  });

  it('boolean FALSE', () => {
    expect(ev('FALSE')).toBe(false);
  });

  it('error literal', () => {
    expect(ev('#N/A')).toBe('#N/A');
  });

  it('decimal number', () => {
    expect(ev('3.14')).toBeCloseTo(3.14, 10);
  });
});

// ---------------------------------------------------------------------------
// Complex expressions
// ---------------------------------------------------------------------------

describe('complex expressions', () => {
  it('(A1+A2)*B1 with cell values', () => {
    expect(ev('(A1+A2)*B1', { '0,1': 3, '0,2': 7, '1,1': 2 })).toBe(20);
  });

  it('nested arithmetic and comparison: (1+2)*3>8', () => {
    expect(ev('(1+2)*3>8')).toBe(true);
  });

  it('chained concatenation: "a"&"b"&"c"', () => {
    expect(ev('"a"&"b"&"c"')).toBe('abc');
  });

  it('percent in expression: 50%*200', () => {
    expect(ev('50%*200')).toBe(100);
  });

  it('mixed arithmetic and concatenation precedence: 1+2&3+4', () => {
    // & has lower precedence than +, so: (1+2) & (3+4) = "3" & "7" = "37"
    expect(ev('1+2&3+4')).toBe('37');
  });
});

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------

describe('edge cases', () => {
  it('empty formula returns null', () => {
    expect(ev('')).toBe(null);
  });

  it('percent of error propagates', () => {
    expect(ev('#REF!%')).toBe('#REF!');
  });

  it('zero to the power zero: 0^0 = #NUM! (as in Excel)', () => {
    expect(ev('0^0')).toBe('#NUM!');
  });

  it('negative exponent: 2^-1 = 0.5', () => {
    expect(ev('2^-1')).toBe(0.5);
  });
});

describe('#REF! range endpoints', () => {
  it('propagates #REF! from a range operator endpoint', () => {
    expect(ev('=#REF!:C1')).toBe('#REF!');
    expect(ev('=SUM(#REF!:C2)', {}, createDefaultFunctions())).toBe('#REF!');
  });
});
