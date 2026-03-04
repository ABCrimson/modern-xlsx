import { describe, expect, it } from 'vitest';
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
    functions,
    getCell(sheet: string, col: number, row: number): CellValue {
      const key = `${sheet}!${col},${row}`;
      if (key in data) return data[key];
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

  it('exponentiation right-associativity: 2^3^2 = 512', () => {
    // 2^(3^2) = 2^9 = 512
    expect(ev('2^3^2')).toBe(512);
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
      const a = evaluate(args[0], ctx);
      const b = evaluate(args[1], ctx);
      if (typeof a === 'number' && typeof b === 'number') return a + b;
      return '#VALUE!';
    });
    expect(ev('add(10,20)', {}, fns)).toBe(30);
  });

  it('function receives raw AST nodes for lazy evaluation', () => {
    const fns = new Map<string, FormulaFunction>();
    fns.set('IF', (args, ctx, evaluate) => {
      const cond = evaluate(args[0], ctx);
      if (cond === true || cond === 1) return evaluate(args[1], ctx);
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
    // Named ranges are parsed as "name" nodes; evaluator returns #NAME?
    // Note: This only works if the tokenizer/parser produces a NameNode.
    // The formula "MyRange" may tokenize as a cell_ref or name depending
    // on the tokenizer. We test via the evaluateNode path directly.
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

  it('exponent of zero: 0^0 = 1', () => {
    expect(ev('0^0')).toBe(1);
  });

  it('negative exponent: 2^-1 = 0.5', () => {
    expect(ev('2^-1')).toBe(0.5);
  });
});
