import { describe, expect, it } from 'vitest';
import { EXCEL_ERROR_LITERALS } from '../src/formula/errors.js';
import { evaluateFormula } from '../src/formula/evaluator.js';
import type { EvalContext } from '../src/formula/resolver.js';
import type { TokenType } from '../src/formula/tokenizer.js';
import { tokenize } from '../src/formula/tokenizer.js';
import { Workbook } from '../src/index.js';

/** One table row: the input text and its exact [type, value] token pairs. */
interface TokenCase {
  readonly formula: string;
  readonly tokens: readonly (readonly [TokenType, string])[];
}

/** Helper: extract [type, value] pairs from tokens */
function tv(formula: string): [TokenType, string][] {
  const { tokens } = tokenize(formula);
  return tokens.map((t) => [t.type, t.value]);
}

describe('tokenize', () => {
  describe('leading = sign', () => {
    it('strips leading = from formulas', () => {
      const result = tokenize('=1+2');
      expect(result.tokens).toHaveLength(3);
      expect(result.tokens[0]?.value).toBe('1');
    });

    it('works without leading =', () => {
      const result = tokenize('1+2');
      expect(result.tokens).toHaveLength(3);
      expect(result.tokens[0]?.value).toBe('1');
    });
  });

  describe('simple arithmetic', () => {
    it.each<TokenCase>([
      {
        formula: '1+2',
        tokens: [
          ['number', '1'],
          ['operator', '+'],
          ['number', '2'],
        ],
      },
      {
        formula: '3.14*2',
        tokens: [
          ['number', '3.14'],
          ['operator', '*'],
          ['number', '2'],
        ],
      },
      {
        formula: '10/3-1',
        tokens: [
          ['number', '10'],
          ['operator', '/'],
          ['number', '3'],
          ['operator', '-'],
          ['number', '1'],
        ],
      },
      // tokenizes power operator
      {
        formula: '2^8',
        tokens: [
          ['number', '2'],
          ['operator', '^'],
          ['number', '8'],
        ],
      },
      // tokenizes concatenation operator
      {
        formula: '"a"&"b"',
        tokens: [
          ['string', 'a'],
          ['operator', '&'],
          ['string', 'b'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('numbers', () => {
    it('tokenizes integers', () => {
      expect(tv('123')[0]).toEqual(['number', '123']);
    });

    it('tokenizes decimals', () => {
      expect(tv('3.14')[0]).toEqual(['number', '3.14']);
    });

    it('tokenizes leading decimal point', () => {
      expect(tv('.5')[0]).toEqual(['number', '.5']);
    });

    it('tokenizes exponent notation', () => {
      expect(tv('1.5E10')[0]).toEqual(['number', '1.5E10']);
    });

    it('tokenizes lowercase exponent', () => {
      expect(tv('2e-3')[0]).toEqual(['number', '2e-3']);
    });

    it('tokenizes exponent with plus sign', () => {
      expect(tv('1E+5')[0]).toEqual(['number', '1E+5']);
    });

    it('tokenizes zero', () => {
      expect(tv('0')[0]).toEqual(['number', '0']);
    });
  });

  describe('cell references', () => {
    it('tokenizes simple cell ref', () => {
      expect(tv('A1')[0]).toEqual(['cell_ref', 'A1']);
    });

    it('tokenizes absolute column', () => {
      expect(tv('$A1')[0]).toEqual(['cell_ref', '$A1']);
    });

    it('tokenizes absolute row', () => {
      expect(tv('A$1')[0]).toEqual(['cell_ref', 'A$1']);
    });

    it('tokenizes fully absolute ref', () => {
      expect(tv('$A$1')[0]).toEqual(['cell_ref', '$A$1']);
    });

    it('tokenizes double-letter column', () => {
      expect(tv('AA100')[0]).toEqual(['cell_ref', 'AA100']);
    });

    it('tokenizes triple-letter column', () => {
      expect(tv('XFD1048576')[0]).toEqual(['cell_ref', 'XFD1048576']);
    });
  });

  describe('sheet-qualified references', () => {
    it('tokenizes unquoted sheet name', () => {
      expect(tv('Sheet1!A1')[0]).toEqual(['cell_ref', 'Sheet1!A1']);
    });

    it('tokenizes quoted sheet name', () => {
      expect(tv("'My Sheet'!B2")[0]).toEqual(['cell_ref', "'My Sheet'!B2"]);
    });

    it('tokenizes quoted sheet with escaped quote', () => {
      expect(tv("'It''s Sheet'!C3")[0]).toEqual(['cell_ref', "'It''s Sheet'!C3"]);
    });

    it('tokenizes sheet-qualified absolute ref', () => {
      expect(tv('Sheet1!$A$1')[0]).toEqual(['cell_ref', 'Sheet1!$A$1']);
    });
  });

  describe('ranges via colon', () => {
    it.each<TokenCase>([
      // tokenizes A1:B2 as cell_ref + colon + cell_ref
      {
        formula: 'A1:B2',
        tokens: [
          ['cell_ref', 'A1'],
          ['colon', ':'],
          ['cell_ref', 'B2'],
        ],
      },
      // tokenizes absolute range
      {
        formula: '$A$1:$B$10',
        tokens: [
          ['cell_ref', '$A$1'],
          ['colon', ':'],
          ['cell_ref', '$B$10'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('cell reference bounds (A1:XFD1048576)', () => {
    it('tokenizes the last cell XFD1048576 as a cell_ref', () => {
      expect(tokenize('XFD1048576')).toEqual({
        tokens: [{ type: 'cell_ref', value: 'XFD1048576', start: 0, end: 10 }],
        errors: [],
      });
    });

    it('tokenizes a column past XFD as a name', () => {
      expect(tokenize('XFE1')).toEqual({
        tokens: [{ type: 'name', value: 'XFE1', start: 0, end: 4 }],
        errors: [],
      });
    });

    it('tokenizes row 0, row 1048577 and ZZZ1 as names', () => {
      expect(tv('A0')).toEqual([['name', 'A0']]);
      expect(tv('A1048577')).toEqual([['name', 'A1048577']]);
      expect(tv('ZZZ1')).toEqual([['name', 'ZZZ1']]);
      expect(tv('xfe1')).toEqual([['name', 'xfe1']]);
    });

    it('reports an out-of-bounds A$1 shape (Excel rejects the formula)', () => {
      expect(tokenize('XFE$1')).toEqual({
        tokens: [],
        errors: ['Invalid cell reference at position 0: XFE$1'],
      });
      expect(tokenize('A$0').errors).toEqual(['Invalid cell reference at position 0: A$0']);
      expect(tv('XFD$1048576')).toEqual([['cell_ref', 'XFD$1048576']]);
    });

    it('reports an out-of-bounds $-prefixed reference', () => {
      const result = tokenize('$XFE$1');
      expect(result.tokens).toEqual([]);
      expect(result.errors).toEqual(['Invalid cell reference at position 0: $XFE$1']);
      expect(tokenize('$A0').errors).toEqual(['Invalid cell reference at position 0: $A0']);
    });

    it('reports an out-of-bounds sheet-qualified reference', () => {
      const result = tokenize('Sheet1!XFE1');
      expect(result.tokens).toEqual([]);
      expect(result.errors).toEqual(['Invalid cell reference at position 0: Sheet1!XFE1']);
      expect(tokenize("1+'My Sheet'!A0").errors).toEqual([
        "Invalid cell reference at position 2: 'My Sheet'!A0",
      ]);
      expect(tv('Sheet1!$XFD$1048576')).toEqual([['cell_ref', 'Sheet1!$XFD$1048576']]);
    });

    it('applies the bound to each range endpoint', () => {
      expect(tv('A1:XFE1')).toEqual([
        ['cell_ref', 'A1'],
        ['colon', ':'],
        ['name', 'XFE1'],
      ]);
      expect(tokenize('Sheet1!A1:Sheet1!A1048577').errors).toEqual([
        'Invalid cell reference at position 10: Sheet1!A1048577',
      ]);
    });
  });

  describe('function calls', () => {
    it.each<TokenCase>([
      {
        formula: 'SUM(A1:A10)',
        tokens: [
          ['function', 'SUM'],
          ['paren_open', '('],
          ['cell_ref', 'A1'],
          ['colon', ':'],
          ['cell_ref', 'A10'],
          ['paren_close', ')'],
        ],
      },
      {
        formula: 'IF(A1>0,1,0)',
        tokens: [
          ['function', 'IF'],
          ['paren_open', '('],
          ['cell_ref', 'A1'],
          ['operator', '>'],
          ['number', '0'],
          ['comma', ','],
          ['number', '1'],
          ['comma', ','],
          ['number', '0'],
          ['paren_close', ')'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('tokenizes VLOOKUP with multiple args', () => {
      const result = tv('VLOOKUP(A1,B1:D10,3,FALSE)');
      expect(result[0]).toEqual(['function', 'VLOOKUP']);
      expect(result[result.length - 2]).toEqual(['boolean', 'FALSE']);
    });
  });

  describe('nested functions', () => {
    it('tokenizes IF(SUM(A1:A5)>10,TRUE,FALSE)', () => {
      const result = tv('IF(SUM(A1:A5)>10,TRUE,FALSE)');
      expect(result).toEqual([
        ['function', 'IF'],
        ['paren_open', '('],
        ['function', 'SUM'],
        ['paren_open', '('],
        ['cell_ref', 'A1'],
        ['colon', ':'],
        ['cell_ref', 'A5'],
        ['paren_close', ')'],
        ['operator', '>'],
        ['number', '10'],
        ['comma', ','],
        ['boolean', 'TRUE'],
        ['comma', ','],
        ['boolean', 'FALSE'],
        ['paren_close', ')'],
      ]);
    });
  });

  describe('string literals', () => {
    it('tokenizes simple string', () => {
      expect(tv('"hello"')[0]).toEqual(['string', 'hello']);
    });

    it('tokenizes string with escaped quotes', () => {
      expect(tv('"say ""hi"""')[0]).toEqual(['string', 'say "hi"']);
    });

    it('tokenizes empty string', () => {
      expect(tv('""')[0]).toEqual(['string', '']);
    });

    it('reports unterminated string', () => {
      const result = tokenize('"oops');
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0]).toContain('Unterminated string');
    });
  });

  describe('boolean values', () => {
    it('tokenizes TRUE', () => {
      expect(tv('TRUE')[0]).toEqual(['boolean', 'TRUE']);
    });

    it('tokenizes FALSE', () => {
      expect(tv('FALSE')[0]).toEqual(['boolean', 'FALSE']);
    });

    it('tokenizes case-insensitive true', () => {
      expect(tv('true')[0]).toEqual(['boolean', 'TRUE']);
    });

    it('tokenizes case-insensitive False', () => {
      expect(tv('False')[0]).toEqual(['boolean', 'FALSE']);
    });
  });

  describe('error values', () => {
    it('tokenizes #N/A', () => {
      expect(tv('#N/A')[0]).toEqual(['error', '#N/A']);
    });

    it('tokenizes #REF!', () => {
      expect(tv('#REF!')[0]).toEqual(['error', '#REF!']);
    });

    it('tokenizes #VALUE!', () => {
      expect(tv('#VALUE!')[0]).toEqual(['error', '#VALUE!']);
    });

    it('tokenizes #DIV/0!', () => {
      expect(tv('#DIV/0!')[0]).toEqual(['error', '#DIV/0!']);
    });

    it('tokenizes #NULL!', () => {
      expect(tv('#NULL!')[0]).toEqual(['error', '#NULL!']);
    });

    it('tokenizes #NAME?', () => {
      expect(tv('#NAME?')[0]).toEqual(['error', '#NAME?']);
    });

    it('tokenizes #NUM!', () => {
      expect(tv('#NUM!')[0]).toEqual(['error', '#NUM!']);
    });

    it.each(EXCEL_ERROR_LITERALS)('tokenizes %s as one error token', (literal) => {
      const result = tokenize(literal);
      expect(result.errors).toEqual([]);
      expect(tv(literal)).toEqual([['error', literal]]);
    });

    it('stops an unknown error literal at the first operator', () => {
      const result = tokenize('#FOO+1');
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0]).toContain(': #FOO');
      expect(result.errors[0]).not.toContain('+1');
      expect(tv('#FOO+1')).toContainEqual(['operator', '+']);
      expect(tv('#FOO+1')).toContainEqual(['number', '1']);
    });

    it('keeps a trailing ! or ? on an unknown error literal', () => {
      expect(tokenize('#FOO!').errors).toEqual(['Unknown error literal at position 0: #FOO!']);
      expect(tokenize('#BAR?,1').errors).toEqual(['Unknown error literal at position 0: #BAR?']);
    });

    it('tokenizes a sheet-qualified error literal as the error alone', () => {
      expect(tokenize('Sheet1!#REF!').errors).toEqual([]);
      expect(tv('Sheet1!#REF!')).toEqual([['error', '#REF!']]);
      expect(tokenize('Sheet1!#REF!').tokens[0]).toMatchObject({ start: 0, end: 12 });
      expect(tv("'My Sheet'!#REF!+1")).toEqual([
        ['error', '#REF!'],
        ['operator', '+'],
        ['number', '1'],
      ]);
    });

    it('reports a sheet prefix with no reference once, without a name token', () => {
      const result = tokenize('Sheet1!');
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0]).toMatch(/Expected reference after 'Sheet1!'/);
      expect(result.tokens.some((t) => t.type === 'name')).toBe(false);
    });
  });

  describe('comparison operators', () => {
    it.each<TokenCase>([
      // tokenizes >=
      {
        formula: 'A1>=10',
        tokens: [
          ['cell_ref', 'A1'],
          ['operator', '>='],
          ['number', '10'],
        ],
      },
      // tokenizes <=
      {
        formula: 'A1<=10',
        tokens: [
          ['cell_ref', 'A1'],
          ['operator', '<='],
          ['number', '10'],
        ],
      },
      // tokenizes <>
      {
        formula: 'A1<>0',
        tokens: [
          ['cell_ref', 'A1'],
          ['operator', '<>'],
          ['number', '0'],
        ],
      },
      // tokenizes = as operator
      {
        formula: 'A1=B1',
        tokens: [
          ['cell_ref', 'A1'],
          ['operator', '='],
          ['cell_ref', 'B1'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('tokenizes < and >', () => {
      expect(tv('A1<B1')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '<'],
        ['cell_ref', 'B1'],
      ]);
      expect(tv('A1>B1')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '>'],
        ['cell_ref', 'B1'],
      ]);
    });
  });

  describe('unary operators', () => {
    it.each<TokenCase>([
      // tokenizes unary minus at start
      {
        formula: '-A1',
        tokens: [
          ['prefix_op', '-'],
          ['cell_ref', 'A1'],
        ],
      },
      // tokenizes unary plus at start
      {
        formula: '+5',
        tokens: [
          ['prefix_op', '+'],
          ['number', '5'],
        ],
      },
      // tokenizes unary after operator
      {
        formula: '1+-2',
        tokens: [
          ['number', '1'],
          ['operator', '+'],
          ['prefix_op', '-'],
          ['number', '2'],
        ],
      },
      // tokenizes unary after open paren
      {
        formula: '(-1)',
        tokens: [
          ['paren_open', '('],
          ['prefix_op', '-'],
          ['number', '1'],
          ['paren_close', ')'],
        ],
      },
      // tokenizes unary after comma
      {
        formula: 'SUM(-1,-2)',
        tokens: [
          ['function', 'SUM'],
          ['paren_open', '('],
          ['prefix_op', '-'],
          ['number', '1'],
          ['comma', ','],
          ['prefix_op', '-'],
          ['number', '2'],
          ['paren_close', ')'],
        ],
      },
      // tokenizes binary minus after cell ref
      {
        formula: 'A1-B1',
        tokens: [
          ['cell_ref', 'A1'],
          ['operator', '-'],
          ['cell_ref', 'B1'],
        ],
      },
      // tokenizes binary plus after number
      {
        formula: '1+2',
        tokens: [
          ['number', '1'],
          ['operator', '+'],
          ['number', '2'],
        ],
      },
      // tokenizes binary minus after close paren
      {
        formula: '(1)-2',
        tokens: [
          ['paren_open', '('],
          ['number', '1'],
          ['paren_close', ')'],
          ['operator', '-'],
          ['number', '2'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('percent', () => {
    it.each<TokenCase>([
      {
        formula: '50%',
        tokens: [
          ['number', '50'],
          ['percent', '%'],
        ],
      },
      // tokenizes percent after cell ref
      {
        formula: 'A1%',
        tokens: [
          ['cell_ref', 'A1'],
          ['percent', '%'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('array constants', () => {
    it.each<TokenCase>([
      {
        formula: '{1,2;3,4}',
        tokens: [
          ['array_open', '{'],
          ['number', '1'],
          ['array_col_sep', ','],
          ['number', '2'],
          ['array_row_sep', ';'],
          ['number', '3'],
          ['array_col_sep', ','],
          ['number', '4'],
          ['array_close', '}'],
        ],
      },
      // uses comma/semicolon outside arrays
      {
        formula: 'SUM(1,2)',
        tokens: [
          ['function', 'SUM'],
          ['paren_open', '('],
          ['number', '1'],
          ['comma', ','],
          ['number', '2'],
          ['paren_close', ')'],
        ],
      },
      // tokenizes array with strings
      {
        formula: '{"a","b";"c","d"}',
        tokens: [
          ['array_open', '{'],
          ['string', 'a'],
          ['array_col_sep', ','],
          ['string', 'b'],
          ['array_row_sep', ';'],
          ['string', 'c'],
          ['array_col_sep', ','],
          ['string', 'd'],
          ['array_close', '}'],
        ],
      },
      // tokenizes unary minus inside array
      {
        formula: '{-1,2}',
        tokens: [
          ['array_open', '{'],
          ['prefix_op', '-'],
          ['number', '1'],
          ['array_col_sep', ','],
          ['number', '2'],
          ['array_close', '}'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('named ranges', () => {
    it('tokenizes named range', () => {
      expect(tv('MyRange')[0]).toEqual(['name', 'MyRange']);
    });

    it('tokenizes underscore-prefixed name', () => {
      expect(tv('_custom')[0]).toEqual(['name', '_custom']);
    });

    it('tokenizes name with dots', () => {
      expect(tv('Data.Total')[0]).toEqual(['name', 'Data.Total']);
    });

    it('distinguishes name from function', () => {
      // Without paren -> name; with paren -> function
      expect(tv('MyFunc')[0]).toEqual(['name', 'MyFunc']);
      expect(tv('MyFunc(')[0]).toEqual(['function', 'MyFunc']);
    });
  });

  describe('complex formulas', () => {
    it('tokenizes IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")', () => {
      const result = tv('IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")');
      expect(result).toEqual([
        ['function', 'IF'],
        ['paren_open', '('],
        ['function', 'AND'],
        ['paren_open', '('],
        ['cell_ref', 'A1'],
        ['operator', '>'],
        ['number', '0'],
        ['comma', ','],
        ['cell_ref', 'B1'],
        ['operator', '<'],
        ['number', '10'],
        ['paren_close', ')'],
        ['comma', ','],
        ['function', 'SUM'],
        ['paren_open', '('],
        ['cell_ref', 'C1'],
        ['colon', ':'],
        ['cell_ref', 'C10'],
        ['paren_close', ')'],
        ['operator', '*'],
        ['number', '1.1'],
        ['comma', ','],
        ['string', 'N/A'],
        ['paren_close', ')'],
      ]);
    });

    it.each<TokenCase>([
      {
        formula: 'IFERROR(VLOOKUP(A1,Data!B:C,2,0),"")',
        tokens: [
          ['function', 'IFERROR'],
          ['paren_open', '('],
          ['function', 'VLOOKUP'],
          ['paren_open', '('],
          ['cell_ref', 'A1'],
          ['comma', ','],
          ['column_range', 'Data!B:C'],
          ['comma', ','],
          ['number', '2'],
          ['comma', ','],
          ['number', '0'],
          ['paren_close', ')'],
          ['comma', ','],
          ['string', ''],
          ['paren_close', ')'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('tokenizes formula with mixed references and operators', () => {
      const result = tv('$A$1+Sheet1!B2*100%');
      expect(result).toEqual([
        ['cell_ref', '$A$1'],
        ['operator', '+'],
        ['cell_ref', 'Sheet1!B2'],
        ['operator', '*'],
        ['number', '100'],
        ['percent', '%'],
      ]);
    });
  });

  describe('whitespace handling', () => {
    it.each<TokenCase>([
      // skips spaces between tokens
      {
        formula: '1 + 2',
        tokens: [
          ['number', '1'],
          ['operator', '+'],
          ['number', '2'],
        ],
      },
      // skips tabs and newlines
      {
        formula: '1\t+\n2',
        tokens: [
          ['number', '1'],
          ['operator', '+'],
          ['number', '2'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('token positions', () => {
    it('tracks start and end positions', () => {
      const { tokens } = tokenize('=A1+B2');
      // Positions index the original string including `=`: A1 is at 1..3, + at 3..4, B2 at 4..6
      expect(tokens[0]).toEqual({ type: 'cell_ref', value: 'A1', start: 1, end: 3 });
      expect(tokens[1]).toEqual({ type: 'operator', value: '+', start: 3, end: 4 });
      expect(tokens[2]).toEqual({ type: 'cell_ref', value: 'B2', start: 4, end: 6 });
    });

    it('tracks positions with spaces', () => {
      const { tokens } = tokenize('1 + 2');
      expect(tokens[0]).toEqual({ type: 'number', value: '1', start: 0, end: 1 });
      expect(tokens[1]).toEqual({ type: 'operator', value: '+', start: 2, end: 3 });
      expect(tokens[2]).toEqual({ type: 'number', value: '2', start: 4, end: 5 });
    });
  });

  describe('error handling', () => {
    it('reports unterminated string', () => {
      const result = tokenize('"unclosed');
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.errors[0]).toContain('Unterminated string');
    });

    it('reports unknown character', () => {
      const result = tokenize('1+~2');
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.errors[0]).toContain('~');
      // Should still tokenize other parts
      expect(result.tokens.length).toBeGreaterThanOrEqual(2);
    });

    it('handles empty formula', () => {
      const result = tokenize('');
      expect(result.tokens).toHaveLength(0);
      expect(result.errors).toHaveLength(0);
    });

    it('handles formula that is just =', () => {
      const result = tokenize('=');
      expect(result.tokens).toHaveLength(0);
      expect(result.errors).toHaveLength(0);
    });

    it('reports unterminated quoted sheet name', () => {
      const result = tokenize("'Sheet");
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.errors[0]).toContain('Unterminated quoted sheet name');
    });
  });

  describe('semicolons outside arrays', () => {
    it.each<TokenCase>([
      // emits semicolon type outside arrays
      {
        formula: 'SUM(1;2)',
        tokens: [
          ['function', 'SUM'],
          ['paren_open', '('],
          ['number', '1'],
          ['semicolon', ';'],
          ['number', '2'],
          ['paren_close', ')'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('reports a semicolon outside an array constant', () => {
      const { errors } = tokenize('SUM(1;2)');
      expect(errors).toHaveLength(1);
      expect(errors[0]).toMatch(/only valid as an array-constant row separator/);
      expect(errors[0]).toContain('position 5');
    });

    it('accepts semicolons inside an array constant', () => {
      expect(tokenize('{1,2;3,4}').errors).toEqual([]);
    });
  });

  describe('whole-column and whole-row ranges', () => {
    it.each([
      ['A:A', 'column_range'],
      ['$A:$A', 'column_range'],
      ["'My Sheet'!A:B", 'column_range'],
      ['a:xfd', 'column_range'],
      ['Sheet1:Sheet3!C:$D', 'column_range'],
      ['1:1', 'row_range'],
      ['$1:$1', 'row_range'],
      ['Sheet1!2:5', 'row_range'],
      ['1:1048576', 'row_range'],
      ["'Q1:Q4'!$3:7", 'row_range'],
    ] as const)('tokenizes %s as one %s token', (formula, type) => {
      expect(tokenize(formula).errors).toEqual([]);
      expect(tv(formula)).toEqual([[type, formula]]);
    });

    it('tokenizes ROW(1:3) and SUM(A:A)', () => {
      expect(tv('ROW(1:3)')).toEqual([
        ['function', 'ROW'],
        ['paren_open', '('],
        ['row_range', '1:3'],
        ['paren_close', ')'],
      ]);
      expect(tv('SUM(A:A)')).toEqual([
        ['function', 'SUM'],
        ['paren_open', '('],
        ['column_range', 'A:A'],
        ['paren_close', ')'],
      ]);
    });

    it('does not read out-of-grid or mixed operands as ranges', () => {
      expect(tv('A:XFE').map(([type]) => type)).toEqual(['name', 'colon', 'name']);
      expect(tv('0:3').map(([type]) => type)).toEqual(['number', 'colon', 'number']);
      expect(tv('1:1048577').map(([type]) => type)).toEqual(['number', 'colon', 'number']);
      expect(tv('1:2.5').map(([type]) => type)).toEqual(['number', 'colon', 'number']);
      expect(tv('A:A1')).toEqual([
        ['name', 'A'],
        ['colon', ':'],
        ['cell_ref', 'A1'],
      ]);
    });

    it('keeps numbers inside array constants and strings', () => {
      expect(tv('{1:2}').map(([type]) => type)).toEqual([
        'array_open',
        'number',
        'colon',
        'number',
        'array_close',
      ]);
      expect(tv('"1:2"')).toEqual([['string', '1:2']]);
    });
  });

  describe('structured references', () => {
    it.each([
      'Table1[Col]',
      'Table1[[#Headers],[Col]]',
      '[@Col]',
      'Table1[[#This Row],[Unit Price]]',
      "Table1[Col']Name]",
    ])('tokenizes %s as one structured_ref token', (formula) => {
      expect(tokenize(formula).errors).toEqual([]);
      expect(tv(formula)).toEqual([['structured_ref', formula]]);
    });

    it.each<TokenCase>([
      // tokenizes a structured reference inside a function call
      {
        formula: 'SUM(Sales[Amount])',
        tokens: [
          ['function', 'SUM'],
          ['paren_open', '('],
          ['structured_ref', 'Sales[Amount]'],
          ['paren_close', ')'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('reports an unbalanced bracket', () => {
      const result = tokenize('Table1[[#Headers],[Col]');
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0]).toMatch(/structured reference/);
      expect(result.tokens).toEqual([]);
    });
  });

  describe('3-D references', () => {
    it('tokenizes an unquoted sheet span as one cell_ref', () => {
      expect(tv('Sheet1:Sheet3!A1')).toEqual([['cell_ref', 'Sheet1:Sheet3!A1']]);
      expect(tokenize('Sheet1:Sheet3!A1').errors).toEqual([]);
    });

    it.each<TokenCase>([
      // tokenizes a quoted sheet span as one cell_ref
      { formula: "'Jan 1:Mar 1'!B2", tokens: [['cell_ref', "'Jan 1:Mar 1'!B2"]] },
      // keeps a range whose start is a plain cell ref
      {
        formula: 'A1:Sheet2!B1',
        tokens: [
          ['cell_ref', 'A1'],
          ['colon', ':'],
          ['cell_ref', 'Sheet2!B1'],
        ],
      },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });
  });

  describe('grid bounds', () => {
    // Excel reads an unqualified cell-shaped identifier outside the grid as a name.
    it.each(['XFE1', 'A0', 'A1048577'])('tokenizes unqualified %s as a name', (formula) => {
      expect(tv(formula)).toEqual([['name', formula]]);
      expect(tokenize(formula).errors).toEqual([]);
    });

    it.each<TokenCase>([
      // keeps the last cell of the grid a cell_ref
      { formula: 'XFD1048576', tokens: [['cell_ref', 'XFD1048576']] },
    ])('tokenizes $formula', ({ formula, tokens }) => {
      expect(tv(formula)).toEqual(tokens);
    });

    it('reports out-of-grid sheet-qualified or absolute refs', () => {
      expect(tokenize('Sheet1!XFE1').errors).toEqual([
        'Invalid cell reference at position 0: Sheet1!XFE1',
      ]);
      expect(tokenize('$A$0').errors).toEqual(['Invalid cell reference at position 0: $A$0']);
      expect(tokenize('A$1048577').errors).toHaveLength(1);
    });
  });

  describe('intersection operator', () => {
    it('emits a single-space operator between two references', () => {
      const result = tv('A1:B5 B2:C3');
      expect(result).toEqual([
        ['cell_ref', 'A1'],
        ['colon', ':'],
        ['cell_ref', 'B5'],
        ['operator', ' '],
        ['cell_ref', 'B2'],
        ['colon', ':'],
        ['cell_ref', 'C3'],
      ]);
    });

    it('collapses a whitespace run to one operator', () => {
      expect(tv('A:A \t 1:1')).toEqual([
        ['column_range', 'A:A'],
        ['operator', ' '],
        ['row_range', '1:1'],
      ]);
      expect(tv('(A1) Sales[Amount]')).toContainEqual(['operator', ' ']);
    });

    it('leaves other whitespace insignificant', () => {
      expect(tv('1 + 2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['number', '2'],
      ]);
      expect(tv('SUM( A1 , B2 )')).not.toContainEqual(['operator', ' ']);
      expect(tv('A1 -B1')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '-'],
        ['cell_ref', 'B1'],
      ]);
      expect(tv('A1 2')).not.toContainEqual(['operator', ' ']);
    });
  });

  describe('edge cases', () => {
    it('tokenizes consecutive operators correctly', () => {
      // 1*-2 should be number, operator, prefix_op, number
      expect(tv('1*-2')).toEqual([
        ['number', '1'],
        ['operator', '*'],
        ['prefix_op', '-'],
        ['number', '2'],
      ]);
    });

    it('tokenizes R1 as a cell_ref', () => {
      // R1 matches the cell reference pattern (row 1, column R); R1C1 tokenizes as a name
      const result = tv('R1');
      expect(result[0]?.[0]).toBe('cell_ref');
    });

    it('tokenizes multiple error values in a formula', () => {
      const result = tv('IF(A1=#N/A,#REF!,0)');
      expect(result).toEqual([
        ['function', 'IF'],
        ['paren_open', '('],
        ['cell_ref', 'A1'],
        ['operator', '='],
        ['error', '#N/A'],
        ['comma', ','],
        ['error', '#REF!'],
        ['comma', ','],
        ['number', '0'],
        ['paren_close', ')'],
      ]);
    });

    it('handles deeply nested formulas', () => {
      const result = tokenize('IF(IF(IF(1,2,3),4,5),6,7)');
      expect(result.errors).toHaveLength(0);
      expect(result.tokens.filter((t) => t.type === 'function')).toHaveLength(3);
    });

    it('handles long cell reference', () => {
      expect(tv('$XFD$1048576')[0]).toEqual(['cell_ref', '$XFD$1048576']);
    });

    it('tokenizes double unary minus', () => {
      // --1 should be prefix_op, prefix_op, number
      expect(tv('--1')).toEqual([
        ['prefix_op', '-'],
        ['prefix_op', '-'],
        ['number', '1'],
      ]);
    });

    it('tokenizes formula with all operator types', () => {
      const result = tokenize('=1+2-3*4/5^6&"x"=1<>2<3>4<=5>=6');
      expect(result.errors).toHaveLength(0);
      const ops = result.tokens.filter((t) => t.type === 'operator').map((t) => t.value);
      expect(ops).toEqual(['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>=']);
    });
  });
});

describe('non-ASCII names and unquoted sheet prefixes', () => {
  it.each([
    ['Données!A1', [['cell_ref', 'Données!A1']]],
    [
      'Лист1!A1:B2',
      [
        ['cell_ref', 'Лист1!A1'],
        ['colon', ':'],
        ['cell_ref', 'B2'],
      ],
    ],
    [
      '売上+1',
      [
        ['name', '売上'],
        ['operator', '+'],
        ['number', '1'],
      ],
    ],
    ['Лист1:Лист3!A1', [['cell_ref', 'Лист1:Лист3!A1']]],
    [
      'SUM(Umsätze)',
      [
        ['function', 'SUM'],
        ['paren_open', '('],
        ['name', 'Umsätze'],
        ['paren_close', ')'],
      ],
    ],
    [
      'A1 Données!A1',
      [
        ['cell_ref', 'A1'],
        ['operator', ' '],
        ['cell_ref', 'Données!A1'],
      ],
    ],
    ['𝔸𝔹!A1', [['cell_ref', '𝔸𝔹!A1']]],
  ] as const)('tokenizes %s without errors', (formula, tokens) => {
    expect(tokenize(formula).errors).toEqual([]);
    expect(tv(formula)).toEqual(tokens);
  });

  it('keeps column letters and whole-column ranges ASCII', () => {
    expect(tv('A:Aé')).toEqual([
      ['name', 'A'],
      ['colon', ':'],
      ['name', 'Aé'],
    ]);
    expect(tokenize('$É1').errors.length).toBeGreaterThan(0);
  });

  it('evaluates a reference on an unquoted non-ASCII sheet', () => {
    const calls: [string, number, number][] = [];
    const ctx: EvalContext = {
      currentSheet: 'Sheet1',
      getCell: (sheet, col, row) => {
        calls.push([sheet, col, row]);
        return 21;
      },
    };
    expect(evaluateFormula('Лист2!A1*2', ctx)).toBe(42);
    expect(calls).toEqual([['Лист2', 0, 1]]);
  });

  it('renameSheet rewrites a defined name with an unquoted non-ASCII prefix', () => {
    const wb = new Workbook();
    wb.addSheet('Лист1');
    wb.addNamedRange('Rev', 'Лист1!$A$1');
    wb.renameSheet('Лист1', 'Data');
    expect(wb.getNamedRange('Rev')?.value).toBe('Data!$A$1');
  });
});
