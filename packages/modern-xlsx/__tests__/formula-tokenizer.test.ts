import { describe, expect, it } from 'vitest';
import type { TokenType } from '../src/formula/tokenizer.js';
import { tokenize } from '../src/formula/tokenizer.js';

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
      expect(result.tokens[0].value).toBe('1');
    });

    it('works without leading =', () => {
      const result = tokenize('1+2');
      expect(result.tokens).toHaveLength(3);
      expect(result.tokens[0].value).toBe('1');
    });
  });

  describe('simple arithmetic', () => {
    it('tokenizes 1+2', () => {
      expect(tv('1+2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['number', '2'],
      ]);
    });

    it('tokenizes 3.14*2', () => {
      expect(tv('3.14*2')).toEqual([
        ['number', '3.14'],
        ['operator', '*'],
        ['number', '2'],
      ]);
    });

    it('tokenizes 10/3-1', () => {
      expect(tv('10/3-1')).toEqual([
        ['number', '10'],
        ['operator', '/'],
        ['number', '3'],
        ['operator', '-'],
        ['number', '1'],
      ]);
    });

    it('tokenizes power operator', () => {
      expect(tv('2^8')).toEqual([
        ['number', '2'],
        ['operator', '^'],
        ['number', '8'],
      ]);
    });

    it('tokenizes concatenation operator', () => {
      expect(tv('"a"&"b"')).toEqual([
        ['string', 'a'],
        ['operator', '&'],
        ['string', 'b'],
      ]);
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
    it('tokenizes A1:B2 as cell_ref + colon + cell_ref', () => {
      expect(tv('A1:B2')).toEqual([
        ['cell_ref', 'A1'],
        ['colon', ':'],
        ['cell_ref', 'B2'],
      ]);
    });

    it('tokenizes absolute range', () => {
      expect(tv('$A$1:$B$10')).toEqual([
        ['cell_ref', '$A$1'],
        ['colon', ':'],
        ['cell_ref', '$B$10'],
      ]);
    });
  });

  describe('function calls', () => {
    it('tokenizes SUM(A1:A10)', () => {
      expect(tv('SUM(A1:A10)')).toEqual([
        ['function', 'SUM'],
        ['paren_open', '('],
        ['cell_ref', 'A1'],
        ['colon', ':'],
        ['cell_ref', 'A10'],
        ['paren_close', ')'],
      ]);
    });

    it('tokenizes IF(A1>0,1,0)', () => {
      expect(tv('IF(A1>0,1,0)')).toEqual([
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
      ]);
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
  });

  describe('comparison operators', () => {
    it('tokenizes >=', () => {
      expect(tv('A1>=10')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '>='],
        ['number', '10'],
      ]);
    });

    it('tokenizes <=', () => {
      expect(tv('A1<=10')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '<='],
        ['number', '10'],
      ]);
    });

    it('tokenizes <>', () => {
      expect(tv('A1<>0')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '<>'],
        ['number', '0'],
      ]);
    });

    it('tokenizes = as operator', () => {
      expect(tv('A1=B1')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '='],
        ['cell_ref', 'B1'],
      ]);
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
    it('tokenizes unary minus at start', () => {
      expect(tv('-A1')).toEqual([
        ['prefix_op', '-'],
        ['cell_ref', 'A1'],
      ]);
    });

    it('tokenizes unary plus at start', () => {
      expect(tv('+5')).toEqual([
        ['prefix_op', '+'],
        ['number', '5'],
      ]);
    });

    it('tokenizes unary after operator', () => {
      expect(tv('1+-2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['prefix_op', '-'],
        ['number', '2'],
      ]);
    });

    it('tokenizes unary after open paren', () => {
      expect(tv('(-1)')).toEqual([
        ['paren_open', '('],
        ['prefix_op', '-'],
        ['number', '1'],
        ['paren_close', ')'],
      ]);
    });

    it('tokenizes unary after comma', () => {
      expect(tv('SUM(-1,-2)')).toEqual([
        ['function', 'SUM'],
        ['paren_open', '('],
        ['prefix_op', '-'],
        ['number', '1'],
        ['comma', ','],
        ['prefix_op', '-'],
        ['number', '2'],
        ['paren_close', ')'],
      ]);
    });

    it('tokenizes binary minus after cell ref', () => {
      expect(tv('A1-B1')).toEqual([
        ['cell_ref', 'A1'],
        ['operator', '-'],
        ['cell_ref', 'B1'],
      ]);
    });

    it('tokenizes binary plus after number', () => {
      expect(tv('1+2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['number', '2'],
      ]);
    });

    it('tokenizes binary minus after close paren', () => {
      expect(tv('(1)-2')).toEqual([
        ['paren_open', '('],
        ['number', '1'],
        ['paren_close', ')'],
        ['operator', '-'],
        ['number', '2'],
      ]);
    });
  });

  describe('percent', () => {
    it('tokenizes 50%', () => {
      expect(tv('50%')).toEqual([
        ['number', '50'],
        ['percent', '%'],
      ]);
    });

    it('tokenizes percent after cell ref', () => {
      expect(tv('A1%')).toEqual([
        ['cell_ref', 'A1'],
        ['percent', '%'],
      ]);
    });
  });

  describe('array constants', () => {
    it('tokenizes {1,2;3,4}', () => {
      expect(tv('{1,2;3,4}')).toEqual([
        ['array_open', '{'],
        ['number', '1'],
        ['array_col_sep', ','],
        ['number', '2'],
        ['array_row_sep', ';'],
        ['number', '3'],
        ['array_col_sep', ','],
        ['number', '4'],
        ['array_close', '}'],
      ]);
    });

    it('uses comma/semicolon outside arrays', () => {
      expect(tv('SUM(1,2)')).toEqual([
        ['function', 'SUM'],
        ['paren_open', '('],
        ['number', '1'],
        ['comma', ','],
        ['number', '2'],
        ['paren_close', ')'],
      ]);
    });

    it('tokenizes array with strings', () => {
      expect(tv('{"a","b";"c","d"}')).toEqual([
        ['array_open', '{'],
        ['string', 'a'],
        ['array_col_sep', ','],
        ['string', 'b'],
        ['array_row_sep', ';'],
        ['string', 'c'],
        ['array_col_sep', ','],
        ['string', 'd'],
        ['array_close', '}'],
      ]);
    });

    it('tokenizes unary minus inside array', () => {
      expect(tv('{-1,2}')).toEqual([
        ['array_open', '{'],
        ['prefix_op', '-'],
        ['number', '1'],
        ['array_col_sep', ','],
        ['number', '2'],
        ['array_close', '}'],
      ]);
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

    it('tokenizes IFERROR(VLOOKUP(A1,Data!B:C,2,0),"")', () => {
      const result = tv('IFERROR(VLOOKUP(A1,Data!B:C,2,0),"")');
      expect(result[0]).toEqual(['function', 'IFERROR']);
      expect(result[4]).toEqual(['cell_ref', 'A1']);
      // Data!B is a sheet-qualified ref
      expect(result[6][0]).toBe('cell_ref');
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
    it('skips spaces between tokens', () => {
      expect(tv('1 + 2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['number', '2'],
      ]);
    });

    it('skips tabs and newlines', () => {
      expect(tv('1\t+\n2')).toEqual([
        ['number', '1'],
        ['operator', '+'],
        ['number', '2'],
      ]);
    });
  });

  describe('token positions', () => {
    it('tracks start and end positions', () => {
      const { tokens } = tokenize('=A1+B2');
      // After stripping `=`, A1 is at positions 1..3, + at 3..4, B2 at 4..6
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
    it('emits semicolon type outside arrays', () => {
      // European-style argument separator
      expect(tv('SUM(1;2)')).toEqual([
        ['function', 'SUM'],
        ['paren_open', '('],
        ['number', '1'],
        ['semicolon', ';'],
        ['number', '2'],
        ['paren_close', ')'],
      ]);
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

    it('handles R1C1-like names as cell_ref or name', () => {
      // R1C1 looks like a cell_ref (matches pattern)
      const result = tv('R1');
      expect(result[0][0]).toBe('cell_ref');
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
