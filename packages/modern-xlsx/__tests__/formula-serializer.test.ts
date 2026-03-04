import { describe, expect, it } from 'vitest';
import type { ASTNode } from '../src/formula/parser.js';
import { parseFormula } from '../src/formula/parser.js';
import { serializeFormula } from '../src/formula/serializer.js';

/** Parse a formula, assert no errors, serialize back. */
function roundtrip(formula: string): string {
  const result = parseFormula(formula);
  expect(result.errors, `parse errors for "${formula}": ${result.errors.join(', ')}`).toHaveLength(
    0,
  );
  expect(result.ast).not.toBeNull();
  return serializeFormula(result.ast as ASTNode);
}

describe('serializeFormula', () => {
  // -----------------------------------------------------------------------
  // Roundtrip tests: parse → serialize → should equal original
  // -----------------------------------------------------------------------
  describe('roundtrip', () => {
    it('simple addition', () => {
      expect(roundtrip('1+2')).toBe('1+2');
    });

    it('function call with range', () => {
      expect(roundtrip('SUM(A1:A10)')).toBe('SUM(A1:A10)');
    });

    it('IF with comparison and strings', () => {
      expect(roundtrip('IF(A1>0,"yes","no")')).toBe('IF(A1>0,"yes","no")');
    });

    it('absolute cell reference', () => {
      expect(roundtrip('$A$1')).toBe('$A$1');
    });

    it('sheet-qualified reference', () => {
      expect(roundtrip('Sheet1!A1')).toBe('Sheet1!A1');
    });

    it('quoted sheet name with space', () => {
      expect(roundtrip("'My Sheet'!B2")).toBe("'My Sheet'!B2");
    });

    it('array constant', () => {
      expect(roundtrip('{1,2;3,4}')).toBe('{1,2;3,4}');
    });

    it('unary minus', () => {
      expect(roundtrip('-A1')).toBe('-A1');
    });

    it('percent', () => {
      expect(roundtrip('50%')).toBe('50%');
    });

    it('parenthesized multiplication', () => {
      expect(roundtrip('(1+2)*3')).toBe('(1+2)*3');
    });

    it('comparison operator >=', () => {
      expect(roundtrip('A1>=10')).toBe('A1>=10');
    });

    it('concatenation operator', () => {
      expect(roundtrip('A1&B1')).toBe('A1&B1');
    });
  });

  // -----------------------------------------------------------------------
  // Precedence preservation
  // -----------------------------------------------------------------------
  describe('precedence', () => {
    it('no unnecessary parens for higher-precedence child', () => {
      // 2*3 has higher precedence than +, so no parens needed around 2*3
      expect(roundtrip('1+2*3')).toBe('1+2*3');
    });

    it('adds parens for lower-precedence child', () => {
      // (1+2)*3 — + has lower precedence than *, so needs parens
      expect(roundtrip('(1+2)*3')).toBe('(1+2)*3');
    });

    it('preserves left-associativity of subtraction', () => {
      // a-(b-c) needs parens on the right for left-associative -
      // parse "A1-(B1-C1)" → subtract(A1, subtract(B1, C1))
      expect(roundtrip('A1-(B1-C1)')).toBe('A1-(B1-C1)');
    });

    it('no parens for left-to-right chain of same precedence', () => {
      // A1+B1+C1 parses as (A1+B1)+C1, which serializes without extra parens
      expect(roundtrip('A1+B1+C1')).toBe('A1+B1+C1');
    });

    it('handles nested comparison and arithmetic', () => {
      // comparison has lower precedence than +, so 1+2=3 means (1+2)=3
      expect(roundtrip('1+2=3')).toBe('1+2=3');
    });

    it('handles exponentiation right-associativity', () => {
      // 2^3^4 parses as 2^(3^4) — right-associative, no parens needed
      expect(roundtrip('2^3^4')).toBe('2^3^4');
    });

    it('concatenation vs addition precedence', () => {
      // & has lower precedence than +
      expect(roundtrip('A1+B1&C1')).toBe('A1+B1&C1');
    });
  });

  // -----------------------------------------------------------------------
  // Complex formulas
  // -----------------------------------------------------------------------
  describe('complex formulas', () => {
    it('nested IF with AND and SUM', () => {
      expect(roundtrip('IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")')).toBe(
        'IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")',
      );
    });

    it('multiple functions and operators', () => {
      expect(roundtrip('ROUND(A1*B1/100,2)')).toBe('ROUND(A1*B1/100,2)');
    });

    it('string with embedded quotes', () => {
      expect(roundtrip('IF(A1=1,"say ""hello""","bye")')).toBe('IF(A1=1,"say ""hello""","bye")');
    });

    it('error literal in formula', () => {
      expect(roundtrip('IFERROR(A1/B1,#N/A)')).toBe('IFERROR(A1/B1,#N/A)');
    });

    it('boolean literal', () => {
      expect(roundtrip('IF(TRUE,1,0)')).toBe('IF(TRUE,1,0)');
    });

    it('named range', () => {
      expect(roundtrip('SUM(myRange)')).toBe('SUM(myRange)');
    });

    it('mixed absolute and relative refs', () => {
      expect(roundtrip('$A1+A$1+$A$1+A1')).toBe('$A1+A$1+$A$1+A1');
    });
  });

  // -----------------------------------------------------------------------
  // Individual node types
  // -----------------------------------------------------------------------
  describe('node types', () => {
    it('serializes number node', () => {
      expect(serializeFormula({ type: 'number', value: 42 })).toBe('42');
    });

    it('serializes decimal number node', () => {
      expect(serializeFormula({ type: 'number', value: 3.14 })).toBe('3.14');
    });

    it('serializes string node', () => {
      expect(serializeFormula({ type: 'string', value: 'hello' })).toBe('"hello"');
    });

    it('serializes string with quotes', () => {
      expect(serializeFormula({ type: 'string', value: 'say "hi"' })).toBe('"say ""hi"""');
    });

    it('serializes boolean true', () => {
      expect(serializeFormula({ type: 'boolean', value: true })).toBe('TRUE');
    });

    it('serializes boolean false', () => {
      expect(serializeFormula({ type: 'boolean', value: false })).toBe('FALSE');
    });

    it('serializes error node', () => {
      expect(serializeFormula({ type: 'error', value: '#REF!' })).toBe('#REF!');
    });

    it('serializes name node', () => {
      expect(serializeFormula({ type: 'name', name: 'myRange' })).toBe('myRange');
    });

    it('serializes empty array', () => {
      expect(serializeFormula({ type: 'array', rows: [] })).toBe('{}');
    });
  });
});
