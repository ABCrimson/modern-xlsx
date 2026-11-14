import { describe, expect, it } from 'vitest';
import type { ASTNode, CellRefNode } from '../src/formula/parser.js';
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

    it('keeps a left-associative ^ chain bare', () => {
      // 2^3^4 parses as (2^3)^4, like Excel, so no parens are needed
      expect(roundtrip('2^3^4')).toBe('2^3^4');
      expect(roundtrip('(2^3)^4')).toBe('2^3^4');
    });

    it('parenthesises a right-nested ^', () => {
      expect(roundtrip('2^(3^4)')).toBe('2^(3^4)');
    });

    it('keeps unary minus under ^ bare: -2^2', () => {
      expect(roundtrip('-2^2')).toBe('-2^2');
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
  // Excel reference forms: exact round trip with no parse errors
  // -----------------------------------------------------------------------
  describe('reference-form corpus', () => {
    it.each([
      'SUM((A1,B2))',
      'SUM(Data!B:C)',
      'IF(A1,,2)',
      'SUM(A1,)',
      "'2024'!A1",
      'IF(A1,#SPILL!,1)',
      'SUM(A1:B5 B2:C3)',
      'SUM($A:$A)',
      'SUM(1:1)',
      'SUM($2:$5)',
      'SUM(Table1[[#Headers],[Col]])',
      'SUM([@Col])',
      'SUM(Sheet1:Sheet3!A1)',
      "SUM('Jan 1:Mar 1'!B2)",
      'SUM(Sheet1:Sheet3!A1:B2)',
      "'a''!b'!A1",
      'A1,B2',
      'SUM((A1,B2,C3))',
      'F()',
    ])('%s', (formula) => {
      expect(roundtrip(formula)).toBe(formula);
    });
  });

  // -----------------------------------------------------------------------
  // Reference operators
  // -----------------------------------------------------------------------
  describe('reference operators', () => {
    const A1: CellRefNode = { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: false };
    const B2: CellRefNode = { type: 'cell_ref', col: 'B', row: 2, absCol: false, absRow: false };
    const C3: CellRefNode = { type: 'cell_ref', col: 'C', row: 3, absCol: false, absRow: false };
    const union = (left: ASTNode, right: ASTNode): ASTNode => ({
      type: 'binary_op',
      op: ',',
      left,
      right,
    });

    it("binds a user-built ':' tighter than every other operator", () => {
      expect(
        serializeFormula({
          type: 'binary_op',
          op: '+',
          left: { type: 'binary_op', op: ':', left: A1, right: B2 },
          right: { type: 'number', value: 1 },
        }),
      ).toBe('A1:B2+1');
    });

    it('emits a union bare at the root and parenthesised once elsewhere', () => {
      expect(serializeFormula(union(union(A1, B2), C3))).toBe('A1,B2,C3');
      expect(
        serializeFormula({ type: 'function', name: 'SUM', args: [union(union(A1, B2), C3)] }),
      ).toBe('SUM((A1,B2,C3))');
      expect(
        serializeFormula({
          type: 'binary_op',
          op: ' ',
          left: union(A1, B2),
          right: C3,
        }),
      ).toBe('(A1,B2) C3');
      expect(serializeFormula({ type: 'unary_op', op: '-', operand: union(A1, B2) })).toBe(
        '-(A1,B2)',
      );
    });

    it('parenthesises a unary or percent operand of intersection', () => {
      expect(
        serializeFormula({
          type: 'binary_op',
          op: ' ',
          left: { type: 'unary_op', op: '-', operand: A1 },
          right: B2,
        }),
      ).toBe('(-A1) B2');
      expect(
        serializeFormula({
          type: 'binary_op',
          op: ' ',
          left: A1,
          right: { type: 'percent', operand: B2 },
        }),
      ).toBe('A1 (B2%)');
    });

    it('serializes a missing argument as empty text', () => {
      expect(serializeFormula({ type: 'missing' })).toBe('');
    });

    it('quotes a 3-D prefix as one unit when either name needs it', () => {
      expect(serializeFormula({ ...A1, sheet: 'Sheet 1', endSheet: 'Sheet3' })).toBe(
        "'Sheet 1:Sheet3'!A1",
      );
      expect(serializeFormula({ ...A1, sheet: 'Sheet1', endSheet: "O'Brien" })).toBe(
        "'Sheet1:O''Brien'!A1",
      );
      expect(serializeFormula({ ...A1, sheet: 'Sheet1', endSheet: 'Sheet3' })).toBe(
        'Sheet1:Sheet3!A1',
      );
      expect(
        serializeFormula({
          type: 'column_range',
          sheet: 'Q1',
          endSheet: 'Q4',
          startCol: 'A',
          endCol: 'B',
          absStart: true,
          absEnd: false,
        }),
      ).toBe("'Q1:Q4'!$A:B");
      expect(
        serializeFormula({
          type: 'row_range',
          sheet: 'Data',
          startRow: 2,
          endRow: 5,
          absStart: false,
          absEnd: true,
        }),
      ).toBe('Data!2:$5');
    });
  });

  // -----------------------------------------------------------------------
  // Sheet-name quoting
  // -----------------------------------------------------------------------
  describe('sheet-name quoting', () => {
    const ref = (sheet: string): CellRefNode => ({
      type: 'cell_ref',
      sheet,
      col: 'B',
      row: 2,
      absCol: false,
      absRow: false,
    });

    it.each([
      ['2024', "'2024'!B2"],
      ['A1', "'A1'!B2"],
      ['TRUE', "'TRUE'!B2"],
      ['false', "'false'!B2"],
      ['R2D2', "'R2D2'!B2"],
      ['C', "'C'!B2"],
      ['RC', "'RC'!B2"],
      ['R1C1', "'R1C1'!B2"],
      ['RC1x', "'RC1x'!B2"],
      ['Données', "'Données'!B2"],
      ["It's", "'It''s'!B2"],
      ['Sheet1', 'Sheet1!B2'],
      ['Data_2.v1', 'Data_2.v1!B2'],
    ])('%s -> %s', (sheet, expected) => {
      const text = serializeFormula(ref(sheet));
      expect(text).toBe(expected);
      const reparsed = parseFormula(text);
      expect(reparsed.errors).toEqual([]);
      expect(reparsed.ast).toMatchObject({ type: 'cell_ref', sheet });
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
