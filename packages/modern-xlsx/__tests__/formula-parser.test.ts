import { describe, expect, it } from 'vitest';
import { evaluateFormula } from '../src/formula/evaluator.js';
import type {
  ArrayNode,
  ASTNode,
  CellRefNode,
  ColumnRangeNode,
  FunctionCallNode,
  RowRangeNode,
} from '../src/formula/parser.js';
import { parseCellRefValue, parseFormula } from '../src/formula/parser.js';
import type { EvalContext } from '../src/formula/resolver.js';
import { rewriteFormula } from '../src/formula/rewriter.js';
import { serializeFormula } from '../src/formula/serializer.js';

/** Matches any element: pads positions a row does not assert, so array lengths stay pinned. */
const ANY = expect.anything();

/** One table row: the input text, the subset of the result it must match and the keys it must not have. */
interface ParseCase {
  readonly formula: string;
  readonly expected: object;
  readonly absent?: readonly string[];
}

/** Parse a formula and assert zero errors, returning the AST node. */
function p(formula: string): ASTNode {
  const result = parseFormula(formula);
  expect(result.errors, `parse errors for "${formula}": ${result.errors.join(', ')}`).toHaveLength(
    0,
  );
  expect(result.ast).not.toBeNull();
  return result.ast as ASTNode;
}

describe('parseFormula', () => {
  // -----------------------------------------------------------------------
  // Literals
  // -----------------------------------------------------------------------
  describe('literals', () => {
    it.each<ParseCase>([
      // parses integer
      { formula: '42', expected: { type: 'number', value: 42 } },
      // parses decimal
      { formula: '3.14', expected: { type: 'number', value: expect.closeTo(3.14) } },
      // parses leading dot decimal
      { formula: '.5', expected: { type: 'number', value: 0.5 } },
      // parses scientific notation
      { formula: '1.5E10', expected: { type: 'number', value: 1.5e10 } },
      // parses string literal
      { formula: '"hello"', expected: { type: 'string', value: 'hello' } },
      // parses empty string
      { formula: '""', expected: { type: 'string', value: '' } },
      { formula: 'TRUE', expected: { type: 'boolean', value: true } },
      { formula: 'FALSE', expected: { type: 'boolean', value: false } },
      // parses error #N/A
      { formula: '#N/A', expected: { type: 'error', value: '#N/A' } },
      // parses error #REF!
      { formula: '#REF!', expected: { type: 'error', value: '#REF!' } },
      // parses error #DIV/0!
      { formula: '#DIV/0!', expected: { type: 'error', value: '#DIV/0!' } },
      // parses error #VALUE!
      { formula: '#VALUE!', expected: { type: 'error', value: '#VALUE!' } },
      // parses error #NAME?
      { formula: '#NAME?', expected: { type: 'error', value: '#NAME?' } },
      // parses error #NUM!
      { formula: '#NUM!', expected: { type: 'error', value: '#NUM!' } },
      // parses error #NULL!
      { formula: '#NULL!', expected: { type: 'error', value: '#NULL!' } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Cell references
  // -----------------------------------------------------------------------
  describe('cell references', () => {
    it.each<ParseCase>([
      // parses simple cell ref A1
      {
        formula: 'A1',
        expected: { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: false },
        absent: ['sheet'],
      },
      // parses absolute column $A1
      {
        formula: '$A1',
        expected: { type: 'cell_ref', col: 'A', row: 1, absCol: true, absRow: false },
      },
      // parses absolute row A$1
      {
        formula: 'A$1',
        expected: { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: true },
      },
      // parses fully absolute $A$1
      {
        formula: '$A$1',
        expected: { type: 'cell_ref', col: 'A', row: 1, absCol: true, absRow: true },
      },
      // parses double-letter column AA100
      { formula: 'AA100', expected: { type: 'cell_ref', col: 'AA', row: 100 } },
      // parses triple-letter column XFD1048576
      { formula: 'XFD1048576', expected: { type: 'cell_ref', col: 'XFD', row: 1048576 } },
      // parses sheet-qualified ref Sheet1!A1
      { formula: 'Sheet1!A1', expected: { type: 'cell_ref', sheet: 'Sheet1', col: 'A', row: 1 } },
      // parses sheet-qualified absolute ref Sheet1!$A$1
      {
        formula: 'Sheet1!$A$1',
        expected: {
          type: 'cell_ref',
          sheet: 'Sheet1',
          col: 'A',
          row: 1,
          absCol: true,
          absRow: true,
        },
      },
      // parses quoted sheet ref
      {
        formula: "'My Sheet'!B2",
        expected: { type: 'cell_ref', sheet: 'My Sheet', col: 'B', row: 2 },
      },
      // parses quoted sheet with escaped quote
      {
        formula: "'It''s Sheet'!C3",
        expected: { type: 'cell_ref', sheet: "It's Sheet", col: 'C', row: 3 },
      },
      // parses quoted sheet with absolute ref
      {
        formula: "'My Sheet'!$B$2",
        expected: {
          type: 'cell_ref',
          sheet: 'My Sheet',
          col: 'B',
          row: 2,
          absCol: true,
          absRow: true,
        },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Ranges
  // -----------------------------------------------------------------------
  describe('ranges', () => {
    it.each<ParseCase>([
      {
        formula: 'A1:B2',
        expected: {
          type: 'range',
          start: { type: 'cell_ref', col: 'A', row: 1 },
          end: { col: 'B', row: 2 },
        },
      },
      // parses absolute range $A$1:$B$10
      {
        formula: '$A$1:$B$10',
        expected: {
          type: 'range',
          start: { absCol: true, absRow: true },
          end: { absCol: true, absRow: true, row: 10 },
        },
      },
      // parses mixed absolute range A$1:$B2
      {
        formula: 'A$1:$B2',
        expected: {
          type: 'range',
          start: { absCol: false, absRow: true },
          end: { absCol: true, absRow: false },
        },
      },
      // parses wide range A1:XFD1048576
      { formula: 'A1:XFD1048576', expected: { type: 'range', end: { col: 'XFD', row: 1048576 } } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Operator precedence
  // -----------------------------------------------------------------------
  describe('operator precedence', () => {
    it.each<ParseCase>([
      // parses 1+2*3 with correct precedence
      // 1 + (2 * 3)
      {
        formula: '1+2*3',
        expected: {
          type: 'binary_op',
          op: '+',
          left: { value: 1 },
          right: { type: 'binary_op', op: '*', left: { value: 2 }, right: { value: 3 } },
        },
      },
      // parses 2*3+1 with correct precedence
      // (2 * 3) + 1
      {
        formula: '2*3+1',
        expected: {
          type: 'binary_op',
          op: '+',
          left: { op: '*', left: { value: 2 }, right: { value: 3 } },
          right: { value: 1 },
        },
      },
      // parses 2^3^4 as left-associative, like Excel
      // (2 ^ 3) ^ 4
      {
        formula: '2^3^4',
        expected: {
          type: 'binary_op',
          op: '^',
          right: { value: 4 },
          left: { type: 'binary_op', op: '^', left: { value: 2 }, right: { value: 3 } },
        },
      },
      // parses concatenation A1&B1
      {
        formula: 'A1&B1',
        expected: { type: 'binary_op', op: '&', left: { col: 'A' }, right: { col: 'B' } },
      },
      // parses comparison A1>=10
      {
        formula: 'A1>=10',
        expected: { type: 'binary_op', op: '>=', left: { type: 'cell_ref' }, right: { value: 10 } },
      },
      // parses comparison A1<>0
      { formula: 'A1<>0', expected: { op: '<>' } },
      // parses comparison A1=B1
      { formula: 'A1=B1', expected: { op: '=' } },
      // parses comparison A1<B1
      { formula: 'A1<B1', expected: { op: '<' } },
      // parses comparison A1>B1
      { formula: 'A1>B1', expected: { op: '>' } },
      // parses comparison A1<=B1
      { formula: 'A1<=B1', expected: { op: '<=' } },
      // parses concatenation lower than additive: "a"&1+2
      // "a" & (1 + 2)
      { formula: '"a"&1+2', expected: { op: '&', left: { value: 'a' }, right: { op: '+' } } },
      // parses division 10/5
      { formula: '10/5', expected: { op: '/', left: { value: 10 }, right: { value: 5 } } },
      // parses subtraction 10-3
      { formula: '10-3', expected: { op: '-' } },
      // parses chained addition left-to-right: 1+2+3
      // (1+2)+3
      {
        formula: '1+2+3',
        expected: {
          op: '+',
          left: { op: '+', left: { value: 1 }, right: { value: 2 } },
          right: { value: 3 },
        },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Unary operators
  // -----------------------------------------------------------------------
  describe('unary operators', () => {
    it.each<ParseCase>([
      { formula: '-A1', expected: { type: 'unary_op', op: '-', operand: { col: 'A' } } },
      { formula: '+5', expected: { type: 'unary_op', op: '+', operand: { value: 5 } } },
      // parses double unary --1
      {
        formula: '--1',
        expected: {
          type: 'unary_op',
          op: '-',
          operand: { type: 'unary_op', op: '-', operand: { value: 1 } },
        },
      },
      {
        formula: '1+-2',
        expected: {
          op: '+',
          left: { value: 1 },
          right: { type: 'unary_op', op: '-', operand: { value: 2 } },
        },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Percent
  // -----------------------------------------------------------------------
  describe('percent', () => {
    it.each<ParseCase>([
      { formula: '50%', expected: { type: 'percent', operand: { value: 50 } } },
      { formula: 'A1%', expected: { type: 'percent', operand: { col: 'A' } } },
      // parses chained percent 50%%
      {
        formula: '50%%',
        expected: { type: 'percent', operand: { type: 'percent', operand: { value: 50 } } },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Parentheses
  // -----------------------------------------------------------------------
  describe('parentheses', () => {
    it.each<ParseCase>([
      {
        formula: '(1+2)*3',
        expected: {
          op: '*',
          left: { op: '+', left: { value: 1 }, right: { value: 2 } },
          right: { value: 3 },
        },
      },
      // parses nested parens ((1+2))
      { formula: '((1+2))', expected: { op: '+', left: { value: 1 }, right: { value: 2 } } },
      { formula: '(-1)', expected: { type: 'unary_op', op: '-', operand: { value: 1 } } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Function calls
  // -----------------------------------------------------------------------
  describe('function calls', () => {
    it.each<ParseCase>([
      {
        formula: 'SUM(A1:A10)',
        expected: {
          type: 'function',
          name: 'SUM',
          args: [{ type: 'range', start: { col: 'A', row: 1 }, end: { col: 'A', row: 10 } }],
        },
      },
      {
        formula: 'IF(A1>0,1,0)',
        expected: { type: 'function', name: 'IF', args: [{ op: '>' }, { value: 1 }, { value: 0 }] },
      },
      // parses no-arg function NOW()
      { formula: 'NOW()', expected: { type: 'function', name: 'NOW', args: [] } },
      // parses no-arg function TODAY()
      { formula: 'TODAY()', expected: { name: 'TODAY', args: [] } },
      // parses single-arg function ABS(-5)
      { formula: 'ABS(-5)', expected: { name: 'ABS', args: [{ type: 'unary_op', op: '-' }] } },
      // parses VLOOKUP with 4 args
      {
        formula: 'VLOOKUP(A1,B1:D10,3,FALSE)',
        expected: {
          name: 'VLOOKUP',
          args: [{ col: 'A' }, { type: 'range' }, { value: 3 }, { value: false }],
        },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Nested functions
  // -----------------------------------------------------------------------
  describe('nested functions', () => {
    it.each<ParseCase>([
      {
        formula: 'IF(SUM(A1:A5)>10,TRUE,FALSE)',
        expected: {
          name: 'IF',
          args: [
            { op: '>', left: { name: 'SUM', args: [{ type: 'range' }] }, right: { value: 10 } },
            { value: true },
            { value: false },
          ],
        },
      },
      {
        formula: 'IFERROR(VLOOKUP(A1,B1:D10,2,0),"")',
        expected: {
          name: 'IFERROR',
          args: [
            {
              name: 'VLOOKUP',
              args: [ANY, ANY, ANY, ANY],
            },
            { value: '' },
          ],
        },
      },
      {
        formula: 'IF(IF(IF(1,2,3),4,5),6,7)',
        expected: {
          name: 'IF',
          args: [
            { name: 'IF', args: [{ name: 'IF', args: [{ value: 1 }, ANY, ANY] }, ANY, ANY] },
            ANY,
            ANY,
          ],
        },
      },
    ])('parses $formula', ({ formula, expected }) => {
      expect(p(formula)).toMatchObject(expected);
    });
  });

  // -----------------------------------------------------------------------
  // Array constants
  // -----------------------------------------------------------------------
  describe('array constants', () => {
    it.each<ParseCase>([
      {
        formula: '{1,2;3,4}',
        expected: {
          type: 'array',
          rows: [
            [{ value: 1 }, { value: 2 }],
            [{ value: 3 }, { value: 4 }],
          ],
        },
      },
      // parses single-row array {1,2,3}
      {
        formula: '{1,2,3}',
        expected: { rows: [[ANY, ANY, ANY]] },
      },
      // parses single-element array {42}
      { formula: '{42}', expected: { rows: [[{ value: 42 }]] } },
      // parses empty array {}
      { formula: '{}', expected: { type: 'array', rows: [] } },
      {
        formula: '{"a","b";"c","d"}',
        expected: {
          rows: [
            [{ value: 'a' }, { value: 'b' }],
            [{ value: 'c' }, { value: 'd' }],
          ],
        },
      },
      { formula: '{TRUE,FALSE}', expected: { rows: [[{ value: true }, { value: false }]] } },
      { formula: '{#N/A,#REF!}', expected: { rows: [[{ value: '#N/A' }, { value: '#REF!' }]] } },
      {
        formula: '{-1,2}',
        expected: { rows: [[{ type: 'unary_op', op: '-', operand: { value: 1 } }, ANY]] },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Named ranges
  // -----------------------------------------------------------------------
  describe('named ranges', () => {
    it.each<ParseCase>([
      // parses MyRange as NameNode
      { formula: 'MyRange', expected: { type: 'name', name: 'MyRange' } },
      // parses underscore-prefixed name
      { formula: '_custom', expected: { type: 'name', name: '_custom' } },
      // parses name with dots
      { formula: 'Data.Total', expected: { type: 'name', name: 'Data.Total' } },
      // parses name in expression: MyRange+1
      { formula: 'MyRange+1', expected: { op: '+', left: { type: 'name', name: 'MyRange' } } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Complex formulas
  // -----------------------------------------------------------------------
  describe('complex formulas', () => {
    it.each<ParseCase>([
      // First arg: AND(A1>0,B1<10)
      // Second arg: SUM(C1:C10)*1.1
      // Third arg: "N/A"
      {
        formula: 'IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")',
        expected: {
          name: 'IF',
          args: [
            { name: 'AND', args: [ANY, ANY] },
            { op: '*', left: { name: 'SUM' }, right: { value: expect.closeTo(1.1) } },
            { value: 'N/A' },
          ],
        },
      },
      // parses formula with mixed references and operators
      // $A$1+Sheet1!B2*100%
      {
        formula: '$A$1+Sheet1!B2*100%',
        expected: {
          op: '+',
          left: { absCol: true, absRow: true },
          right: {
            op: '*',
            left: { sheet: 'Sheet1' },
            right: { type: 'percent', operand: { value: 100 } },
          },
        },
      },
      // parses formula with leading =
      { formula: '=1+2', expected: { op: '+', left: { value: 1 }, right: { value: 2 } } },
      {
        formula: 'SUMPRODUCT((A1:A10>0)*(B1:B10))',
        expected: { name: 'SUMPRODUCT', args: [{ op: '*' }] },
      },
      // parses INDEX(MATCH()) pattern
      {
        formula: 'INDEX(A1:A10,MATCH(B1,C1:C10,0))',
        expected: {
          name: 'INDEX',
          args: [ANY, { name: 'MATCH', args: [ANY, ANY, ANY] }],
        },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Error handling
  // -----------------------------------------------------------------------
  describe('error handling', () => {
    it('returns null ast for empty formula', () => {
      const result = parseFormula('');
      expect(result.ast).toBeNull();
      expect(result.errors).toHaveLength(0);
    });

    it('returns null ast for just =', () => {
      const result = parseFormula('=');
      expect(result.ast).toBeNull();
      expect(result.errors).toHaveLength(0);
    });

    it('reports mismatched paren: (1+2', () => {
      const result = parseFormula('(1+2');
      expect(result.errors.length).toBeGreaterThan(0);
      // Should still produce a partial AST
      expect(result.ast).not.toBeNull();
    });

    it('reports extra close paren: 1+2)', () => {
      const result = parseFormula('1+2)');
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.ast).not.toBeNull();
    });

    it('reports unexpected token', () => {
      const result = parseFormula('1+*2');
      // The * after + is unexpected (no operand between)
      expect(result.errors.length).toBeGreaterThan(0);
    });

    it('propagates tokenizer errors', () => {
      const result = parseFormula('"unterminated');
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.errors.some((e) => e.includes('Unterminated'))).toBe(true);
    });

    it('recovers from missing function close paren', () => {
      const result = parseFormula('SUM(1,2');
      expect(result.errors.length).toBeGreaterThan(0);
      expect(result.ast).not.toBeNull();
      const fn = result.ast as FunctionCallNode;
      expect(fn.type).toBe('function');
      expect(fn.name).toBe('SUM');
    });
  });

  // -----------------------------------------------------------------------
  // parseCellRefValue standalone tests
  // -----------------------------------------------------------------------
  describe('parseCellRefValue', () => {
    it.each<ParseCase>([
      {
        formula: 'A1',
        expected: { col: 'A', row: 1, absCol: false, absRow: false },
        absent: ['sheet'],
      },
      { formula: '$A$1', expected: { col: 'A', row: 1, absCol: true, absRow: true } },
      { formula: 'Sheet1!B2', expected: { sheet: 'Sheet1', col: 'B', row: 2 } },
      {
        formula: "'My Sheet'!$C$3",
        expected: { sheet: 'My Sheet', col: 'C', row: 3, absCol: true, absRow: true },
      },
      { formula: "'It''s Sheet'!D4", expected: { sheet: "It's Sheet", col: 'D', row: 4 } },
      { formula: 'AA100', expected: { col: 'AA', row: 100 } },
      {
        formula: '$XFD$1048576',
        expected: { col: 'XFD', row: 1048576, absCol: true, absRow: true },
      },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = parseCellRefValue(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });

    it("parses a 3-D prefix: Sheet1:Sheet3!A1 and 'Q 1:Q''4'!B2", () => {
      expect(parseCellRefValue('Sheet1:Sheet3!A1')).toMatchObject({
        sheet: 'Sheet1',
        endSheet: 'Sheet3',
        col: 'A',
        row: 1,
      });
      expect(parseCellRefValue("'Q 1:Q''4'!B2")).toMatchObject({
        sheet: 'Q 1',
        endSheet: "Q'4",
        col: 'B',
        row: 2,
      });
    });
  });

  // -----------------------------------------------------------------------
  // Semicolons: array-constant row separators only (stored formulas are en-US)
  // -----------------------------------------------------------------------
  describe('semicolon argument separators', () => {
    const ctx: EvalContext = { currentSheet: 'Sheet1', getCell: () => null };

    it('rejects ; as a function-argument separator', () => {
      expect(parseFormula('SUM(1;2)').errors.length).toBeGreaterThan(0);
      expect(evaluateFormula('SUM(1;2)', ctx)).toBe('#VALUE!');
    });

    it('keeps ; as the array-constant row separator', () => {
      const result = parseFormula('{1,2;3,4}');
      expect(result.errors).toEqual([]);
      expect((result.ast as ArrayNode).rows).toHaveLength(2);
    });
  });

  // -----------------------------------------------------------------------
  // Nesting limits: the parser never throws
  // -----------------------------------------------------------------------
  describe('nesting limits', () => {
    const nestedAbs = (n: number): string => `${'ABS('.repeat(n)}1${')'.repeat(n)}`;

    it('reports 100k nested parentheses without throwing', () => {
      const formula = `${'('.repeat(100_000)}1${')'.repeat(100_000)}`;
      expect(() => parseFormula(formula)).not.toThrow();
      expect(parseFormula(formula).errors.length).toBeGreaterThan(0);
    });

    it('accepts 64 nested functions (Excel limit)', () => {
      expect(parseFormula(nestedAbs(64)).errors).toEqual([]);
    });

    it('rejects 65 nested functions', () => {
      expect(parseFormula(nestedAbs(65)).errors.length).toBeGreaterThan(0);
    });

    it('puts #VALUE! in place of the 65th nested call and reports one error', () => {
      const { ast, errors } = parseFormula(nestedAbs(65));
      let expected: ASTNode = { type: 'error', value: '#VALUE!' };
      for (let i = 0; i < 64; i++) expected = { type: 'function', name: 'ABS', args: [expected] };
      expect(ast).toEqual(expected);
      expect(errors).toEqual(['More than 64 nested function levels']);
    });

    it('replaces the abandoned subtree with #VALUE!', () => {
      const { ast } = parseFormula(`-${'('.repeat(300)}1${')'.repeat(300)}`);
      expect(ast).toEqual({
        type: 'unary_op',
        op: '-',
        operand: { type: 'error', value: '#VALUE!' },
      });
    });
  });

  // -----------------------------------------------------------------------
  // Missing arguments
  // -----------------------------------------------------------------------
  describe('missing arguments', () => {
    it('parses IF(A1,,2) with a missing middle argument', () => {
      const fn = p('IF(A1,,2)');
      expect(fn).toMatchObject({ args: [ANY, ANY, ANY] });
      expect(fn).toHaveProperty(['args', 1], { type: 'missing' });
    });

    it('parses a trailing empty argument: SUM(A1,)', () => {
      const fn = p('SUM(A1,)');
      expect(fn).toMatchObject({ args: [ANY, ANY] });
      expect(fn).toHaveProperty(['args', 1], { type: 'missing' });
    });

    it('parses a leading empty argument: F(,1)', () => {
      expect(p('F(,1)')).toHaveProperty('args', [
        { type: 'missing' },
        { type: 'number', value: 1 },
      ]);
    });

    it('parses F(,) as two missing arguments', () => {
      expect(p('F(,)')).toHaveProperty('args', [{ type: 'missing' }, { type: 'missing' }]);
    });

    it('keeps F() at zero arguments', () => {
      expect(p('F()')).toHaveProperty('args', []);
    });
  });

  // -----------------------------------------------------------------------
  // Sheet prefixes and whole-row/column, structured and 3-D references
  // -----------------------------------------------------------------------
  describe('reference forms', () => {
    it("splits a quoted sheet name at the real closing quote: 'a''!b'!A1", () => {
      expect(parseFormula("'a''!b'!A1").ast).toEqual({
        type: 'cell_ref',
        sheet: "a'!b",
        col: 'A',
        row: 1,
        absCol: false,
        absRow: false,
      } satisfies CellRefNode);
    });

    it('parses a sheet-qualified column range: SUM(Data!B:C)', () => {
      const result = parseFormula('SUM(Data!B:C)');
      expect(result.errors).toEqual([]);
      expect((result.ast as FunctionCallNode).args[0]).toEqual({
        type: 'column_range',
        sheet: 'Data',
        startCol: 'B',
        endCol: 'C',
        absStart: false,
        absEnd: false,
      } satisfies ColumnRangeNode);
    });

    it('parses an absolute column range: $A:$A', () => {
      expect(p('$A:$A')).toEqual({
        type: 'column_range',
        startCol: 'A',
        endCol: 'A',
        absStart: true,
        absEnd: true,
      });
    });

    it('parses row ranges: 1:1 and Sheet1!$2:5', () => {
      expect(p('1:1')).toEqual({
        type: 'row_range',
        startRow: 1,
        endRow: 1,
        absStart: false,
        absEnd: false,
      } satisfies RowRangeNode);
      expect(p('Sheet1!$2:5')).toEqual({
        type: 'row_range',
        sheet: 'Sheet1',
        startRow: 2,
        endRow: 5,
        absStart: true,
        absEnd: false,
      });
    });

    it('parses structured references verbatim', () => {
      expect(p('SUM(Table1[[#Headers],[Col]])')).toHaveProperty(['args', 0], {
        type: 'structured_ref',
        text: 'Table1[[#Headers],[Col]]',
      });
      expect(p('[@Col]')).toEqual({ type: 'structured_ref', text: '[@Col]' });
    });

    it('puts a 3-D prefix on the start endpoint only', () => {
      expect(p('Sheet1:Sheet3!A1')).toEqual({
        type: 'cell_ref',
        sheet: 'Sheet1',
        endSheet: 'Sheet3',
        col: 'A',
        row: 1,
        absCol: false,
        absRow: false,
      });
      expect(p("'Jan 1:Mar 1'!B2")).toMatchObject({ sheet: 'Jan 1', endSheet: 'Mar 1' });
      const range = p('Sheet1:Sheet3!A1:B2');
      expect(range).toMatchObject({ start: { sheet: 'Sheet1', endSheet: 'Sheet3', col: 'A' } });
      expect(range).toHaveProperty('end', {
        type: 'cell_ref',
        col: 'B',
        row: 2,
        absCol: false,
        absRow: false,
      });
      expect(p('Sheet1:Sheet3!A:B')).toMatchObject({
        type: 'column_range',
        sheet: 'Sheet1',
        endSheet: 'Sheet3',
      });
    });

    it('parses a sheet-qualified error literal as the error alone', () => {
      expect(parseFormula('Sheet1!#REF!').ast).toEqual({ type: 'error', value: '#REF!' });
    });

    it('parses an out-of-grid cell-shaped identifier as a name', () => {
      expect(parseFormula('XFE1').ast?.type).toBe('name');
      expect(parseFormula('A0')).toEqual({ ast: { type: 'name', name: 'A0' }, errors: [] });
    });

    it('keeps the range operator when an endpoint is a name', () => {
      const a1: CellRefNode = { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: false };
      expect(parseFormula('A1:XFE1')).toEqual({
        ast: { type: 'binary_op', op: ':', left: a1, right: { type: 'name', name: 'XFE1' } },
        errors: [],
      });
      expect(parseFormula('XFE1:XFE2')).toEqual({
        ast: {
          type: 'binary_op',
          op: ':',
          left: { type: 'name', name: 'XFE1' },
          right: { type: 'name', name: 'XFE2' },
        },
        errors: [],
      });
      expect(parseFormula('Total:A1')).toEqual({
        ast: { type: 'binary_op', op: ':', left: { type: 'name', name: 'Total' }, right: a1 },
        errors: [],
      });
      expect(parseFormula('Total:1').errors.length).toBeGreaterThan(0);
    });

    it.each(['A1:XFE1', 'SUM(XFE1:XFE2)', 'Total:A1+1'])(
      'serializes and rewrites %s with the name endpoint intact',
      (formula) => {
        const { ast } = parseFormula(formula);
        expect(ast && serializeFormula(ast)).toBe(formula);
        expect(rewriteFormula(formula, { type: 'insert_rows', start: 0, count: 1 })).toBe(
          formula.replace('A1', 'A2'),
        );
      },
    );

    it('parses modern error literals', () => {
      expect(p('IF(A1,#SPILL!,1)')).toHaveProperty(['args', 1], {
        type: 'error',
        value: '#SPILL!',
      });
    });
  });

  // -----------------------------------------------------------------------
  // Reference operators: ':' > ' ' (intersection) > ',' (union)
  // -----------------------------------------------------------------------
  describe('reference operators', () => {
    it('parses intersection as binary_op " "', () => {
      expect(p('SUM(A1:B5 B2:C3)')).toMatchObject({
        args: [{ op: ' ', left: { type: 'range' }, right: { type: 'range' } }],
      });
    });

    it('binds intersection tighter than unary minus and percent', () => {
      expect(p('-A1 B1')).toMatchObject({ type: 'unary_op', operand: { op: ' ' } });
      expect(p('A1 B1%')).toMatchObject({ type: 'percent', operand: { op: ' ' } });
    });

    it.each<ParseCase>([
      // parses a parenthesised union inside a function argument
      {
        formula: 'SUM((A1,B2))',
        expected: { args: [{ op: ',', left: { col: 'A' }, right: { col: 'B' } }] },
      },
      // folds a top-level union left to right
      { formula: 'A1,B2,C3', expected: { op: ',', left: { op: ',' }, right: { col: 'C' } } },
      // binds intersection tighter than union
      { formula: '(A1:B2 B1,C3)', expected: { op: ',', left: { op: ' ' } } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });

    it('never reads function-argument commas as union', () => {
      expect(p('SUM(A1,B2)')).toMatchObject({ args: [ANY, ANY] });
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------
  describe('edge cases', () => {
    it('handles formula with all operator types', () => {
      // This is a complex formula just to ensure no crashes
      const result = parseFormula('=1+2-3*4/5^6&"x"');
      expect(result.errors).toHaveLength(0);
      expect(result.ast).not.toBeNull();
    });

    it.each<ParseCase>([
      // handles whitespace in formulas
      { formula: '1 + 2', expected: { op: '+', left: { value: 1 }, right: { value: 2 } } },
      // handles percent in arithmetic: 50%+1
      { formula: '50%+1', expected: { op: '+', left: { type: 'percent' }, right: { value: 1 } } },
      // handles unary minus with exponentiation: -2^3
      // Unary has higher precedence than exponentiation in our parser
      // so -2^3 parses as (-2)^3
      { formula: '-2^3', expected: { op: '^', left: { type: 'unary_op', op: '-' } } },
    ])('parses $formula', ({ formula, expected, absent }) => {
      const result = p(formula);
      expect(result).toMatchObject(expected);
      for (const key of absent ?? []) expect(result).not.toHaveProperty(key);
    });
  });

  // -----------------------------------------------------------------------
  // Range operator with any reference-valued endpoint
  // -----------------------------------------------------------------------
  describe('range operator endpoints', () => {
    const cell = (col: string, row: number): CellRefNode => ({
      type: 'cell_ref',
      col,
      row,
      absCol: false,
      absRow: false,
    });

    it('parses A1:B1:C1 as a left-folded chain and writes it back', () => {
      const result = parseFormula('A1:B1:C1');
      expect(result.errors).toEqual([]);
      expect(result.ast).toEqual({
        type: 'binary_op',
        op: ':',
        left: { type: 'range', start: cell('A', 1), end: cell('B', 1) },
        right: cell('C', 1),
      });
      expect(serializeFormula(result.ast as ASTNode)).toBe('A1:B1:C1');
    });

    it('parses a function endpoint on either side', () => {
      const right = p('SUM(A2:INDEX(A:A,B1))');
      expect(right).toMatchObject({
        type: 'function',
        args: [{ type: 'binary_op', op: ':', left: cell('A', 2), right: { type: 'function' } }],
      });
      expect(serializeFormula(right)).toBe('SUM(A2:INDEX(A:A,B1))');
      const left = p('INDEX(A:A,1):A3');
      expect(left).toMatchObject({ type: 'binary_op', op: ':', right: cell('A', 3) });
      expect(serializeFormula(left)).toBe('INDEX(A:A,1):A3');
      expect(serializeFormula(p('OFFSET(A1,1,1):INDIRECT("B5")'))).toBe(
        'OFFSET(A1,1,1):INDIRECT("B5")',
      );
    });

    it('keeps today’s nodes for two cell endpoints and for name endpoints', () => {
      expect(p('A1:B2')).toEqual({ type: 'range', start: cell('A', 1), end: cell('B', 2) });
      expect(p('(A1):B2')).toEqual({ type: 'range', start: cell('A', 1), end: cell('B', 2) });
      expect(p('Total:A1')).toEqual({
        type: 'binary_op',
        op: ':',
        left: { type: 'name', name: 'Total' },
        right: cell('A', 1),
      });
    });

    it('evaluates a range with a function endpoint to #VALUE! (not resolved yet)', () => {
      const ctx: EvalContext = { currentSheet: 'Sheet1', getCell: () => 1 };
      expect(evaluateFormula('A1:INDEX(A:A,1)', ctx)).toBe('#VALUE!');
      expect(evaluateFormula('Total:A1:B1', ctx)).toBe('#NAME?');
    });

    it('still reports a colon with no reference endpoint', () => {
      expect(parseFormula('A1:1').errors).toContain('Expected cell reference after ":"');
      expect(parseFormula('A1:"x"').errors.length).toBeGreaterThan(0);
    });
  });
});

describe('#REF! range endpoints', () => {
  it('parses #REF! as an endpoint of the range operator', () => {
    expect(parseFormula('SUM(#REF!:C2)').errors).toEqual([]);
    expect(parseFormula('SUM(A1:#REF!)').errors).toEqual([]);
  });

  it('still rejects other error literals as endpoints', () => {
    expect(parseFormula('SUM(#N/A:C2)').errors.length).toBeGreaterThan(0);
  });
});
