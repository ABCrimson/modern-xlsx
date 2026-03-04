import { describe, expect, it } from 'vitest';
import type {
  ArrayNode,
  ASTNode,
  BinaryOpNode,
  BooleanNode,
  CellRefNode,
  ErrorNode,
  FunctionCallNode,
  NameNode,
  NumberNode,
  PercentNode,
  RangeNode,
  StringNode,
  UnaryOpNode,
} from '../src/formula/parser.js';
import { parseCellRefValue, parseFormula } from '../src/formula/parser.js';

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
    it('parses integer', () => {
      const node = p('42') as NumberNode;
      expect(node.type).toBe('number');
      expect(node.value).toBe(42);
    });

    it('parses decimal', () => {
      const node = p('3.14') as NumberNode;
      expect(node.type).toBe('number');
      expect(node.value).toBeCloseTo(3.14);
    });

    it('parses leading dot decimal', () => {
      const node = p('.5') as NumberNode;
      expect(node.type).toBe('number');
      expect(node.value).toBe(0.5);
    });

    it('parses scientific notation', () => {
      const node = p('1.5E10') as NumberNode;
      expect(node.type).toBe('number');
      expect(node.value).toBe(1.5e10);
    });

    it('parses string literal', () => {
      const node = p('"hello"') as StringNode;
      expect(node.type).toBe('string');
      expect(node.value).toBe('hello');
    });

    it('parses empty string', () => {
      const node = p('""') as StringNode;
      expect(node.type).toBe('string');
      expect(node.value).toBe('');
    });

    it('parses TRUE', () => {
      const node = p('TRUE') as BooleanNode;
      expect(node.type).toBe('boolean');
      expect(node.value).toBe(true);
    });

    it('parses FALSE', () => {
      const node = p('FALSE') as BooleanNode;
      expect(node.type).toBe('boolean');
      expect(node.value).toBe(false);
    });

    it('parses error #N/A', () => {
      const node = p('#N/A') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#N/A');
    });

    it('parses error #REF!', () => {
      const node = p('#REF!') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#REF!');
    });

    it('parses error #DIV/0!', () => {
      const node = p('#DIV/0!') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#DIV/0!');
    });

    it('parses error #VALUE!', () => {
      const node = p('#VALUE!') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#VALUE!');
    });

    it('parses error #NAME?', () => {
      const node = p('#NAME?') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#NAME?');
    });

    it('parses error #NUM!', () => {
      const node = p('#NUM!') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#NUM!');
    });

    it('parses error #NULL!', () => {
      const node = p('#NULL!') as ErrorNode;
      expect(node.type).toBe('error');
      expect(node.value).toBe('#NULL!');
    });
  });

  // -----------------------------------------------------------------------
  // Cell references
  // -----------------------------------------------------------------------
  describe('cell references', () => {
    it('parses simple cell ref A1', () => {
      const node = p('A1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(false);
      expect(node.absRow).toBe(false);
      expect(node.sheet).toBeUndefined();
    });

    it('parses absolute column $A1', () => {
      const node = p('$A1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(true);
      expect(node.absRow).toBe(false);
    });

    it('parses absolute row A$1', () => {
      const node = p('A$1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(false);
      expect(node.absRow).toBe(true);
    });

    it('parses fully absolute $A$1', () => {
      const node = p('$A$1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(true);
      expect(node.absRow).toBe(true);
    });

    it('parses double-letter column AA100', () => {
      const node = p('AA100') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('AA');
      expect(node.row).toBe(100);
    });

    it('parses triple-letter column XFD1048576', () => {
      const node = p('XFD1048576') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.col).toBe('XFD');
      expect(node.row).toBe(1048576);
    });

    it('parses sheet-qualified ref Sheet1!A1', () => {
      const node = p('Sheet1!A1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.sheet).toBe('Sheet1');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
    });

    it('parses sheet-qualified absolute ref Sheet1!$A$1', () => {
      const node = p('Sheet1!$A$1') as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.sheet).toBe('Sheet1');
      expect(node.col).toBe('A');
      expect(node.row).toBe(1);
      expect(node.absCol).toBe(true);
      expect(node.absRow).toBe(true);
    });

    it('parses quoted sheet ref', () => {
      const node = p("'My Sheet'!B2") as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.sheet).toBe('My Sheet');
      expect(node.col).toBe('B');
      expect(node.row).toBe(2);
    });

    it('parses quoted sheet with escaped quote', () => {
      const node = p("'It''s Sheet'!C3") as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.sheet).toBe("It's Sheet");
      expect(node.col).toBe('C');
      expect(node.row).toBe(3);
    });

    it('parses quoted sheet with absolute ref', () => {
      const node = p("'My Sheet'!$B$2") as CellRefNode;
      expect(node.type).toBe('cell_ref');
      expect(node.sheet).toBe('My Sheet');
      expect(node.col).toBe('B');
      expect(node.row).toBe(2);
      expect(node.absCol).toBe(true);
      expect(node.absRow).toBe(true);
    });
  });

  // -----------------------------------------------------------------------
  // Ranges
  // -----------------------------------------------------------------------
  describe('ranges', () => {
    it('parses A1:B2', () => {
      const node = p('A1:B2') as RangeNode;
      expect(node.type).toBe('range');
      expect(node.start.type).toBe('cell_ref');
      expect(node.start.col).toBe('A');
      expect(node.start.row).toBe(1);
      expect(node.end.col).toBe('B');
      expect(node.end.row).toBe(2);
    });

    it('parses absolute range $A$1:$B$10', () => {
      const node = p('$A$1:$B$10') as RangeNode;
      expect(node.type).toBe('range');
      expect(node.start.absCol).toBe(true);
      expect(node.start.absRow).toBe(true);
      expect(node.end.absCol).toBe(true);
      expect(node.end.absRow).toBe(true);
      expect(node.end.row).toBe(10);
    });

    it('parses mixed absolute range A$1:$B2', () => {
      const node = p('A$1:$B2') as RangeNode;
      expect(node.type).toBe('range');
      expect(node.start.absCol).toBe(false);
      expect(node.start.absRow).toBe(true);
      expect(node.end.absCol).toBe(true);
      expect(node.end.absRow).toBe(false);
    });

    it('parses wide range A1:XFD1048576', () => {
      const node = p('A1:XFD1048576') as RangeNode;
      expect(node.type).toBe('range');
      expect(node.end.col).toBe('XFD');
      expect(node.end.row).toBe(1048576);
    });
  });

  // -----------------------------------------------------------------------
  // Operator precedence
  // -----------------------------------------------------------------------
  describe('operator precedence', () => {
    it('parses 1+2*3 with correct precedence', () => {
      // 1 + (2 * 3)
      const node = p('1+2*3') as BinaryOpNode;
      expect(node.type).toBe('binary_op');
      expect(node.op).toBe('+');
      expect((node.left as NumberNode).value).toBe(1);

      const right = node.right as BinaryOpNode;
      expect(right.type).toBe('binary_op');
      expect(right.op).toBe('*');
      expect((right.left as NumberNode).value).toBe(2);
      expect((right.right as NumberNode).value).toBe(3);
    });

    it('parses 2*3+1 with correct precedence', () => {
      // (2 * 3) + 1
      const node = p('2*3+1') as BinaryOpNode;
      expect(node.type).toBe('binary_op');
      expect(node.op).toBe('+');

      const left = node.left as BinaryOpNode;
      expect(left.op).toBe('*');
      expect((left.left as NumberNode).value).toBe(2);
      expect((left.right as NumberNode).value).toBe(3);

      expect((node.right as NumberNode).value).toBe(1);
    });

    it('parses 2^3^4 as right-associative', () => {
      // 2 ^ (3 ^ 4)
      const node = p('2^3^4') as BinaryOpNode;
      expect(node.type).toBe('binary_op');
      expect(node.op).toBe('^');
      expect((node.left as NumberNode).value).toBe(2);

      const right = node.right as BinaryOpNode;
      expect(right.op).toBe('^');
      expect((right.left as NumberNode).value).toBe(3);
      expect((right.right as NumberNode).value).toBe(4);
    });

    it('parses concatenation A1&B1', () => {
      const node = p('A1&B1') as BinaryOpNode;
      expect(node.type).toBe('binary_op');
      expect(node.op).toBe('&');
      expect((node.left as CellRefNode).col).toBe('A');
      expect((node.right as CellRefNode).col).toBe('B');
    });

    it('parses comparison A1>=10', () => {
      const node = p('A1>=10') as BinaryOpNode;
      expect(node.type).toBe('binary_op');
      expect(node.op).toBe('>=');
      expect((node.left as CellRefNode).type).toBe('cell_ref');
      expect((node.right as NumberNode).value).toBe(10);
    });

    it('parses comparison A1<>0', () => {
      const node = p('A1<>0') as BinaryOpNode;
      expect(node.op).toBe('<>');
    });

    it('parses comparison A1=B1', () => {
      const node = p('A1=B1') as BinaryOpNode;
      expect(node.op).toBe('=');
    });

    it('parses comparison A1<B1', () => {
      const node = p('A1<B1') as BinaryOpNode;
      expect(node.op).toBe('<');
    });

    it('parses comparison A1>B1', () => {
      const node = p('A1>B1') as BinaryOpNode;
      expect(node.op).toBe('>');
    });

    it('parses comparison A1<=B1', () => {
      const node = p('A1<=B1') as BinaryOpNode;
      expect(node.op).toBe('<=');
    });

    it('parses concatenation lower than additive: "a"&1+2', () => {
      // "a" & (1 + 2)
      const node = p('"a"&1+2') as BinaryOpNode;
      expect(node.op).toBe('&');
      expect((node.left as StringNode).value).toBe('a');
      const right = node.right as BinaryOpNode;
      expect(right.op).toBe('+');
    });

    it('parses division 10/5', () => {
      const node = p('10/5') as BinaryOpNode;
      expect(node.op).toBe('/');
      expect((node.left as NumberNode).value).toBe(10);
      expect((node.right as NumberNode).value).toBe(5);
    });

    it('parses subtraction 10-3', () => {
      const node = p('10-3') as BinaryOpNode;
      expect(node.op).toBe('-');
    });

    it('parses chained addition left-to-right: 1+2+3', () => {
      // (1+2)+3
      const node = p('1+2+3') as BinaryOpNode;
      expect(node.op).toBe('+');
      const left = node.left as BinaryOpNode;
      expect(left.op).toBe('+');
      expect((left.left as NumberNode).value).toBe(1);
      expect((left.right as NumberNode).value).toBe(2);
      expect((node.right as NumberNode).value).toBe(3);
    });
  });

  // -----------------------------------------------------------------------
  // Unary operators
  // -----------------------------------------------------------------------
  describe('unary operators', () => {
    it('parses -A1', () => {
      const node = p('-A1') as UnaryOpNode;
      expect(node.type).toBe('unary_op');
      expect(node.op).toBe('-');
      expect((node.operand as CellRefNode).col).toBe('A');
    });

    it('parses +5', () => {
      const node = p('+5') as UnaryOpNode;
      expect(node.type).toBe('unary_op');
      expect(node.op).toBe('+');
      expect((node.operand as NumberNode).value).toBe(5);
    });

    it('parses double unary --1', () => {
      const node = p('--1') as UnaryOpNode;
      expect(node.type).toBe('unary_op');
      expect(node.op).toBe('-');
      const inner = node.operand as UnaryOpNode;
      expect(inner.type).toBe('unary_op');
      expect(inner.op).toBe('-');
      expect((inner.operand as NumberNode).value).toBe(1);
    });

    it('parses 1+-2', () => {
      const node = p('1+-2') as BinaryOpNode;
      expect(node.op).toBe('+');
      expect((node.left as NumberNode).value).toBe(1);
      const right = node.right as UnaryOpNode;
      expect(right.type).toBe('unary_op');
      expect(right.op).toBe('-');
      expect((right.operand as NumberNode).value).toBe(2);
    });
  });

  // -----------------------------------------------------------------------
  // Percent
  // -----------------------------------------------------------------------
  describe('percent', () => {
    it('parses 50%', () => {
      const node = p('50%') as PercentNode;
      expect(node.type).toBe('percent');
      expect((node.operand as NumberNode).value).toBe(50);
    });

    it('parses A1%', () => {
      const node = p('A1%') as PercentNode;
      expect(node.type).toBe('percent');
      expect((node.operand as CellRefNode).col).toBe('A');
    });

    it('parses chained percent 50%%', () => {
      const node = p('50%%') as PercentNode;
      expect(node.type).toBe('percent');
      const inner = node.operand as PercentNode;
      expect(inner.type).toBe('percent');
      expect((inner.operand as NumberNode).value).toBe(50);
    });
  });

  // -----------------------------------------------------------------------
  // Parentheses
  // -----------------------------------------------------------------------
  describe('parentheses', () => {
    it('parses (1+2)*3', () => {
      const node = p('(1+2)*3') as BinaryOpNode;
      expect(node.op).toBe('*');

      const left = node.left as BinaryOpNode;
      expect(left.op).toBe('+');
      expect((left.left as NumberNode).value).toBe(1);
      expect((left.right as NumberNode).value).toBe(2);

      expect((node.right as NumberNode).value).toBe(3);
    });

    it('parses nested parens ((1+2))', () => {
      const node = p('((1+2))') as BinaryOpNode;
      expect(node.op).toBe('+');
      expect((node.left as NumberNode).value).toBe(1);
      expect((node.right as NumberNode).value).toBe(2);
    });

    it('parses (-1)', () => {
      const node = p('(-1)') as UnaryOpNode;
      expect(node.type).toBe('unary_op');
      expect(node.op).toBe('-');
      expect((node.operand as NumberNode).value).toBe(1);
    });
  });

  // -----------------------------------------------------------------------
  // Function calls
  // -----------------------------------------------------------------------
  describe('function calls', () => {
    it('parses SUM(A1:A10)', () => {
      const node = p('SUM(A1:A10)') as FunctionCallNode;
      expect(node.type).toBe('function');
      expect(node.name).toBe('SUM');
      expect(node.args).toHaveLength(1);

      const arg = node.args[0] as RangeNode;
      expect(arg.type).toBe('range');
      expect(arg.start.col).toBe('A');
      expect(arg.start.row).toBe(1);
      expect(arg.end.col).toBe('A');
      expect(arg.end.row).toBe(10);
    });

    it('parses IF(A1>0,1,0)', () => {
      const node = p('IF(A1>0,1,0)') as FunctionCallNode;
      expect(node.type).toBe('function');
      expect(node.name).toBe('IF');
      expect(node.args).toHaveLength(3);

      const cond = node.args[0] as BinaryOpNode;
      expect(cond.op).toBe('>');

      expect((node.args[1] as NumberNode).value).toBe(1);
      expect((node.args[2] as NumberNode).value).toBe(0);
    });

    it('parses no-arg function NOW()', () => {
      const node = p('NOW()') as FunctionCallNode;
      expect(node.type).toBe('function');
      expect(node.name).toBe('NOW');
      expect(node.args).toHaveLength(0);
    });

    it('parses no-arg function TODAY()', () => {
      const node = p('TODAY()') as FunctionCallNode;
      expect(node.name).toBe('TODAY');
      expect(node.args).toHaveLength(0);
    });

    it('parses single-arg function ABS(-5)', () => {
      const node = p('ABS(-5)') as FunctionCallNode;
      expect(node.name).toBe('ABS');
      expect(node.args).toHaveLength(1);
      const arg = node.args[0] as UnaryOpNode;
      expect(arg.type).toBe('unary_op');
      expect(arg.op).toBe('-');
    });

    it('parses VLOOKUP with 4 args', () => {
      const node = p('VLOOKUP(A1,B1:D10,3,FALSE)') as FunctionCallNode;
      expect(node.name).toBe('VLOOKUP');
      expect(node.args).toHaveLength(4);
      expect((node.args[0] as CellRefNode).col).toBe('A');
      expect((node.args[1] as RangeNode).type).toBe('range');
      expect((node.args[2] as NumberNode).value).toBe(3);
      expect((node.args[3] as BooleanNode).value).toBe(false);
    });
  });

  // -----------------------------------------------------------------------
  // Nested functions
  // -----------------------------------------------------------------------
  describe('nested functions', () => {
    it('parses IF(SUM(A1:A5)>10,TRUE,FALSE)', () => {
      const node = p('IF(SUM(A1:A5)>10,TRUE,FALSE)') as FunctionCallNode;
      expect(node.name).toBe('IF');
      expect(node.args).toHaveLength(3);

      const cond = node.args[0] as BinaryOpNode;
      expect(cond.op).toBe('>');

      const sum = cond.left as FunctionCallNode;
      expect(sum.name).toBe('SUM');
      expect(sum.args).toHaveLength(1);
      expect((sum.args[0] as RangeNode).type).toBe('range');

      expect((cond.right as NumberNode).value).toBe(10);
      expect((node.args[1] as BooleanNode).value).toBe(true);
      expect((node.args[2] as BooleanNode).value).toBe(false);
    });

    it('parses IFERROR(VLOOKUP(A1,B1:D10,2,0),"")', () => {
      const node = p('IFERROR(VLOOKUP(A1,B1:D10,2,0),"")') as FunctionCallNode;
      expect(node.name).toBe('IFERROR');
      expect(node.args).toHaveLength(2);

      const vlookup = node.args[0] as FunctionCallNode;
      expect(vlookup.name).toBe('VLOOKUP');
      expect(vlookup.args).toHaveLength(4);

      const fallback = node.args[1] as StringNode;
      expect(fallback.value).toBe('');
    });

    it('parses deeply nested IF(IF(IF(1,2,3),4,5),6,7)', () => {
      const node = p('IF(IF(IF(1,2,3),4,5),6,7)') as FunctionCallNode;
      expect(node.name).toBe('IF');
      const inner1 = node.args[0] as FunctionCallNode;
      expect(inner1.name).toBe('IF');
      const inner2 = inner1.args[0] as FunctionCallNode;
      expect(inner2.name).toBe('IF');
      expect((inner2.args[0] as NumberNode).value).toBe(1);
    });
  });

  // -----------------------------------------------------------------------
  // Array constants
  // -----------------------------------------------------------------------
  describe('array constants', () => {
    it('parses {1,2;3,4}', () => {
      const node = p('{1,2;3,4}') as ArrayNode;
      expect(node.type).toBe('array');
      expect(node.rows).toHaveLength(2);
      expect(node.rows[0]).toHaveLength(2);
      expect(node.rows[1]).toHaveLength(2);
      expect((node.rows[0][0] as NumberNode).value).toBe(1);
      expect((node.rows[0][1] as NumberNode).value).toBe(2);
      expect((node.rows[1][0] as NumberNode).value).toBe(3);
      expect((node.rows[1][1] as NumberNode).value).toBe(4);
    });

    it('parses single-row array {1,2,3}', () => {
      const node = p('{1,2,3}') as ArrayNode;
      expect(node.rows).toHaveLength(1);
      expect(node.rows[0]).toHaveLength(3);
    });

    it('parses single-element array {42}', () => {
      const node = p('{42}') as ArrayNode;
      expect(node.rows).toHaveLength(1);
      expect(node.rows[0]).toHaveLength(1);
      expect((node.rows[0][0] as NumberNode).value).toBe(42);
    });

    it('parses array with strings {"a","b";"c","d"}', () => {
      const node = p('{"a","b";"c","d"}') as ArrayNode;
      expect(node.rows).toHaveLength(2);
      expect((node.rows[0][0] as StringNode).value).toBe('a');
      expect((node.rows[0][1] as StringNode).value).toBe('b');
      expect((node.rows[1][0] as StringNode).value).toBe('c');
      expect((node.rows[1][1] as StringNode).value).toBe('d');
    });

    it('parses array with booleans {TRUE,FALSE}', () => {
      const node = p('{TRUE,FALSE}') as ArrayNode;
      expect(node.rows[0]).toHaveLength(2);
      expect((node.rows[0][0] as BooleanNode).value).toBe(true);
      expect((node.rows[0][1] as BooleanNode).value).toBe(false);
    });

    it('parses array with errors {#N/A,#REF!}', () => {
      const node = p('{#N/A,#REF!}') as ArrayNode;
      expect(node.rows[0]).toHaveLength(2);
      expect((node.rows[0][0] as ErrorNode).value).toBe('#N/A');
      expect((node.rows[0][1] as ErrorNode).value).toBe('#REF!');
    });

    it('parses array with unary minus {-1,2}', () => {
      const node = p('{-1,2}') as ArrayNode;
      expect(node.rows[0]).toHaveLength(2);
      const first = node.rows[0][0] as UnaryOpNode;
      expect(first.type).toBe('unary_op');
      expect(first.op).toBe('-');
      expect((first.operand as NumberNode).value).toBe(1);
    });

    it('parses empty array {}', () => {
      const node = p('{}') as ArrayNode;
      expect(node.type).toBe('array');
      expect(node.rows).toHaveLength(0);
    });
  });

  // -----------------------------------------------------------------------
  // Named ranges
  // -----------------------------------------------------------------------
  describe('named ranges', () => {
    it('parses MyRange as NameNode', () => {
      const node = p('MyRange') as NameNode;
      expect(node.type).toBe('name');
      expect(node.name).toBe('MyRange');
    });

    it('parses underscore-prefixed name', () => {
      const node = p('_custom') as NameNode;
      expect(node.type).toBe('name');
      expect(node.name).toBe('_custom');
    });

    it('parses name with dots', () => {
      const node = p('Data.Total') as NameNode;
      expect(node.type).toBe('name');
      expect(node.name).toBe('Data.Total');
    });

    it('parses name in expression: MyRange+1', () => {
      const node = p('MyRange+1') as BinaryOpNode;
      expect(node.op).toBe('+');
      expect((node.left as NameNode).type).toBe('name');
      expect((node.left as NameNode).name).toBe('MyRange');
    });
  });

  // -----------------------------------------------------------------------
  // Complex formulas
  // -----------------------------------------------------------------------
  describe('complex formulas', () => {
    it('parses IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")', () => {
      const node = p('IF(AND(A1>0,B1<10),SUM(C1:C10)*1.1,"N/A")') as FunctionCallNode;
      expect(node.name).toBe('IF');
      expect(node.args).toHaveLength(3);

      // First arg: AND(A1>0,B1<10)
      const andCall = node.args[0] as FunctionCallNode;
      expect(andCall.name).toBe('AND');
      expect(andCall.args).toHaveLength(2);

      // Second arg: SUM(C1:C10)*1.1
      const mul = node.args[1] as BinaryOpNode;
      expect(mul.op).toBe('*');
      const sumCall = mul.left as FunctionCallNode;
      expect(sumCall.name).toBe('SUM');
      expect((mul.right as NumberNode).value).toBeCloseTo(1.1);

      // Third arg: "N/A"
      expect((node.args[2] as StringNode).value).toBe('N/A');
    });

    it('parses formula with mixed references and operators', () => {
      // $A$1+Sheet1!B2*100%
      const node = p('$A$1+Sheet1!B2*100%') as BinaryOpNode;
      expect(node.op).toBe('+');

      const left = node.left as CellRefNode;
      expect(left.absCol).toBe(true);
      expect(left.absRow).toBe(true);

      const right = node.right as BinaryOpNode;
      expect(right.op).toBe('*');
      expect((right.left as CellRefNode).sheet).toBe('Sheet1');

      const pct = right.right as PercentNode;
      expect(pct.type).toBe('percent');
      expect((pct.operand as NumberNode).value).toBe(100);
    });

    it('parses formula with leading =', () => {
      const node = p('=1+2') as BinaryOpNode;
      expect(node.op).toBe('+');
      expect((node.left as NumberNode).value).toBe(1);
      expect((node.right as NumberNode).value).toBe(2);
    });

    it('parses SUMPRODUCT((A1:A10>0)*(B1:B10))', () => {
      const node = p('SUMPRODUCT((A1:A10>0)*(B1:B10))') as FunctionCallNode;
      expect(node.name).toBe('SUMPRODUCT');
      expect(node.args).toHaveLength(1);

      const mul = node.args[0] as BinaryOpNode;
      expect(mul.op).toBe('*');
    });

    it('parses INDEX(MATCH()) pattern', () => {
      const node = p('INDEX(A1:A10,MATCH(B1,C1:C10,0))') as FunctionCallNode;
      expect(node.name).toBe('INDEX');
      expect(node.args).toHaveLength(2);
      const matchCall = node.args[1] as FunctionCallNode;
      expect(matchCall.name).toBe('MATCH');
      expect(matchCall.args).toHaveLength(3);
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
    it('parses A1', () => {
      const ref = parseCellRefValue('A1');
      expect(ref.col).toBe('A');
      expect(ref.row).toBe(1);
      expect(ref.absCol).toBe(false);
      expect(ref.absRow).toBe(false);
      expect(ref.sheet).toBeUndefined();
    });

    it('parses $A$1', () => {
      const ref = parseCellRefValue('$A$1');
      expect(ref.col).toBe('A');
      expect(ref.row).toBe(1);
      expect(ref.absCol).toBe(true);
      expect(ref.absRow).toBe(true);
    });

    it('parses Sheet1!B2', () => {
      const ref = parseCellRefValue('Sheet1!B2');
      expect(ref.sheet).toBe('Sheet1');
      expect(ref.col).toBe('B');
      expect(ref.row).toBe(2);
    });

    it("parses 'My Sheet'!$C$3", () => {
      const ref = parseCellRefValue("'My Sheet'!$C$3");
      expect(ref.sheet).toBe('My Sheet');
      expect(ref.col).toBe('C');
      expect(ref.row).toBe(3);
      expect(ref.absCol).toBe(true);
      expect(ref.absRow).toBe(true);
    });

    it("parses 'It''s Sheet'!D4", () => {
      const ref = parseCellRefValue("'It''s Sheet'!D4");
      expect(ref.sheet).toBe("It's Sheet");
      expect(ref.col).toBe('D');
      expect(ref.row).toBe(4);
    });

    it('parses AA100', () => {
      const ref = parseCellRefValue('AA100');
      expect(ref.col).toBe('AA');
      expect(ref.row).toBe(100);
    });

    it('parses $XFD$1048576', () => {
      const ref = parseCellRefValue('$XFD$1048576');
      expect(ref.col).toBe('XFD');
      expect(ref.row).toBe(1048576);
      expect(ref.absCol).toBe(true);
      expect(ref.absRow).toBe(true);
    });
  });

  // -----------------------------------------------------------------------
  // Semicolons as argument separators (European locales)
  // -----------------------------------------------------------------------
  describe('semicolon argument separators', () => {
    it('parses SUM(1;2) with semicolons as separators', () => {
      const node = p('SUM(1;2)') as FunctionCallNode;
      expect(node.name).toBe('SUM');
      expect(node.args).toHaveLength(2);
      expect((node.args[0] as NumberNode).value).toBe(1);
      expect((node.args[1] as NumberNode).value).toBe(2);
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

    it('handles whitespace in formulas', () => {
      const node = p('1 + 2') as BinaryOpNode;
      expect(node.op).toBe('+');
      expect((node.left as NumberNode).value).toBe(1);
      expect((node.right as NumberNode).value).toBe(2);
    });

    it('handles percent in arithmetic: 50%+1', () => {
      const node = p('50%+1') as BinaryOpNode;
      expect(node.op).toBe('+');
      const left = node.left as PercentNode;
      expect(left.type).toBe('percent');
      expect((node.right as NumberNode).value).toBe(1);
    });

    it('handles unary minus with exponentiation: -2^3', () => {
      // Unary has higher precedence than exponentiation in our parser
      // so -2^3 parses as (-2)^3
      const node = p('-2^3') as BinaryOpNode;
      expect(node.op).toBe('^');
      const left = node.left as UnaryOpNode;
      expect(left.type).toBe('unary_op');
      expect(left.op).toBe('-');
    });
  });
});
