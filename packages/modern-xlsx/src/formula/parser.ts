/**
 * Recursive descent parser for Excel formula strings.
 *
 * Consumes tokens from the tokenizer and builds a typed AST.
 * Handles operator precedence via layered parse functions (lowest to highest):
 *   comparison → concatenation → additive → multiplicative → exponentiation → unary → postfix → atom
 *
 * Error-recovering: collects errors rather than throwing, returning partial ASTs
 * for malformed input when possible.
 *
 * @module formula/parser
 */

import type { Token } from './tokenizer.js';
import { tokenize } from './tokenizer.js';

// ---------------------------------------------------------------------------
// AST node types
// ---------------------------------------------------------------------------

export interface NumberNode {
  type: 'number';
  value: number;
}

export interface StringNode {
  type: 'string';
  value: string;
}

export interface BooleanNode {
  type: 'boolean';
  value: boolean;
}

export interface ErrorNode {
  type: 'error';
  value: string;
}

export interface CellRefNode {
  type: 'cell_ref';
  sheet?: string;
  col: string;
  row: number;
  absCol: boolean;
  absRow: boolean;
}

export interface RangeNode {
  type: 'range';
  start: CellRefNode;
  end: CellRefNode;
}

export interface NameNode {
  type: 'name';
  name: string;
}

export interface FunctionCallNode {
  type: 'function';
  name: string;
  args: ASTNode[];
}

export interface BinaryOpNode {
  type: 'binary_op';
  op: string;
  left: ASTNode;
  right: ASTNode;
}

export interface UnaryOpNode {
  type: 'unary_op';
  op: string;
  operand: ASTNode;
}

export interface PercentNode {
  type: 'percent';
  operand: ASTNode;
}

export interface ArrayNode {
  type: 'array';
  rows: ASTNode[][];
}

export type ASTNode =
  | NumberNode
  | StringNode
  | BooleanNode
  | ErrorNode
  | CellRefNode
  | RangeNode
  | NameNode
  | FunctionCallNode
  | BinaryOpNode
  | UnaryOpNode
  | PercentNode
  | ArrayNode;

export interface ParseResult {
  ast: ASTNode | null;
  errors: string[];
}

// ---------------------------------------------------------------------------
// Cell reference value parser
// ---------------------------------------------------------------------------

/** Extract sheet prefix from a cell ref value. Returns [sheet, rest]. */
function extractSheet(value: string): [string | undefined, string] {
  if (value.startsWith("'")) {
    const bangIdx = value.indexOf("'!");
    if (bangIdx >= 0) {
      const raw = value.slice(1, bangIdx);
      return [raw.replace(/''/g, "'"), value.slice(bangIdx + 2)];
    }
    return [undefined, value];
  }
  const bangIdx = value.indexOf('!');
  if (bangIdx >= 0) {
    return [value.slice(0, bangIdx), value.slice(bangIdx + 1)];
  }
  return [undefined, value];
}

/** Scan column/row components from a bare cell ref string (e.g. "$A$1", "B2"). */
function scanCellParts(s: string): { col: string; row: number; absCol: boolean; absRow: boolean } {
  let pos = 0;

  let absCol = false;
  if (pos < s.length && s.charAt(pos) === '$') {
    absCol = true;
    pos++;
  }

  const colStart = pos;
  while (pos < s.length && s.charAt(pos) >= 'A' && s.charAt(pos) <= 'Z') pos++;
  // Fallback: lowercase letters
  if (pos === colStart) {
    while (pos < s.length && s.charAt(pos) >= 'a' && s.charAt(pos) <= 'z') pos++;
  }
  const col = s.slice(colStart, pos).toUpperCase();

  let absRow = false;
  if (pos < s.length && s.charAt(pos) === '$') {
    absRow = true;
    pos++;
  }

  const rowStart = pos;
  while (pos < s.length && s.charAt(pos) >= '0' && s.charAt(pos) <= '9') pos++;
  const row = Number.parseInt(s.slice(rowStart, pos), 10) || 0;

  return { col, row, absCol, absRow };
}

/**
 * Parse a tokenizer `cell_ref` value into a `CellRefNode`.
 *
 * Accepted formats:
 *   A1, $A1, A$1, $A$1
 *   Sheet1!A1, Sheet1!$A$1
 *   'My Sheet'!A1, 'My Sheet'!$A$1
 */
export function parseCellRefValue(value: string): CellRefNode {
  const [sheet, rest] = extractSheet(value);
  const { col, row, absCol, absRow } = scanCellParts(rest);
  const node: CellRefNode = { type: 'cell_ref', col, row, absCol, absRow };
  if (sheet !== undefined) {
    node.sheet = sheet;
  }
  return node;
}

// ---------------------------------------------------------------------------
// Comparison operators set
// ---------------------------------------------------------------------------

const COMPARISON_OPS: ReadonlySet<string> = new Set(['=', '<>', '<', '>', '<=', '>=']);

// ---------------------------------------------------------------------------
// Parser class (internal)
// ---------------------------------------------------------------------------

class Parser {
  private readonly tokens: Token[];
  private pos = 0;
  readonly errors: string[];

  constructor(tokens: Token[], errors: string[]) {
    this.tokens = tokens;
    this.errors = errors;
  }

  // -- cursor helpers -------------------------------------------------------

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private advance(): Token {
    const tok = this.tokens[this.pos];
    this.pos++;
    if (!tok) {
      return { type: 'number', value: '0', start: 0, end: 0 };
    }
    return tok;
  }

  private expect(type: string): Token | undefined {
    const tok = this.peek();
    if (tok && tok.type === type) {
      return this.advance();
    }
    const actual = tok ? `${tok.type} '${tok.value}'` : 'end of input';
    this.errors.push(`Expected ${type} but found ${actual}`);
    return undefined;
  }

  // -- entry ---------------------------------------------------------------

  parse(): ASTNode | null {
    if (this.tokens.length === 0) {
      return null;
    }
    const node = this.parseComparison();
    if (this.pos < this.tokens.length) {
      const tok = this.peek();
      if (tok) {
        this.errors.push(`Unexpected token ${tok.type} '${tok.value}' at position ${tok.start}`);
      }
    }
    return node;
  }

  // -- precedence layers (lowest → highest) --------------------------------

  /** Level 1: comparison operators = <> < > <= >= */
  private parseComparison(): ASTNode {
    let left = this.parseConcatenation();
    while (true) {
      const tok = this.peek();
      if (tok && tok.type === 'operator' && COMPARISON_OPS.has(tok.value)) {
        const op = this.advance().value;
        const right = this.parseConcatenation();
        left = { type: 'binary_op', op, left, right };
      } else {
        break;
      }
    }
    return left;
  }

  /** Level 2: concatenation & */
  private parseConcatenation(): ASTNode {
    let left = this.parseAdditive();
    while (true) {
      const tok = this.peek();
      if (tok && tok.type === 'operator' && tok.value === '&') {
        this.advance();
        const right = this.parseAdditive();
        left = { type: 'binary_op', op: '&', left, right };
      } else {
        break;
      }
    }
    return left;
  }

  /** Level 3: additive + - */
  private parseAdditive(): ASTNode {
    let left = this.parseMultiplicative();
    while (true) {
      const tok = this.peek();
      if (tok && tok.type === 'operator' && (tok.value === '+' || tok.value === '-')) {
        const op = this.advance().value;
        const right = this.parseMultiplicative();
        left = { type: 'binary_op', op, left, right };
      } else {
        break;
      }
    }
    return left;
  }

  /** Level 4: multiplicative * / */
  private parseMultiplicative(): ASTNode {
    let left = this.parseExponentiation();
    while (true) {
      const tok = this.peek();
      if (tok && tok.type === 'operator' && (tok.value === '*' || tok.value === '/')) {
        const op = this.advance().value;
        const right = this.parseExponentiation();
        left = { type: 'binary_op', op, left, right };
      } else {
        break;
      }
    }
    return left;
  }

  /** Level 5: exponentiation ^ (right-associative) */
  private parseExponentiation(): ASTNode {
    const left = this.parseUnary();
    const tok = this.peek();
    if (tok && tok.type === 'operator' && tok.value === '^') {
      this.advance();
      const right = this.parseExponentiation(); // right-recursive → right-associative
      return { type: 'binary_op', op: '^', left, right };
    }
    return left;
  }

  /** Level 6: unary prefix +, - */
  private parseUnary(): ASTNode {
    const tok = this.peek();
    if (tok && tok.type === 'prefix_op') {
      const op = this.advance().value;
      const operand = this.parseUnary(); // allow chaining: --1
      return { type: 'unary_op', op, operand };
    }
    return this.parsePostfix();
  }

  /** Level 7: postfix % */
  private parsePostfix(): ASTNode {
    let node = this.parseAtom();
    while (true) {
      const tok = this.peek();
      if (tok && tok.type === 'percent') {
        this.advance();
        node = { type: 'percent', operand: node };
      } else {
        break;
      }
    }
    return node;
  }

  // -- atoms ---------------------------------------------------------------

  private parseAtom(): ASTNode {
    const tok = this.peek();
    if (!tok) {
      this.errors.push('Unexpected end of formula');
      return { type: 'number', value: 0 };
    }

    switch (tok.type) {
      case 'number': {
        this.advance();
        return { type: 'number', value: Number(tok.value) };
      }

      case 'string': {
        this.advance();
        return { type: 'string', value: tok.value };
      }

      case 'boolean': {
        this.advance();
        return { type: 'boolean', value: tok.value === 'TRUE' };
      }

      case 'error': {
        this.advance();
        return { type: 'error', value: tok.value };
      }

      case 'cell_ref': {
        this.advance();
        const cellRef = parseCellRefValue(tok.value);

        // Check for range: cell_ref COLON cell_ref
        const next = this.peek();
        if (next && next.type === 'colon') {
          this.advance(); // consume ':'
          const endTok = this.peek();
          if (endTok && endTok.type === 'cell_ref') {
            this.advance();
            const endRef = parseCellRefValue(endTok.value);
            return { type: 'range', start: cellRef, end: endRef };
          }
          // Colon but no second cell ref — error recovery
          this.errors.push('Expected cell reference after ":"');
          return cellRef;
        }

        return cellRef;
      }

      case 'name': {
        this.advance();
        return { type: 'name', name: tok.value };
      }

      case 'function': {
        return this.parseFunctionCall();
      }

      case 'paren_open': {
        this.advance(); // consume '('
        const expr = this.parseComparison();
        this.expect('paren_close');
        return expr;
      }

      case 'array_open': {
        return this.parseArray();
      }

      default: {
        this.advance();
        this.errors.push(`Unexpected token ${tok.type} '${tok.value}' at position ${tok.start}`);
        return { type: 'number', value: 0 };
      }
    }
  }

  // -- function call -------------------------------------------------------

  private parseFunctionCall(): FunctionCallNode {
    const nameTok = this.advance(); // function name
    this.expect('paren_open'); // consume '('

    const args: ASTNode[] = [];
    const closeParen = this.peek();

    if (!closeParen || closeParen.type !== 'paren_close') {
      // Parse first argument
      args.push(this.parseComparison());

      // Parse remaining comma-separated arguments
      while (true) {
        const tok = this.peek();
        if (!tok || tok.type === 'paren_close') break;
        if (tok.type === 'comma' || tok.type === 'semicolon') {
          this.advance(); // consume separator
          args.push(this.parseComparison());
        } else {
          break;
        }
      }
    }

    this.expect('paren_close');
    return { type: 'function', name: nameTok.value, args };
  }

  // -- array constant ------------------------------------------------------

  private parseArray(): ArrayNode {
    this.advance(); // consume '{'
    const rows: ASTNode[][] = [];
    let currentRow: ASTNode[] = [];

    // Handle empty array
    const first = this.peek();
    if (first && first.type === 'array_close') {
      this.advance();
      return { type: 'array', rows: [] };
    }

    // Parse first value
    currentRow.push(this.parseArrayValue());

    while (true) {
      const tok = this.peek();
      if (!tok || tok.type === 'array_close') break;

      if (tok.type === 'array_col_sep') {
        this.advance();
        currentRow.push(this.parseArrayValue());
      } else if (tok.type === 'array_row_sep') {
        this.advance();
        rows.push(currentRow);
        currentRow = [];
        currentRow.push(this.parseArrayValue());
      } else {
        break;
      }
    }

    rows.push(currentRow);
    this.expect('array_close');
    return { type: 'array', rows };
  }

  /** Parse a single array element (literal + optional unary prefix) */
  private parseArrayValue(): ASTNode {
    const tok = this.peek();
    if (!tok) {
      this.errors.push('Unexpected end of array');
      return { type: 'number', value: 0 };
    }

    // Handle unary prefix inside arrays
    if (tok.type === 'prefix_op') {
      const op = this.advance().value;
      const operand = this.parseArrayValue();
      return { type: 'unary_op', op, operand };
    }

    switch (tok.type) {
      case 'number': {
        this.advance();
        return { type: 'number', value: Number(tok.value) };
      }
      case 'string': {
        this.advance();
        return { type: 'string', value: tok.value };
      }
      case 'boolean': {
        this.advance();
        return { type: 'boolean', value: tok.value === 'TRUE' };
      }
      case 'error': {
        this.advance();
        return { type: 'error', value: tok.value };
      }
      default: {
        this.advance();
        this.errors.push(`Unexpected token in array: ${tok.type} '${tok.value}'`);
        return { type: 'number', value: 0 };
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse an Excel formula string into a typed AST.
 *
 * Tokenizes the input first, then builds the AST via recursive descent.
 * Errors from both the tokenizer and parser are collected in `errors`.
 *
 * @param formula - Formula string, optionally starting with `=`.
 * @returns Parse result with AST and collected errors.
 */
export function parseFormula(formula: string): ParseResult {
  const { tokens, errors } = tokenize(formula);
  const parser = new Parser(tokens, [...errors]);
  const ast = parser.parse();
  return { ast, errors: parser.errors };
}
