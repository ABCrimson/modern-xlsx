/**
 * Recursive descent parser for Excel formula strings.
 *
 * Consumes tokens from the tokenizer and builds a typed AST.
 * Handles operator precedence via layered parse functions (lowest to highest):
 *   union `,` → comparison → concatenation `&` → additive → multiplicative →
 *   exponentiation `^` → unary prefix → postfix `%` → intersection ` ` → atom
 * (an atom includes the range `A1:B2`, so `:` binds tightest; its endpoints
 * may be cell references, names, whole columns or rows, function calls or
 * parenthesised groups or `#REF!`, and only two plain cell references make a
 * `range` node: `A1:XFE1`, `A2:INDEX(A:A,B1)` and the chain `A1:B1:C1`, folded left,
 * are `binary_op` `':'`). Every binary
 * operator is left-associative, `^` included (`2^3^2` is `(2^3)^2`, as in
 * Excel). The union layer exists only at the top level and inside a
 * parenthesised group; a comma between function arguments is never union.
 *
 * Error-recovering: collects errors rather than throwing, returning partial ASTs
 * for malformed input when possible.  Hostile input is bounded the same way:
 * a body over 8192 characters is rejected before tokenizing, nesting deeper
 * than 256 levels (parentheses, function calls, array constants, unary
 * prefixes, `^` chains) or more than 64 nested function levels (Excel's
 * limit) stops the parse with one error and a `#VALUE!` node in place of the
 * abandoned subtree.  The other left-folded chains (`+ - * / &`, comparisons,
 * `%`, `:`, intersection and union) are not counted: within 8192 characters
 * they build left spines thousands of levels deep, which the evaluator,
 * serializer, reference walker and resolver follow with loops rather than
 * recursion.
 *
 * @module formula/parser
 */

import type { Token, TokenType } from './tokenizer.js';
import { tokenize } from './tokenizer.js';

// ---------------------------------------------------------------------------
// AST node types
// ---------------------------------------------------------------------------

/** A numeric literal: `1`, `2.5`, `1E3`; also the `0` placeholder the parser substitutes for a missing or unexpected operand. */
export interface NumberNode {
  type: 'number';
  value: number;
}

/** A text literal, with `""` unescaped. */
export interface StringNode {
  type: 'string';
  value: string;
}

/** A `TRUE` or `FALSE` literal. */
export interface BooleanNode {
  type: 'boolean';
  value: boolean;
}

/** An error literal such as `#REF!` (`value` is its text); also the `#VALUE!` placeholder for input the parser abandons. */
export interface ErrorNode {
  type: 'error';
  value: string;
}

/** A single cell reference: `A1`, `$B$2`, `Sheet1!C3`, `Sheet1:Sheet3!D4`. */
export interface CellRefNode {
  type: 'cell_ref';
  /** Sheet name (unquoted, `''` unescaped); the first sheet of a 3-D span. */
  sheet?: string;
  /** Last sheet of a 3-D span (`Sheet3` in `Sheet1:Sheet3!A1`); set only with `sheet`. */
  endSheet?: string;
  /** Column letters, upper-case. */
  col: string;
  /** 1-based row number. */
  row: number;
  absCol: boolean;
  absRow: boolean;
}

/**
 * A cell range `start:end`. A sheet or 3-D prefix lives on `start` only;
 * `end` carries a sheet only when the source names one after the colon.
 */
export interface RangeNode {
  type: 'range';
  start: CellRefNode;
  end: CellRefNode;
}

/** A defined name, as written (`MyRange`). The parser does not resolve it. */
export interface NameNode {
  type: 'name';
  name: string;
}

/** A function call: the name as written and its argument nodes. */
export interface FunctionCallNode {
  type: 'function';
  name: string;
  args: ASTNode[];
}

/**
 * An infix operation. `op` is an arithmetic, text or comparison operator
 * (`+ - * / ^ & = <> < > <= >=`) or a reference operator: `':'` (range),
 * `' '` (intersection) or `','` (union).
 */
export interface BinaryOpNode {
  type: 'binary_op';
  op: string;
  left: ASTNode;
  right: ASTNode;
}

/** An empty function argument, as in `IF(A1,,2)` or `SUM(A1,)`. */
export interface MissingArgNode {
  type: 'missing';
}

/** A whole-column range: `A:A`, `$B:$D`, `Data!B:C`, `Sheet1:Sheet3!A:A`. */
export interface ColumnRangeNode {
  type: 'column_range';
  /** Sheet name; the first sheet of a 3-D span. */
  sheet?: string;
  /** Last sheet of a 3-D span; set only with `sheet`. */
  endSheet?: string;
  /** First column letters, upper-case. */
  startCol: string;
  /** Last column letters, upper-case. */
  endCol: string;
  absStart: boolean;
  absEnd: boolean;
}

/** A whole-row range: `1:1`, `$2:$5`, `Sheet1!3:4`, `Sheet1:Sheet3!1:1`. */
export interface RowRangeNode {
  type: 'row_range';
  /** Sheet name; the first sheet of a 3-D span. */
  sheet?: string;
  /** Last sheet of a 3-D span; set only with `sheet`. */
  endSheet?: string;
  /** First 1-based row number. */
  startRow: number;
  /** Last 1-based row number. */
  endRow: number;
  absStart: boolean;
  absEnd: boolean;
}

/** A structured table reference, kept as its verbatim source text: `Table1[[#Headers],[Col]]`, `[@Col]`. */
export interface StructuredRefNode {
  type: 'structured_ref';
  text: string;
}

/** A prefix operator (`op`, e.g. `-` or `+`) applied to `operand`. */
export interface UnaryOpNode {
  type: 'unary_op';
  op: string;
  operand: ASTNode;
}

/** A postfix `%` applied to `operand`. */
export interface PercentNode {
  type: 'percent';
  operand: ASTNode;
}

/** An array constant `{1,2;3,4}`: `rows` holds one node list per row. */
export interface ArrayNode {
  type: 'array';
  rows: ASTNode[][];
}

/** Any node the parser produces. */
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
  | ArrayNode
  | MissingArgNode
  | ColumnRangeNode
  | RowRangeNode
  | StructuredRefNode;

/** The outcome of parseFormula(): the tree (`null` when none was built) and every tokenizer and parser message. */
export interface ParseResult {
  ast: ASTNode | null;
  errors: string[];
}

// ---------------------------------------------------------------------------
// Cell reference value parser
// ---------------------------------------------------------------------------

/** A reference token's sheet prefix, split from the reference text after the `!`. */
interface SheetPrefix {
  /** Sheet name, or the first sheet of a 3-D span. */
  sheet?: string;
  /** Last sheet of a 3-D span. */
  endSheet?: string;
  /** The text after the `!` (the whole value when there is no prefix). */
  rest: string;
}

/**
 * Split the sheet prefix off a reference token value: none, `Sheet1!`,
 * `'My Sheet'!`, or a 3-D `Sheet1:Sheet3!` / `'Jan 1:Mar 1'!`. A quoted
 * prefix ends at the first `'` that is not part of a doubled `''`, so a name
 * such as `a'!b` (written `'a''!b'!`) splits correctly. A sheet name cannot
 * contain `:`, so a colon inside the prefix always separates a 3-D span.
 */
function splitSheetPrefix(value: string): SheetPrefix {
  let names: string;
  let rest: string;
  if (value.startsWith("'")) {
    let close = 1;
    while (close < value.length) {
      if (value.charAt(close) === "'") {
        if (value.charAt(close + 1) !== "'") break;
        close += 2;
      } else {
        close++;
      }
    }
    if (value.charAt(close + 1) !== '!') return { rest: value };
    names = value.slice(1, close).replaceAll("''", "'");
    rest = value.slice(close + 2);
  } else {
    const bang = value.indexOf('!');
    if (bang < 0) return { rest: value };
    names = value.slice(0, bang);
    rest = value.slice(bang + 1);
  }
  const colon = names.indexOf(':');
  if (colon < 0) return { sheet: names, rest };
  return { sheet: names.slice(0, colon), endSheet: names.slice(colon + 1), rest };
}

/** Copy a prefix's sheet names onto a reference node, leaving absent ones unset. */
function applySheetPrefix<T extends { sheet?: string; endSheet?: string }>(
  node: T,
  prefix: SheetPrefix,
): T {
  if (prefix.sheet !== undefined) node.sheet = prefix.sheet;
  if (prefix.endSheet !== undefined) node.endSheet = prefix.endSheet;
  return node;
}

/** Split `$?X:$?Y` into its two endpoints, without `$`, and their absolute flags. */
function splitWholeRange(text: string): [string, boolean, string, boolean] {
  const colon = text.indexOf(':');
  const start = text.slice(0, colon);
  const end = text.slice(colon + 1);
  return [start.replace('$', ''), start.startsWith('$'), end.replace('$', ''), end.startsWith('$')];
}

/** Parse a tokenizer `column_range` value (`$A:C`, `Data!B:C`) into a {@link ColumnRangeNode}. */
function parseColumnRangeValue(value: string): ColumnRangeNode {
  const prefix = splitSheetPrefix(value);
  const [startCol, absStart, endCol, absEnd] = splitWholeRange(prefix.rest);
  const node: ColumnRangeNode = {
    type: 'column_range',
    startCol: startCol.toUpperCase(),
    endCol: endCol.toUpperCase(),
    absStart,
    absEnd,
  };
  return applySheetPrefix(node, prefix);
}

/** Parse a tokenizer `row_range` value (`1:1`, `Sheet1!$2:5`) into a {@link RowRangeNode}. */
function parseRowRangeValue(value: string): RowRangeNode {
  const prefix = splitSheetPrefix(value);
  const [startRow, absStart, endRow, absEnd] = splitWholeRange(prefix.rest);
  const node: RowRangeNode = {
    type: 'row_range',
    startRow: Number(startRow),
    endRow: Number(endRow),
    absStart,
    absEnd,
  };
  return applySheetPrefix(node, prefix);
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
 *   'My Sheet'!A1, 'It''s'!$A$1 (`''` is an escaped quote)
 *   Sheet1:Sheet3!A1, 'Jan 1:Mar 1'!A1 (3-D: `sheet` and `endSheet`)
 */
export function parseCellRefValue(value: string): CellRefNode {
  const prefix = splitSheetPrefix(value);
  const node: CellRefNode = { type: 'cell_ref', ...scanCellParts(prefix.rest) };
  return applySheetPrefix(node, prefix);
}

// ---------------------------------------------------------------------------
// Comparison operators set
// ---------------------------------------------------------------------------

const COMPARISON_OPS: ReadonlySet<string> = new Set(['=', '<>', '<', '>', '<=', '>=']);

/** Token types that start an endpoint of the range operator `:`; the `#REF!` error literal is one too (see `isRangeEndpoint`). */
const RANGE_ENDPOINT_START: ReadonlySet<TokenType> = new Set<TokenType>([
  'cell_ref',
  'name',
  'column_range',
  'row_range',
  'function',
  'paren_open',
]);

/** Whether `tok` can be an endpoint of the range operator `:`: a start type, or the `#REF!` constant (the only error that is a reference). */
function isRangeEndpoint(tok: Token | undefined): boolean {
  return (
    tok !== undefined &&
    (RANGE_ENDPOINT_START.has(tok.type) || (tok.type === 'error' && tok.value === '#REF!'))
  );
}

// ---------------------------------------------------------------------------
// Limits
// ---------------------------------------------------------------------------

/** Longest formula body (text after the leading `=`) Excel accepts. */
const MAX_FORMULA_LENGTH = 8192;
/** Deepest recursive nesting the parser follows before giving up. */
const MAX_NESTING_DEPTH = 256;
/** Excel's limit on nested function levels. */
const MAX_FUNCTION_DEPTH = 64;

// ---------------------------------------------------------------------------
// Parser class (internal)
// ---------------------------------------------------------------------------

class Parser {
  private readonly tokens: Token[];
  private pos = 0;
  readonly errors: string[];
  /** Current recursion depth across every nesting construct. */
  #depth = 0;
  /** Current number of enclosing function calls. */
  #functionDepth = 0;
  /**
   * Set once nesting exceeded {@link MAX_NESTING_DEPTH} or function nesting
   * exceeded {@link MAX_FUNCTION_DEPTH}; the rest of the input is abandoned.
   */
  #abandoned = false;

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

  private expect(type: TokenType): Token | undefined {
    const tok = this.peek();
    if (tok && tok.type === type) {
      return this.advance();
    }
    // After a limit breach the input is abandoned; report it only once.
    if (this.#abandoned) return undefined;
    const actual = tok ? `${tok.type} '${tok.value}'` : 'end of input';
    this.errors.push(`Expected ${type} but found ${actual}`);
    return undefined;
  }

  /**
   * Run `parse` one nesting level deeper.  Past {@link MAX_NESTING_DEPTH} it
   * gives up through {@link Parser.#tooDeepNode} instead of recursing.
   */
  #nested(parse: () => ASTNode): ASTNode {
    this.#depth++;
    try {
      return this.#depth > MAX_NESTING_DEPTH ? this.#tooDeepNode() : parse();
    } finally {
      this.#depth--;
    }
  }

  /** Abandon the input because nesting exceeded {@link MAX_NESTING_DEPTH}. */
  #tooDeepNode(): ErrorNode {
    return this.#abandon('Formula nested too deeply');
  }

  /**
   * Record `message` once, skip the remaining tokens and return the `#VALUE!`
   * placeholder that stands in for the abandoned subtree.
   */
  #abandon(message: string): ErrorNode {
    if (!this.#abandoned) {
      this.#abandoned = true;
      this.errors.push(message);
      this.pos = this.tokens.length;
    }
    return { type: 'error', value: '#VALUE!' };
  }

  // -- entry ---------------------------------------------------------------

  parse(): ASTNode | null {
    if (this.tokens.length === 0) {
      return null;
    }
    const node = this.parseUnion();
    if (this.pos < this.tokens.length) {
      const tok = this.peek();
      if (tok) {
        this.errors.push(`Unexpected token ${tok.type} '${tok.value}' at position ${tok.start}`);
      }
    }
    return node;
  }

  // -- precedence layers (lowest → highest) --------------------------------

  /**
   * Level 0: the union operator `,`, folded left into `binary_op` `','`.
   * Reached only from the top level and from a parenthesised group, so the
   * commas between function arguments never mean union.
   */
  private parseUnion(): ASTNode {
    let left = this.parseComparison();
    while (this.peek()?.type === 'comma') {
      this.advance();
      const right = this.parseComparison();
      left = { type: 'binary_op', op: ',', left, right };
    }
    return left;
  }

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

  /**
   * Level 5: exponentiation ^ (left-associative, like Excel: `2^3^2` is
   * `(2^3)^2` = 64). Each `^` in a chain counts one nesting level, because
   * the left-nested AST it builds is one level deeper for every walker.
   */
  private parseExponentiation(): ASTNode {
    let left = this.parseUnary();
    const depth = this.#depth;
    try {
      while (true) {
        const tok = this.peek();
        if (!(tok && tok.type === 'operator' && tok.value === '^')) break;
        this.advance();
        this.#depth++;
        const right = this.#depth > MAX_NESTING_DEPTH ? this.#tooDeepNode() : this.parseUnary();
        left = { type: 'binary_op', op: '^', left, right };
      }
    } finally {
      this.#depth = depth;
    }
    return left;
  }

  /** Level 6: unary prefix +, - */
  private parseUnary(): ASTNode {
    const tok = this.peek();
    if (tok && tok.type === 'prefix_op') {
      const op = this.advance().value;
      const operand = this.#nested(() => this.parseUnary()); // allow chaining: --1
      return { type: 'unary_op', op, operand };
    }
    return this.parsePostfix();
  }

  /** Level 7: postfix % */
  private parsePostfix(): ASTNode {
    let node = this.parseIntersection();
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

  /**
   * Level 8: the intersection operator (the tokenizer's single-space
   * `operator` between two references), folded left into `binary_op` `' '`.
   * It binds tighter than every operator but `:`, which the atom layer
   * handles inside a range.
   */
  private parseIntersection(): ASTNode {
    let left = this.parseAtom();
    while (true) {
      const tok = this.peek();
      if (!(tok && tok.type === 'operator' && tok.value === ' ')) break;
      this.advance();
      const right = this.parseAtom();
      left = { type: 'binary_op', op: ' ', left, right };
    }
    return left;
  }

  // -- atoms ---------------------------------------------------------------

  /**
   * An atom and any range operators after it, folded left. The range operator
   * `:` joins two reference-valued endpoints: a cell reference, name, whole
   * column or row, function call, parenthesised group or `#REF!` (and, along a chain,
   * the range built so far). Two plain cell references make a `range` node
   * (`A1:B2`); anything else is a `binary_op` `':'`: a name endpoint
   * (`A1:XFE1`, `Total:A1`, since names are not resolved), a `#REF!` endpoint
   * (`#REF!:C2`, as Excel writes a deleted reference), a function endpoint
   * (`A2:INDEX(A:A,B1)`) or a chain (`A1:B1:C1` is `(A1:B1):C1`). A colon
   * after a cell reference with no such endpoint is reported and skipped.
   */
  private parseAtom(): ASTNode {
    const startsRange = isRangeEndpoint(this.peek());
    let node = this.parsePrimary();
    if (!startsRange) return node;
    while (this.peek()?.type === 'colon') {
      if (!isRangeEndpoint(this.tokens[this.pos + 1])) {
        if (node.type === 'cell_ref') {
          // Colon but no second endpoint — error recovery
          this.advance();
          this.errors.push('Expected cell reference after ":"');
        }
        break;
      }
      this.advance(); // consume ':'
      const end = this.parsePrimary();
      node =
        node.type === 'cell_ref' && end.type === 'cell_ref'
          ? { type: 'range', start: node, end }
          : { type: 'binary_op', op: ':', left: node, right: end };
    }
    return node;
  }

  /** One atom: a literal, reference, name, function call, parenthesised group or array constant. */
  private parsePrimary(): ASTNode {
    const tok = this.peek();
    if (!tok) {
      if (!this.#abandoned) this.errors.push('Unexpected end of formula');
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
        return parseCellRefValue(tok.value);
      }

      case 'column_range': {
        this.advance();
        return parseColumnRangeValue(tok.value);
      }

      case 'row_range': {
        this.advance();
        return parseRowRangeValue(tok.value);
      }

      case 'structured_ref': {
        this.advance();
        return { type: 'structured_ref', text: tok.value };
      }

      case 'name': {
        this.advance();
        return { type: 'name', name: tok.value };
      }

      case 'function': {
        return this.#nested(() => this.parseFunctionCall());
      }

      case 'paren_open': {
        return this.#nested(() => {
          this.advance(); // consume '('
          const expr = this.parseUnion(); // `(A1,B2)` is a union
          this.expect('paren_close');
          return expr;
        });
      }

      case 'array_open': {
        return this.#nested(() => this.parseArray());
      }

      default: {
        this.advance();
        this.errors.push(`Unexpected token ${tok.type} '${tok.value}' at position ${tok.start}`);
        return { type: 'number', value: 0 };
      }
    }
  }

  // -- function call -------------------------------------------------------

  /**
   * Parse one function call, or abandon the input with a `#VALUE!` node in
   * its place when it would be nested past {@link MAX_FUNCTION_DEPTH}.
   */
  private parseFunctionCall(): ASTNode {
    if (this.#functionDepth >= MAX_FUNCTION_DEPTH) {
      return this.#abandon(`More than ${MAX_FUNCTION_DEPTH} nested function levels`);
    }
    this.#functionDepth++;
    try {
      return this.parseFunctionBody();
    } finally {
      this.#functionDepth--;
    }
  }

  private parseFunctionBody(): FunctionCallNode {
    const nameTok = this.advance(); // function name
    this.expect('paren_open'); // consume '('

    const args: ASTNode[] = [];

    // `F()` has no arguments; otherwise one argument per comma plus one.
    if (this.peek()?.type !== 'paren_close') {
      args.push(this.parseArgument(false));
      // Only a comma separates arguments: a `;` (already reported by the
      // tokenizer) ends the list and fails the closing-paren check.
      while (this.peek()?.type === 'comma') {
        this.advance();
        args.push(this.parseArgument(true));
      }
    }

    this.expect('paren_close');
    return { type: 'function', name: nameTok.value, args };
  }

  /**
   * Parse one function argument. An empty one (`IF(A1,,2)`, `F(,1)`,
   * `SUM(A1,)`) becomes a {@link MissingArgNode} and consumes nothing: the
   * next token is a comma, or the closing parenthesis after a separator.
   */
  private parseArgument(afterSeparator: boolean): ASTNode {
    const next = this.peek()?.type;
    if (next === 'comma' || (afterSeparator && next === 'paren_close')) {
      return { type: 'missing' };
    }
    return this.parseComparison();
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
      const operand = this.#nested(() => this.parseArrayValue());
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
 * A body (the text after an optional leading `=`) longer than 8192
 * characters is not tokenized: the result is `{ ast: null, errors:
 * ['Formula exceeds 8192 characters'] }`.
 *
 * @param formula - Formula string, optionally starting with `=`.
 * @returns Parse result with AST and collected errors.
 */
export function parseFormula(formula: string): ParseResult {
  const bodyLength = formula.length - (formula.startsWith('=') ? 1 : 0);
  if (bodyLength > MAX_FORMULA_LENGTH) {
    return { ast: null, errors: [`Formula exceeds ${MAX_FORMULA_LENGTH} characters`] };
  }
  const { tokens, errors } = tokenize(formula);
  const parser = new Parser(tokens, [...errors]);
  const ast = parser.parse();
  return { ast, errors: parser.errors };
}
