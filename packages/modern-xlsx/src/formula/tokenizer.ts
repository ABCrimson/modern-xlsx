/**
 * Formula tokenizer (lexer) for Excel formula strings.
 *
 * Single-pass character scanner that breaks a stored (en-US, A1-style) formula
 * into typed tokens: number, string, boolean and error literals; cell
 * references; whole-column (`A:A`) and whole-row (`1:1`) ranges; structured
 * table references (`Table1[Col]`, `[@Col]`); names and function names;
 * operators, including the space intersection operator; and array constants.
 *
 * References may carry a sheet prefix: `Sheet1!`, `'My Sheet'!`, or a 3-D
 * span `Sheet1:Sheet3!` / `'Jan 1:Mar 1'!`. Names and unquoted sheet names
 * may use letters of any script (`Données!A1`, `Лист1!A1`, `売上`), as Excel
 * writes them; column letters and error literals stay ASCII. Each reference
 * token's value is the verbatim source text, prefix included. A
 * sheet-qualified or `$`-marked cell reference outside the grid
 * (A1:XFD1048576) is an error; an unqualified one such as `XFE1` or `A0` is
 * a `name`, as in Excel, and the evaluator,
 * which has no name table, answers it with `#NAME?`. Whole-column
 * and whole-row ranges are recognised only inside the grid.
 *
 * The tokenizer never throws: malformed input is reported in
 * {@link TokenizeResult.errors} and scanning continues.
 *
 * @module formula/tokenizer
 */

import { columnIndexOrNull } from '../format/cell-ref.js';
import { EXCEL_ERROR_LITERALS } from './errors.js';
import { MAX_ROW } from './grid.js';

/**
 * The kind of a {@link Token}.
 *
 * `'range'` is never produced by {@link tokenize}: a range such as `A1:B2`
 * lexes as `cell_ref`, `colon`, `cell_ref`. The member is deprecated and is
 * removed in 2.0.
 */
export type TokenType =
  /** Numeric literal (`1`, `3.14`, `.5`, `1E+3`); the value is the source text. */
  | 'number'
  /** String literal; the value is the unescaped content without quotes. */
  | 'string'
  /** `TRUE` or `FALSE` (any case); the value is upper-cased. */
  | 'boolean'
  /**
   * One of the Excel error literals (`#N/A`, `#SPILL!`, ...). A sheet-qualified
   * literal such as `Sheet1!#REF!` is one token whose value is the literal alone
   * and whose span covers the prefix.
   */
  | 'error'
  /**
   * Cell reference with any sheet or 3-D prefix: `A1`, `$B$2`, `Sheet1!A1`,
   * `Sheet1:Sheet3!A1`. Every one lies inside the grid A1:XFD1048576.
   */
  | 'cell_ref'
  /**
   * Deprecated: never emitted by tokenize(); removed in 2.0. `A1:B2` tokenizes
   * as `cell_ref`, `colon`, `cell_ref`; parseFormula() joins them into an AST
   * node of type `'range'`, so match that node type instead.
   */
  | 'range'
  /**
   * Defined name: an identifier that is not a function, boolean or cell
   * reference inside the grid (so `XFE1` and `A0` are names, as in Excel),
   * or quoted text with no `!` after it.
   */
  | 'name'
  /** Identifier immediately followed by `(`; the value excludes the parenthesis. */
  | 'function'
  /**
   * Binary operator: `+ - * / ^ & = <> < > <= >=`, or `' '` (a single space) for
   * the intersection operator, emitted for a whitespace run between a reference
   * and the start of another reference.
   */
  | 'operator'
  | 'paren_open'
  | 'paren_close'
  /** `,` outside an array constant. */
  | 'comma'
  /**
   * `;` outside an array constant. Stored formulas use `;` only as the
   * array-constant row separator, so this token is always accompanied by the
   * error "';' at position N is only valid as an array-constant row separator".
   */
  | 'semicolon'
  /** `:` that is not part of a `column_range`, `row_range` or 3-D prefix. */
  | 'colon'
  | 'percent'
  /** Unary `+` or `-`. */
  | 'prefix_op'
  | 'array_open'
  | 'array_close'
  /** `;` inside an array constant. */
  | 'array_row_sep'
  /** `,` inside an array constant. */
  | 'array_col_sep'
  /** Whole-column range with any sheet prefix: `A:A`, `$A:$C`, `Data!B:C`; both columns within XFD. */
  | 'column_range'
  /** Whole-row range with any sheet prefix: `1:1`, `$2:$5`, `Sheet1!2:5`; both rows within 1..1048576. */
  | 'row_range'
  /** Structured table reference: `Table1[Col]`, `Table1[[#Headers],[Col]]`, `[@Col]`. */
  | 'structured_ref';

/** One lexical token of a formula. */
export interface Token {
  /** Token kind. */
  type: TokenType;
  /** Token text; see {@link TokenType} for what each kind holds. */
  value: string;
  /** Offset of the first source character, counting a leading `=`. */
  start: number;
  /** Offset just past the last source character. */
  end: number;
}

/** Result of {@link tokenize}. */
export interface TokenizeResult {
  /** Tokens in source order. */
  tokens: Token[];
  /** Messages for malformed input, in source order; empty when the formula is well formed. */
  errors: string[];
}

/** Cell reference shape, any case: optional `$`, 1-3 letters, optional `$`, digits. */
const CELL_REF_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/;

/** Token types that indicate the next +/- is unary (prefix) */
const UNARY_CONTEXT: ReadonlySet<TokenType> = new Set<TokenType>([
  'operator',
  'prefix_op',
  'paren_open',
  'comma',
  'semicolon',
  'array_open',
  'array_col_sep',
  'array_row_sep',
]);

/** Token types that can end a reference, so whitespace after them may be an intersection. */
const REFERENCE_END: ReadonlySet<TokenType> = new Set<TokenType>([
  'cell_ref',
  'column_range',
  'row_range',
  'structured_ref',
  'name',
  'paren_close',
]);

function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9';
}

function isAlpha(ch: string): boolean {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
}

function isIdentChar(ch: string): boolean {
  return isAlpha(ch) || isDigit(ch) || ch === '_' || ch === '.';
}

/** A non-ASCII character that may start a name or unquoted sheet name: a letter or letter number of any script. */
const NON_ASCII_NAME_START = /^[\p{L}\p{Nl}]$/u;

/** A non-ASCII character that may continue a name: {@link NON_ASCII_NAME_START}, combining marks and decimal digits. */
const NON_ASCII_NAME_CHAR = /^[\p{L}\p{Nl}\p{Mn}\p{Mc}\p{Nd}]$/u;

/**
 * `true` when the code point `cp` may start a defined name, function name or
 * unquoted sheet name: an ASCII letter, `_`, or a letter of any script
 * (`Données`, `Лист1`, `売上`). Column letters stay ASCII ({@link isAlpha}).
 */
function isNameStart(cp: number | undefined): boolean {
  if (cp === undefined) return false;
  if (cp < 0x80) return cp === 0x5f || isAlpha(String.fromCharCode(cp));
  return NON_ASCII_NAME_START.test(String.fromCodePoint(cp));
}

/** `true` when the code point `cp` may continue a name: {@link isIdentChar}, or a letter, mark or digit of any script. */
function isNameChar(cp: number | undefined): boolean {
  if (cp === undefined) return false;
  if (cp < 0x80) return isIdentChar(String.fromCharCode(cp));
  return NON_ASCII_NAME_CHAR.test(String.fromCodePoint(cp));
}

/** First index at or after `pos` that does not continue a name, stepping over astral code points whole. */
function scanName(formula: string, pos: number): number {
  let i = pos;
  while (i < formula.length) {
    const cp = formula.codePointAt(i);
    if (cp === undefined || !isNameChar(cp)) break;
    i += cp > 0xffff ? 2 : 1;
  }
  return i;
}

function isWhitespace(ch: string): boolean {
  return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r';
}

/** Characters an unknown `#...` error literal may contain before its optional `!`/`?`. */
function isErrorLiteralChar(ch: string): boolean {
  return isAlpha(ch) || isDigit(ch) || ch === '_' || ch === '/';
}

/** First index at or after `pos` whose character fails `test`. */
function scanWhile(formula: string, pos: number, test: (ch: string) => boolean): number {
  let i = pos;
  while (i < formula.length && test(formula.charAt(i))) {
    i++;
  }
  return i;
}

/** Skip one `$` absolute marker at `pos`, if present. */
function skipAbsolute(formula: string, pos: number): number {
  return formula.charAt(pos) === '$' ? pos + 1 : pos;
}

/** `true` when `letters` names a column in A..XFD (any case). */
function isColumnInGrid(letters: string): boolean {
  return columnIndexOrNull(letters) !== null;
}

/** `true` when `digits` names a row in 1..{@link MAX_ROW}. */
function isRowInGrid(digits: string): boolean {
  const row = Number(digits);
  return digits.length > 0 && row >= 1 && row <= MAX_ROW;
}

/** `true` when `text` has the cell-reference shape and lies inside the grid. */
function isCellRefInGrid(text: string): boolean {
  const match = CELL_REF_RE.exec(text);
  if (match === null) return false;
  const [, letters = '', digits = ''] = match;
  return isColumnInGrid(letters) && isRowInGrid(digits);
}

/** The error literal starting at `pos`, if any. No literal is a prefix of another. */
function matchErrorLiteral(formula: string, pos: number): string | undefined {
  return EXCEL_ERROR_LITERALS.find((literal) => formula.startsWith(literal, pos));
}

/**
 * End of a `$?LETTERS:$?LETTERS` column range at `pos` with both columns in
 * A..XFD and no digit, identifier character or `$` after it; otherwise -1.
 */
function matchColumnRange(formula: string, pos: number): number {
  const startLetters = skipAbsolute(formula, pos);
  const colon = scanWhile(formula, startLetters, isAlpha);
  if (formula.charAt(colon) !== ':' || !isColumnInGrid(formula.slice(startLetters, colon))) {
    return -1;
  }
  const endLetters = skipAbsolute(formula, colon + 1);
  const end = scanWhile(formula, endLetters, isAlpha);
  const next = formula.charAt(end);
  if (
    isNameChar(formula.codePointAt(end)) ||
    next === '$' ||
    !isColumnInGrid(formula.slice(endLetters, end))
  ) {
    return -1;
  }
  return end;
}

/**
 * End of a `$?DIGITS:$?DIGITS` row range at `pos` with both rows in
 * 1..{@link MAX_ROW} and no `.`, exponent, identifier character or `$` after
 * it; otherwise -1.
 */
function matchRowRange(formula: string, pos: number): number {
  const startDigits = skipAbsolute(formula, pos);
  const colon = scanWhile(formula, startDigits, isDigit);
  if (formula.charAt(colon) !== ':' || !isRowInGrid(formula.slice(startDigits, colon))) {
    return -1;
  }
  const endDigits = skipAbsolute(formula, colon + 1);
  const end = scanWhile(formula, endDigits, isDigit);
  const next = formula.charAt(end);
  if (
    isNameChar(formula.codePointAt(end)) ||
    next === '$' ||
    !isRowInGrid(formula.slice(endDigits, end))
  ) {
    return -1;
  }
  return end;
}

/**
 * End of the bracketed part of a structured reference whose `[` is at `pos`,
 * or -1 when the brackets never balance. Inside the brackets `'` escapes the
 * next character, as Excel does for `[`, `]`, `#` and `'` in column names.
 */
function matchBrackets(formula: string, pos: number): number {
  let depth = 0;
  for (let i = pos; i < formula.length; i++) {
    const ch = formula.charAt(i);
    if (ch === "'") {
      i++;
    } else if (ch === '[') {
      depth++;
    } else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

/**
 * Index just past the `!` of a 3-D prefix's `:Sheet3!` tail at `pos` (which
 * must hold the `:`), or -1 when no unquoted sheet name and `!` follow.
 */
function match3dTail(formula: string, pos: number): number {
  if (formula.charAt(pos) !== ':') return -1;
  if (!isNameStart(formula.codePointAt(pos + 1))) return -1;
  const bang = scanName(formula, pos + 1);
  return formula.charAt(bang) === '!' ? bang + 1 : -1;
}

/** `true` when the text at `pos` starts a reference operand (for the intersection operator). */
function startsReference(formula: string, pos: number): boolean {
  const ch = formula.charAt(pos);
  return (
    isNameStart(formula.codePointAt(pos)) ||
    ch === '$' ||
    ch === "'" ||
    ch === '(' ||
    ch === '[' ||
    (isDigit(ch) && matchRowRange(formula, pos) >= 0)
  );
}

/**
 * Tokenize an Excel formula string into an array of typed tokens.
 *
 * Strips a leading `=` if present. Collects errors for malformed input
 * rather than throwing, allowing partial tokenization of broken formulas.
 * Whitespace is insignificant except between two references, where a run of
 * it becomes one `{ type: 'operator', value: ' ' }` intersection token.
 */
export function tokenize(formula: string): TokenizeResult {
  const tokens: Token[] = [];
  const errors: string[] = [];

  // Strip leading `=`
  let offset = 0;
  if (formula.length > 0 && formula.charAt(0) === '=') {
    offset = 1;
  }

  let pos = offset;
  let arrayDepth = 0;
  // End of the last operand that produced an error instead of a token, so a
  // following +/- still reads as binary (`#FOO+1`).
  let droppedOperandEnd = -1;

  /**
   * Scan the reference after a sheet prefix: `start` is where the prefix
   * begins and `refStart` is just past its `!`. Returns the new position.
   */
  const scanQualifiedRef = (start: number, refStart: number): number => {
    const literal = matchErrorLiteral(formula, refStart);
    if (literal !== undefined) {
      const end = refStart + literal.length;
      tokens.push({ type: 'error', value: literal, start, end });
      return end;
    }
    const columnEnd = matchColumnRange(formula, refStart);
    if (columnEnd >= 0) {
      tokens.push({
        type: 'column_range',
        value: formula.slice(start, columnEnd),
        start,
        end: columnEnd,
      });
      return columnEnd;
    }
    const rowEnd = matchRowRange(formula, refStart);
    if (rowEnd >= 0) {
      tokens.push({ type: 'row_range', value: formula.slice(start, rowEnd), start, end: rowEnd });
      return rowEnd;
    }
    let end = skipAbsolute(formula, refStart);
    end = scanWhile(formula, end, isAlpha);
    end = skipAbsolute(formula, end);
    end = scanWhile(formula, end, isDigit);
    if (end === refStart) {
      errors.push(
        `Expected reference after '${formula.slice(start, refStart)}' at position ${refStart}`,
      );
      droppedOperandEnd = refStart;
      return refStart;
    }
    if (isCellRefInGrid(formula.slice(refStart, end))) {
      tokens.push({ type: 'cell_ref', value: formula.slice(start, end), start, end });
      return end;
    }
    // Report the whole identifier-like run (`Sheet1!Print_Area`), not a fragment.
    end = scanName(formula, end);
    errors.push(`Invalid cell reference at position ${start}: ${formula.slice(start, end)}`);
    droppedOperandEnd = end;
    return end;
  };

  /** Scan a structured reference whose optional table name starts at `start` and `[` sits at `bracket`. */
  const scanStructuredRef = (start: number, bracket: number): number => {
    const end = matchBrackets(formula, bracket);
    if (end < 0) {
      errors.push(
        `Unbalanced '[' in structured reference at position ${bracket}: ${formula.slice(start)}`,
      );
      droppedOperandEnd = formula.length;
      return formula.length;
    }
    tokens.push({ type: 'structured_ref', value: formula.slice(start, end), start, end });
    return end;
  };

  while (pos < formula.length) {
    const ch = formula.charAt(pos);

    // (a) Whitespace: insignificant, except between two references (intersection)
    if (isWhitespace(ch)) {
      const start = pos;
      pos = scanWhile(formula, pos, isWhitespace);
      const prev = tokens.at(-1);
      if (
        arrayDepth === 0 &&
        prev !== undefined &&
        prev.end === start &&
        REFERENCE_END.has(prev.type) &&
        startsReference(formula, pos)
      ) {
        tokens.push({ type: 'operator', value: ' ', start, end: pos });
      }
      continue;
    }

    // (b) String literal
    if (ch === '"') {
      const start = pos;
      pos++; // skip opening quote
      let value = '';
      let terminated = false;
      while (pos < formula.length) {
        if (formula.charAt(pos) === '"') {
          // Check for escaped quote ""
          if (pos + 1 < formula.length && formula.charAt(pos + 1) === '"') {
            value += '"';
            pos += 2;
          } else {
            pos++; // skip closing quote
            terminated = true;
            break;
          }
        } else {
          value += formula.charAt(pos);
          pos++;
        }
      }
      if (!terminated) {
        errors.push(`Unterminated string literal at position ${start}`);
      }
      tokens.push({ type: 'string', value, start, end: pos });
      continue;
    }

    // (c) Row range (`1:3`, never inside an array constant), else a number:
    //     digit, or `.` followed by digit
    if (isDigit(ch) && arrayDepth === 0) {
      const end = matchRowRange(formula, pos);
      if (end >= 0) {
        tokens.push({ type: 'row_range', value: formula.slice(pos, end), start: pos, end });
        pos = end;
        continue;
      }
    }
    if (
      isDigit(ch) ||
      (ch === '.' && pos + 1 < formula.length && isDigit(formula.charAt(pos + 1)))
    ) {
      const start = pos;
      // Consume integer part
      while (pos < formula.length && isDigit(formula.charAt(pos))) {
        pos++;
      }
      // Consume decimal part
      if (pos < formula.length && formula.charAt(pos) === '.') {
        pos++;
        while (pos < formula.length && isDigit(formula.charAt(pos))) {
          pos++;
        }
      }
      // Consume exponent part
      if (pos < formula.length && (formula.charAt(pos) === 'e' || formula.charAt(pos) === 'E')) {
        const expStart = pos;
        pos++;
        if (pos < formula.length && (formula.charAt(pos) === '+' || formula.charAt(pos) === '-')) {
          pos++;
        }
        const digitStart = pos;
        while (pos < formula.length && isDigit(formula.charAt(pos))) {
          pos++;
        }
        if (pos === digitStart) {
          // No digits after exponent — backtrack
          pos = expStart;
        }
      }
      tokens.push({ type: 'number', value: formula.slice(start, pos), start, end: pos });
      continue;
    }

    // (d) Error literal
    if (ch === '#') {
      const start = pos;
      const literal = matchErrorLiteral(formula, pos);
      if (literal !== undefined) {
        pos += literal.length;
        tokens.push({ type: 'error', value: literal, start, end: pos });
      } else {
        // Unknown literal: `#` + [A-Za-z0-9_/]* + one optional trailing `!` or `?`
        pos = scanWhile(formula, pos + 1, isErrorLiteralChar);
        if (formula.charAt(pos) === '!' || formula.charAt(pos) === '?') {
          pos++;
        }
        errors.push(`Unknown error literal at position ${start}: ${formula.slice(start, pos)}`);
        droppedOperandEnd = pos;
      }
      continue;
    }

    // (e) Array braces
    if (ch === '{') {
      arrayDepth++;
      tokens.push({ type: 'array_open', value: '{', start: pos, end: pos + 1 });
      pos++;
      continue;
    }
    if (ch === '}') {
      arrayDepth = Math.max(0, arrayDepth - 1);
      tokens.push({ type: 'array_close', value: '}', start: pos, end: pos + 1 });
      pos++;
      continue;
    }

    // (f) Parentheses
    if (ch === '(') {
      tokens.push({ type: 'paren_open', value: '(', start: pos, end: pos + 1 });
      pos++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'paren_close', value: ')', start: pos, end: pos + 1 });
      pos++;
      continue;
    }

    // (g) Comma
    if (ch === ',') {
      if (arrayDepth > 0) {
        tokens.push({ type: 'array_col_sep', value: ',', start: pos, end: pos + 1 });
      } else {
        tokens.push({ type: 'comma', value: ',', start: pos, end: pos + 1 });
      }
      pos++;
      continue;
    }

    // (h) Semicolon: the array-constant row separator; anywhere else an error
    if (ch === ';') {
      if (arrayDepth > 0) {
        tokens.push({ type: 'array_row_sep', value: ';', start: pos, end: pos + 1 });
      } else {
        tokens.push({ type: 'semicolon', value: ';', start: pos, end: pos + 1 });
        errors.push(`';' at position ${pos} is only valid as an array-constant row separator`);
      }
      pos++;
      continue;
    }

    // (i) Colon
    if (ch === ':') {
      tokens.push({ type: 'colon', value: ':', start: pos, end: pos + 1 });
      pos++;
      continue;
    }

    // (j) Percent
    if (ch === '%') {
      tokens.push({ type: 'percent', value: '%', start: pos, end: pos + 1 });
      pos++;
      continue;
    }

    // (k) Two-char operators: <>, <=, >=
    if (pos + 1 < formula.length) {
      const next = formula.charAt(pos + 1);
      if ((ch === '<' && (next === '>' || next === '=')) || (ch === '>' && next === '=')) {
        const value = ch + next;
        tokens.push({ type: 'operator', value, start: pos, end: pos + 2 });
        pos += 2;
        continue;
      }
    }

    // Single-char operators (excluding +/- which need unary check)
    if (
      ch === '=' ||
      ch === '<' ||
      ch === '>' ||
      ch === '*' ||
      ch === '/' ||
      ch === '^' ||
      ch === '&'
    ) {
      tokens.push({ type: 'operator', value: ch, start: pos, end: pos + 1 });
      pos++;
      continue;
    }

    // (l) Plus or minus — check unary vs binary
    if (ch === '+' || ch === '-') {
      const lastToken = tokens.at(-1);
      const afterDroppedOperand = droppedOperandEnd > (lastToken?.end ?? -1);
      const isUnary =
        !afterDroppedOperand && (lastToken === undefined || UNARY_CONTEXT.has(lastToken.type));
      tokens.push({
        type: isUnary ? 'prefix_op' : 'operator',
        value: ch,
        start: pos,
        end: pos + 1,
      });
      pos++;
      continue;
    }

    // (m) Single-quoted sheet name or 3-D span: 'Sheet Name'!A1, 'Jan 1:Mar 1'!B2
    if (ch === "'") {
      const start = pos;
      pos++; // skip opening quote
      let terminated = false;
      while (pos < formula.length) {
        if (formula.charAt(pos) === "'") {
          // Check for escaped quote ''
          if (pos + 1 < formula.length && formula.charAt(pos + 1) === "'") {
            pos += 2;
          } else {
            pos++; // skip closing quote
            terminated = true;
            break;
          }
        } else {
          pos++;
        }
      }
      if (!terminated) {
        errors.push(`Unterminated quoted sheet name at position ${start}`);
        continue;
      }
      if (formula.charAt(pos) === '!') {
        pos = scanQualifiedRef(start, pos + 1);
      } else {
        // Quoted name without `!` — treat as a name token
        tokens.push({ type: 'name', value: formula.slice(start, pos), start, end: pos });
      }
      continue;
    }

    // (n) Bare structured reference: [@Col], [[#Headers],[Col]]
    if (ch === '[') {
      pos = scanStructuredRef(pos, pos);
      continue;
    }

    // (o) Letter of any script or underscore — identifier, function, boolean,
    //     cell ref, column range, structured ref, sheet or 3-D prefix
    if (isNameStart(formula.codePointAt(pos))) {
      const start = pos;
      pos = scanName(formula, pos);
      const ident = formula.slice(start, pos);
      const next = formula.charAt(pos);

      // Table1[...] → structured reference
      if (next === '[') {
        pos = scanStructuredRef(start, pos);
        continue;
      }

      // Sheet1!... → sheet-qualified reference
      if (next === '!') {
        pos = scanQualifiedRef(start, pos + 1);
        continue;
      }

      if (next === ':') {
        // Sheet1:Sheet3!... → 3-D reference. A plain cell ref before the colon
        // starts a range instead (`A1:Sheet2!B1`).
        const refStart = CELL_REF_RE.test(ident) ? -1 : match3dTail(formula, pos);
        if (refStart >= 0) {
          pos = scanQualifiedRef(start, refStart);
          continue;
        }
        // A:C → whole-column range
        const columnEnd = matchColumnRange(formula, start);
        if (columnEnd >= 0) {
          tokens.push({
            type: 'column_range',
            value: formula.slice(start, columnEnd),
            start,
            end: columnEnd,
          });
          pos = columnEnd;
          continue;
        }
      }

      // Check if followed by `(` → function
      if (next === '(') {
        tokens.push({ type: 'function', value: ident, start, end: pos });
        continue;
      }

      // Check for boolean (case-insensitive)
      const upper = ident.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ type: 'boolean', value: upper, start, end: pos });
        continue;
      }

      // Cell reference (A1, AA10). A cell-shaped identifier outside the grid
      // (XFE1, ZZZ1, A0, A1048577) is a name in Excel, so it evaluates to #NAME?.
      if (CELL_REF_RE.test(ident)) {
        const type = isCellRefInGrid(ident) ? 'cell_ref' : 'name';
        tokens.push({ type, value: ident, start, end: pos });
        continue;
      }

      // Check for absolute-row cell ref: letters followed by $digits (e.g. A$1)
      if (next === '$' && /^[A-Z]{1,3}$/i.test(ident) && isDigit(formula.charAt(pos + 1))) {
        pos = scanWhile(formula, pos + 1, isDigit);
        const value = formula.slice(start, pos);
        if (isCellRefInGrid(value)) {
          tokens.push({ type: 'cell_ref', value, start, end: pos });
        } else {
          errors.push(`Invalid cell reference at position ${start}: ${value}`);
          droppedOperandEnd = pos;
        }
        continue;
      }

      // Otherwise it is a name (named range, etc.)
      tokens.push({ type: 'name', value: ident, start, end: pos });
      continue;
    }

    // (p) `$` — absolute cell reference, column range ($A:$C) or row range ($1:$1)
    if (ch === '$') {
      const start = pos;
      const rangeEnd = Math.max(matchRowRange(formula, start), matchColumnRange(formula, start));
      if (rangeEnd >= 0) {
        const type = isDigit(formula.charAt(start + 1)) ? 'row_range' : 'column_range';
        tokens.push({ type, value: formula.slice(start, rangeEnd), start, end: rangeEnd });
        pos = rangeEnd;
        continue;
      }
      pos = scanWhile(formula, pos + 1, isAlpha);
      pos = skipAbsolute(formula, pos);
      pos = scanWhile(formula, pos, isDigit);
      const value = formula.slice(start, pos);
      if (isCellRefInGrid(value)) {
        tokens.push({ type: 'cell_ref', value, start, end: pos });
      } else {
        errors.push(`Invalid cell reference at position ${start}: ${value}`);
        droppedOperandEnd = pos;
      }
      continue;
    }

    // (q) Unknown character
    errors.push(`Unexpected character '${ch}' at position ${pos}`);
    pos++;
  }

  return { tokens, errors };
}
