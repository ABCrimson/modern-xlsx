/**
 * Formula tokenizer (lexer) for Excel formula strings.
 *
 * Single-pass character scanner that breaks a formula string into typed tokens.
 * Handles cell references, sheet-qualified references, function calls, operators,
 * string/number/boolean/error literals, array constants, and more.
 *
 * @module formula/tokenizer
 */

export type TokenType =
  | 'number'
  | 'string'
  | 'boolean'
  | 'error'
  | 'cell_ref'
  | 'range'
  | 'name'
  | 'function'
  | 'operator'
  | 'paren_open'
  | 'paren_close'
  | 'comma'
  | 'semicolon'
  | 'colon'
  | 'percent'
  | 'prefix_op'
  | 'array_open'
  | 'array_close'
  | 'array_row_sep'
  | 'array_col_sep';

export interface Token {
  type: TokenType;
  value: string;
  start: number;
  end: number;
}

export interface TokenizeResult {
  tokens: Token[];
  errors: string[];
}

/** Cell reference pattern: optional $ + 1-3 uppercase letters + optional $ + 1+ digits */
const CELL_REF_RE = /^\$?[A-Z]{1,3}\$?[0-9]+$/;

/** Known Excel error literals */
const ERROR_LITERALS = [
  '#DIV/0!',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#VALUE!',
] as const satisfies readonly string[];

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

function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9';
}

function isAlpha(ch: string): boolean {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');
}

function isIdentChar(ch: string): boolean {
  return isAlpha(ch) || isDigit(ch) || ch === '_' || ch === '.';
}

/**
 * Tokenize an Excel formula string into an array of typed tokens.
 *
 * Strips a leading `=` if present. Collects errors for malformed input
 * rather than throwing, allowing partial tokenization of broken formulas.
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

  while (pos < formula.length) {
    const ch = formula.charAt(pos);

    // (a) Whitespace — skip
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      pos++;
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

    // (c) Number: digit, or `.` followed by digit
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
      const remainingLen = formula.length - pos;
      let matched = false;
      for (const err of ERROR_LITERALS) {
        if (remainingLen >= err.length && formula.startsWith(err, pos)) {
          tokens.push({ type: 'error', value: err, start, end: pos + err.length });
          pos += err.length;
          matched = true;
          break;
        }
      }
      if (!matched) {
        // Scan to end of error-like token for a better error message
        let end = pos + 1;
        while (
          end < formula.length &&
          formula.charAt(end) !== ' ' &&
          formula.charAt(end) !== ',' &&
          formula.charAt(end) !== ')'
        ) {
          end++;
        }
        errors.push(`Unknown error literal at position ${start}: ${formula.slice(start, end)}`);
        pos = end;
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

    // (h) Semicolon
    if (ch === ';') {
      if (arrayDepth > 0) {
        tokens.push({ type: 'array_row_sep', value: ';', start: pos, end: pos + 1 });
      } else {
        tokens.push({ type: 'semicolon', value: ';', start: pos, end: pos + 1 });
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
      const lastToken = tokens.length > 0 ? tokens[tokens.length - 1] : undefined;
      const isUnary = lastToken === undefined || UNARY_CONTEXT.has(lastToken.type);
      if (isUnary) {
        tokens.push({ type: 'prefix_op', value: ch, start: pos, end: pos + 1 });
      } else {
        tokens.push({ type: 'operator', value: ch, start: pos, end: pos + 1 });
      }
      pos++;
      continue;
    }

    // (m) Single-quoted sheet name: 'Sheet Name'!A1
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
      // Expect `!` after closing quote
      if (pos < formula.length && formula.charAt(pos) === '!') {
        pos++; // skip `!`
        // Scan cell reference after `!`
        // Allow $ prefix
        if (pos < formula.length && formula.charAt(pos) === '$') {
          pos++;
        }
        // Column letters
        while (pos < formula.length && isAlpha(formula.charAt(pos))) {
          pos++;
        }
        // Allow $ before row
        if (pos < formula.length && formula.charAt(pos) === '$') {
          pos++;
        }
        // Row digits
        while (pos < formula.length && isDigit(formula.charAt(pos))) {
          pos++;
        }
        const fullValue = formula.slice(start, pos);
        tokens.push({ type: 'cell_ref', value: fullValue, start, end: pos });
      } else {
        // Quoted name without `!` — treat as a name token
        const fullValue = formula.slice(start, pos);
        tokens.push({ type: 'name', value: fullValue, start, end: pos });
      }
      continue;
    }

    // (n) Letter or underscore — identifier, function, boolean, cell ref, sheet ref
    if (isAlpha(ch) || ch === '_') {
      const start = pos;
      // Consume identifier characters
      while (pos < formula.length && isIdentChar(formula.charAt(pos))) {
        pos++;
      }
      const ident = formula.slice(start, pos);

      // Check if followed by `!` → sheet-qualified reference
      if (pos < formula.length && formula.charAt(pos) === '!') {
        pos++; // skip `!`
        const refStart = pos;
        // Allow $ prefix
        if (pos < formula.length && formula.charAt(pos) === '$') {
          pos++;
        }
        // Column letters
        while (pos < formula.length && isAlpha(formula.charAt(pos))) {
          pos++;
        }
        // Allow $ before row
        if (pos < formula.length && formula.charAt(pos) === '$') {
          pos++;
        }
        // Row digits
        while (pos < formula.length && isDigit(formula.charAt(pos))) {
          pos++;
        }
        const refPart = formula.slice(refStart, pos);
        if (refPart.length > 0) {
          tokens.push({ type: 'cell_ref', value: `${ident}!${refPart}`, start, end: pos });
        } else {
          // `Sheet!` with nothing after — emit as name
          tokens.push({ type: 'name', value: ident, start, end: pos - 1 });
          // Back up past the `!`
          pos--;
        }
        continue;
      }

      // Check if followed by `(` → function
      if (pos < formula.length && formula.charAt(pos) === '(') {
        tokens.push({ type: 'function', value: ident, start, end: pos });
        continue;
      }

      // Check for boolean (case-insensitive)
      const upper = ident.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ type: 'boolean', value: upper, start, end: pos });
        continue;
      }

      // Check for cell reference pattern (e.g. A1, AA10)
      if (CELL_REF_RE.test(upper)) {
        tokens.push({ type: 'cell_ref', value: ident, start, end: pos });
        continue;
      }

      // Check for absolute-row cell ref: letters followed by $digits (e.g. A$1)
      if (pos < formula.length && formula.charAt(pos) === '$' && /^[A-Z]{1,3}$/i.test(ident)) {
        const savedPos = pos;
        pos++; // skip $
        if (pos < formula.length && isDigit(formula.charAt(pos))) {
          while (pos < formula.length && isDigit(formula.charAt(pos))) {
            pos++;
          }
          const value = formula.slice(start, pos);
          tokens.push({ type: 'cell_ref', value, start, end: pos });
          continue;
        }
        // Not a valid cell ref after all; restore position
        pos = savedPos;
      }

      // Otherwise it is a name (named range, etc.)
      tokens.push({ type: 'name', value: ident, start, end: pos });
      continue;
    }

    // (o) `$` — absolute cell reference
    if (ch === '$') {
      const start = pos;
      pos++; // skip `$`
      // Column letters
      while (pos < formula.length && isAlpha(formula.charAt(pos))) {
        pos++;
      }
      // Allow $ before row
      if (pos < formula.length && formula.charAt(pos) === '$') {
        pos++;
      }
      // Row digits
      while (pos < formula.length && isDigit(formula.charAt(pos))) {
        pos++;
      }
      const value = formula.slice(start, pos);
      if (CELL_REF_RE.test(value.toUpperCase())) {
        tokens.push({ type: 'cell_ref', value, start, end: pos });
      } else {
        errors.push(`Invalid cell reference at position ${start}: ${value}`);
      }
      continue;
    }

    // (p) Unknown character
    errors.push(`Unexpected character '${ch}' at position ${pos}`);
    pos++;
  }

  return { tokens, errors };
}
