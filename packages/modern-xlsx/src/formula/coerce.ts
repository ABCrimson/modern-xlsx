/**
 * Excel value coercion shared by the evaluator and the built-in functions.
 *
 * Internal (not exported from the package entries); `DECIMAL_LITERAL` is also used by `io/sheet-utils.ts`
 * for CSV formula escaping.
 *
 * @module formula/coerce
 */

import { numberToExcelText } from '../format/number-text.js';
// ERROR_SET.has, not the isError type guard: inside a `string` branch the
// guard's false side would narrow `val` to `never`.
import { ERROR_SET } from './errors.js';
import type { CellValue } from './resolver.js';

/**
 * A decimal literal: optional sign, digits with an optional point, optional
 * exponent. The point and its fraction digits sit in one optional group so
 * the digit runs never overlap and a rejecting tail cannot make the match
 * backtrack quadratically. Shared with `io/sheet-utils.ts` (CSV formula escaping);
 * not global, so `test` keeps no state.
 */
export const DECIMAL_LITERAL: RegExp = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;

/**
 * Longest text value Excel allows in a cell or formula result (32,767
 * characters). `&` and CONCATENATE keep the first this many characters, as
 * Excel does; REPT and SUBSTITUTE return `#VALUE!` instead.
 */
export const MAX_TEXT_LENGTH = 32_767;

/** A decimal literal with en-US thousands grouping (`1,234`, `-1,234,567.5`). */
const GROUPED_LITERAL = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?(?:e[+-]?\d+)?$/i;

/**
 * Coerce a cell value to a number following Excel rules:
 *   null    -> 0
 *   number  -> itself
 *   boolean -> 1 / 0
 *   error   -> propagated as-is
 *   string  -> decimal literal (surrounding spaces and en-US grouping
 *              allowed), else `#VALUE!`
 *
 * Empty or blank text, hex/binary/octal, `Infinity` and overflowing literals
 * give `#VALUE!`, as in Excel. Known gap: Excel also converts currency
 * (`"$5"`), percent (`"5%"`), parenthesized negatives (`"(5)"`) and date
 * text; this engine returns `#VALUE!`.
 */
export function toNumber(val: CellValue): number | string {
  if (val === null) return 0;
  switch (typeof val) {
    case 'number':
      return val;
    case 'boolean':
      return val ? 1 : 0;
    case 'string':
      return ERROR_SET.has(val) ? val : parseDecimal(val);
  }
}

/** Parse text as a decimal literal, or return `#VALUE!`. */
function parseDecimal(text: string): number | string {
  const t = text.trim();
  let n: number;
  if (DECIMAL_LITERAL.test(t)) n = Number(t);
  else if (GROUPED_LITERAL.test(t)) n = Number(t.replaceAll(',', ''));
  else return '#VALUE!';
  return Number.isFinite(n) ? n : '#VALUE!';
}

/**
 * Coerce a cell value to text following Excel rules: null -> `''`, booleans
 * -> `'TRUE'`/`'FALSE'`, numbers -> {@link numberToExcelText} (15
 * significant digits, `1E+21`), text unchanged.
 */
export function coerceToString(val: CellValue): string {
  if (val === null) return '';
  switch (typeof val) {
    case 'boolean':
      return val ? 'TRUE' : 'FALSE';
    case 'number':
      return numberToExcelText(val);
    case 'string':
      return val;
  }
}

/**
 * Coerce a cell value to a boolean following Excel rules: null -> false,
 * numbers -> non-zero, errors propagated, `'TRUE'`/`'FALSE'` in any case,
 * other text -> `#VALUE!`.
 */
export function toBool(val: CellValue): boolean | string {
  if (val === null) return false;
  switch (typeof val) {
    case 'boolean':
      return val;
    case 'number':
      return val !== 0;
    case 'string': {
      if (ERROR_SET.has(val)) return val;
      const upper = val.toUpperCase();
      if (upper === 'TRUE') return true;
      if (upper === 'FALSE') return false;
      return '#VALUE!';
    }
  }
}

/**
 * `a ^ b` as Excel computes it: `0^0` is `#NUM!`, zero to a negative power is
 * `#DIV/0!`, and a negative base with a fractional exponent whose reciprocal
 * is an odd integer gives the real odd root (`(-8)^(1/3)` is `-2`,
 * `(-1)^(1/3)` is `-1`, as in Excel 16.0). Any other result that is not a
 * finite number (another negative base with a fractional exponent, or
 * overflow) is `#NUM!`.
 */
export function excelPower(a: number, b: number): number | string {
  if (a === 0 && b === 0) return '#NUM!';
  if (a === 0 && b < 0) return '#DIV/0!';
  const r = a < 0 && !Number.isInteger(b) && isOddInteger(1 / b) ? -((-a) ** b) : a ** b;
  return Number.isFinite(r) ? r : '#NUM!';
}

/** Whether `n` is an odd integer (`3`, `-5`). */
function isOddInteger(n: number): boolean {
  return Number.isInteger(n) && Math.abs(n) % 2 === 1;
}
