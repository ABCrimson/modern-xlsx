/**
 * The closed set of Excel error values.
 *
 * Internal to the formula engine (not exported from the package entries).
 *
 * @module formula/errors
 */

import type { CellValue } from './resolver.js';

/**
 * Every error literal Excel can produce: the seven classic codes plus the
 * dynamic-array, data-type, connection and Python codes (the values of
 * Office JS `Excel.ErrorCellValueType`). No literal is a prefix of another,
 * so a first-match scan over this list is unambiguous.
 */
export const EXCEL_ERROR_LITERALS = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CONNECT!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#FIELD!',
  '#CALC!',
  '#EXTERNAL!',
  '#PYTHON!',
  '#BUSY!',
] as const;

/** {@link EXCEL_ERROR_LITERALS} as a set for O(1) membership tests. */
export const ERROR_SET: ReadonlySet<string> = new Set(EXCEL_ERROR_LITERALS);

/**
 * `true` when `val` is exactly one of {@link EXCEL_ERROR_LITERALS}.
 *
 * Text such as `'#1'` or `'#hashtag'` is ordinary text. Limitation: errors
 * travel as strings because `CellValue` stays `number | string | boolean |
 * null` (a distinct error type would break the exported `FormulaFunction`
 * contract), so text exactly equal to an error literal, e.g. the string
 * `"#N/A"`, is treated as that error.
 */
export function isError(val: CellValue): val is string {
  return typeof val === 'string' && ERROR_SET.has(val);
}
