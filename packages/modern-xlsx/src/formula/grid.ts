/**
 * Worksheet grid limits used by the formula engine.
 *
 * Excel 2007 and later: 1,048,576 rows by 16,384 columns (A to XFD), per
 * Microsoft's "Excel specifications and limits". Internal only.
 *
 * @module formula/grid
 */

/** Last row number (1-based). */
export const MAX_ROW = 1_048_576;

/** Last column index (0-based): 16,383 is column XFD. */
export const MAX_COL_INDEX = 16_383;
