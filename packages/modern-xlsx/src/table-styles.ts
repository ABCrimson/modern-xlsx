/** All 60 built-in Excel table style names, grouped by category. */
export const TABLE_STYLES = {
  light: Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  medium: Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  dark: Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
} as const;

/** Set of all valid built-in table style names for validation. */
export const VALID_TABLE_STYLES: ReadonlySet<string> = new Set([
  ...TABLE_STYLES.light,
  ...TABLE_STYLES.medium,
  ...TABLE_STYLES.dark,
]);

/** Totals row aggregate functions supported by Excel tables. */
export type TotalsRowFunction =
  | 'none'
  | 'sum'
  | 'min'
  | 'max'
  | 'average'
  | 'count'
  | 'countNums'
  | 'stdDev'
  | 'var'
  | 'custom';
