/**
 * Locale-independent text comparison for the formula engine.
 *
 * `String#localeCompare` with no locale follows the host's default ICU
 * locale, so the same formula could order text differently on a Swedish and
 * an English machine.  The engine compares with one fixed `en-US` collator,
 * created once at module load.
 *
 * @module formula/collation
 */

const TEXT_COLLATOR = new Intl.Collator('en-US');

/**
 * Compare two strings with the engine's fixed `en-US` collator.
 *
 * Callers lower-case both operands first; this function only pins the locale.
 * Returns negative / zero / positive like an `Array#sort` comparator.
 */
export function compareText(a: string, b: string): number {
  return TEXT_COLLATOR.compare(a, b);
}
