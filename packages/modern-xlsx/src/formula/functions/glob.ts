/**
 * Excel wildcard patterns for the criteria functions (SUMIF, COUNTIF, AVERAGEIF)
 * and exact lookups (VLOOKUP, HLOOKUP, MATCH 0); searchGlob is the unanchored
 * first-match form a wildcard search needs.
 *
 * `*` matches any run of characters (including none), `?` matches exactly one
 * UTF-16 code unit, and `~` makes the next `*`, `?` or `~` a literal.  A `~`
 * before any other character, or at the end of the pattern, is itself a
 * literal `~`.  Consecutive `*` collapse into one.
 *
 * Matching is case-insensitive one code unit at a time
 * (`a.toLowerCase() === b.toLowerCase()` on single characters), so the text is
 * never lower-cased as a whole and positions keep the original text's length
 * (`İ` folds to two units on its own but still occupies one position).
 *
 * No RegExp: matching is an iterative two-pointer scan that backtracks only
 * to the most recent `*`, O(pattern × text), so a hostile pattern such as
 * `*a*a*a*…b` cannot trigger catastrophic backtracking.
 *
 * Internal to the formula engine (not exported from the package entries).
 *
 * @module formula/functions/glob
 */

/** One pattern unit: a collapsed run of `*`, a `?`, or a literal code unit (folded). */
type GlobToken = { kind: 'star' } | { kind: 'any' } | { kind: 'lit'; folded: string };

const STAR: GlobToken = { kind: 'star' };
const ANY: GlobToken = { kind: 'any' };

/** Split `pattern` into tokens, resolving `~` escapes and collapsing runs of `*`. */
function tokenize(pattern: string): GlobToken[] {
  const tokens: GlobToken[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern.charAt(i);
    if (ch === '~') {
      const next = pattern.charAt(i + 1);
      if (next === '*' || next === '?' || next === '~') {
        tokens.push({ kind: 'lit', folded: next });
        i++;
      } else {
        tokens.push({ kind: 'lit', folded: '~' });
      }
    } else if (ch === '*') {
      if (tokens.at(-1) !== STAR) tokens.push(STAR);
    } else if (ch === '?') {
      tokens.push(ANY);
    } else {
      tokens.push({ kind: 'lit', folded: ch.toLowerCase() });
    }
  }
  return tokens;
}

/** Whether `token` is a `?` or a literal that accepts the code unit `ch`. */
function accepts(token: GlobToken | undefined, ch: string): boolean {
  return (
    token?.kind === 'any' ||
    (token?.kind === 'lit' && (token.folded === ch || token.folded === ch.toLowerCase()))
  );
}

/**
 * Match `tokens[first..]` against `text` from index `start`.  With
 * `prefix`, the pattern may end before the text does (it matches a prefix of
 * `text.slice(start)`); otherwise it must consume the text to its end.
 *
 * Two pointers with last-star backtracking: on a mismatch the most recent
 * `*` swallows one more unit and the scan resumes just after it.  Earlier
 * stars never need revisiting, which bounds the work at O(tokens × text).
 */
function matchFrom(
  tokens: readonly GlobToken[],
  first: number,
  text: string,
  start: number,
  prefix: boolean,
): boolean {
  let pi = first;
  let ti = start;
  // Token index of the last `*` passed and the text index it has consumed
  // up to; -1 while no `*` has been passed.
  let starPi = -1;
  let starTi = 0;

  while (ti < text.length) {
    if (prefix && pi === tokens.length) return true;
    const token = tokens[pi];
    if (token === STAR) {
      starPi = pi++;
      starTi = ti;
    } else if (accepts(token, text.charAt(ti))) {
      pi++;
      ti++;
    } else if (starPi < 0) {
      return false;
    } else {
      pi = starPi + 1;
      ti = ++starTi;
    }
  }
  while (tokens[pi] === STAR) pi++;
  return pi === tokens.length;
}

/**
 * Compile an Excel wildcard pattern into an anchored, case-insensitive
 * full-string matcher: `compileGlob('a*')('Apple')` is `true`,
 * `compileGlob('a~*b')('acb')` is `false`.
 *
 * Compile once per criterion and call the matcher per cell; each call runs in
 * O(pattern × text) time.
 */
export function compileGlob(pattern: string): (text: string) => boolean {
  const tokens = tokenize(pattern);
  return (text) => matchFrom(tokens, 0, text, 0, false);
}

/**
 * Find the first match of an Excel wildcard pattern in `text`, unanchored:
 * the smallest start `>= from` at which the pattern matches a prefix of
 * `text.slice(start)`, or `-1` when there is none.  Case folding is the same
 * as {@link compileGlob}, per code unit, so the returned index is a position
 * in `text` itself.
 *
 * `searchGlob('b?d', 'abcd', 0)` is `1`; `searchGlob('~*', 'a*b', 0)` is `1`.
 *
 * The fixed-width head before the first `*` is tried at each start; the
 * star-led rest runs at most once, because if it cannot match a prefix of the
 * text from one position it cannot from any later one.  Worst case
 * O(pattern × text).
 */
export function searchGlob(pattern: string, text: string, from: number): number {
  const tokens = tokenize(pattern);
  const firstStar = tokens.indexOf(STAR);
  const headLength = firstStar < 0 ? tokens.length : firstStar;

  for (let start = Math.max(0, from); start + headLength <= text.length; start++) {
    let i = 0;
    while (i < headLength && accepts(tokens[i], text.charAt(start + i))) i++;
    if (i < headLength) continue;
    if (firstStar < 0) return start;
    return matchFrom(tokens, firstStar, text, start + headLength, true) ? start : -1;
  }
  return -1;
}

/** `true` when `pattern` holds a `*` or `?` that is not escaped by `~`. */
export function hasWildcard(pattern: string): boolean {
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern.charAt(i);
    if (ch === '~') i++;
    else if (ch === '*' || ch === '?') return true;
  }
  return false;
}
