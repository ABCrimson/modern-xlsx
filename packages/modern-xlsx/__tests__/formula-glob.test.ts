import { describe, expect, it } from 'vitest';
import { compileGlob, hasWildcard, searchGlob } from '../src/formula/functions/glob.js';

// ---------------------------------------------------------------------------
// compileGlob: anchored, case-insensitive Excel wildcard match
// ---------------------------------------------------------------------------

describe('compileGlob', () => {
  it('matches * as any run, case-insensitively', () => {
    expect(compileGlob('a*')('Apple')).toBe(true);
    expect(compileGlob('a*')('banana')).toBe(false);
    expect(compileGlob('***')('')).toBe(true);
  });

  it('matches ? as exactly one character', () => {
    expect(compileGlob('b?nana')('banana')).toBe(true);
    expect(compileGlob('b?nana')('bnana')).toBe(false);
  });

  it('treats ~*, ~? and ~~ as literals', () => {
    expect(compileGlob('a~*b')('a*b')).toBe(true);
    expect(compileGlob('a~*b')('acb')).toBe(false);
    expect(compileGlob('a~?')('a?')).toBe(true);
    expect(compileGlob('a~?')('ab')).toBe(false);
    expect(compileGlob('~~')('~')).toBe(true);
  });

  it('keeps a ~ before any other character, or at the end, as a literal ~', () => {
    expect(compileGlob('~a')('~a')).toBe(true);
    expect(compileGlob('a~')('a~')).toBe(true);
  });

  it('is anchored at both ends', () => {
    expect(compileGlob('b')('abc')).toBe(false);
    expect(compileGlob('ab')('abc')).toBe(false);
  });

  it('finishes a ReDoS-shaped pattern in linear-ish time', () => {
    const match = compileGlob(`${'*a'.repeat(10)}b`);
    // The 10 s testTimeout (vitest.config.ts) bounds catastrophic backtracking or quadratic work.
    expect(match('a'.repeat(40))).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// searchGlob: unanchored first match, per-unit case folding
// ---------------------------------------------------------------------------

describe('searchGlob', () => {
  it('returns the smallest start where the pattern matches a prefix', () => {
    expect(searchGlob('b?d', 'abcd', 0)).toBe(1);
    expect(searchGlob('~*', 'a*b', 0)).toBe(1);
    expect(searchGlob('B', 'abcb', 0)).toBe(1);
  });

  it('folds case per code unit, so indexes keep the text length', () => {
    expect(searchGlob('b', 'İabc', 0)).toBe(2);
  });

  it('starts at from and returns -1 when nothing matches', () => {
    expect(searchGlob('b', 'abcb', 2)).toBe(3);
    expect(searchGlob('x', 'abc', 0)).toBe(-1);
    expect(searchGlob('a*c', 'abd', 0)).toBe(-1);
  });

  it('matches an empty pattern or a leading * at from', () => {
    expect(searchGlob('', 'abc', 1)).toBe(1);
    expect(searchGlob('*c', 'abc', 0)).toBe(0);
    expect(searchGlob('b*d', 'abxbcd', 0)).toBe(1);
  });

  it('finishes a ReDoS-shaped pattern quickly', () => {
    // The 10 s testTimeout (vitest.config.ts) bounds only the `*a` x10 `b` assertion; the `a?a..*b` assertion has no time bound and is a bench candidate.
    expect(searchGlob(`${'*a'.repeat(10)}b`, 'a'.repeat(2000), 0)).toBe(-1);
    expect(searchGlob(`a${'?a'.repeat(5)}*b`, 'a'.repeat(2000), 0)).toBe(-1);
  });
});

// ---------------------------------------------------------------------------
// hasWildcard
// ---------------------------------------------------------------------------

describe('hasWildcard', () => {
  it('detects an unescaped * or ?', () => {
    expect(hasWildcard('a*')).toBe(true);
    expect(hasWildcard('a?')).toBe(true);
    expect(hasWildcard('~~*')).toBe(true);
  });

  it('ignores escaped wildcards and plain text', () => {
    expect(hasWildcard('a~*')).toBe(false);
    expect(hasWildcard('~?')).toBe(false);
    expect(hasWildcard('abc')).toBe(false);
  });
});
