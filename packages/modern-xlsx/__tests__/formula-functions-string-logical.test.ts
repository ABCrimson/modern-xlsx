import { describe, expect, it } from 'vitest';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import type { CellValue } from '../src/index.js';
import { evaluateFormula } from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const functions = createDefaultFunctions();

function ev(formula: string, cells?: Record<string, CellValue>): CellValue {
  return evaluateFormula(formula, {
    currentSheet: 'Sheet1',
    functions,
    getCell: (sheet: string, col: number, row: number): CellValue => {
      const letter = String.fromCharCode(65 + col);
      const key = `${sheet}!${letter}${row}`;
      return cells?.[key] ?? null;
    },
  });
}

// ---------------------------------------------------------------------------
// IF
// ---------------------------------------------------------------------------

describe('IF', () => {
  it('returns true branch when condition is true', () => {
    expect(ev('IF(TRUE,1,2)')).toBe(1);
  });

  it('returns false branch when condition is false', () => {
    expect(ev('IF(FALSE,1,2)')).toBe(2);
  });

  it('returns false when else branch missing and condition is false', () => {
    expect(ev('IF(FALSE,1)')).toBe(false);
  });

  it('short-circuits: does not evaluate unused branch', () => {
    // 1/0 would be #DIV/0! but should never be evaluated
    expect(ev('IF(TRUE,42,1/0)')).toBe(42);
  });

  it('propagates error from condition', () => {
    expect(ev('IF(#REF!,1,2)')).toBe('#REF!');
  });

  it('coerces number 1 to true', () => {
    expect(ev('IF(1,"yes","no")')).toBe('yes');
  });

  it('coerces number 0 to false', () => {
    expect(ev('IF(0,"yes","no")')).toBe('no');
  });

  it('nested IF', () => {
    expect(ev('IF(FALSE,1,IF(TRUE,2,3))')).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// AND / OR / NOT
// ---------------------------------------------------------------------------

describe('AND', () => {
  it('returns true when all args are truthy', () => {
    expect(ev('AND(TRUE,TRUE,TRUE)')).toBe(true);
  });

  it('returns false when any arg is falsy', () => {
    expect(ev('AND(TRUE,FALSE,TRUE)')).toBe(false);
  });

  it('coerces numbers to booleans', () => {
    expect(ev('AND(1,2,3)')).toBe(true);
  });

  it('0 is falsy', () => {
    expect(ev('AND(1,0,1)')).toBe(false);
  });

  it('propagates errors', () => {
    expect(ev('AND(TRUE,#N/A)')).toBe('#N/A');
  });
});

describe('OR', () => {
  it('returns true when any arg is truthy', () => {
    expect(ev('OR(FALSE,TRUE,FALSE)')).toBe(true);
  });

  it('returns false when all args are falsy', () => {
    expect(ev('OR(FALSE,FALSE,FALSE)')).toBe(false);
  });

  it('coerces 1 to true', () => {
    expect(ev('OR(0,0,1)')).toBe(true);
  });

  it('propagates errors', () => {
    expect(ev('OR(#DIV/0!,TRUE)')).toBe('#DIV/0!');
  });
});

describe('NOT', () => {
  it('negates true', () => {
    expect(ev('NOT(TRUE)')).toBe(false);
  });

  it('negates false', () => {
    expect(ev('NOT(FALSE)')).toBe(true);
  });

  it('negates number: NOT(0) = true', () => {
    expect(ev('NOT(0)')).toBe(true);
  });

  it('negates number: NOT(1) = false', () => {
    expect(ev('NOT(1)')).toBe(false);
  });

  it('propagates errors', () => {
    expect(ev('NOT(#VALUE!)')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// IFERROR
// ---------------------------------------------------------------------------

describe('IFERROR', () => {
  it('returns value when not an error', () => {
    expect(ev('IFERROR(42,"fallback")')).toBe(42);
  });

  it('returns fallback when value is error', () => {
    expect(ev('IFERROR(1/0,"oops")')).toBe('oops');
  });

  it('returns fallback for #N/A', () => {
    expect(ev('IFERROR(#N/A,0)')).toBe(0);
  });

  it('passes through non-error strings', () => {
    expect(ev('IFERROR("hello","fallback")')).toBe('hello');
  });
});

// ---------------------------------------------------------------------------
// CONCATENATE
// ---------------------------------------------------------------------------

describe('CONCATENATE', () => {
  it('joins two strings', () => {
    expect(ev('CONCATENATE("hello"," world")')).toBe('hello world');
  });

  it('joins multiple strings', () => {
    expect(ev('CONCATENATE("a","b","c","d")')).toBe('abcd');
  });

  it('coerces numbers to strings', () => {
    expect(ev('CONCATENATE("val=",42)')).toBe('val=42');
  });

  it('propagates errors', () => {
    expect(ev('CONCATENATE("x",#REF!)')).toBe('#REF!');
  });

  it('empty string for no args', () => {
    expect(ev('CONCATENATE()')).toBe('');
  });
});

// ---------------------------------------------------------------------------
// LEFT / RIGHT / MID
// ---------------------------------------------------------------------------

describe('LEFT', () => {
  it('returns first character by default', () => {
    expect(ev('LEFT("Hello")')).toBe('H');
  });

  it('returns first n characters', () => {
    expect(ev('LEFT("Hello",3)')).toBe('Hel');
  });

  it('returns empty for n=0', () => {
    expect(ev('LEFT("Hello",0)')).toBe('');
  });

  it('returns full string if n > length', () => {
    expect(ev('LEFT("Hi",10)')).toBe('Hi');
  });

  it('error for negative n', () => {
    expect(ev('LEFT("Hi",-1)')).toBe('#VALUE!');
  });
});

describe('RIGHT', () => {
  it('returns last character by default', () => {
    expect(ev('RIGHT("Hello")')).toBe('o');
  });

  it('returns last n characters', () => {
    expect(ev('RIGHT("Hello",3)')).toBe('llo');
  });

  it('returns empty for n=0', () => {
    expect(ev('RIGHT("Hello",0)')).toBe('');
  });

  it('returns full string if n > length', () => {
    expect(ev('RIGHT("Hi",10)')).toBe('Hi');
  });
});

describe('MID', () => {
  it('extracts substring', () => {
    expect(ev('MID("Hello",2,3)')).toBe('ell');
  });

  it('1-based start: MID("ABC",1,1) = "A"', () => {
    expect(ev('MID("ABC",1,1)')).toBe('A');
  });

  it('returns what is available if n exceeds length', () => {
    expect(ev('MID("Hi",1,10)')).toBe('Hi');
  });

  it('error for start < 1', () => {
    expect(ev('MID("Hi",0,1)')).toBe('#VALUE!');
  });

  it('error for negative count', () => {
    expect(ev('MID("Hi",1,-1)')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// LEN / TRIM / UPPER / LOWER
// ---------------------------------------------------------------------------

describe('LEN', () => {
  it('returns string length', () => {
    expect(ev('LEN("Hello")')).toBe(5);
  });

  it('empty string = 0', () => {
    expect(ev('LEN("")')).toBe(0);
  });

  it('null = 0', () => {
    expect(ev('LEN(A1)')).toBe(0);
  });

  it('propagates errors', () => {
    expect(ev('LEN(#REF!)')).toBe('#REF!');
  });
});

describe('TRIM', () => {
  it('removes leading and trailing spaces', () => {
    expect(ev('TRIM("  hello  ")')).toBe('hello');
  });

  it('collapses multiple interior spaces', () => {
    expect(ev('TRIM("  hello   world  ")')).toBe('hello world');
  });
});

describe('UPPER', () => {
  it('converts to uppercase', () => {
    expect(ev('UPPER("hello")')).toBe('HELLO');
  });

  it('mixed case', () => {
    expect(ev('UPPER("Hello World")')).toBe('HELLO WORLD');
  });
});

describe('LOWER', () => {
  it('converts to lowercase', () => {
    expect(ev('LOWER("HELLO")')).toBe('hello');
  });

  it('mixed case', () => {
    expect(ev('LOWER("Hello World")')).toBe('hello world');
  });
});

// ---------------------------------------------------------------------------
// TEXT / VALUE
// ---------------------------------------------------------------------------

describe('TEXT', () => {
  it('formats number with decimal places', () => {
    expect(ev('TEXT(3.14159,"0.00")')).toBe('3.14');
  });

  it('formats number with more decimal places', () => {
    expect(ev('TEXT(1.5,"0.000")')).toBe('1.500');
  });

  it('formats as percent', () => {
    expect(ev('TEXT(0.75,"0%")')).toBe('75%');
  });

  it('propagates errors', () => {
    expect(ev('TEXT(#REF!,"0.00")')).toBe('#REF!');
  });
});

describe('VALUE', () => {
  it('parses numeric string', () => {
    expect(ev('VALUE("42")')).toBe(42);
  });

  it('parses float string', () => {
    expect(ev('VALUE("3.14")')).toBe(3.14);
  });

  it('returns #VALUE! for non-numeric string', () => {
    expect(ev('VALUE("abc")')).toBe('#VALUE!');
  });

  it('returns number as-is', () => {
    expect(ev('VALUE(42)')).toBe(42);
  });

  it('coerces boolean', () => {
    expect(ev('VALUE(TRUE)')).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// EXACT
// ---------------------------------------------------------------------------

describe('EXACT', () => {
  it('true for identical strings', () => {
    expect(ev('EXACT("hello","hello")')).toBe(true);
  });

  it('false for different case', () => {
    expect(ev('EXACT("Hello","hello")')).toBe(false);
  });

  it('false for different strings', () => {
    expect(ev('EXACT("abc","xyz")')).toBe(false);
  });

  it('propagates errors', () => {
    expect(ev('EXACT(#N/A,"test")')).toBe('#N/A');
  });
});

// ---------------------------------------------------------------------------
// SUBSTITUTE
// ---------------------------------------------------------------------------

describe('SUBSTITUTE', () => {
  it('replaces all occurrences', () => {
    expect(ev('SUBSTITUTE("aabaa","a","x")')).toBe('xxbxx');
  });

  it('replaces specific instance', () => {
    expect(ev('SUBSTITUTE("aabaa","a","x",2)')).toBe('axbaa');
  });

  it('replaces first instance', () => {
    expect(ev('SUBSTITUTE("aabaa","a","x",1)')).toBe('xabaa');
  });

  it('replaces third instance', () => {
    expect(ev('SUBSTITUTE("aabaa","a","x",3)')).toBe('aabxa');
  });

  it('returns original if old_text not found', () => {
    expect(ev('SUBSTITUTE("hello","z","x")')).toBe('hello');
  });

  it('returns original if instance not found', () => {
    expect(ev('SUBSTITUTE("aabaa","a","x",10)')).toBe('aabaa');
  });

  it('handles empty old_text', () => {
    expect(ev('SUBSTITUTE("hello","","x")')).toBe('hello');
  });
});

// ---------------------------------------------------------------------------
// REPT
// ---------------------------------------------------------------------------

describe('REPT', () => {
  it('repeats string', () => {
    expect(ev('REPT("ab",3)')).toBe('ababab');
  });

  it('0 repeats returns empty', () => {
    expect(ev('REPT("x",0)')).toBe('');
  });

  it('1 repeat returns original', () => {
    expect(ev('REPT("hello",1)')).toBe('hello');
  });

  it('negative repeats returns error', () => {
    expect(ev('REPT("x",-1)')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// FIND / SEARCH
// ---------------------------------------------------------------------------

describe('FIND', () => {
  it('finds substring (case-sensitive)', () => {
    expect(ev('FIND("lo","Hello")')).toBe(4);
  });

  it('returns #VALUE! when not found', () => {
    expect(ev('FIND("XY","Hello")')).toBe('#VALUE!');
  });

  it('case-sensitive: does not find lowercase', () => {
    expect(ev('FIND("h","Hello")')).toBe('#VALUE!');
  });

  it('with start position', () => {
    expect(ev('FIND("l","Hello World",5)')).toBe(10);
  });

  it('error for start < 1', () => {
    expect(ev('FIND("a","abc",0)')).toBe('#VALUE!');
  });
});

describe('SEARCH', () => {
  it('finds substring (case-insensitive)', () => {
    expect(ev('SEARCH("lo","Hello")')).toBe(4);
  });

  it('case-insensitive match', () => {
    expect(ev('SEARCH("h","Hello")')).toBe(1);
  });

  it('returns #VALUE! when not found', () => {
    expect(ev('SEARCH("XY","Hello")')).toBe('#VALUE!');
  });

  it('with start position', () => {
    expect(ev('SEARCH("l","Hello World",5)')).toBe(10);
  });
});
