import { describe, expect, it } from 'vitest';
import { compareText } from '../src/formula/collation.js';
import { createDefaultFunctions } from '../src/formula/functions/index.js';
import { mapRefs } from '../src/formula/walk.js';
import type { ASTNode, CellValue, EvalContext, RangeNode } from '../src/index.js';
import {
  evaluateFormula,
  evaluateNode,
  expandSharedFormula,
  INVALID_ARGUMENT,
  ModernXlsxError,
  parseFormula,
  resolveRange,
  rewriteFormula,
  serializeFormula,
} from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const functions = createDefaultFunctions();

interface CountingContext extends EvalContext {
  calls: number;
}

/** Build a context over `cells` (keys like `A1`, columns A-Z only) that counts getCell calls. */
function makeCtx(
  cells: Record<string, CellValue> = {},
  bounds?: { maxRow: number; maxCol: number },
): CountingContext {
  const ctx: CountingContext = {
    calls: 0,
    currentSheet: 'Sheet1',
    functions,
    getCell(_sheet: string, col: number, row: number): CellValue {
      ctx.calls++;
      const key = `${String.fromCharCode(65 + col)}${row}`;
      return cells[key] ?? null;
    },
  };
  if (bounds) ctx.getBounds = () => bounds;
  return ctx;
}

const ev = (formula: string, cells?: Record<string, CellValue>): CellValue =>
  evaluateFormula(formula, makeCtx(cells));

function rangeOf(formula: string): RangeNode {
  const { ast } = parseFormula(formula);
  if (ast?.type !== 'range') throw new Error(`not a range: ${formula}`);
  return ast;
}

// ---------------------------------------------------------------------------
// Range bounds
// ---------------------------------------------------------------------------

describe('range bounds', () => {
  const abc = { A1: 1, A2: 2, A3: 3 };

  it('clamps a whole-sheet range to getBounds', () => {
    const ctx = makeCtx(abc, { maxRow: 3, maxCol: 0 });
    expect(evaluateFormula('=SUM(A1:XFD1048576)', ctx)).toBe(6);
    expect(ctx.calls).toBeLessThanOrEqual(3);
  });

  it('returns #NUM! fast for a whole-sheet range without bounds', () => {
    const ctx = makeCtx(abc);
    // ctx.calls === 0 below is the deterministic guard: no cell is read for an unbounded range.
    expect(evaluateFormula('=SUM(A1:XFD1048576)', ctx)).toBe('#NUM!');
    expect(ctx.calls).toBe(0);
  });

  it('exported resolveRange throws INVALID_ARGUMENT without bounds', () => {
    const range = rangeOf('=A1:XFD1048576');
    let caught: unknown;
    try {
      resolveRange(range, makeCtx(abc));
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(ModernXlsxError);
    expect((caught as ModernXlsxError).code).toBe(INVALID_ARGUMENT);
    expect((caught as ModernXlsxError).message).toContain('Range spans 17179869184 cells');
  });

  it('exported resolveRange clamps with bounds and keeps the top-left', () => {
    expect(resolveRange(rangeOf('=A2:C9'), makeCtx(abc, { maxRow: 3, maxCol: 0 }))).toEqual([
      [2],
      [3],
    ]);
  });

  // Excel reads an unqualified endpoint outside the grid as a name (#NAME?).
  it('returns #NAME? for endpoints outside the grid', () => {
    expect(ev('=A1048577')).toBe('#NAME?');
    expect(ev('=SUM(A1:A99999999999)')).toBe('#NAME?');
    expect(ev('=A1:A99999999999')).toBe('#NAME?');
    expect(ev('=SUM(XFE1:XFE2)')).toBe('#NAME?');
  });

  it('COUNTBLANK counts the declared area, not the clamped one', () => {
    const ctx = makeCtx(abc, { maxRow: 3, maxCol: 0 });
    expect(evaluateFormula('=COUNTBLANK(A1:A10)', ctx)).toBe(7);
  });

  it('a fully outside range resolves empty', () => {
    const ctx = makeCtx(abc, { maxRow: 3, maxCol: 0 });
    expect(evaluateFormula('=SUM(C5:D9)', ctx)).toBe(0);
    expect(evaluateFormula('=COUNTBLANK(C5:D9)', ctx)).toBe(10);
    expect(ctx.calls).toBe(0);
  });

  it('COUNTIF counts blanks beyond the clamp when the criteria matches blank', () => {
    const ctx = makeCtx(abc, { maxRow: 3, maxCol: 0 });
    expect(evaluateFormula('=COUNTIF(A1:A10,"<>2")', ctx)).toBe(9);
    expect(evaluateFormula('=COUNTIF(A1:A10,"")', ctx)).toBe(7);
  });

  it('SUMIF and AVERAGEIF visit sum cells whose criteria cells are clamped away', () => {
    const cells = { B1: 1, B2: 1, B3: 1, B4: 1, B5: 1 };
    const ctx = makeCtx(cells, { maxRow: 5, maxCol: 1 });
    expect(evaluateFormula('=SUMIF(D1:D5,"",B1:B5)', ctx)).toBe(5);
    expect(evaluateFormula('=AVERAGEIF(D1:D5,"<>x",B1:B5)', ctx)).toBe(1);
  });

  it('clamping does not shift INDEX, MATCH or lookup positions', () => {
    const cells = { A1: 'a', A2: 'b', A3: 'c', B1: 10, B2: 20, B3: 30 };
    const ctx = makeCtx(cells, { maxRow: 3, maxCol: 1 });
    expect(evaluateFormula('=MATCH("c",A1:A1048576,0)', ctx)).toBe(3);
    expect(evaluateFormula('=INDEX(A1:C1048576,2,2)', ctx)).toBe(20);
    expect(evaluateFormula('=INDEX(A1:C1048576,500,3)', ctx)).toBe(null);
    expect(evaluateFormula('=INDEX(A1:C10,11,1)', ctx)).toBe('#REF!');
    expect(evaluateFormula('=VLOOKUP("b",A1:C1048576,3,FALSE)', ctx)).toBe(null);
    expect(evaluateFormula('=VLOOKUP("b",A1:C1048576,4,FALSE)', ctx)).toBe('#REF!');
    expect(evaluateFormula('=SUMPRODUCT(B1:B10,B1:B10)', ctx)).toBe(1400);
  });

  it('a lookup table without bounds shares the resolver read cap of 1,048,576 cells', () => {
    const cells = { A1: 'a', A2: 'b', B1: 1, B2: 2 };
    expect(ev('=VLOOKUP("b",A1:B524289,2,FALSE)', cells)).toBe('#NUM!');
    expect(ev('=VLOOKUP("b",A1:A1048576,1,FALSE)', cells)).toBe('b');
  });

  it('a scalar range reads at most one cell (implicit intersection)', () => {
    const ctx = makeCtx(abc);
    expect(evaluateFormula('=A2:XFD1048576', ctx)).toBe('#VALUE!');
    expect(ctx.calls).toBe(0);
    const at = makeCtx(abc);
    at.currentCell = { col: 3, row: 2 };
    expect(evaluateFormula('=A1:A1048576', at)).toBe(2);
    expect(at.calls).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Parser limits
// ---------------------------------------------------------------------------

describe('parser limits', () => {
  const nestedParens = (n: number): string => `=${'('.repeat(n)}1${')'.repeat(n)}`;

  it('100k nested parens are rejected without throwing', () => {
    const formula = nestedParens(100_000);
    expect(() => parseFormula(formula)).not.toThrow();
    expect(parseFormula(formula).errors.length).toBeGreaterThan(0);
    expect(evaluateFormula(formula, makeCtx())).toBe('#VALUE!');
  });

  it('nesting past 256 levels reports "Formula nested too deeply" once', () => {
    const { ast, errors } = parseFormula(nestedParens(4000));
    expect(ast).not.toBeNull();
    expect(errors).toEqual(['Formula nested too deeply']);
    expect(parseFormula(nestedParens(256)).errors).toEqual([]);
  });

  it('100k unary minus are rejected without throwing', () => {
    expect(() => parseFormula(`=${'-'.repeat(100_000)}1`)).not.toThrow();
    expect(parseFormula(`=${'-'.repeat(8000)}1`).errors).toEqual(['Formula nested too deeply']);
    expect(parseFormula(`=${'2^'.repeat(4000)}1`).errors).toEqual(['Formula nested too deeply']);
    expect(parseFormula(`={${'-'.repeat(8000)}1}`).errors).toEqual(['Formula nested too deeply']);
  });

  it('a 66-term chain of 1+1+... evaluates to 66', () => {
    expect(ev(`=${'1+'.repeat(65)}1`)).toBe(66);
  });

  it('enforces the 64-level function nesting limit', () => {
    const nestedSum = (n: number): string => `=${'SUM('.repeat(n)}1${')'.repeat(n)}`;
    expect(parseFormula(nestedSum(64)).errors).toEqual([]);
    expect(ev(nestedSum(64))).toBe(1);
    expect(parseFormula(nestedSum(65)).errors).toEqual(['More than 64 nested function levels']);
    expect(ev(nestedSum(65))).toBe('#VALUE!');
  });

  it('evaluateNode turns stack exhaustion into #NUM! and lets other errors through', () => {
    let ast: ASTNode = { type: 'number', value: 1 };
    for (let i = 0; i < 200_000; i++) ast = { type: 'unary_op', op: '-', operand: ast };
    expect(evaluateNode(ast, makeCtx())).toBe('#NUM!');

    const throwing: EvalContext = {
      currentSheet: 'Sheet1',
      getCell: () => {
        throw new TypeError('boom');
      },
    };
    expect(() => evaluateFormula('=A1', throwing)).toThrow(TypeError);
  });

  it('rejects a body over 8192 characters', () => {
    const at = `=${'1+'.repeat(4095)}12`;
    expect(at.length - 1).toBe(8192);
    expect(parseFormula(at).errors).toEqual([]);
    const over = `${at}3`;
    expect(parseFormula(over)).toEqual({
      ast: null,
      errors: ['Formula exceeds 8192 characters'],
    });
  });
});

// ---------------------------------------------------------------------------
// Left-folded operator chains (thousands of levels deep within 8192 characters)
// ---------------------------------------------------------------------------

describe('left-folded operator chains', () => {
  const ones: EvalContext = { currentSheet: 'Sheet1', functions, getCell: () => 1 };

  it('a 4095-term + chain evaluates to 4095', () => {
    expect(evaluateFormula(`1${'+1'.repeat(4094)}`, ones)).toBe(4095);
  });

  it('long &, comparison and % chains evaluate', () => {
    expect(evaluateFormula(`"a"${'&"a"'.repeat(2000)}`, ones)).toBe('a'.repeat(2001));
    expect(evaluateFormula(`A1${'<1'.repeat(4000)}`, ones)).toBe(false);
    expect(evaluateFormula(`A1${'%'.repeat(8000)}`, ones)).toBe(0);
    expect(evaluateFormula(`A1${' A1'.repeat(2700)}`, ones)).toBe(1);
  });

  it('keeps left-first error propagation along a chain', () => {
    expect(evaluateFormula(`1/0${'+#N/A'.repeat(1600)}`, ones)).toBe('#DIV/0!');
    expect(evaluateFormula(`1${'+1'.repeat(3000)}+#N/A+1/0`, ones)).toBe('#N/A');
    expect(evaluateFormula(`#N/A%${'%'.repeat(3000)}+1`, ones)).toBe('#N/A');
  });

  it('expandSharedFormula and rewriteFormula shift a 4095-term chain', () => {
    const chain = `A1${'+1'.repeat(4094)}`;
    let shared = '';
    expect(() => {
      shared = expandSharedFormula(chain, 'B1', 'B2');
    }).not.toThrow();
    expect(shared.startsWith('A2+1')).toBe(true);
    expect(shared).toHaveLength(chain.length);
    const rewritten = rewriteFormula(chain, { type: 'insert_rows', start: 1, count: 1 });
    expect(rewritten.startsWith('A2+1')).toBe(true);
    const percent = `A1${'%'.repeat(8000)}`;
    expect(expandSharedFormula(percent, 'B1', 'B2')).toBe(`A2${'%'.repeat(8000)}`);
  });

  it('serializeFormula writes a 4095-term chain back unchanged', () => {
    const chain = `A1${'&1'.repeat(4094)}`;
    const { ast } = parseFormula(chain);
    expect(ast).not.toBeNull();
    expect(serializeFormula(ast as ASTNode)).toBe(chain);
  });

  it('SUM of a 2700-area union reads every area', () => {
    const union = `SUM((${Array.from({ length: 2700 }, () => 'A1').join(',')}))`;
    expect(evaluateFormula(union, ones)).toBe(2700);
    expect(evaluateFormula(union.replace('SUM', 'COUNT'), ones)).toBe(2700);
    expect(evaluateFormula(union.replace('SUM', 'MAX'), ones)).toBe(1);
  });

  // The string tests above overflow the ~1 MB main-thread stack before the fix,
  // but not always a worker thread's larger one. These hand-built trees are deep
  // enough to overflow any default stack, so they pin the loops in every pool.
  describe('hand-built 100,000-level trees', () => {
    const DEPTH = 100_000;
    const a1: ASTNode = { type: 'cell_ref', col: 'A', row: 1, absCol: false, absRow: false };
    const chain = (op: string, operand: ASTNode): ASTNode => {
      let node: ASTNode = operand;
      for (let i = 0; i < DEPTH; i++) node = { type: 'binary_op', op, left: node, right: operand };
      return node;
    };

    it('evaluateNode folds + and ^ chains, and a percent chain, without #NUM!', () => {
      expect(evaluateNode(chain('+', { type: 'number', value: 1 }), ones)).toBe(DEPTH + 1);
      expect(evaluateNode(chain('^', { type: 'number', value: 1 }), ones)).toBe(1);
      let percent: ASTNode = { type: 'number', value: 5 };
      for (let i = 0; i < DEPTH; i++) percent = { type: 'percent', operand: percent };
      expect(evaluateNode(percent, ones)).toBe(0);
    });

    it('evaluateNode reads intersection and union chains', () => {
      expect(evaluateNode(chain(' ', a1), ones)).toBe(1);
      const sum: ASTNode = { type: 'function', name: 'SUM', args: [chain(',', a1)] };
      expect(evaluateNode(sum, ones)).toBe(DEPTH + 1);
      const count: ASTNode = { type: 'function', name: 'COUNT', args: [chain(',', a1)] };
      expect(evaluateNode(count, ones)).toBe(DEPTH + 1);
      const max: ASTNode = { type: 'function', name: 'MAX', args: [chain(',', a1)] };
      expect(evaluateNode(max, ones)).toBe(1);
    });

    it('serializeFormula writes operator and percent chains', () => {
      expect(serializeFormula(chain('&', a1))).toBe(`A1${'&A1'.repeat(DEPTH)}`);
      expect(serializeFormula(chain(',', a1))).toBe(`A1${',A1'.repeat(DEPTH)}`);
      let percent: ASTNode = a1;
      for (let i = 0; i < DEPTH; i++) percent = { type: 'percent', operand: percent };
      expect(serializeFormula(percent)).toBe(`A1${'%'.repeat(DEPTH)}`);
    });

    it('mapRefs rebuilds a chain', () => {
      const shifted = mapRefs(chain('-', a1), {
        cellRef: (ref) => ({ ...ref, row: ref.row + 1 }),
        range: (node) => node,
        columnRange: (node) => node,
        rowRange: (node) => node,
      });
      expect(serializeFormula(shifted)).toBe(`A2${'-A2'.repeat(DEPTH)}`);
    });
  });
});

// ---------------------------------------------------------------------------
// Criteria
// ---------------------------------------------------------------------------

describe('criteria', () => {
  it('applies wildcards after <> and negates', () => {
    expect(ev('=COUNTIF(A1:A3,"<>*x*")', { A1: 'axb', A2: 'abc', A3: 5 })).toBe(2);
  });

  it('applies wildcards after = and honours ~ escapes', () => {
    const cells = { A1: 'xyz', A2: 'abc', A3: 5, A4: 'a*b', A5: 'aXb' };
    expect(ev('=COUNTIF(A1:A5,"=*b")', cells)).toBe(2);
    expect(ev('=COUNTIF(A1:A5,"a~*b")', cells)).toBe(1);
  });

  it('compares bare numeric criteria numerically', () => {
    expect(ev('=COUNTIF(A1:A3,"5.0")', { A1: 5, A2: 'five', A3: 6 })).toBe(1);
  });

  it('does not parse loose numbers', () => {
    expect(ev('=COUNTIF(A1:A2," ")', { A1: 0, A2: 1 })).toBe(0);
    expect(ev('=COUNTIF(A1:A2,"0x10")', { A1: 16, A2: 1 })).toBe(0);
    expect(ev('=COUNTIF(A1:A2,"= ")', { A1: 0, A2: 1 })).toBe(0);
    expect(ev('=COUNTIF(A1:A2,"=0x10")', { A1: 16, A2: 1 })).toBe(0);
  });

  it('finishes a ReDoS-shaped criteria quickly', () => {
    // The 10 s testTimeout (vitest.config.ts) bounds catastrophic backtracking or quadratic work.
    expect(ev(`=COUNTIF(A1,"${'*a'.repeat(12)}*b")`, { A1: 'a'.repeat(40) })).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// REPT
// ---------------------------------------------------------------------------

describe('REPT cap', () => {
  it('rejects a result over 32767 characters', () => {
    expect(ev('=REPT("x",1e9)')).toBe('#VALUE!');
    expect(ev('=REPT("x",32768)')).toBe('#VALUE!');
  });

  it('allows exactly 32767 characters', () => {
    const out = ev('=REPT("x",32767)');
    expect(typeof out).toBe('string');
    expect((out as string).length).toBe(32767);
  });

  it('rejects a non-finite count without throwing', () => {
    expect(() => ev('=REPT("",1E308*10)')).not.toThrow();
    expect(ev('=REPT("",1E308*10)')).toBe('#VALUE!');
  });
});

// ---------------------------------------------------------------------------
// Text length cap (32,767 characters)
// ---------------------------------------------------------------------------

// Expected values from Excel (Microsoft 365, COM probe on 2026-11-12): `&` and
// CONCATENATE keep the first 32,767 characters; SUBSTITUTE gives #VALUE!.
describe('text length cap', () => {
  const long = { A2: 'x'.repeat(30_000) };

  it('& keeps the first 32767 characters', () => {
    expect(ev('=LEN(A2&A2)', long)).toBe(32767);
    expect(ev('=LEN(A2&A2&A2)', long)).toBe(32767);
    expect(ev('=LEN(REPT("a",32767)&"")')).toBe(32767);
    expect(ev('=RIGHT(REPT("a",32767)&"b")')).toBe('a');
    expect(ev('=MID(A2&REPT("b",5000),32767,1)', long)).toBe('b');
  });

  it('CONCATENATE keeps the first 32767 characters and still propagates errors', () => {
    expect(ev('=LEN(CONCATENATE(A2,A2,A2))', long)).toBe(32767);
    expect(ev('=RIGHT(CONCATENATE(REPT("a",32767),"b"))')).toBe('a');
    expect(ev('=CONCATENATE(REPT("a",32767),"b",#N/A)')).toBe('#N/A');
  });

  it('SUBSTITUTE over 32767 characters is #VALUE!', () => {
    expect(ev('=LEN(SUBSTITUTE(REPT("a",32767),"a",REPT("b",16000)))')).toBe('#VALUE!');
    expect(ev('=LEN(SUBSTITUTE(REPT("a",16384),"a","bb"))')).toBe('#VALUE!');
    expect(ev('=LEN(SUBSTITUTE(REPT("a",16383),"a","bb"))')).toBe(32766);
    expect(ev('=LEN(SUBSTITUTE(REPT("ab",10000),"a","xyz"))')).toBe('#VALUE!');
    expect(ev('=LEN(SUBSTITUTE(REPT("a",32767),"a",""))')).toBe(0);
  });

  it('SUBSTITUTE with instance_num over 32767 characters is #VALUE!', () => {
    expect(ev('=LEN(SUBSTITUTE(REPT("a",32767),"a","bb",1))')).toBe('#VALUE!');
    expect(ev('=LEN(SUBSTITUTE(REPT("a",32766),"a","bb",1))')).toBe(32767);
  });
});

// ---------------------------------------------------------------------------
// Collation
// ---------------------------------------------------------------------------

describe('collation', () => {
  it('compares text with a fixed en-US collator', () => {
    expect(compareText('ä', 'z')).toBeLessThan(0);
    expect(ev('="ä"<"z"')).toBe(true);
    expect(ev('="a"="A"')).toBe(true);
    expect(ev('="a"="á"')).toBe(false);
  });
});
