import { afterEach, describe, expect, it, vi } from 'vitest';
import { INVALID_ARGUMENT, Workbook } from '../src/index.js';

const UUID_V4 = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

/** Runs `fn` and returns the `code` of the error it throws (fails when it does not throw). */
function thrownCode(fn: () => unknown): unknown {
  try {
    fn();
  } catch (err) {
    return (err as { code?: unknown }).code;
  }
  throw new Error('expected the call to throw');
}

describe('threaded comments outside a secure context (no crypto.randomUUID)', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('addThreadedComment, replyToComment and new persons get v4 UUIDs from getRandomValues', () => {
    const real = globalThis.crypto;
    vi.stubGlobal('crypto', { getRandomValues: real.getRandomValues.bind(real) });

    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'Looks good!', 'Alice');
    const reply = ws.replyToComment(id, 'Thanks!', 'Bob');

    expect(id).toMatch(UUID_V4);
    expect(reply).toMatch(UUID_V4);
    expect(reply).not.toBe(id);
    expect(ws.threadedComments.map((c) => c.parentId)).toEqual([undefined, id]);
    expect(wb.toJSON().persons?.map((p) => p.id)).toEqual([
      expect.stringMatching(UUID_V4),
      expect.stringMatching(UUID_V4),
    ]);
  });
});

describe('row and column index bounds', () => {
  for (const bad of [0, -1, 1.5, 1_048_577]) {
    it(`setRowHeight(${bad}) and setRowHidden(${bad}) throw INVALID_ARGUMENT and add no row`, () => {
      const ws = new Workbook().addSheet('S');
      expect(thrownCode(() => ws.setRowHeight(bad, 20))).toBe(INVALID_ARGUMENT);
      expect(thrownCode(() => ws.setRowHidden(bad, true))).toBe(INVALID_ARGUMENT);
      expect(ws.rows).toEqual([]);
    });
  }

  it('accepts the first and last rows', () => {
    const ws = new Workbook().addSheet('S');
    ws.setRowHeight(1, 20);
    ws.setRowHidden(1_048_576, true);
    expect(ws.rows.map((r) => r.index)).toEqual([1, 1_048_576]);
  });

  for (const bad of [0, 16_385]) {
    it(`setColumnWidth(${bad}) throws INVALID_ARGUMENT and adds no column`, () => {
      const ws = new Workbook().addSheet('S');
      expect(thrownCode(() => ws.setColumnWidth(bad, 12))).toBe(INVALID_ARGUMENT);
      expect(ws.columns).toEqual([]);
    });
  }

  it('accepts the first and last columns', () => {
    const ws = new Workbook().addSheet('S');
    ws.setColumnWidth(1, 12);
    ws.setColumnWidth(16_384, 12);
    expect(ws.columns.map((c) => [c.min, c.max])).toEqual([
      [1, 1],
      [16_384, 16_384],
    ]);
  });

  it('groupRows(1048576, 1048577) throws before touching row 1048576', () => {
    const ws = new Workbook().addSheet('S');
    expect(thrownCode(() => ws.groupRows(1_048_576, 1_048_577))).toBe(INVALID_ARGUMENT);
    expect(ws.rows).toEqual([]);
  });

  it('groupRows(0, 2) throws before touching any row', () => {
    const ws = new Workbook().addSheet('S');
    expect(thrownCode(() => ws.groupRows(0, 2))).toBe(INVALID_ARGUMENT);
    expect(ws.rows).toEqual([]);
  });

  it('groupColumns(16384, 16385) throws before touching column 16384', () => {
    const ws = new Workbook().addSheet('S');
    expect(thrownCode(() => ws.groupColumns(16_384, 16_385))).toBe(INVALID_ARGUMENT);
    expect(ws.columns).toEqual([]);
  });

  it('collapseRows(1048570, 1048576) does not throw and marks no row past the sheet', () => {
    const ws = new Workbook().addSheet('S');
    ws.groupRows(1_048_570, 1_048_576);
    expect(() => ws.collapseRows(1_048_570, 1_048_576)).not.toThrow();
    expect(ws.rows.at(-1)?.index).toBe(1_048_576);
    expect(ws.rows.every((r) => r.hidden)).toBe(true);
  });
});

describe('print area and print titles sheet index', () => {
  it('setPrintArea and setPrintTitles throw INVALID_ARGUMENT for an index naming no sheet', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    expect(thrownCode(() => wb.setPrintArea(5, 'S!$A$1:$B$2'))).toBe(INVALID_ARGUMENT);
    expect(thrownCode(() => wb.setPrintTitles(5, 'S!$1:$1'))).toBe(INVALID_ARGUMENT);
    expect(thrownCode(() => wb.setPrintArea(-1, 'S!$A$1:$B$2'))).toBe(INVALID_ARGUMENT);
    expect(wb.namedRanges).toEqual([]);
  });

  it('still sets the print area of an existing sheet', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    wb.setPrintArea(0, 'S!$A$1:$B$2');
    expect(wb.getPrintArea(0)).toBe('S!$A$1:$B$2');
  });
});

describe('sheet names with control characters', () => {
  for (const name of ['a\nb', 'a\u0001b']) {
    it(`addSheet(${JSON.stringify(name)}) throws INVALID_ARGUMENT`, () => {
      const wb = new Workbook();
      expect(thrownCode(() => wb.addSheet(name))).toBe(INVALID_ARGUMENT);
      expect(wb.sheetNames).toEqual([]);
    });
  }

  it("renameSheet(existing, 'x\\u009Fy') throws INVALID_ARGUMENT", () => {
    const wb = new Workbook();
    wb.addSheet('Data');
    expect(thrownCode(() => wb.renameSheet('Data', 'x\u009Fy'))).toBe(INVALID_ARGUMENT);
    expect(wb.sheetNames).toEqual(['Data']);
  });

  it("the Worksheet name setter rejects 'a\\tb'", () => {
    const ws = new Workbook().addSheet('Data');
    expect(
      thrownCode(() => {
        ws.name = 'a\tb';
      }),
    ).toBe(INVALID_ARGUMENT);
    expect(ws.name).toBe('Data');
  });

  it("accepts 'a b'", () => {
    const wb = new Workbook();
    wb.addSheet('a b');
    expect(wb.sheetNames).toEqual(['a b']);
  });
});
