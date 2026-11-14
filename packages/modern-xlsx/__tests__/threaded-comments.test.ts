import { assert, describe, expect, it } from 'vitest';
import { INVALID_ARGUMENT, readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

/** The error `fn` throws, or `undefined` when it returns. */
function thrown(fn: () => unknown): unknown {
  try {
    fn();
  } catch (error) {
    return error;
  }
  return undefined;
}

/** Longest threaded comment text Excel 16 opens, in UTF-16 code units (Excel 16 probed 2026-11-12). */
const MAX_TEXT = 32_767;

/** One astral character: one code point, two UTF-16 code units. */
const ASTRAL = '\u{1F600}';

describe('Threaded Comments', () => {
  it('empty sheet has no threaded comments', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.threadedComments).toHaveLength(0);
  });

  it('addThreadedComment creates comment', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'Hello', 'John');
    expect(id).toBeTruthy();
    expect(ws.threadedComments).toHaveLength(1);
    expect(ws.threadedComments[0]?.text).toBe('Hello');
    expect(ws.threadedComments[0]?.refCell).toBe('A1');
  });

  it('replyToComment creates reply chain', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id1 = ws.addThreadedComment('A1', 'Original', 'Alice');
    const id2 = ws.replyToComment(id1, 'Reply 1', 'Bob');
    const id3 = ws.replyToComment(id1, 'Reply 2', 'Alice');
    expect(ws.threadedComments).toHaveLength(3);
    expect(ws.threadedComments[1]?.parentId).toBe(id1);
    expect(ws.threadedComments[2]?.parentId).toBe(id1);
    // Replies should use the same cell ref as parent
    expect(ws.threadedComments[1]?.refCell).toBe('A1');
    expect(ws.threadedComments[2]?.refCell).toBe('A1');
    // IDs should be unique
    expect(id1).not.toBe(id2);
    expect(id2).not.toBe(id3);
  });

  it('shared person across comments', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'First', 'Alice');
    ws.addThreadedComment('B1', 'Second', 'Alice');
    // Same author should reuse person ID
    expect(ws.threadedComments[0]?.personId).toBe(ws.threadedComments[1]?.personId);
  });

  it('replyToComment throws for invalid commentId', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(() => ws.replyToComment('nonexistent', 'text', 'author')).toThrow(
      'Comment nonexistent not found',
    );
  });

  it('different authors create different persons', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'First', 'Alice');
    ws.addThreadedComment('B1', 'Second', 'Bob');
    expect(ws.threadedComments[0]?.personId).not.toBe(ws.threadedComments[1]?.personId);
  });

  it('persons are shared across sheets', () => {
    const wb = new Workbook();
    const ws1 = wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws1.addThreadedComment('A1', 'Hello from 1', 'Alice');
    ws2.addThreadedComment('A1', 'Hello from 2', 'Alice');
    // Same person should be used across sheets
    expect(ws1.threadedComments[0]?.personId).toBe(ws2.threadedComments[0]?.personId);
  });

  it('comment has timestamp', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'Hello', 'John');
    const timestamp = ws.threadedComments[0]?.timestamp;
    expect(timestamp).toBeTruthy();
    assert.isDefined(timestamp);
    // Must be a canonical ISO 8601 string: it survives a Date round trip unchanged.
    expect(new Date(timestamp).toISOString()).toBe(timestamp);
  });

  it('threaded comments and persons survive a write/read round trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'Original', 'Alice');
    ws.replyToComment(id, 'Reply', 'Bob');

    const wb2 = await readBuffer(await wb.toBuffer());
    const ws2 = wb2.getSheet('Sheet1');
    assert.isDefined(ws2);
    // The timestamp is written in Excel's form (UTC, hundredths, no `Z`).
    expect(ws2.threadedComments).toEqual(
      ws.threadedComments.map((c) => ({ ...c, timestamp: excelForm(c.timestamp) })),
    );
    expect(wb2.toJSON().persons).toEqual(wb.toJSON().persons);
  });
});

/** `dT` in the form Excel writes: UTC, two fractional digits (hundredths), no `Z`. */
const EXCEL_DT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{2}$/;

/** An ISO 8601 `…Z` timestamp in Excel's `dT` form (truncated to hundredths). */
function excelForm(iso: string): string {
  return iso.replace(/^(.*\.\d{2})\d*Z$/, '$1');
}

/** Milliseconds since the epoch of a model timestamp; one without a zone is UTC, as Excel writes it. */
function stampMs(timestamp: string): number {
  return Date.parse(/Z$|[+-]\d{2}:\d{2}$/.test(timestamp) ? timestamp : `${timestamp}Z`);
}

/** The `dT` and text of each threaded comment the written sheet 1 holds, in element order. */
async function writtenThreads(wb: Workbook): Promise<{ dT: string; text: string }[]> {
  const xml = readZipText(await wb.toBuffer())['xl/threadedComments/threadedComment1.xml'];
  assert.isDefined(xml);
  return [...xml.matchAll(/ dT="([^"]*)"[^>]*><text>([^<]*)<\/text>/g)].map((m) => ({
    dT: m[1] ?? '',
    text: m[2] ?? '',
  }));
}

describe('replies keep the order they were added in (Excel orders a thread by dT)', () => {
  const TEXTS = Array.from({ length: 10 }, (_, i) => `reply${i + 1}`);

  it('ten replies added in a loop are stamped at least 10 ms apart and written in order', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const root = ws.addThreadedComment('B2', 'root', 'Ann');
    for (const text of TEXTS) ws.replyToComment(root, text, 'Ann');

    const stamps = ws.threadedComments.map((c) => stampMs(c.timestamp));
    for (let i = 1; i < stamps.length; i++) {
      expect((stamps[i] ?? 0) - (stamps[i - 1] ?? 0)).toBeGreaterThanOrEqual(10);
    }
    // The model keeps the ISO 8601 form the API stamps.
    for (const c of ws.threadedComments)
      expect(new Date(c.timestamp).toISOString()).toBe(c.timestamp);

    const written = await writtenThreads(wb);
    expect(written.map((t) => t.text)).toEqual(['root', ...TEXTS]);
    for (const t of written) expect(t.dT).toMatch(EXCEL_DT);
    for (let i = 1; i < written.length; i++) {
      expect((written[i]?.dT ?? '') > (written[i - 1]?.dT ?? '')).toBe(true);
    }
  });

  it('a second generation adds replies after the latest one read, and keeps the read dT values', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const root = ws.addThreadedComment('B2', 'root', 'Ann');
    for (const text of TEXTS) ws.replyToComment(root, text, 'Ann');
    const first = await writtenThreads(wb);

    const read = await readBuffer(await wb.toBuffer());
    const rs = read.getSheet('Sheet1');
    assert.isDefined(rs);
    // A read gives each dT back as the file holds it.
    expect(rs.threadedComments.map((c) => c.timestamp)).toEqual(first.map((t) => t.dT));
    rs.replyToComment(root, 'reply11', 'Bob');
    rs.replyToComment(root, 'reply12', 'Ann');

    const second = await writtenThreads(read);
    expect(second.map((t) => t.text)).toEqual(['root', ...TEXTS, 'reply11', 'reply12']);
    expect(second.slice(0, first.length)).toEqual(first);
    for (let i = 1; i < second.length; i++) {
      expect((second[i]?.dT ?? '') > (second[i - 1]?.dT ?? '')).toBe(true);
    }
  });

  it('a reply to a thread whose latest dT is in the future is stamped 10 ms after it, read as UTC', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    const root = ws.addThreadedComment('A1', 'root', 'Ann');
    const reply = ws.replyToComment(root, 'reply', 'Ann');
    const [r0, r1] = wb.toJSON().sheets[0]?.worksheet.threadedComments ?? [];
    assert.isDefined(r0);
    assert.isDefined(r1);
    r0.timestamp = '2099-01-01T00:00:00.00';
    r1.timestamp = '2099-01-01T00:00:00.50';
    // Read from a file, the thread's dT values are in Excel's form: UTC without a zone.
    const read = await readBuffer(await wb.toBuffer());
    const rs = read.getSheet('S');
    assert.isDefined(rs);
    const id = rs.replyToComment(reply, 'next', 'Ann');
    expect(rs.threadedComments.find((c) => c.id === id)?.timestamp).toBe(
      '2099-01-01T00:00:00.510Z',
    );
    // A new thread on another cell is not held back by this one.
    rs.addThreadedComment('C3', 'other thread', 'Ann');
    expect(stampMs(rs.threadedComments.at(-1)?.timestamp ?? '')).toBeLessThan(
      Date.parse('2099-01-01'),
    );
  });
});

describe('thread lookups stay linear and follow the model array', () => {
  it('20,000 threads on distinct cells and 20,000 replies on one root are quick', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    let t = performance.now();
    for (let i = 1; i <= 20_000; i++) ws.addThreadedComment(`A${i}`, 't', 'Ann');
    const adds = performance.now() - t;
    const root = ws.addThreadedComment('B1', 'root', 'Ann');
    t = performance.now();
    for (let i = 1; i <= 20_000; i++) ws.replyToComment(root, 'r', 'Ann');
    const replies = performance.now() - t;
    // Quadratic scans took 4.4 s for the adds on this host (1.3.0 candidate of 2026-11-12, UTC).
    expect(adds).toBeLessThan(2_000);
    expect(replies).toBeLessThan(2_000);
    expect(ws.threadedComments).toHaveLength(40_001);
  });

  it('a getSheet() handle per call sees the threads added through another handle', () => {
    const wb = new Workbook();
    wb.addSheet('S');
    for (let i = 1; i <= 50; i++) wb.getSheet('S')?.addThreadedComment(`A${i}`, 't', 'Ann');
    expect(thrown(() => wb.getSheet('S')?.addThreadedComment('a7', 't', 'Ann'))).toMatchObject({
      code: INVALID_ARGUMENT,
    });
  });

  it('a thread pushed onto the model array blocks its cell and takes replies', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'root', 'Ann');
    const threads = wb.toJSON().sheets[0]?.worksheet.threadedComments;
    assert.isDefined(threads);
    const [root] = threads;
    assert.isDefined(root);
    const pushed = { ...root, id: '{00000000-0000-4000-8000-0000000000B2}', refCell: 'B2' };
    threads.push(pushed);

    expect(thrown(() => ws.addThreadedComment('$b$2', 'second', 'Ann'))).toMatchObject({
      code: INVALID_ARGUMENT,
    });
    const reply = ws.replyToComment(pushed.id, 'reply', 'Ann');
    expect(ws.threadedComments.find((c) => c.id === reply)?.parentId).toBe(pushed.id);
  });

  it('a thread removed from the model array frees its cell and its id', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const a1 = ws.addThreadedComment('A1', 'root', 'Ann');
    ws.addThreadedComment('B2', 'root', 'Ann');
    const model = wb.toJSON().sheets[0]?.worksheet;
    assert.isDefined(model?.threadedComments);
    // Removed in place, and by a new array.
    model.threadedComments.splice(0, 1);
    expect(() => ws.addThreadedComment('A1', 'again', 'Ann')).not.toThrow();
    model.threadedComments = model.threadedComments.filter((c) => c.refCell !== 'B2');
    expect(() => ws.addThreadedComment('B2', 'again', 'Ann')).not.toThrow();
    expect(thrown(() => ws.replyToComment(a1, 'x', 'Ann'))).toMatchObject({
      code: 'COMMENT_NOT_FOUND',
    });
  });

  it('a same-length in-place replacement is seen by a reply and by the cell it vacates', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const a1 = ws.addThreadedComment('A1', 'root', 'Ann');
    const threads = wb.toJSON().sheets[0]?.worksheet.threadedComments;
    assert.isDefined(threads);
    const [root] = threads;
    assert.isDefined(root);
    // The root moves to C3 under a new id, without changing the array or its length.
    threads[0] = { ...root, id: '{00000000-0000-4000-8000-0000000000C3}', refCell: 'C3' };

    expect(thrown(() => ws.replyToComment(a1, 'x', 'Ann'))).toMatchObject({
      code: 'COMMENT_NOT_FOUND',
    });
    expect(() => ws.addThreadedComment('A1', 'free again', 'Ann')).not.toThrow();
    expect(thrown(() => ws.addThreadedComment('C3', 'taken', 'Ann'))).toMatchObject({
      code: INVALID_ARGUMENT,
    });
    expect(ws.replyToComment('{00000000-0000-4000-8000-0000000000C3}', 'y', 'Ann')).toBeTruthy();
  });

  it('cloneSheet and a read do not share lookups with the source sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Src');
    ws.addThreadedComment('A1', 'root', 'Ann');
    const copy = wb.cloneSheet(0, 'Copy');
    expect(() => copy.addThreadedComment('B1', 'only on the copy', 'Ann')).not.toThrow();
    expect(() => ws.addThreadedComment('B1', 'only on the source', 'Ann')).not.toThrow();
    expect(thrown(() => copy.addThreadedComment('A1', 'taken on the copy', 'Ann'))).toMatchObject({
      code: INVALID_ARGUMENT,
    });

    const read = await readBuffer(await wb.toBuffer());
    const rs = read.getSheet('Src');
    assert.isDefined(rs);
    expect(() => rs.addThreadedComment('C1', 'read copy', 'Ann')).not.toThrow();
    expect(ws.threadedComments.map((c) => c.refCell)).toEqual(['A1', 'B1']);
  });
});

describe('a cell holds one thread', () => {
  it('a second addThreadedComment on a cell throws INVALID_ARGUMENT and changes nothing', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const rootId = ws.addThreadedComment('A1', 'first root', 'Alice');
    const before = structuredClone(wb.toJSON());

    const error = thrown(() => ws.addThreadedComment('A1', 'second root', 'Bob'));
    expect(error).toMatchObject({ name: 'ModernXlsxError', code: INVALID_ARGUMENT });
    const message = (error as Error).message;
    expect(message).toContain('A1');
    expect(message).toContain('replyToComment');
    expect(message).toContain(rootId);
    // No second root and no person for Bob.
    expect(wb.toJSON()).toEqual(before);
    expect(ws.threadedComments.filter((c) => c.parentId === undefined)).toHaveLength(1);
  });

  it.each([
    ['A1', 'a1'],
    ['a1', '$A$1'],
    ['$A$1', 'A1'],
    ['$a1', 'A$1'],
  ])('%j and %j name one cell', (first, second) => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment(first, 'first root', 'Alice');
    expect(thrown(() => ws.addThreadedComment(second, 'second root', 'Alice'))).toMatchObject({
      code: INVALID_ARGUMENT,
    });
    expect(ws.threadedComments).toHaveLength(1);
  });

  it('replies and threads on other cells are not affected', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const a1 = ws.addThreadedComment('A1', 'root', 'Alice');
    ws.replyToComment(a1, 'reply', 'Bob');
    ws.addThreadedComment('B1', 'root on B1', 'Bob');
    wb.addSheet('Sheet2').addThreadedComment('A1', 'root on another sheet', 'Alice');
    expect(ws.threadedComments.map((c) => c.refCell)).toEqual(['A1', 'A1', 'B1']);
    expect(wb.validate().isValid).toBe(true);
  });

  it('validate() reports two thread roots on one cell in a raw model', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addThreadedComment('A1', 'first root', 'Alice');
    ws.addThreadedComment('B2', 'other root', 'Alice');
    expect(wb.validate().isValid).toBe(true);
    // Raw-model callers can point a second root at A1, in another spelling.
    const second = wb.toJSON().sheets[0]?.worksheet.threadedComments?.[1];
    assert.isDefined(second);
    second.refCell = '$a$1';
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    const errors = report.issues.filter((i) => i.severity === 'error');
    expect(errors.map((i) => [i.category, i.location])).toEqual([
      ['structure', 'Sheet1!threadedComments[1]'],
    ]);
    expect(errors[0]?.message).toContain('A1');
  });
});

describe('the threaded comment text limit', () => {
  it(`a root or a reply of ${MAX_TEXT} UTF-16 code units is accepted`, () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'a'.repeat(MAX_TEXT), 'Alice');
    ws.replyToComment(id, 'a'.repeat(MAX_TEXT), 'Bob');
    // 16,383 astral characters and one ASCII: 32,767 code units, 16,384 code points.
    ws.addThreadedComment('B1', `${ASTRAL.repeat(16_383)}a`, 'Alice');
    expect(ws.threadedComments).toHaveLength(3);
    expect(wb.validate().isValid).toBe(true);
  });

  it.each([
    ['ASCII', 'a'.repeat(MAX_TEXT + 1)],
    // 16,384 code points but 32,768 code units: Excel counts code units.
    ['astral', ASTRAL.repeat(16_384)],
  ])(
    'a longer %s text throws INVALID_ARGUMENT naming the limit, for a root and a reply',
    (_, text) => {
      const wb = new Workbook();
      const ws = wb.addSheet('Sheet1');
      const id = ws.addThreadedComment('A1', 'root', 'Alice');
      const before = structuredClone(wb.toJSON());
      for (const fn of [
        () => ws.addThreadedComment('B1', text, 'Bob'),
        () => ws.replyToComment(id, text, 'Bob'),
      ]) {
        const error = thrown(fn);
        expect(error).toMatchObject({ name: 'ModernXlsxError', code: INVALID_ARGUMENT });
        expect((error as Error).message).toContain('32767 UTF-16 code units');
      }
      expect(wb.toJSON()).toEqual(before);
    },
  );

  it('validate() reports a longer text in a raw model', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const id = ws.addThreadedComment('A1', 'root', 'Alice');
    ws.replyToComment(id, 'reply', 'Bob');
    const [root, reply] = wb.toJSON().sheets[0]?.worksheet.threadedComments ?? [];
    assert.isDefined(root);
    assert.isDefined(reply);
    root.text = ASTRAL.repeat(16_384);
    reply.text = 'a'.repeat(MAX_TEXT + 1);
    const errors = wb.validate().issues.filter((i) => i.severity === 'error');
    expect(errors.map((i) => [i.category, i.location])).toEqual([
      ['structure', 'Sheet1!threadedComments[0]'],
      ['structure', 'Sheet1!threadedComments[1]'],
    ]);
    for (const issue of errors) expect(issue.message).toContain('32768 UTF-16 code units');
    root.text = `${ASTRAL.repeat(16_383)}a`;
    reply.text = 'a'.repeat(MAX_TEXT);
    expect(wb.validate().isValid).toBe(true);
  });
});
