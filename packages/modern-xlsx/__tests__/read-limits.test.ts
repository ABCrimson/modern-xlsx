import { assert, describe, expect, it } from 'vitest';
import type { ReadLimits, Worksheet } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { serializeReadLimits } from '../src/read-limits.js';

/** Returns a sheet that must exist, narrowing away `undefined`. */
function sheetOf(wb: Workbook, name: string): Worksheet {
  const ws = wb.getSheet(name);
  assert.isDefined(ws);
  return ws;
}

async function smallWorkbook(rows: number): Promise<Uint8Array> {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  for (let r = 1; r <= rows; r++) {
    ws.cell(`A${r}`).value = `row ${r} text that makes the JSON longer`;
    ws.cell(`B${r}`).value = r;
  }
  return wb.toBuffer();
}

describe('ReadOptions.limits', () => {
  it('rejects with SIZE_LIMIT when the JSON exceeds maxJsonLength', async () => {
    const buf = await smallWorkbook(50);
    await expect(readBuffer(buf, { limits: { maxJsonLength: 1024 } })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'SIZE_LIMIT',
    });
    const wb = await readBuffer(buf);
    expect(sheetOf(wb, 'Data').cell('B50').value).toBe(50);
  });

  it('rejects with SIZE_LIMIT when one shared string expands past maxJsonLength', async () => {
    // The writer stores the string once; the read copies it into all 200 cells (200,000 bytes).
    const text = 'x'.repeat(1_000);
    const wb = new Workbook();
    const ws = wb.addSheet('Shared');
    for (let r = 1; r <= 200; r++) ws.cell(`A${r}`).value = text;
    const buf = await wb.toBuffer();
    expect(buf.length).toBeLessThan(10_000);

    await expect(readBuffer(buf, { limits: { maxJsonLength: 100_000 } })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'SIZE_LIMIT',
    });
    const back = await readBuffer(buf);
    expect(sheetOf(back, 'Shared').cell('A200').value).toBe(text);
  });

  it('rejects with SIZE_LIMIT when one comment author expands past maxJsonLength', async () => {
    // One author stored once, copied into all 200 notes (200,000 bytes) before any JSON is written.
    const author = 'a'.repeat(1_000);
    const wb = new Workbook();
    const ws = wb.addSheet('Notes');
    for (let r = 1; r <= 200; r++) ws.addComment(`A${r}`, author, 'n');
    const buf = await wb.toBuffer();

    const err = await readBuffer(buf, { limits: { maxJsonLength: 100_000 } }).catch(
      (e: unknown) => e,
    );
    expect(err).toMatchObject({ name: 'ModernXlsxError', code: 'SIZE_LIMIT' });
    expect(String((err as Error).message)).toMatch(
      /shared text expands past the 100000-byte limit/,
    );
    const back = await readBuffer(buf);
    expect(sheetOf(back, 'Notes').comments).toHaveLength(200);
  });

  it('rejects a whitespace-padded cell by ratio unless maxCompressionRatio is raised', async () => {
    const wb = new Workbook();
    wb.addSheet('Pad').cell('A1').value = ' '.repeat(2_000_000);
    const buf = await wb.toBuffer();

    await expect(readBuffer(buf)).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'SECURITY',
    });

    const back = await readBuffer(buf, { limits: { maxCompressionRatio: 100000 } });
    expect(String(sheetOf(back, 'Pad').cell('A1').value).length).toBe(2_000_000);
  });

  it('rejects maxEntries and maxDecompressedSize set below the archive', async () => {
    const buf = await smallWorkbook(3);
    await expect(readBuffer(buf, { limits: { maxEntries: 2 } })).rejects.toMatchObject({
      code: 'SECURITY',
    });
    await expect(readBuffer(buf, { limits: { maxDecompressedSize: 100 } })).rejects.toMatchObject({
      code: 'SECURITY',
    });
  });

  it('accepts values above a ceiling by clamping them', async () => {
    const buf = await smallWorkbook(3);
    const wb = await readBuffer(buf, {
      limits: {
        maxDecompressedSize: Number.MAX_SAFE_INTEGER,
        maxEntries: Number.MAX_SAFE_INTEGER,
        maxJsonLength: Number.MAX_SAFE_INTEGER,
      },
    });
    expect(sheetOf(wb, 'Data').cell('B3').value).toBe(3);
  });

  it.each<[string, unknown]>([
    ['maxEntries', -1],
    ['maxJsonLength', 1.5],
    ['maxCompressionRatio', Number.NaN],
    ['maxDecompressedSize', 0],
    ['maxEntries', '5'],
    ['bogus', 1],
  ])('rejects invalid limits.%s = %s with INVALID_ARGUMENT', async (key, value) => {
    const buf = await smallWorkbook(1);
    const limits = { [key]: value } as ReadLimits;
    await expect(readBuffer(buf, { limits })).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'INVALID_ARGUMENT',
    });
  });

  it('enforces maxJsonLength on the password path', async () => {
    const wb = new Workbook();
    wb.addSheet('Secret').cell('A1').value = 'hello';
    const enc = await wb.toBuffer({ password: 'pw' });
    await expect(
      readBuffer(enc, { password: 'pw', limits: { maxJsonLength: 64 } }),
    ).rejects.toMatchObject({ name: 'ModernXlsxError', code: 'SIZE_LIMIT' });
    const back = await readBuffer(enc, { password: 'pw' });
    expect(sheetOf(back, 'Secret').cell('A1').value).toBe('hello');
  });
});

describe('serializeReadLimits', () => {
  it('returns undefined for undefined and an empty object for {}', () => {
    expect(serializeReadLimits(undefined)).toBeUndefined();
    expect(serializeReadLimits({})).toBe('{}');
  });

  it('serializes only the provided fields and skips undefined ones', () => {
    expect(serializeReadLimits({ maxEntries: 5, maxCompressionRatio: 2.5 })).toBe(
      '{"maxEntries":5,"maxCompressionRatio":2.5}',
    );
    // A JavaScript caller may pass undefined explicitly; the type forbids it, so go through Reflect.apply.
    expect(Reflect.apply(serializeReadLimits, undefined, [{ maxEntries: undefined }])).toBe('{}');
  });

  it('keeps sizes above 2^32 exact', () => {
    expect(serializeReadLimits({ maxDecompressedSize: 2 ** 40 })).toBe(
      '{"maxDecompressedSize":1099511627776}',
    );
  });

  it('rejects Infinity and unsafe integers', () => {
    expect(() => serializeReadLimits({ maxCompressionRatio: Number.POSITIVE_INFINITY })).toThrow(
      'Invalid ReadOptions.limits.maxCompressionRatio: Infinity',
    );
    expect(() => serializeReadLimits({ maxEntries: 2 ** 60 })).toThrow(/maxEntries/);
  });
});
