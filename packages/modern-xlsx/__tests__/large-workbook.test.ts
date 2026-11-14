import { assert, describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { fillGrid } from './helpers/fixtures.js';

describe('Large workbook round-trips', () => {
  it.for([1_000, 10_000, 100_000])(
    'writes and reads %i rows x 5 columns',
    { timeout: 60_000 },
    async (rows) => {
      const wb = new Workbook();
      fillGrid(wb.addSheet('Data'), rows, 5);

      const buffer = await wb.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      const ws2 = (await readBuffer(buffer)).getSheet('Data');
      expect(ws2).toBeDefined();
      expect(ws2?.rows).toHaveLength(rows);

      const mid = rows / 2;
      expect(ws2?.cell('A1').value).toBe(1);
      expect(ws2?.cell('E1').value).toBe(5);
      expect(ws2?.cell(`A${mid}`).value).toBe(mid);
      expect(ws2?.cell(`C${mid}`).value).toBe(3 * mid);
      expect(ws2?.cell(`A${rows}`).value).toBe(rows);
      expect(ws2?.cell(`E${rows}`).value).toBe(5 * rows);
    },
  );

  it('writes and reads 10K rows with mixed types', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Mixed');
    const ROWS = 10_000;

    for (let r = 1; r <= ROWS; r++) {
      // Column A: number
      ws.cell(`A${r}`).value = r * 3.14;
      // Column B: string (shared string)
      ws.cell(`B${r}`).value = `Row ${r} text`;
      // Column C: boolean
      ws.cell(`C${r}`).value = r % 2 === 0;
      // Column D: formula
      ws.cell(`D${r}`).formula = `A${r}*2`;
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Mixed');
    expect(ws2).toBeDefined();
    expect(ws2?.rows).toHaveLength(ROWS);

    // Check first row
    expect(ws2?.cell('A1').value).toBeCloseTo(3.14);
    expect(ws2?.cell('A1').type).toBe('number');
    expect(ws2?.cell('B1').value).toBe('Row 1 text');
    expect(ws2?.cell('B1').type).toBe('sharedString');
    expect(ws2?.cell('C1').value).toBe(false);
    expect(ws2?.cell('C1').type).toBe('boolean');
    expect(ws2?.cell('D1').formula).toBe('A1*2');
    expect(ws2?.cell('D1').type).toBe('number');

    // Check middle row (row 5000)
    expect(ws2?.cell('A5000').value).toBeCloseTo(5000 * 3.14);
    expect(ws2?.cell('B5000').value).toBe('Row 5000 text');
    expect(ws2?.cell('C5000').value).toBe(true); // 5000 is even
    expect(ws2?.cell('D5000').formula).toBe('A5000*2');

    // Check last row
    expect(ws2?.cell('A10000').value).toBeCloseTo(10_000 * 3.14);
    expect(ws2?.cell('B10000').value).toBe('Row 10000 text');
    expect(ws2?.cell('C10000').value).toBe(true); // 10000 is even
    expect(ws2?.cell('D10000').formula).toBe('A10000*2');
  });

  it('writes and reads 5 sheets each with 1K rows', async () => {
    const wb = new Workbook();
    const SHEETS = 5;
    const ROWS = 1000;
    const COLS = 5;

    for (let s = 0; s < SHEETS; s++) {
      const ws = wb.addSheet(`Sheet${s + 1}`);
      for (let r = 1; r <= ROWS; r++) {
        for (let c = 0; c < COLS; c++) {
          ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = (s + 1) * 1000 + r * (c + 1);
        }
      }
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(SHEETS);
    expect(wb2.sheetNames).toEqual(['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']);

    for (let s = 0; s < SHEETS; s++) {
      const ws2 = wb2.getSheet(`Sheet${s + 1}`);
      expect(ws2).toBeDefined();
      expect(ws2?.rows).toHaveLength(ROWS);

      // Verify first row of each sheet
      expect(ws2?.cell('A1').value).toBe((s + 1) * 1000 + 1);
      expect(ws2?.cell('E1').value).toBe((s + 1) * 1000 + 5);

      // Verify last row of each sheet
      expect(ws2?.cell('A1000').value).toBe((s + 1) * 1000 + 1000);
    }
  });

  it('handles heavy SST deduplication with 10K rows and 100 unique strings', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Dedup');
    const ROWS = 10_000;
    const UNIQUE_COUNT = 100;
    const uniqueStrings = Array.from({ length: UNIQUE_COUNT }, (_, i) => `UniqueValue_${i}`);

    for (let r = 1; r <= ROWS; r++) {
      // Column A: repeating string from pool of 100
      const str = uniqueStrings[r % UNIQUE_COUNT];
      if (!str) throw new Error(`uniqueStrings[${r % UNIQUE_COUNT}] not found`);
      ws.cell(`A${r}`).value = str;
      // Column B: row number for verification
      ws.cell(`B${r}`).value = r;
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Dedup');
    expect(ws2).toBeDefined();
    expect(ws2?.rows).toHaveLength(ROWS);

    // Verify values are correctly deduplicated and resolved
    for (const sampleRow of [1, 50, 100, 500, 1000, 5000, 9999, 10_000]) {
      const expectedStr = uniqueStrings[sampleRow % UNIQUE_COUNT];
      if (!expectedStr) throw new Error(`uniqueStrings[${sampleRow % UNIQUE_COUNT}] not found`);
      expect(ws2?.cell(`A${sampleRow}`).value).toBe(expectedStr);
      expect(ws2?.cell(`A${sampleRow}`).type).toBe('sharedString');
      expect(ws2?.cell(`B${sampleRow}`).value).toBe(sampleRow);
    }
  });

  it('keeps buffer size growth below 11x for 10x rows', async () => {
    const sizes: { rows: number; bytes: number }[] = [];

    for (const rowCount of [1000, 5000, 10_000]) {
      const wb = new Workbook();
      fillGrid(wb.addSheet('Data'), rowCount, 5);
      const buffer = await wb.toBuffer();
      sizes.push({ rows: rowCount, bytes: buffer.length });
    }

    // Verify sizes are ordered (more rows = larger file)
    const [small, medium, large] = sizes;
    assert.isDefined(small);
    assert.isDefined(medium);
    assert.isDefined(large);
    expect(medium.bytes).toBeGreaterThan(small.bytes);
    expect(large.bytes).toBeGreaterThan(medium.bytes);

    // 10x rows must produce less than 11x bytes: growth is roughly linear
    // Numeric-only data has limited compressibility; the ratio is typically ~10x
    const ratio = large.bytes / small.bytes;
    expect(ratio).toBeLessThan(11);
  });

  it('preserves sparse row indices through roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sparse');
    const sparseIndices = [1, 100, 1000, 10_000];

    for (const idx of sparseIndices) {
      ws.cell(`A${idx}`).value = idx;
      ws.cell(`B${idx}`).value = `Row ${idx}`;
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sparse');
    expect(ws2).toBeDefined();

    // Should only have the 4 rows we created
    expect(ws2?.rows).toHaveLength(sparseIndices.length);

    // Verify all sparse row values survived
    for (const idx of sparseIndices) {
      expect(ws2?.cell(`A${idx}`).value).toBe(idx);
      expect(ws2?.cell(`B${idx}`).value).toBe(`Row ${idx}`);
    }

    // Verify row indices are correct
    const rowIndices = ws2?.rows.map((r) => r.index);
    expect(rowIndices).toEqual(sparseIndices);
  });
});
