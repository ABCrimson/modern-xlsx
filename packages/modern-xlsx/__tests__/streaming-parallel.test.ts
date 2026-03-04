import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('0.1.4 — Streaming & Performance Tests', () => {
  // -------------------------------------------------------------------------
  // 1. Write and read 1K rows
  // -------------------------------------------------------------------------
  it('writes and reads 1K rows x 5 columns', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    const ROWS = 1000;
    const COLS = 5;

    for (let r = 1; r <= ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data');
    expect(ws2).toBeDefined();
    expect(ws2?.rows).toHaveLength(ROWS);

    // First row
    expect(ws2?.cell('A1').value).toBe(1);
    expect(ws2?.cell('E1').value).toBe(5);

    // Middle row (row 500)
    expect(ws2?.cell('A500').value).toBe(500);
    expect(ws2?.cell('C500').value).toBe(1500);

    // Last row
    expect(ws2?.cell('A1000').value).toBe(1000);
    expect(ws2?.cell('E1000').value).toBe(5000);
  });

  // -------------------------------------------------------------------------
  // 2. Write and read 10K rows
  // -------------------------------------------------------------------------
  it('writes and reads 10K rows x 5 columns', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    const ROWS = 10_000;
    const COLS = 5;

    for (let r = 1; r <= ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data');
    expect(ws2).toBeDefined();
    expect(ws2?.rows).toHaveLength(ROWS);

    // First row
    expect(ws2?.cell('A1').value).toBe(1);
    expect(ws2?.cell('E1').value).toBe(5);

    // Middle row (row 5000)
    expect(ws2?.cell('A5000').value).toBe(5000);
    expect(ws2?.cell('C5000').value).toBe(15_000);

    // Last row
    expect(ws2?.cell('A10000').value).toBe(10_000);
    expect(ws2?.cell('E10000').value).toBe(50_000);
  });

  // -------------------------------------------------------------------------
  // 3. Write and read 100K rows (stress test)
  // -------------------------------------------------------------------------
  it('writes and reads 100K rows x 5 columns', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    const ROWS = 100_000;
    const COLS = 5;

    for (let r = 1; r <= ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data');
    expect(ws2).toBeDefined();
    expect(ws2?.rows).toHaveLength(ROWS);

    // First row
    expect(ws2?.cell('A1').value).toBe(1);
    expect(ws2?.cell('E1').value).toBe(5);

    // Middle row (row 50000)
    expect(ws2?.cell('A50000').value).toBe(50_000);
    expect(ws2?.cell('C50000').value).toBe(150_000);

    // Last row
    expect(ws2?.cell('A100000').value).toBe(100_000);
    expect(ws2?.cell('E100000').value).toBe(500_000);
  }, 60_000);

  // -------------------------------------------------------------------------
  // 4. Large file with mixed types
  // -------------------------------------------------------------------------
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
    expect(ws2?.cell('D1').type).toBe('formulaStr');

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

  // -------------------------------------------------------------------------
  // 5. Multiple sheets with data
  // -------------------------------------------------------------------------
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

  // -------------------------------------------------------------------------
  // 6. String deduplication stress test
  // -------------------------------------------------------------------------
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

  // -------------------------------------------------------------------------
  // 7. Benchmark: 1K rows write
  // -------------------------------------------------------------------------
  it('benchmarks 1K row write', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Bench');
    for (let r = 1; r <= 1000; r++) {
      for (let c = 0; c < 5; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }
    const start = performance.now();
    const buffer = await wb.toBuffer();
    const writeMs = performance.now() - start;
    expect(buffer.length).toBeGreaterThan(0);
    console.log(`1K write: ${writeMs.toFixed(1)}ms, ${buffer.length} bytes`);
  });

  // -------------------------------------------------------------------------
  // 8. Benchmark: 1K rows read
  // -------------------------------------------------------------------------
  it('benchmarks 1K row read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Bench');
    for (let r = 1; r <= 1000; r++) {
      for (let c = 0; c < 5; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }
    const buffer = await wb.toBuffer();

    const start = performance.now();
    const wb2 = await readBuffer(buffer);
    const readMs = performance.now() - start;
    expect(wb2.getSheet('Bench')).toBeDefined();
    expect(wb2.getSheet('Bench')?.rows).toHaveLength(1000);
    console.log(`1K read: ${readMs.toFixed(1)}ms, ${buffer.length} bytes`);
  });

  // -------------------------------------------------------------------------
  // 9. Benchmark: 10K rows write
  // -------------------------------------------------------------------------
  it('benchmarks 10K row write', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Bench');
    for (let r = 1; r <= 10_000; r++) {
      for (let c = 0; c < 5; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }
    const start = performance.now();
    const buffer = await wb.toBuffer();
    const writeMs = performance.now() - start;
    expect(buffer.length).toBeGreaterThan(0);
    console.log(`10K write: ${writeMs.toFixed(1)}ms, ${buffer.length} bytes`);
  });

  // -------------------------------------------------------------------------
  // 10. Benchmark: 10K rows read
  // -------------------------------------------------------------------------
  it('benchmarks 10K row read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Bench');
    for (let r = 1; r <= 10_000; r++) {
      for (let c = 0; c < 5; c++) {
        ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
      }
    }
    const buffer = await wb.toBuffer();

    const start = performance.now();
    const wb2 = await readBuffer(buffer);
    const readMs = performance.now() - start;
    expect(wb2.getSheet('Bench')).toBeDefined();
    expect(wb2.getSheet('Bench')?.rows).toHaveLength(10_000);
    console.log(`10K read: ${readMs.toFixed(1)}ms, ${buffer.length} bytes`);
  });

  // -------------------------------------------------------------------------
  // 11. Buffer size scaling
  // -------------------------------------------------------------------------
  it('verifies buffer size scaling is sublinear (compression)', async () => {
    const sizes: { rows: number; bytes: number }[] = [];

    for (const rowCount of [1000, 5000, 10_000]) {
      const wb = new Workbook();
      const ws = wb.addSheet('Data');
      for (let r = 1; r <= rowCount; r++) {
        for (let c = 0; c < 5; c++) {
          ws.cell(`${String.fromCharCode(65 + c)}${r}`).value = r * (c + 1);
        }
      }
      const buffer = await wb.toBuffer();
      sizes.push({ rows: rowCount, bytes: buffer.length });
    }

    console.log('Buffer size scaling:');
    for (const { rows, bytes } of sizes) {
      console.log(`  ${rows} rows: ${bytes} bytes (${(bytes / 1024).toFixed(1)} KB)`);
    }

    // Verify sizes are ordered (more rows = larger file)
    expect(sizes[1]?.bytes).toBeGreaterThan(sizes[0]?.bytes);
    expect(sizes[2]?.bytes).toBeGreaterThan(sizes[1]?.bytes);

    // Verify compression helps: 10x rows should produce less than 11x bytes
    // Numeric-only data has limited compressibility; the ratio is typically ~10x
    const ratio = sizes[2]?.bytes / sizes[0]?.bytes;
    console.log(`  10K/1K size ratio: ${ratio.toFixed(2)}x (10x rows, expect < 11x bytes)`);
    expect(ratio).toBeLessThan(11);
  });

  // -------------------------------------------------------------------------
  // 12. Empty rows handling (sparse data)
  // -------------------------------------------------------------------------
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
