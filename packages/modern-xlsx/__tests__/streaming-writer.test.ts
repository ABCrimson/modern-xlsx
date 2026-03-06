import { describe, expect, it } from 'vitest';
import { readBuffer, StreamingXlsxWriter } from '../src/index.js';
import type { StreamingCellInput } from '../src/streaming-writer.js';

describe('StreamingXlsxWriter', () => {
  // -----------------------------------------------------------------------
  // 1. Basic: write 1K rows and verify output is valid XLSX
  // -----------------------------------------------------------------------
  it('writes 1K rows and produces a valid XLSX', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Data');
    const ROWS = 1000;
    const COLS = 5;

    for (let r = 0; r < ROWS; r++) {
      const cells: StreamingCellInput[] = [];
      for (let c = 0; c < COLS; c++) {
        cells.push({
          value: String(r * COLS + c),
          cellType: 'number',
        });
      }
      writer.writeRow(cells);
    }

    const xlsx = writer.finish();
    expect(xlsx).toBeInstanceOf(Uint8Array);
    expect(xlsx.length).toBeGreaterThan(0);

    // Read back and verify
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Data');
    expect(ws).toBeDefined();
    expect(ws?.rows).toHaveLength(ROWS);
    // First row, first cell
    expect(ws?.cell('A1').value).toBe(0);
    // Last row, last cell
    expect(ws?.cell('E1000').value).toBe(999 * 5 + 4);
  });

  // -----------------------------------------------------------------------
  // 2. Large: write 100K rows without OOM
  // -----------------------------------------------------------------------
  it('writes 100K rows without OOM', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('BigData');
    const ROWS = 100_000;

    for (let r = 0; r < ROWS; r++) {
      writer.writeRow([
        { value: String(r), cellType: 'number' },
        { value: `row_${r}`, cellType: 'sharedString' },
      ]);
    }

    const xlsx = writer.finish();
    expect(xlsx).toBeInstanceOf(Uint8Array);
    // The output should be non-trivial in size
    expect(xlsx.length).toBeGreaterThan(1_000_000);
  });

  // -----------------------------------------------------------------------
  // 3. Multi-sheet streaming
  // -----------------------------------------------------------------------
  it('writes multiple sheets', async () => {
    const writer = StreamingXlsxWriter.create();

    writer.startSheet('First');
    writer.writeRow([{ value: 'hello', cellType: 'sharedString' }]);
    writer.writeRow([{ value: '42', cellType: 'number' }]);

    writer.startSheet('Second');
    writer.writeRow([{ value: 'world', cellType: 'sharedString' }]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);

    expect(wb.sheetNames).toEqual(['First', 'Second']);

    const s1 = wb.getSheet('First');
    expect(s1?.rows).toHaveLength(2);
    expect(s1?.cell('A1').value).toBe('hello');
    expect(s1?.cell('A2').value).toBe(42);

    const s2 = wb.getSheet('Second');
    expect(s2?.rows).toHaveLength(1);
    expect(s2?.cell('A1').value).toBe('world');
  });

  // -----------------------------------------------------------------------
  // 4. SST deduplication
  // -----------------------------------------------------------------------
  it('builds SST correctly with deduplication', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('SST');

    // Write duplicate strings
    writer.writeRow([
      { value: 'Alpha', cellType: 'sharedString' },
      { value: 'Beta', cellType: 'sharedString' },
    ]);
    writer.writeRow([
      { value: 'Alpha', cellType: 'sharedString' },
      { value: 'Gamma', cellType: 'sharedString' },
    ]);
    writer.writeRow([
      { value: 'Beta', cellType: 'sharedString' },
      { value: 'Beta', cellType: 'sharedString' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);

    // Only 3 unique strings in the SST
    expect(wb.toJSON().sharedStrings?.strings).toHaveLength(3);

    // Verify cell values
    const ws = wb.getSheet('SST');
    expect(ws).toBeDefined();
    expect(ws?.cell('A1').value).toBe('Alpha');
    expect(ws?.cell('B1').value).toBe('Beta');
    expect(ws?.cell('A2').value).toBe('Alpha');
    expect(ws?.cell('B2').value).toBe('Gamma');
    expect(ws?.cell('A3').value).toBe('Beta');
    expect(ws?.cell('B3').value).toBe('Beta');
  });

  // -----------------------------------------------------------------------
  // 5. Mixed cell types
  // -----------------------------------------------------------------------
  it('supports mixed cell types in a row', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Mixed');

    writer.writeRow([
      { value: '3.14', cellType: 'number' },
      { value: 'text', cellType: 'sharedString' },
      { value: 'true', cellType: 'boolean' },
      { value: 'inline text', cellType: 'inlineStr' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Mixed');
    expect(ws).toBeDefined();

    expect(ws?.cell('A1').value).toBe(3.14);
    expect(ws?.cell('B1').value).toBe('text');
    expect(ws?.cell('C1').value).toBe(true);
    expect(ws?.cell('D1').value).toBe('inline text');
  });

  // -----------------------------------------------------------------------
  // 6. Empty cells are skipped (sparse rows)
  // -----------------------------------------------------------------------
  it('skips empty cells', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Sparse');

    writer.writeRow([
      { value: 'A', cellType: 'sharedString' },
      { value: undefined },
      { value: 'C', cellType: 'sharedString' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Sparse');
    expect(ws).toBeDefined();

    expect(ws?.cell('A1').value).toBe('A');
    expect(ws?.cell('B1').value).toBeNull();
    expect(ws?.cell('C1').value).toBe('C');
  });

  // -----------------------------------------------------------------------
  // 7. Error: finish without sheets
  // -----------------------------------------------------------------------
  it('throws when finishing without any sheets', () => {
    const writer = StreamingXlsxWriter.create();
    expect(() => writer.finish()).toThrow();
  });

  // -----------------------------------------------------------------------
  // 8. Error: use after finish
  // -----------------------------------------------------------------------
  it('throws when using writer after finish', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S1');
    writer.writeRow([{ value: '1', cellType: 'number' }]);
    writer.finish();

    expect(() => writer.startSheet('S2')).toThrow();
    expect(() => writer.writeRow([{ value: '2' }])).toThrow();
    expect(() => writer.finish()).toThrow();
  });

  // -----------------------------------------------------------------------
  // 9. Wide columns (beyond Z -> AA, AB, ...)
  // -----------------------------------------------------------------------
  it('handles wide rows (30+ columns)', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Wide');

    const cells: StreamingCellInput[] = [];
    for (let c = 0; c < 30; c++) {
      cells.push({ value: String(c), cellType: 'number' });
    }
    writer.writeRow(cells);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Wide');
    expect(ws).toBeDefined();

    // Column 0 = A, column 25 = Z, column 26 = AA, column 27 = AB, etc.
    expect(ws?.cell('A1').value).toBe(0);
    expect(ws?.cell('Z1').value).toBe(25);
    expect(ws?.cell('AA1').value).toBe(26);
    expect(ws?.cell('AD1').value).toBe(29);
  });

  // -----------------------------------------------------------------------
  // 10. Performance: 100K rows with 5 columns
  // -----------------------------------------------------------------------
  it('writes 100K x 5 cols in reasonable time', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Perf');
    const ROWS = 100_000;

    const start = performance.now();
    for (let r = 0; r < ROWS; r++) {
      writer.writeRow([
        { value: String(r), cellType: 'number' },
        { value: String(r * 2), cellType: 'number' },
        { value: String(r * 3), cellType: 'number' },
        { value: `name_${r % 100}`, cellType: 'sharedString' },
        { value: r % 2 === 0 ? 'true' : 'false', cellType: 'boolean' },
      ]);
    }
    const xlsx = writer.finish();
    const elapsed = performance.now() - start;

    expect(xlsx.length).toBeGreaterThan(0);
    // Should complete within 30 seconds even in debug/CI
    expect(elapsed).toBeLessThan(30_000);
  });
});
