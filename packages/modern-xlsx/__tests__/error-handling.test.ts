import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('0.1.5 — Error Handling & Edge Cases', () => {
  // 1. Empty Uint8Array throws
  it('rejects empty Uint8Array', async () => {
    await expect(readBuffer(new Uint8Array(0))).rejects.toThrow();
  });

  // 2. Random bytes throw
  it('rejects random bytes gracefully', async () => {
    const garbage = new Uint8Array(1024);
    crypto.getRandomValues(garbage);
    await expect(readBuffer(garbage)).rejects.toThrow();
  });

  // 3. Truncated ZIP throws
  it('rejects truncated ZIP', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');
    const buffer = await wb.toBuffer();
    const truncated = buffer.slice(0, Math.floor(buffer.length / 2));
    await expect(readBuffer(truncated)).rejects.toThrow();
  });

  // 4. OLE2 magic bytes detection
  it('detects OLE2 compound documents (encrypted/legacy)', async () => {
    const ole2 = new Uint8Array([
      0xd0,
      0xcf,
      0x11,
      0xe0,
      0xa1,
      0xb1,
      0x1a,
      0xe1,
      ...new Uint8Array(100),
    ]);
    await expect(readBuffer(ole2)).rejects.toThrow();
  });

  // 5. Empty workbook roundtrip
  it('roundtrips empty workbook (single empty sheet)', async () => {
    const wb = new Workbook();
    wb.addSheet('Empty');
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(1);
    expect(wb2.sheetNames).toEqual(['Empty']);
    const ws2 = wb2.getSheet('Empty');
    expect(ws2?.rows).toHaveLength(0);
  });

  // 6. Empty sheet preserves metadata (merge cells only)
  it('roundtrips sheet with only merge cells', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('MergeOnly');
    ws.addMergeCell('A1:C3');
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('MergeOnly');
    expect(ws2?.mergeCells).toContain('A1:C3');
  });

  // 7. Sheet with only frozen pane
  it('roundtrips sheet with only frozen pane', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Frozen');
    ws.frozenPane = { rows: 2, cols: 1 };
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Frozen');
    expect(ws2?.frozenPane).toEqual({ rows: 2, cols: 1 });
  });

  // 8. Sheet with only auto-filter
  it('roundtrips sheet with only auto-filter', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Filtered');
    ws.autoFilter = 'A1:D10';
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Filtered');
    expect(ws2?.autoFilter).toEqual({ range: 'A1:D10' });
  });

  // 9. Sheet with only validation
  it('roundtrips sheet with only data validation', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('ValidOnly');
    ws.addValidation('A1:A100', {
      validationType: 'list',
      operator: null,
      formula1: '"Yes,No,Maybe"',
      formula2: null,
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: null,
      errorMessage: null,
    });
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('ValidOnly');
    expect(ws2?.validations).toHaveLength(1);
    expect(ws2?.validations[0]?.sqref).toBe('A1:A100');
    expect(ws2?.validations[0]?.validationType).toBe('list');
    expect(ws2?.validations[0]?.formula1).toBe('"Yes,No,Maybe"');
  });

  // 10. Minimum valid XLSX: one sheet, one cell
  it('roundtrips minimum valid XLSX (one cell)', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Tiny');
    ws.cell('A1').value = 1;
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(1);
    expect(wb2.getSheet('Tiny')?.cell('A1').value).toBe(1);
  });

  // 11. Double roundtrip
  it('survives double roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Double');
    ws.cell('A1').value = 42;
    ws.cell('B1').value = 'hello';
    ws.cell('C1').value = true;
    // First roundtrip
    const buf1 = await wb.toBuffer();
    const wb2 = await readBuffer(buf1);
    // Second roundtrip
    const buf2 = await wb2.toBuffer();
    const wb3 = await readBuffer(buf2);
    expect(wb3.getSheet('Double')?.cell('A1').value).toBe(42);
    expect(wb3.getSheet('Double')?.cell('B1').value).toBe('hello');
    expect(wb3.getSheet('Double')?.cell('C1').value).toBe(true);
  });

  // 12. Concurrent reads from same buffer
  it('handles concurrent reads from same buffer', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'concurrent';
    const buffer = await wb.toBuffer();
    const results = await Promise.all([readBuffer(buffer), readBuffer(buffer), readBuffer(buffer)]);
    for (const r of results) {
      expect(r.getSheet('Sheet1')?.cell('A1').value).toBe('concurrent');
    }
  });
});
