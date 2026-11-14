import { describe, expect, it } from 'vitest';
import {
  INVALID_WORKBOOK,
  JSON_ERROR,
  ModernXlsxError,
  readBuffer,
  SIZE_LIMIT,
  StreamingXlsxWriter,
  Workbook,
} from '../src/index.js';
import {
  INVALID_WORKBOOK as LITE_INVALID_WORKBOOK,
  JSON_ERROR as LITE_JSON_ERROR,
  SIZE_LIMIT as LITE_SIZE_LIMIT,
} from '../src/index-lite.js';
import type { StreamingCellInput } from '../src/io/streaming-writer.js';
import { write as wasmWrite } from '../wasm/modern_xlsx_wasm.js';

describe('Error Handling & Edge Cases', () => {
  // 1. Empty Uint8Array throws
  it('rejects empty Uint8Array', async () => {
    await expect(readBuffer(new Uint8Array(0))).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'UNRECOGNIZED_FORMAT',
    });
  });

  // 2. Random bytes throw
  it('rejects random bytes gracefully', async () => {
    // Fixed pattern: reproducible, and starts with neither PK nor the OLE2 signature.
    const garbage = Uint8Array.from({ length: 1024 }, (_, i) => (i * 37 + 11) & 0xff);
    await expect(readBuffer(garbage)).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'UNRECOGNIZED_FORMAT',
    });
  });

  // 3. Truncated ZIP throws
  it('rejects truncated ZIP', async () => {
    const wb = new Workbook();
    wb.addSheet('Test');
    const buffer = await wb.toBuffer();
    const truncated = buffer.slice(0, Math.floor(buffer.length / 2));
    await expect(readBuffer(truncated)).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'ZIP_READ',
    });
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
    await expect(readBuffer(ole2)).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: 'UNRECOGNIZED_FORMAT',
    });
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

describe('SIZE_LIMIT error code', () => {
  it('exports SIZE_LIMIT from the full and lite entries', () => {
    expect(SIZE_LIMIT).toBe('SIZE_LIMIT');
    expect(LITE_SIZE_LIMIT).toBe('SIZE_LIMIT');
  });

  it('parses a forwarded [SIZE_LIMIT] WASM error', () => {
    const err = ModernXlsxError.fromWasmError(new Error('[SIZE_LIMIT] size limit exceeded: x'));
    expect(err.code).toBe('SIZE_LIMIT');
    expect(err.message).toBe('size limit exceeded: x');
  });
});

describe('JSON and INVALID_WORKBOOK error codes', () => {
  it('exports JSON_ERROR and INVALID_WORKBOOK from the full and lite entries', () => {
    expect(JSON_ERROR).toBe('JSON');
    expect(LITE_JSON_ERROR).toBe('JSON');
    expect(INVALID_WORKBOOK).toBe('INVALID_WORKBOOK');
    expect(LITE_INVALID_WORKBOOK).toBe('INVALID_WORKBOOK');
  });

  it('rejects writing a workbook without sheets with INVALID_WORKBOOK', async () => {
    await expect(new Workbook().toBuffer()).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: INVALID_WORKBOOK,
    });
  });

  it('codes malformed workbook JSON as JSON, with the position printed once', () => {
    let thrown: unknown;
    try {
      wasmWrite('{bad');
    } catch (e) {
      thrown = e;
    }
    const err = ModernXlsxError.fromWasmError(thrown);
    expect(err.code).toBe(JSON_ERROR);
    expect(err.message.match(/line/g)).toHaveLength(1);
  });

  it('wraps a streaming row with an unknown cellType in a ModernXlsxError coded JSON', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S');
    const cells = JSON.parse('[{"value":"1","cellType":"bogus"}]') as StreamingCellInput[];
    let thrown: unknown;
    try {
      writer.writeRow(cells);
    } catch (e) {
      thrown = e;
    }
    expect(thrown).toBeInstanceOf(ModernXlsxError);
    expect(thrown).toMatchObject({ code: JSON_ERROR });
    writer[Symbol.dispose]();
  });

  it('rejects finishing a streaming workbook without sheets with INVALID_WORKBOOK', () => {
    const writer = StreamingXlsxWriter.create();
    expect(() => writer.finish()).toThrow(
      expect.objectContaining({ name: 'ModernXlsxError', code: INVALID_WORKBOOK }),
    );
  });
});
