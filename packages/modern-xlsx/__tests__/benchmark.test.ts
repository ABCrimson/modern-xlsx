import { describe, expect, it } from 'vitest';
import XLSX from 'xlsx';
import {
  aoaToSheet,
  columnToLetter,
  readBuffer,
  sheetToCsv,
  sheetToJson,
  Workbook,
} from '../src/index.js';

// Regression baselines: modern-xlsx's structural advantages (WASM read + compact output)
// must not regress below these conservative ratios vs SheetJS. Observed on Node 26:
// read 10K ~3.4x, read 100K ~4.2x, aoaToSheet 50K ~1.7x. Gates are set well below the
// observed values so they catch real regressions without flaking on loaded CI runners.
const BASELINE = {
  read10kRatio: 2.0,
  read100kRatio: 2.5,
  aoa50kRatio: 1.2,
} as const;

function generateLargeData(rows: number, cols: number): unknown[][] {
  const data: unknown[][] = [];
  const header: string[] = [];
  for (let c = 0; c < cols; c++) {
    header.push(`Col${c}`);
  }
  data.push(header);

  for (let r = 0; r < rows; r++) {
    const row: unknown[] = [];
    for (let c = 0; c < cols; c++) {
      if (c % 3 === 0) row.push(`String value ${r}-${c}`);
      else if (c % 3 === 1) row.push(r * 1000 + c);
      else row.push(r % 2 === 0);
    }
    data.push(row);
  }
  return data;
}

/** Populate a modern-xlsx worksheet from 2D data using cell-by-cell API. */
function populateSheet(
  ws: ReturnType<Workbook['addSheet']>,
  data: unknown[][],
  colLetters: string[],
): void {
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val === undefined || val === null) continue;
      const cell = ws.cell(`${colLetters[c]}${r + 1}`);
      if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
        cell.value = val;
      }
    }
  }
}

function colLettersForCount(n: number): string[] {
  return Array.from({ length: n }, (_, i) => columnToLetter(i));
}

/** Create an XLSX buffer using SheetJS (used as input for read benchmarks). */
function createXlsxBuffer(data: unknown[][]): Uint8Array {
  const xlsxWs = XLSX.utils.aoa_to_sheet(data);
  const xlsxWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Data');
  return new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));
}

describe('Performance benchmarks', () => {
  it('write 10K rows x 10 cols (cell-by-cell)', async () => {
    const data = generateLargeData(10_000, 10);
    const cols = colLettersForCount(10);

    // modern-xlsx (cell-by-cell API)
    const modernStart = performance.now();
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    populateSheet(ws, data, cols);
    const buf = await wb.toBuffer();
    const modernTime = performance.now() - modernStart;

    // SheetJS/xlsx
    const xlsxStart = performance.now();
    const xlsxBuf = createXlsxBuffer(data);
    const xlsxTime = performance.now() - xlsxStart;

    console.log('[write 10K rows x 10 cols -- cell-by-cell]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms (${buf.length} bytes)`);
    console.log(`  SheetJS:     ${xlsxTime.toFixed(1)}ms (${xlsxBuf.length} bytes)`);
    console.log(`  Ratio:       ${(xlsxTime / modernTime).toFixed(2)}x`);

    expect(buf.length).toBeGreaterThan(0);
    expect(modernTime).toBeLessThan(5000);
  });

  it('read 10K rows x 10 cols', async () => {
    const data = generateLargeData(10_000, 10);

    // Use SheetJS to create the XLSX buffer (fast, reliable)
    const xlsxBuf = createXlsxBuffer(data);

    // Read with modern-xlsx
    const readStart = performance.now();
    const wb2 = await readBuffer(xlsxBuf);
    const modernTime = performance.now() - readStart;

    // Read with SheetJS
    const xlsxReadStart = performance.now();
    const xlsxWb2 = XLSX.read(xlsxBuf, { type: 'buffer' });
    const xlsxReadTime = performance.now() - xlsxReadStart;

    console.log('[read 10K rows x 10 cols]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms`);
    console.log(`  SheetJS:     ${xlsxReadTime.toFixed(1)}ms`);
    console.log(`  Ratio:       ${(xlsxReadTime / modernTime).toFixed(2)}x`);

    expect(wb2.sheetCount).toBe(1);
    expect(xlsxWb2.SheetNames).toHaveLength(1);
    expect(modernTime).toBeLessThan(2000);
    // Regression gate: WASM read must stay faster than SheetJS by the baseline ratio.
    expect(xlsxReadTime / modernTime).toBeGreaterThan(BASELINE.read10kRatio);
  });

  it('write 100K rows x 10 cols via aoaToSheet (batch API)', async () => {
    const data = generateLargeData(100_000, 10);

    // modern-xlsx via aoaToSheet (batch API -- fair comparison with SheetJS aoa_to_sheet)
    const modernStart = performance.now();
    const ws = aoaToSheet(data);
    const modernBuildTime = performance.now() - modernStart;

    // SheetJS/xlsx full pipeline: aoa_to_sheet + write
    const xlsxStart = performance.now();
    const xlsxBuf = createXlsxBuffer(data);
    const xlsxTime = performance.now() - xlsxStart;

    console.log('[write 100K rows x 10 cols -- aoaToSheet batch build]');
    console.log(`  modern-xlsx (build sheet): ${modernBuildTime.toFixed(1)}ms`);
    console.log(`  SheetJS (build + write):   ${xlsxTime.toFixed(1)}ms`);
    console.log(`  modern-xlsx rows:          ${ws.rows.length}`);
    console.log(`  SheetJS output:            ${(xlsxBuf.length / 1024 / 1024).toFixed(1)}MB`);

    expect(ws.rows.length).toBeGreaterThan(0);
  }, 120_000);

  it('read 100K rows x 10 cols', async () => {
    const data = generateLargeData(100_000, 10);

    // Use SheetJS to create the XLSX buffer (avoids slow cell-by-cell write)
    const xlsxBuf = createXlsxBuffer(data);

    // Read with modern-xlsx (WASM-powered)
    const readStart = performance.now();
    const wb2 = await readBuffer(xlsxBuf);
    const modernTime = performance.now() - readStart;

    // Read with SheetJS
    const xlsxReadStart = performance.now();
    XLSX.read(xlsxBuf, { type: 'buffer' });
    const xlsxReadTime = performance.now() - xlsxReadStart;

    const sheet = wb2.getSheet('Data');

    console.log('[read 100K rows x 10 cols]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms`);
    console.log(`  SheetJS:     ${xlsxReadTime.toFixed(1)}ms`);
    console.log(`  Ratio:       ${(xlsxReadTime / modernTime).toFixed(2)}x`);
    console.log(`  Rows: ${sheet?.rows.length}, First cell: ${sheet?.rows[0]?.cells[0]?.value}`);

    // The SheetJS-generated file is ~40MB (uncompressed cell format) which is
    // significantly larger than a modern-xlsx-generated file (~5MB). Reading
    // time is proportional to file size. We use a generous threshold here.
    expect(modernTime).toBeLessThan(120_000);
    // Regression gate: large-file read advantage must hold.
    expect(xlsxReadTime / modernTime).toBeGreaterThan(BASELINE.read100kRatio);
  }, 120_000);

  it('aoaToSheet performance with 50K rows', () => {
    const data = generateLargeData(50_000, 5);

    // modern-xlsx
    const start = performance.now();
    const ws = aoaToSheet(data);
    const modernTime = performance.now() - start;

    // SheetJS
    const xlsxStart = performance.now();
    XLSX.utils.aoa_to_sheet(data);
    const xlsxTime = performance.now() - xlsxStart;

    console.log('[aoaToSheet 50K rows x 5 cols]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms`);
    console.log(`  SheetJS:     ${xlsxTime.toFixed(1)}ms`);
    console.log(`  Ratio:       ${(xlsxTime / modernTime).toFixed(2)}x`);

    expect(ws.rows.length).toBeGreaterThan(0);
    // Regression gate: batch sheet build must stay ahead of SheetJS aoa_to_sheet.
    expect(xlsxTime / modernTime).toBeGreaterThan(BASELINE.aoa50kRatio);
  }, 30_000);

  it('sheetToCsv performance with 10K rows', () => {
    const data = generateLargeData(10_000, 10);
    const cols = colLettersForCount(10);

    // Build modern-xlsx sheet
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    populateSheet(ws, data, cols);

    // Build SheetJS sheet
    const xlsxWs = XLSX.utils.aoa_to_sheet(data);

    // modern-xlsx CSV
    const start = performance.now();
    const csv = sheetToCsv(ws);
    const modernTime = performance.now() - start;

    // SheetJS CSV
    const xlsxStart = performance.now();
    const xlsxCsv = XLSX.utils.sheet_to_csv(xlsxWs);
    const xlsxTime = performance.now() - xlsxStart;

    const modernKB = (csv.length / 1024).toFixed(0);
    const xlsxKB = (xlsxCsv.length / 1024).toFixed(0);

    console.log('[sheetToCsv 10K rows x 10 cols]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms (${modernKB}KB)`);
    console.log(`  SheetJS:     ${xlsxTime.toFixed(1)}ms (${xlsxKB}KB)`);
    console.log(`  Ratio:       ${(xlsxTime / modernTime).toFixed(2)}x`);

    expect(csv.length).toBeGreaterThan(0);
  });

  it('sheetToJson performance with 10K rows', () => {
    const data = generateLargeData(10_000, 10);
    const cols = colLettersForCount(10);

    // Build modern-xlsx sheet
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    populateSheet(ws, data, cols);

    // Build SheetJS sheet
    const xlsxWs = XLSX.utils.aoa_to_sheet(data);

    // modern-xlsx JSON
    const start = performance.now();
    const json = sheetToJson(ws);
    const modernTime = performance.now() - start;

    // SheetJS JSON
    const xlsxStart = performance.now();
    const xlsxJson = XLSX.utils.sheet_to_json(xlsxWs);
    const xlsxTime = performance.now() - xlsxStart;

    console.log('[sheetToJson 10K rows x 10 cols]');
    console.log(`  modern-xlsx: ${modernTime.toFixed(1)}ms (${json.length} objects)`);
    console.log(`  SheetJS:     ${xlsxTime.toFixed(1)}ms (${xlsxJson.length} objects)`);
    console.log(`  Ratio:       ${(xlsxTime / modernTime).toFixed(2)}x`);

    expect(json.length).toBeGreaterThan(0);
  });
});
