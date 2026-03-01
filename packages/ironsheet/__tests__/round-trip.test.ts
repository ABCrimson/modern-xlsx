import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { initSync } from '../wasm/ironsheet_wasm.js';
import { initWasm, Workbook, readBuffer, VERSION } from '../src/index.js';
import { version as wasmVersion, read as wasmRead } from '../wasm/ironsheet_wasm.js';
import type { WorkbookData } from '../src/types.js';

beforeAll(async () => {
  const wasmBytes = readFileSync(
    new URL('../wasm/ironsheet_wasm_bg.wasm', import.meta.url),
  );
  initSync({ module: wasmBytes });
  // Also call initWasm so the wasm-loader's `initialized` flag is set.
  // Since the WASM module is already loaded by initSync, init() returns immediately.
  await initWasm();
});

describe('version', () => {
  it('wasmVersion() returns a non-empty string', () => {
    const v = wasmVersion();
    expect(typeof v).toBe('string');
    expect(v.length).toBeGreaterThan(0);
  });

  it('VERSION constant is a semver string', () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+$/);
  });
});

describe('Workbook API', () => {
  it('creates an empty workbook with defaults', () => {
    const wb = new Workbook();
    expect(wb.sheetNames).toEqual([]);
    expect(wb.dateSystem).toBe('date1900');
  });

  it('addSheet() creates a named sheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.name).toBe('Sheet1');
    expect(wb.sheetNames).toEqual(['Sheet1']);
  });

  it('getSheet() returns the correct sheet', () => {
    const wb = new Workbook();
    wb.addSheet('Alpha');
    wb.addSheet('Beta');
    const ws = wb.getSheet('Beta');
    expect(ws).toBeDefined();
    expect(ws!.name).toBe('Beta');
  });

  it('getSheet() returns undefined for non-existent sheet', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(wb.getSheet('Missing')).toBeUndefined();
  });

  it('getSheetByIndex() returns the correct sheet', () => {
    const wb = new Workbook();
    wb.addSheet('First');
    wb.addSheet('Second');
    const ws = wb.getSheetByIndex(1);
    expect(ws).toBeDefined();
    expect(ws!.name).toBe('Second');
  });

  it('getSheetByIndex() returns undefined for out-of-range index', () => {
    const wb = new Workbook();
    wb.addSheet('Only');
    expect(wb.getSheetByIndex(5)).toBeUndefined();
  });

  it('toJSON() returns the internal WorkbookData', () => {
    const wb = new Workbook();
    wb.addSheet('Test');
    const json = wb.toJSON();
    expect(json.sheets).toHaveLength(1);
    expect(json.sheets[0]!.name).toBe('Test');
    expect(json.dateSystem).toBe('date1900');
    expect(json.styles).toBeDefined();
  });
});

describe('Cell API', () => {
  it('cell() creates a cell at the given reference', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    expect(cell.reference).toBe('A1');
  });

  it('cell() returns the same cell on repeated access', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('B2').value = 42;
    expect(ws.cell('B2').value).toBe(42);
  });

  it('sets and gets number values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.value = 3.14;
    expect(cell.type).toBe('number');
    expect(cell.value).toBeCloseTo(3.14);
  });

  it('sets and gets string values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.value = 'Hello';
    expect(cell.type).toBe('sharedString');
    expect(cell.value).toBe('Hello');
  });

  it('sets and gets boolean values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cellT = ws.cell('A1');
    cellT.value = true;
    expect(cellT.type).toBe('boolean');
    expect(cellT.value).toBe(true);

    const cellF = ws.cell('A2');
    cellF.value = false;
    expect(cellF.type).toBe('boolean');
    expect(cellF.value).toBe(false);
  });

  it('sets and gets null values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.value = 42;
    cell.value = null;
    expect(cell.value).toBeNull();
  });

  it('sets and gets formulas', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.formula = 'SUM(B1:B10)';
    expect(cell.formula).toBe('SUM(B1:B10)');
    expect(cell.type).toBe('formulaStr');
  });

  it('clearing a formula sets it to null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.formula = 'A2+A3';
    cell.formula = null;
    expect(cell.formula).toBeNull();
  });

  it('throws on invalid cell reference', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(() => ws.cell('1A')).toThrow('Invalid cell reference');
    expect(() => ws.cell('')).toThrow('Invalid cell reference');
  });
});

describe('Worksheet', () => {
  it('rows returns row data', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 1;
    ws.cell('A2').value = 2;
    expect(ws.rows).toHaveLength(2);
  });

  it('mergeCells is initially empty', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.mergeCells).toEqual([]);
  });
});

describe('Roundtrip: number cells', () => {
  it('writes and reads back number cells correctly', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Numbers');
    ws.cell('A1').value = 42;
    ws.cell('B1').value = 3.14159;
    ws.cell('C1').value = 0;
    ws.cell('A2').value = -100;
    ws.cell('B2').value = 1e10;

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetNames).toContain('Numbers');

    const ws2 = wb2.getSheet('Numbers')!;
    expect(ws2).toBeDefined();

    expect(ws2.cell('A1').value).toBe(42);
    expect(ws2.cell('B1').value).toBeCloseTo(3.14159);
    expect(ws2.cell('C1').value).toBe(0);
    expect(ws2.cell('A2').value).toBe(-100);
    expect(ws2.cell('B2').value).toBe(1e10);
  });
});

describe('Roundtrip: boolean cells', () => {
  it('writes and reads back boolean cells correctly', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Booleans');
    ws.cell('A1').value = true;
    ws.cell('A2').value = false;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Booleans')!;
    expect(ws2).toBeDefined();

    expect(ws2.cell('A1').type).toBe('boolean');
    expect(ws2.cell('A1').value).toBe(true);
    expect(ws2.cell('A2').type).toBe('boolean');
    expect(ws2.cell('A2').value).toBe(false);
  });
});

describe('Roundtrip: string cells', () => {
  it('writes strings and reads back with SST indices', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Strings');
    ws.cell('A1').value = 'Hello';
    ws.cell('A2').value = 'World';
    ws.cell('A3').value = 'Hello'; // duplicate string

    const buffer = await wb.toBuffer();

    // Use raw WASM read to access sharedStrings (the Workbook constructor drops it)
    const raw = wasmRead(buffer) as WorkbookData;

    // The raw reader output should include shared strings
    expect(raw.sharedStrings).toBeDefined();
    expect(raw.sharedStrings!.strings).toContain('Hello');
    expect(raw.sharedStrings!.strings).toContain('World');

    // Wrap in Workbook to use the cell API
    const wb2 = new Workbook(raw);
    const ws2 = wb2.getSheet('Strings')!;
    expect(ws2).toBeDefined();

    const a1 = ws2.cell('A1');
    expect(a1.type).toBe('sharedString');

    // Resolve the shared string: the cell value is the SST index
    const sst = raw.sharedStrings!.strings;
    const a1Index = Number(a1.value);
    expect(sst[a1Index]).toBe('Hello');

    const a2Index = Number(ws2.cell('A2').value);
    expect(sst[a2Index]).toBe('World');

    // A3 should reference the same SST index as A1 since it's a duplicate
    const a3Index = Number(ws2.cell('A3').value);
    expect(sst[a3Index]).toBe('Hello');
    expect(a3Index).toBe(a1Index);
  });
});

describe('Roundtrip: multiple sheets', () => {
  it('preserves multiple sheet names and content', async () => {
    const wb = new Workbook();
    const ws1 = wb.addSheet('Sheet1');
    ws1.cell('A1').value = 100;
    const ws2 = wb.addSheet('Sheet2');
    ws2.cell('A1').value = 200;
    const ws3 = wb.addSheet('Data');
    ws3.cell('B3').value = 999;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    expect(wb2.sheetNames).toEqual(['Sheet1', 'Sheet2', 'Data']);

    expect(wb2.getSheet('Sheet1')!.cell('A1').value).toBe(100);
    expect(wb2.getSheet('Sheet2')!.cell('A1').value).toBe(200);
    expect(wb2.getSheet('Data')!.cell('B3').value).toBe(999);
  });
});

describe('Roundtrip: empty worksheet', () => {
  it('roundtrips a workbook with an empty sheet', async () => {
    const wb = new Workbook();
    wb.addSheet('Empty');

    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetNames).toContain('Empty');

    const ws = wb2.getSheet('Empty')!;
    expect(ws).toBeDefined();
    expect(ws.rows).toHaveLength(0);
  });
});

describe('Roundtrip: readBuffer produces valid Workbook', () => {
  it('returns a Workbook with expected structure', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    ws.cell('A1').value = 1;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    // Should be an instance with the expected API
    expect(wb2.sheetNames).toBeDefined();
    expect(typeof wb2.dateSystem).toBe('string');
    expect(typeof wb2.getSheet).toBe('function');
    expect(typeof wb2.getSheetByIndex).toBe('function');
    expect(typeof wb2.addSheet).toBe('function');
    expect(typeof wb2.toBuffer).toBe('function');
    expect(typeof wb2.toJSON).toBe('function');

    const json = wb2.toJSON();
    expect(json.sheets).toBeDefined();
    expect(json.styles).toBeDefined();
  });
});
