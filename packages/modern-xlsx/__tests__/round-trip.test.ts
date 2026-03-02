import { describe, expect, it } from 'vitest';
import { readBuffer, StyleBuilder, VERSION, Workbook } from '../src/index.js';
import type { WorkbookData } from '../src/types.js';
import { read as wasmRead, version as wasmVersion } from '../wasm/modern_xlsx_wasm.js';

describe('version', () => {
  it('wasmVersion() returns a non-empty string', () => {
    const v = wasmVersion();
    expect(v).toBeTypeOf('string');
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
    expect(ws?.name).toBe('Beta');
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
    expect(ws?.name).toBe('Second');
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
    expect(json.sheets[0]?.name).toBe('Test');
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

const TEST_FLOAT = 1.23456;

describe('Roundtrip: number cells', () => {
  it('writes and reads back number cells correctly', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Numbers');
    ws.cell('A1').value = 42;
    ws.cell('B1').value = TEST_FLOAT;
    ws.cell('C1').value = 0;
    ws.cell('A2').value = -100;
    ws.cell('B2').value = 1e10;

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);

    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetNames).toContain('Numbers');

    const ws2 = wb2.getSheet('Numbers');
    expect(ws2).toBeDefined();

    expect(ws2?.cell('A1').value).toBe(42);
    expect(ws2?.cell('B1').value).toBeCloseTo(TEST_FLOAT);
    expect(ws2?.cell('C1').value).toBe(0);
    expect(ws2?.cell('A2').value).toBe(-100);
    expect(ws2?.cell('B2').value).toBe(1e10);
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
    const ws2 = wb2.getSheet('Booleans');
    expect(ws2).toBeDefined();

    expect(ws2?.cell('A1').type).toBe('boolean');
    expect(ws2?.cell('A1').value).toBe(true);
    expect(ws2?.cell('A2').type).toBe('boolean');
    expect(ws2?.cell('A2').value).toBe(false);
  });
});

describe('Roundtrip: string cells', () => {
  it('readBuffer() resolves SST indices to actual strings', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Strings');
    ws.cell('A1').value = 'Hello';
    ws.cell('A2').value = 'World';
    ws.cell('A3').value = 'Hello'; // duplicate string

    const buffer = await wb.toBuffer();

    // readBuffer() should resolve SST indices automatically
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Strings');
    expect(ws2).toBeDefined();

    // After SST resolution, cell values should be actual strings
    expect(ws2?.cell('A1').type).toBe('sharedString');
    expect(ws2?.cell('A1').value).toBe('Hello');
    expect(ws2?.cell('A2').value).toBe('World');
    expect(ws2?.cell('A3').value).toBe('Hello');
  });

  it('WASM read resolves SST indices and preserves sharedStrings table', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Strings');
    ws.cell('A1').value = 'Hello';
    ws.cell('A2').value = 'World';
    ws.cell('A3').value = 'Hello'; // duplicate string

    const buffer = await wb.toBuffer();

    // Use raw WASM read to verify SST structure is still present
    // WASM now returns a JSON string; parse it to get WorkbookData
    const raw = JSON.parse(wasmRead(buffer)) as WorkbookData;

    // The raw reader output should include shared strings
    expect(raw.sharedStrings).toBeDefined();
    expect(raw.sharedStrings?.strings).toContain('Hello');
    expect(raw.sharedStrings?.strings).toContain('World');

    // Cell values are now resolved in Rust (not raw SST indices)
    const rawSheet = raw.sheets[0];
    const rawA1 = rawSheet?.worksheet.rows[0]?.cells[0];
    expect(rawA1?.cellType).toBe('sharedString');
    expect(rawA1?.value).toBe('Hello');

    const rawA3 = rawSheet?.worksheet.rows[2]?.cells[0];
    expect(rawA3?.value).toBe('Hello');
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

    expect(wb2.getSheet('Sheet1')?.cell('A1').value).toBe(100);
    expect(wb2.getSheet('Sheet2')?.cell('A1').value).toBe(200);
    expect(wb2.getSheet('Data')?.cell('B3').value).toBe(999);
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

    const ws = wb2.getSheet('Empty');
    expect(ws).toBeDefined();
    expect(ws?.rows).toHaveLength(0);
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
    expect(wb2.dateSystem).toBeTypeOf('string');
    expect(wb2.getSheet).toBeTypeOf('function');
    expect(wb2.getSheetByIndex).toBeTypeOf('function');
    expect(wb2.addSheet).toBeTypeOf('function');
    expect(wb2.toBuffer).toBeTypeOf('function');
    expect(wb2.toJSON).toBeTypeOf('function');

    const json = wb2.toJSON();
    expect(json.sheets).toBeDefined();
    expect(json.styles).toBeDefined();
  });
});

describe('Roundtrip: formula cells', () => {
  it('writes and reads back formula cells correctly', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Formulas');
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('A3').formula = 'SUM(A1:A2)';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Formulas');
    expect(ws2).toBeDefined();
    expect(ws2?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(ws2?.cell('A3').type).toBe('formulaStr');
  });

  it('preserves formula with cached value', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('B1');
    cell.formula = 'A1*2';
    cell.value = 100; // cached value

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.cell('B1').formula).toBe('A1*2');
  });
});

describe('StyleBuilder API', () => {
  it('creates a style with font, fill, and border', () => {
    const wb = new Workbook();
    const styleIndex = wb
      .createStyle()
      .font({ bold: true, size: 14, color: '#FF0000' })
      .fill({ pattern: 'solid', fgColor: '#FFFF00' })
      .border({ bottom: { style: 'thin', color: '#000000' } })
      .numberFormat('0.00%')
      .build(wb.styles);

    expect(styleIndex).toBeGreaterThan(0);
    expect(wb.styles.cellXfs[styleIndex]).toBeDefined();
    expect(wb.styles.fonts[wb.styles.cellXfs[styleIndex]?.fontId]?.bold).toBe(true);
  });

  it('reuses existing number format codes', () => {
    const wb = new Workbook();
    const idx1 = wb.createStyle().numberFormat('0.00%').build(wb.styles);
    const idx2 = wb.createStyle().numberFormat('0.00%').build(wb.styles);

    // Both should reference the same numFmtId
    expect(wb.styles.cellXfs[idx1]?.numFmtId).toBe(wb.styles.cellXfs[idx2]?.numFmtId);
    // Only one custom numFmt should exist
    const pctFmts = wb.styles.numFmts.filter((f) => f.formatCode === '0.00%');
    expect(pctFmts).toHaveLength(1);
  });

  it('StyleBuilder is exported from index', () => {
    expect(StyleBuilder).toBeDefined();
    expect(new StyleBuilder()).toBeInstanceOf(StyleBuilder);
  });

  it('Cell.styleIndex getter/setter works', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Styled');
    const cell = ws.cell('A1');
    expect(cell.styleIndex).toBeNull();
    cell.styleIndex = 5;
    expect(cell.styleIndex).toBe(5);
    cell.styleIndex = null;
    expect(cell.styleIndex).toBeNull();
  });

  it('style survives round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Styled');
    const styleIndex = wb.createStyle().font({ bold: true }).build(wb.styles);

    ws.cell('A1').value = 'Bold';
    ws.cell('A1').styleIndex = styleIndex;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.styles.cellXfs.length).toBeGreaterThan(1);
  });
});

describe('Roundtrip: named ranges', () => {
  it('preserves named ranges through write/read', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addNamedRange('MyRange', 'Sheet1!$A$1:$B$10');
    wb.addNamedRange('LocalRange', 'Sheet1!$C$1', 0);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    expect(wb2.namedRanges).toHaveLength(2);
    expect(wb2.getNamedRange('MyRange')).toBeDefined();
    expect(wb2.getNamedRange('MyRange')?.value).toBe('Sheet1!$A$1:$B$10');
    expect(wb2.getNamedRange('LocalRange')?.sheetId).toBe(0);
  });

  it('namedRanges is empty by default', () => {
    const wb = new Workbook();
    expect(wb.namedRanges).toHaveLength(0);
  });

  it('getNamedRange returns undefined for missing name', () => {
    const wb = new Workbook();
    expect(wb.getNamedRange('NonExistent')).toBeUndefined();
  });
});

describe('Roundtrip: alignment style', () => {
  it('alignment style round-trips', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const idx = wb
      .createStyle()
      .alignment({ horizontal: 'center', vertical: 'top', wrapText: true })
      .build(wb.styles);
    ws.cell('A1').value = 'centered';
    ws.cell('A1').styleIndex = idx;

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const xf = wb2.styles.cellXfs[idx];
    expect(xf?.alignment?.horizontal).toBe('center');
    expect(xf?.alignment?.vertical).toBe('top');
    expect(xf?.alignment?.wrapText).toBe(true);
  });
});

describe('Roundtrip: protection style', () => {
  it('protection style round-trips', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const idx = wb.createStyle().protection({ locked: true, hidden: true }).build(wb.styles);
    ws.cell('A1').value = 'protected';
    ws.cell('A1').styleIndex = idx;

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const xf = wb2.styles.cellXfs[idx];
    expect(xf?.protection?.locked).toBe(true);
    expect(xf?.protection?.hidden).toBe(true);
  });
});

describe('Worksheet data validations', () => {
  it('validations is empty by default', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.validations).toHaveLength(0);
  });

  it('addValidation adds a rule', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addValidation('A1:A10', {
      validationType: 'list',
      operator: null,
      formula1: '"Yes,No,Maybe"',
      formula2: null,
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: null,
      errorMessage: null,
    });
    expect(ws.validations).toHaveLength(1);
    expect(ws.validations[0]?.sqref).toBe('A1:A10');
    expect(ws.validations[0]?.validationType).toBe('list');
    expect(ws.validations[0]?.formula1).toBe('"Yes,No,Maybe"');
  });

  it('data validations survive round-trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Validated');
    ws.cell('A1').value = 'test';
    ws.addValidation('A1:A10', {
      validationType: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '100',
      allowBlank: null,
      showErrorMessage: true,
      errorTitle: 'Invalid',
      errorMessage: 'Must be 1-100',
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Validated');
    expect(ws2).toBeDefined();
    expect(ws2?.validations).toHaveLength(1);
    expect(ws2?.validations[0]?.sqref).toBe('A1:A10');
    expect(ws2?.validations[0]?.validationType).toBe('whole');
    expect(ws2?.validations[0]?.operator).toBe('between');
    expect(ws2?.validations[0]?.formula1).toBe('1');
    expect(ws2?.validations[0]?.formula2).toBe('100');
    expect(ws2?.validations[0]?.showErrorMessage).toBe(true);
    expect(ws2?.validations[0]?.errorTitle).toBe('Invalid');
    expect(ws2?.validations[0]?.errorMessage).toBe('Must be 1-100');
  });
});
