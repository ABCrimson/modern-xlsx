/**
 * Full in-depth feature comparison: modern-xlsx vs SheetJS (xlsx)
 *
 * This test suite demonstrates side-by-side feature parity, highlights
 * modern-xlsx advantages (styling, barcodes, tables, validation and repair);
 * speed is measured in __tests__/bench/read-write.bench.ts.
 *
 * Both libraries write an XLSX buffer and cross-read each other's output
 * where applicable — proving interoperability.
 */

import { assert, describe, expect, it } from 'vitest';
import XLSX from 'xlsx';
import {
  aoaToSheet,
  columnToLetter,
  dateToSerial,
  decodeCellRef,
  drawTableFromData,
  encodeCellRef,
  encodeQR,
  formatCell,
  getBuiltinFormat,
  isDateFormatCode,
  jsonToSheet,
  letterToColumn,
  RichTextBuilder,
  readBuffer,
  renderBarcodePNG,
  serialToDate,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  sheetToJson,
  Workbook,
} from '../src/index.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Create an XLSX buffer with SheetJS from a 2D array. */
function xlsxBufferFromAoa(data: unknown[][], sheetName = 'Sheet1'): Uint8Array {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return new Uint8Array(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}

/** Create an XLSX buffer with modern-xlsx from a 2D array. */
async function modernBufferFromAoa(data: unknown[][], sheetName = 'Sheet1'): Promise<Uint8Array> {
  const wb = new Workbook();
  const ws = wb.addSheet(sheetName);
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val === undefined || val === null) continue;
      const cell = ws.cell(`${columnToLetter(c)}${r + 1}`);
      if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
        cell.value = val;
      }
    }
  }
  return wb.toBuffer();
}

// ==========================================================================
// 1. FILE I/O — BUFFER CREATION & READING
// ==========================================================================

describe('1. File I/O', () => {
  const data = [
    ['Name', 'Age', 'Active'],
    ['Alice', 30, true],
    ['Bob', 25, false],
  ];

  it('both produce valid XLSX buffers', async () => {
    const xlsxBuf = xlsxBufferFromAoa(data);
    const modernBuf = await modernBufferFromAoa(data);

    // Both produce non-empty Uint8Array starting with ZIP magic bytes (PK)
    expect(xlsxBuf[0]).toBe(0x50); // 'P'
    expect(xlsxBuf[1]).toBe(0x4b); // 'K'
    expect(modernBuf[0]).toBe(0x50);
    expect(modernBuf[1]).toBe(0x4b);
  });

  it('modern-xlsx reads SheetJS-generated XLSX', async () => {
    const xlsxBuf = xlsxBufferFromAoa(data);
    const wb = await readBuffer(xlsxBuf);
    expect(wb.sheetCount).toBe(1);
    expect(wb.sheetNames).toContain('Sheet1');

    const ws = wb.getSheet('Sheet1');
    expect(ws?.cell('A1').value).toBe('Name');
    expect(ws?.cell('B2').value).toBe(30);
  });

  it('SheetJS reads modern-xlsx-generated XLSX', async () => {
    const modernBuf = await modernBufferFromAoa(data);
    const wb = XLSX.read(modernBuf, { type: 'buffer' });
    expect(wb.SheetNames).toContain('Sheet1');

    const ws = wb.Sheets['Sheet1'];
    if (!ws) throw new Error('Sheet1 not found');
    expect(ws['A1']?.v).toBe('Name');
    expect(ws['B2']?.v).toBe(30);
  });

  it('cross-read roundtrip: SheetJS -> modern-xlsx -> SheetJS', async () => {
    // SheetJS creates buffer
    const xlsxBuf = xlsxBufferFromAoa(data);
    // modern-xlsx reads and re-writes
    const wb = await readBuffer(xlsxBuf);
    const rewritten = await wb.toBuffer();
    // SheetJS reads the re-written buffer
    const wb2 = XLSX.read(rewritten, { type: 'buffer' });
    expect(wb2.Sheets['Sheet1']?.['A2']?.v).toBe('Alice');
    expect(wb2.Sheets['Sheet1']?.['B2']?.v).toBe(30);
  });
});

// ==========================================================================
// 2. WORKBOOK MANAGEMENT
// ==========================================================================

describe('2. Workbook Management', () => {
  it('both support multiple sheets', async () => {
    // SheetJS
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, XLSX.utils.aoa_to_sheet([['A']]), 'Alpha');
    XLSX.utils.book_append_sheet(xlsxWb, XLSX.utils.aoa_to_sheet([['B']]), 'Beta');
    XLSX.utils.book_append_sheet(xlsxWb, XLSX.utils.aoa_to_sheet([['C']]), 'Gamma');
    expect(xlsxWb.SheetNames).toEqual(['Alpha', 'Beta', 'Gamma']);

    // modern-xlsx
    const modernWb = new Workbook();
    modernWb.addSheet('Alpha').cell('A1').value = 'A';
    modernWb.addSheet('Beta').cell('A1').value = 'B';
    modernWb.addSheet('Gamma').cell('A1').value = 'C';
    expect(modernWb.sheetNames).toEqual(['Alpha', 'Beta', 'Gamma']);
  });

  it('both preserve sheet order through roundtrip', async () => {
    const names = ['First', 'Second', 'Third'];

    // SheetJS
    const xlsxWb = XLSX.utils.book_new();
    for (const n of names) XLSX.utils.book_append_sheet(xlsxWb, XLSX.utils.aoa_to_sheet([[n]]), n);
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));
    const xlsxWb2 = XLSX.read(xlsxBuf, { type: 'buffer' });
    expect(xlsxWb2.SheetNames).toEqual(names);

    // modern-xlsx
    const modernWb = new Workbook();
    for (const n of names) modernWb.addSheet(n).cell('A1').value = n;
    const modernBuf = await modernWb.toBuffer();
    const modernWb2 = await readBuffer(modernBuf);
    expect(modernWb2.sheetNames).toEqual(names);
  });

  it('modern-xlsx addSheet throws on duplicate name', () => {
    const wb = new Workbook();
    wb.addSheet('Dup');
    expect(() => wb.addSheet('Dup')).toThrow('already exists');
  });

  it('modern-xlsx removeSheet works', () => {
    const wb = new Workbook();
    wb.addSheet('Keep');
    wb.addSheet('Remove');
    expect(wb.removeSheet('Remove')).toBe(true);
    expect(wb.sheetNames).toEqual(['Keep']);
    expect(wb.removeSheet('Nonexistent')).toBe(false);
  });

  it('modern-xlsx validates sheet names per ECMA-376', () => {
    const wb = new Workbook();
    expect(() => wb.addSheet('')).toThrow();
    expect(() => wb.addSheet('x'.repeat(32))).toThrow();
    expect(() => wb.addSheet('Bad/Name')).toThrow();
    expect(() => wb.addSheet("'Leading")).toThrow();
  });
});

// ==========================================================================
// 3. CELL OPERATIONS — VALUE TYPES
// ==========================================================================

describe('3. Cell Operations', () => {
  it('both handle string cells', async () => {
    const data = [['hello'], ['world']];

    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet(data);
    expect(xlsxWs['A1']?.t).toBe('s');
    expect(xlsxWs['A1']?.v).toBe('hello');

    // modern-xlsx
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'hello';
    expect(ws.cell('A1').value).toBe('hello');
    expect(ws.cell('A1').type).toBe('sharedString');

    // Roundtrip
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('S')?.cell('A1').value).toBe('hello');
  });

  it('both handle number cells', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([[42], [3.14], [-100]]);
    expect(xlsxWs['A1']?.t).toBe('n');
    expect(xlsxWs['A1']?.v).toBe(42);

    // modern-xlsx
    const wb = new Workbook();
    const ws = wb.addSheet('N');
    ws.cell('A1').value = 42;
    ws.cell('A2').value = 3.14;
    ws.cell('A3').value = -100;
    expect(ws.cell('A1').type).toBe('number');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('N')?.cell('A1').value).toBe(42);
    expect(wb2.getSheet('N')?.cell('A2').value).toBeCloseTo(3.14);
    expect(wb2.getSheet('N')?.cell('A3').value).toBe(-100);
  });

  it('both handle boolean cells', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([[true], [false]]);
    expect(xlsxWs['A1']?.t).toBe('b');
    expect(xlsxWs['A1']?.v).toBe(true);

    // modern-xlsx
    const wb = new Workbook();
    const ws = wb.addSheet('B');
    ws.cell('A1').value = true;
    ws.cell('A2').value = false;
    expect(ws.cell('A1').type).toBe('boolean');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('B')?.cell('A1').value).toBe(true);
    expect(wb2.getSheet('B')?.cell('A2').value).toBe(false);
  });

  it('modern-xlsx handles null values', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 42;
    ws.cell('A1').value = null;
    expect(ws.cell('A1').value).toBeNull();
  });

  it('both handle special string characters (XML entities)', async () => {
    const specials = ['Tom & Jerry', '2 < 3', '5 > 4', 'He said "hello"', "It's fine"];

    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet(specials.map((s) => [s]));
    for (let i = 0; i < specials.length; i++) {
      expect(xlsxWs[`A${i + 1}`]?.v).toBe(specials[i]);
    }

    // modern-xlsx
    const wb = new Workbook();
    const ws = wb.addSheet('X');
    for (let i = 0; i < specials.length; i++) {
      const val = specials[i];
      if (val !== undefined) ws.cell(`A${i + 1}`).value = val;
    }
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('X');
    for (let i = 0; i < specials.length; i++) {
      expect(ws2?.cell(`A${i + 1}`).value).toBe(specials[i]);
    }
  });

  it('both handle unicode strings', async () => {
    const unicodes = ['🎉 celebration', '中文测试', 'café résumé'];

    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet(unicodes.map((s) => [s]));
    expect(xlsxWs['A1']?.v).toBe('🎉 celebration');

    // modern-xlsx
    const wb = new Workbook();
    const ws = wb.addSheet('U');
    for (let i = 0; i < unicodes.length; i++) {
      const val = unicodes[i];
      if (val !== undefined) ws.cell(`A${i + 1}`).value = val;
    }
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    for (let i = 0; i < unicodes.length; i++) {
      expect(wb2.getSheet('U')?.cell(`A${i + 1}`).value).toBe(unicodes[i]);
    }
  });
});

// ==========================================================================
// 4. FORMULAS
// ==========================================================================

describe('4. Formulas', () => {
  it('both write formulas that survive roundtrip', async () => {
    // SheetJS writes formula — verify it stores the formula in-memory
    const xlsxWs = XLSX.utils.aoa_to_sheet([[10], [20]]);
    xlsxWs['A3'] = { t: 'n', f: 'SUM(A1:A2)' };
    expect(xlsxWs['A3']?.f).toBe('SUM(A1:A2)');
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');

    // modern-xlsx writes formula
    const modernWb = new Workbook();
    const ws = modernWb.addSheet('Formulas');
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('A3').formula = 'SUM(A1:A2)';
    ws.cell('B1').formula = 'AVERAGE(A1:A2)';
    ws.cell('C1').formula = 'MAX(A1:A2)';
    // Cached results: a formula cell without a <v> is not read as a formula by SheetJS.
    ws.cell('A3').value = 30;
    ws.cell('B1').value = 15;
    ws.cell('C1').value = 20;

    const buf = await modernWb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Formulas')?.cell('A3').formula).toBe('SUM(A1:A2)');
    expect(wb2.getSheet('Formulas')?.cell('B1').formula).toBe('AVERAGE(A1:A2)');
    expect(wb2.getSheet('Formulas')?.cell('C1').formula).toBe('MAX(A1:A2)');

    // SheetJS can read modern-xlsx formulas
    const xlsxWb3 = XLSX.read(buf, { type: 'buffer' });
    expect(xlsxWb3.Sheets['Formulas']?.['A3']?.f).toBe('SUM(A1:A2)');
  });

  it('modern-xlsx formula with cached value', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 50;
    const cell = ws.cell('A2');
    cell.formula = 'A1*2';
    cell.value = 100; // cached value
    expect(cell.formula).toBe('A1*2');
    expect(cell.type).toBe('number');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('S')?.cell('A2').formula).toBe('A1*2');
  });
});

// ==========================================================================
// 5. MERGE CELLS
// ==========================================================================

describe('5. Merge Cells', () => {
  it('both create and roundtrip merged cells', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([['Merged Title', '', '']]);
    xlsxWs['!merges'] = [XLSX.utils.decode_range('A1:C1')];
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));

    // Read SheetJS merges with modern-xlsx
    const wb = await readBuffer(xlsxBuf);
    expect(wb.getSheet('Sheet1')?.mergeCells).toContain('A1:C1');

    // modern-xlsx creates merges
    const modernWb = new Workbook();
    const ws = modernWb.addSheet('Merges');
    ws.cell('A1').value = 'Wide Header';
    ws.addMergeCell('A1:D1');
    ws.cell('A2').value = 'Tall Cell';
    ws.addMergeCell('A2:A4');
    expect(ws.mergeCells).toEqual(['A1:D1', 'A2:A4']);

    // Roundtrip
    const buf = await modernWb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Merges')?.mergeCells).toContain('A1:D1');
    expect(wb2.getSheet('Merges')?.mergeCells).toContain('A2:A4');

    // SheetJS can read modern-xlsx merges
    const xlsxWb2 = XLSX.read(buf, { type: 'buffer' });
    const merges = xlsxWb2.Sheets['Merges']?.['!merges'] ?? [];
    expect(merges.length).toBe(2);
  });

  it('modern-xlsx removeMergeCell works', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addMergeCell('A1:B1');
    ws.addMergeCell('C1:D1');
    expect(ws.removeMergeCell('A1:B1')).toBe(true);
    expect(ws.mergeCells).toEqual(['C1:D1']);
    expect(ws.removeMergeCell('NONEXISTENT')).toBe(false);
  });
});

// ==========================================================================
// 6. COLUMN WIDTHS & ROW HEIGHTS
// ==========================================================================

describe('6. Column Widths & Row Heights', () => {
  it('both set column widths', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([['Wide', 'Narrow']]);
    xlsxWs['!cols'] = [{ wch: 30 }, { wch: 10 }];
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));

    // modern-xlsx reads SheetJS column widths
    const wb = await readBuffer(xlsxBuf);
    const cols = wb.getSheet('Sheet1')?.columns ?? [];
    expect(cols.length).toBeGreaterThan(0);

    // modern-xlsx sets column widths
    const modernWb = new Workbook();
    const ws = modernWb.addSheet('Layout');
    ws.cell('A1').value = 'Wide Column';
    ws.setColumnWidth(1, 30);
    ws.setColumnWidth(2, 15);
    expect(ws.columns.length).toBe(2);

    const buf = await modernWb.toBuffer();
    const wb2 = await readBuffer(buf);
    const cols2 = wb2.getSheet('Layout')?.columns ?? [];
    expect(cols2.length).toBe(2);
    expect(cols2[0]?.width).toBe(30);
  });

  it('both set row heights', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([['Tall']]);
    xlsxWs['!rows'] = [{ hpt: 40 }];
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));

    // modern-xlsx reads SheetJS row heights
    const wb = await readBuffer(xlsxBuf);
    const rows = wb.getSheet('Sheet1')?.rows ?? [];
    expect(rows.length).toBeGreaterThan(0);

    // modern-xlsx sets row heights
    const modernWb = new Workbook();
    const ws = modernWb.addSheet('Heights');
    ws.cell('A1').value = 'Tall';
    ws.setRowHeight(1, 40);
    ws.cell('A5').value = 'Hidden';
    ws.setRowHidden(5, true);

    const buf = await modernWb.toBuffer();
    const wb2 = await readBuffer(buf);
    const rows2 = wb2.getSheet('Heights')?.rows ?? [];
    const row1 = rows2.find((r) => r.index === 1);
    expect(row1?.height).toBe(40);
    const row5 = rows2.find((r) => r.index === 5);
    expect(row5?.hidden).toBe(true);
  });
});

// ==========================================================================
// 7. FROZEN PANES
// ==========================================================================

describe('7. Frozen Panes', () => {
  it('modern-xlsx freezes rows and columns', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Frozen');
    ws.cell('A1').value = 'Header';
    ws.frozenPane = { rows: 1, cols: 1 };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Frozen')?.frozenPane).toEqual({ rows: 1, cols: 1 });

    // SheetJS reads the file and the cell data, but not <pane>
    const xlsxWb = XLSX.read(buf, { type: 'buffer' });
    expect(xlsxWb.SheetNames).toContain('Frozen');
    expect(xlsxWb.Sheets['Frozen']?.['A1']?.v).toBe('Header');
  });

  it('SheetJS output carries no frozen pane (CE 0.20.3 ignores !freeze)', async () => {
    // '!freeze' is not a SheetJS key: SheetJS CE 0.20.3 writes no <pane>
    const xlsxWs = XLSX.utils.aoa_to_sheet([
      ['H1', 'H2'],
      ['A', 'B'],
    ]);
    xlsxWs['!freeze'] = { xSplit: '1', ySplit: '1' };
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));

    const wb = await readBuffer(xlsxBuf);
    // SheetJS CE 0.20.3 does not write <pane> ('!freeze' is not a SheetJS key),
    // so there is no frozen pane to read back.
    expect(wb.sheetCount).toBe(1);
    expect(wb.getSheet('Sheet1')?.frozenPane).toBeUndefined();
  });
});

// ==========================================================================
// 8. AUTO FILTER
// ==========================================================================

describe('8. Auto Filter', () => {
  it('both set auto filter ranges', async () => {
    // SheetJS
    const xlsxWs = XLSX.utils.aoa_to_sheet([
      ['Name', 'Score'],
      ['Alice', 95],
      ['Bob', 87],
    ]);
    xlsxWs['!autofilter'] = { ref: 'A1:B3' };
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, xlsxWs, 'Sheet1');
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));

    // modern-xlsx reads SheetJS autofilter
    const wb = await readBuffer(xlsxBuf);
    expect(wb.getSheet('Sheet1')?.autoFilter?.range).toBe('A1:B3');

    // modern-xlsx sets autofilter
    const modernWb = new Workbook();
    const ws = modernWb.addSheet('Filtered');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Score';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 95;
    ws.autoFilter = 'A1:B2';

    const buf = await modernWb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.getSheet('Filtered')?.autoFilter?.range).toBe('A1:B2');

    // SheetJS reads modern-xlsx autofilter
    const xlsxWb2 = XLSX.read(buf, { type: 'buffer' });
    expect(xlsxWb2.Sheets['Filtered']?.['!autofilter']?.ref).toBe('A1:B2');
  });

  it('modern-xlsx autoFilter can be cleared', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.autoFilter = 'A1:C10';
    expect(ws.autoFilter?.range).toBe('A1:C10');
    ws.autoFilter = null;
    expect(ws.autoFilter).toBeNull();
  });
});

// ==========================================================================
// 9. STYLING (modern-xlsx ADVANTAGE — free vs SheetJS Pro)
// ==========================================================================

describe('9. Styling (modern-xlsx advantage)', () => {
  it('modern-xlsx: fluent style builder creates font, fill, border, alignment', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Styled');

    const headerStyle = wb
      .createStyle()
      .font({ name: 'Arial', size: 14, bold: true, color: 'FFFFFF' })
      .fill({ pattern: 'solid', fgColor: '4472C4' })
      .alignment({ horizontal: 'center', vertical: 'center', wrapText: true })
      .border({ bottom: { style: 'medium', color: '000000' } })
      .numberFormat('#,##0.00')
      .build(wb.styles);

    ws.cell('A1').value = 'Styled Header';
    ws.cell('A1').styleIndex = headerStyle;

    expect(headerStyle).toBeGreaterThan(0);
    expect(wb.styles.cellXfs.length).toBeGreaterThan(1);

    // Roundtrip
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const xf = wb2.styles.cellXfs[headerStyle];
    assert.isDefined(xf);
    expect(wb2.styles.fonts[xf.fontId]?.bold).toBe(true);
    expect(wb2.styles.fonts[xf.fontId]?.size).toBe(14);
    expect(xf.alignment?.horizontal).toBe('center');
    expect(xf.alignment?.wrapText).toBe(true);

    // SheetJS community edition CANNOT set styles (requires Pro)
    // This demonstrates modern-xlsx's key advantage
  });

  it('modern-xlsx: protection style roundtrips', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Protected');
    const idx = wb.createStyle().protection({ locked: true, hidden: true }).build(wb.styles);
    ws.cell('A1').value = 'locked';
    ws.cell('A1').styleIndex = idx;

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const xf = wb2.styles.cellXfs[idx];
    expect(xf?.protection?.locked).toBe(true);
    expect(xf?.protection?.hidden).toBe(true);
  });

  it('modern-xlsx: number format reuse', () => {
    const wb = new Workbook();
    const idx1 = wb.createStyle().numberFormat('0.00%').build(wb.styles);
    const idx2 = wb.createStyle().numberFormat('0.00%').build(wb.styles);

    // Both reference the same numFmtId
    expect(wb.styles.cellXfs[idx1]?.numFmtId).toBe(wb.styles.cellXfs[idx2]?.numFmtId);
    // Only one custom numFmt entry
    const pctFmts = wb.styles.numFmts.filter((f) => f.formatCode === '0.00%');
    expect(pctFmts).toHaveLength(1);
  });

  it('SheetJS community: no styling API (Pro-only feature)', () => {
    // SheetJS community edition stores cell data but NOT styles
    const ws = XLSX.utils.aoa_to_sheet([['Hello']]);
    // ws['A1'].s would be the style object, but it's not available in community
    expect(ws['A1']?.s).toBeUndefined();
  });
});

// ==========================================================================
// 10. DATA VALIDATION
// ==========================================================================

describe('10. Data Validation', () => {
  it('modern-xlsx creates and roundtrips data validations', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Validated');
    ws.cell('A1').value = 'Pick one';

    // Dropdown list
    ws.addValidation('B1', {
      validationType: 'list',
      operator: null,
      formula1: '"Yes,No,Maybe"',
      formula2: null,
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: 'Invalid',
      errorMessage: 'Please select from the list',
    });

    // Number range
    ws.addValidation('C1', {
      validationType: 'whole',
      operator: 'between',
      formula1: '1',
      formula2: '100',
      allowBlank: null,
      showErrorMessage: true,
      errorTitle: 'Out of range',
      errorMessage: 'Enter 1-100',
    });

    expect(ws.validations).toHaveLength(2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Validated');
    expect(ws2?.validations).toHaveLength(2);
    expect(ws2?.validations[0]?.validationType).toBe('list');
    expect(ws2?.validations[0]?.formula1).toBe('"Yes,No,Maybe"');
    expect(ws2?.validations[1]?.operator).toBe('between');
  });

  it('modern-xlsx removeValidation works', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addValidation('A1', {
      validationType: 'list',
      operator: null,
      formula1: '"X,Y"',
      formula2: null,
      allowBlank: null,
      showErrorMessage: null,
      errorTitle: null,
      errorMessage: null,
    });
    expect(ws.removeValidation('A1')).toBe(true);
    expect(ws.validations).toHaveLength(0);
  });
});

// ==========================================================================
// 11. HYPERLINKS
// ==========================================================================

describe('11. Hyperlinks', () => {
  it('modern-xlsx creates and roundtrips hyperlinks', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Links');
    ws.cell('A1').value = 'Visit Example';
    ws.addHyperlink('A1', 'https://example.com', {
      display: 'Visit Example',
      tooltip: 'Opens example.com',
    });
    ws.cell('A2').value = 'Internal Link';
    ws.addHyperlink('A2', 'Sheet2!A1', { display: 'Go to Sheet2' });

    expect(ws.hyperlinks).toHaveLength(2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Links');
    expect(ws2?.hyperlinks).toHaveLength(2);
    const [external, internal] = ws2?.hyperlinks ?? [];
    expect(external).toMatchObject({
      cellRef: 'A1',
      display: 'Visit Example',
      tooltip: 'Opens example.com',
    });
    expect(internal).toMatchObject({
      cellRef: 'A2',
      location: 'Sheet2!A1',
      display: 'Go to Sheet2',
    });
  });

  it('modern-xlsx removeHyperlink works', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addHyperlink('A1', 'https://example.com');
    expect(ws.removeHyperlink('A1')).toBe(true);
    expect(ws.hyperlinks).toHaveLength(0);
  });
});

// ==========================================================================
// 12. COMMENTS
// ==========================================================================

describe('12. Comments', () => {
  it('modern-xlsx creates and roundtrips comments', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Comments');
    ws.cell('A1').value = 'Has comment';
    ws.addComment('A1', 'Alice', 'This is a note');

    expect(ws.comments).toHaveLength(1);
    expect(ws.comments[0]?.author).toBe('Alice');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Comments');
    expect(ws2?.comments).toEqual([{ cellRef: 'A1', author: 'Alice', text: 'This is a note' }]);
  });

  it('modern-xlsx removeComment works', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addComment('A1', 'Author', 'Text');
    expect(ws.removeComment('A1')).toBe(true);
    expect(ws.comments).toHaveLength(0);
  });
});

// ==========================================================================
// 13. NAMED RANGES
// ==========================================================================

describe('13. Named Ranges', () => {
  it('modern-xlsx creates and roundtrips named ranges', async () => {
    const wb = new Workbook();
    wb.addSheet('Data').cell('A1').value = 100;
    wb.addNamedRange('SalesTotal', 'Data!$A$1:$A$10');
    wb.addNamedRange('LocalName', 'Data!$B$1', 0);

    expect(wb.namedRanges).toHaveLength(2);
    expect(wb.getNamedRange('SalesTotal')?.value).toBe('Data!$A$1:$A$10');

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.namedRanges).toHaveLength(2);
    expect(wb2.getNamedRange('SalesTotal')?.value).toBe('Data!$A$1:$A$10');
    expect(wb2.getNamedRange('LocalName')?.sheetId).toBe(0);
  });

  it('SheetJS creates named ranges (different API)', () => {
    const xlsxWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(xlsxWb, XLSX.utils.aoa_to_sheet([[1]]), 'Data');
    // SheetJS uses Workbook.Names array
    xlsxWb.Workbook = { Names: [{ Name: 'TestRange', Ref: 'Data!$A$1' }] };
    const xlsxBuf = new Uint8Array(XLSX.write(xlsxWb, { type: 'buffer', bookType: 'xlsx' }));
    const xlsxWb2 = XLSX.read(xlsxBuf, { type: 'buffer' });
    expect(xlsxWb2.Workbook?.Names?.[0]?.Name).toBe('TestRange');
  });

  it('modern-xlsx removeNamedRange works', () => {
    const wb = new Workbook();
    wb.addNamedRange('Temp', 'Sheet1!$A$1');
    expect(wb.removeNamedRange('Temp')).toBe(true);
    expect(wb.namedRanges).toHaveLength(0);
    expect(wb.removeNamedRange('Nonexistent')).toBe(false);
  });
});

// ==========================================================================
// 14. PAGE SETUP & MARGINS
// ==========================================================================

describe('14. Page Setup & Margins', () => {
  it('modern-xlsx sets and roundtrips page setup', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Print');
    ws.cell('A1').value = 'Printable';
    ws.pageSetup = {
      orientation: 'landscape',
      paperSize: 1,
      fitToWidth: 1,
      fitToHeight: 0,
    };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Print');
    expect(ws2?.pageSetup?.orientation).toBe('landscape');
    expect(ws2?.pageSetup?.paperSize).toBe(1);
    expect(ws2?.pageSetup?.fitToWidth).toBe(1);
    expect(ws2?.pageSetup?.fitToHeight).toBe(0);
  });

  it('modern-xlsx sets and round-trips page margins', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Margins');
    ws.pageMargins = {
      top: 0.75,
      bottom: 0.75,
      left: 0.7,
      right: 0.7,
      header: 0.3,
      footer: 0.3,
    };
    const wb2 = await readBuffer(await wb.toBuffer());
    expect(wb2.getSheet('Margins')?.pageMargins).toEqual({
      top: 0.75,
      bottom: 0.75,
      left: 0.7,
      right: 0.7,
      header: 0.3,
      footer: 0.3,
    });
  });
});

// ==========================================================================
// 15. SHEET PROTECTION
// ==========================================================================

describe('15. Sheet Protection', () => {
  it('modern-xlsx sets and roundtrips sheet protection', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Protected');
    ws.cell('A1').value = 'Locked';
    ws.sheetProtection = {
      sheet: true,
      objects: false,
      scenarios: false,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
    };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Protected');
    expect(ws2?.sheetProtection?.sheet).toBe(true);
  });

  it('modern-xlsx clears sheet protection', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.sheetProtection = {
      sheet: true,
      objects: false,
      scenarios: false,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
    };
    expect(ws.sheetProtection?.sheet).toBe(true);
    ws.sheetProtection = null;
    expect(ws.sheetProtection).toBeNull();
  });
});

// ==========================================================================
// 16. DOCUMENT PROPERTIES
// ==========================================================================

describe('16. Document Properties', () => {
  it('modern-xlsx sets and roundtrips document properties', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 1;
    wb.docProperties = {
      title: 'Sales Report Q4',
      creator: 'Finance Team',
      description: 'Quarterly sales data',
      created: '2026-01-01T00:00:00Z',
      modified: '2026-03-01T00:00:00Z',
    };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    expect(wb2.docProperties?.title).toBe('Sales Report Q4');
    expect(wb2.docProperties?.creator).toBe('Finance Team');
  });
});

// ==========================================================================
// 17. CELL REFERENCE UTILITIES
// ==========================================================================

describe('17. Cell Reference Utilities', () => {
  it('both convert column indices to letters', () => {
    // modern-xlsx: 0-based
    expect(columnToLetter(0)).toBe('A');
    expect(columnToLetter(25)).toBe('Z');
    expect(columnToLetter(26)).toBe('AA');
    expect(columnToLetter(701)).toBe('ZZ');
    expect(columnToLetter(702)).toBe('AAA');
    expect(columnToLetter(16383)).toBe('XFD');

    // SheetJS: uses encode_col (0-based)
    expect(XLSX.utils.encode_col(0)).toBe('A');
    expect(XLSX.utils.encode_col(25)).toBe('Z');
    expect(XLSX.utils.encode_col(26)).toBe('AA');
    expect(XLSX.utils.encode_col(16383)).toBe('XFD');
  });

  it('both convert letters to column indices', () => {
    // modern-xlsx
    expect(letterToColumn('A')).toBe(0);
    expect(letterToColumn('Z')).toBe(25);
    expect(letterToColumn('AA')).toBe(26);
    expect(letterToColumn('XFD')).toBe(16383);

    // SheetJS
    expect(XLSX.utils.decode_col('A')).toBe(0);
    expect(XLSX.utils.decode_col('Z')).toBe(25);
    expect(XLSX.utils.decode_col('AA')).toBe(26);
  });

  it('both decode cell references', () => {
    // modern-xlsx: returns { row, col } 0-based
    const modern = decodeCellRef('B3');
    expect(modern.row).toBe(2);
    expect(modern.col).toBe(1);

    // SheetJS: decode_cell returns { r, c } 0-based
    const sheetjs = XLSX.utils.decode_cell('B3');
    expect(sheetjs.r).toBe(2);
    expect(sheetjs.c).toBe(1);
  });

  it('both encode cell references', () => {
    // modern-xlsx: encodeCellRef(row, col) 0-based
    expect(encodeCellRef(2, 1)).toBe('B3');

    // SheetJS: encode_cell({ r, c }) 0-based
    expect(XLSX.utils.encode_cell({ r: 2, c: 1 })).toBe('B3');
  });

  it('modern-xlsx roundtrips boundary values', () => {
    const boundaries = [0, 25, 26, 51, 52, 701, 702, 16383];
    for (const col of boundaries) {
      expect(letterToColumn(columnToLetter(col))).toBe(col);
    }
  });
});

// ==========================================================================
// 18. DATE UTILITIES
// ==========================================================================

describe('18. Date Utilities', () => {
  it('modern-xlsx: dateToSerial converts Temporal-like input', () => {
    const serial = dateToSerial({ year: 2026, month: 3, day: 1 });
    expect(serial).toBe(46082);
  });

  it('modern-xlsx: serialToDate converts back', () => {
    const d = serialToDate(46082);
    expect(d.getUTCFullYear()).toBe(2026);
    expect(d.getUTCMonth()).toBe(2); // March, 0-based
    expect(d.getUTCDate()).toBe(1);
  });

  it('modern-xlsx: isDateFormatCode detects date patterns', () => {
    expect(isDateFormatCode('yyyy-mm-dd')).toBe(true);
    expect(isDateFormatCode('dd/mm/yyyy')).toBe(true);
    expect(isDateFormatCode('#,##0.00')).toBe(false);
    expect(isDateFormatCode('0.00%')).toBe(false);
  });

  it('modern-xlsx: getBuiltinFormat returns known formats', () => {
    expect(getBuiltinFormat(0)).toBe('General');
    expect(getBuiltinFormat(1)).toBe('0');
    expect(getBuiltinFormat(14)).toBe('mm-dd-yy');
  });

  it('SheetJS: SSF handles date serial numbers', () => {
    // SheetJS uses XLSX.SSF for number formatting
    const formatted = XLSX.SSF.format('yyyy-mm-dd', 46113);
    expect(formatted).toContain('2026');
  });

  it('Lotus 1-2-3 boundary (libraries differ at 60)', () => {
    // Serial 60 is the Lotus 1-2-3 bug date (Feb 29, 1900 — doesn't exist).
    // modern-xlsx maps it to Feb 28; SheetJS SSF renders Feb 29.
    const iso = (serial: number) => serialToDate(serial).toISOString().slice(0, 10);
    expect([59, 60, 61].map(iso)).toEqual(['1900-02-28', '1900-02-28', '1900-03-01']);
    expect([59, 60, 61].map((n) => XLSX.SSF.format('yyyy-mm-dd', n))).toEqual([
      '1900-02-28',
      '1900-02-29',
      '1900-03-01',
    ]);
  });
});

// ==========================================================================
// 19. SHEET CONVERSION UTILITIES
// ==========================================================================

describe('19. Sheet Conversion Utilities', () => {
  const data = [
    { name: 'Alice', age: 30, city: 'NYC' },
    { name: 'Bob', age: 25, city: 'LA' },
    { name: 'Charlie', age: 35, city: 'SF' },
  ];

  describe('19a. JSON <-> Sheet', () => {
    it('both convert JSON to sheet', () => {
      // modern-xlsx
      const modernWs = jsonToSheet(data);
      expect(modernWs.cell('A1').value).toBe('name');
      expect(modernWs.cell('A2').value).toBe('Alice');
      expect(modernWs.cell('B3').value).toBe(25);

      // SheetJS
      const xlsxWs = XLSX.utils.json_to_sheet(data);
      expect(xlsxWs['A1']?.v).toBe('name');
      expect(xlsxWs['A2']?.v).toBe('Alice');
      expect(xlsxWs['B3']?.v).toBe(25);
    });

    it('both convert sheet to JSON', () => {
      // modern-xlsx
      const modernWs = jsonToSheet(data);
      const modernJson = sheetToJson(modernWs);
      expect(modernJson).toHaveLength(3);
      expect(modernJson[0]).toEqual({ name: 'Alice', age: 30, city: 'NYC' });

      // SheetJS
      const xlsxWs = XLSX.utils.json_to_sheet(data);
      const xlsxJson = XLSX.utils.sheet_to_json(xlsxWs);
      expect(xlsxJson).toHaveLength(3);
      expect(xlsxJson[0]).toEqual({ name: 'Alice', age: 30, city: 'NYC' });
    });

    it('modern-xlsx jsonToSheet with custom header', () => {
      const ws = jsonToSheet(data, { header: ['city', 'name'] });
      expect(ws.cell('A1').value).toBe('city');
      expect(ws.cell('B1').value).toBe('name');
    });

    it('modern-xlsx jsonToSheet with skipHeader', () => {
      const ws = jsonToSheet(data, { skipHeader: true });
      // First row should be data, not headers
      expect(ws.cell('A1').value).toBe('Alice');
    });
  });

  describe('19b. AOA <-> Sheet', () => {
    const aoa = [
      ['Name', 'Age'],
      ['Alice', 30],
      ['Bob', 25],
    ];

    it('both convert AOA to sheet', () => {
      // modern-xlsx
      const modernWs = aoaToSheet(aoa);
      expect(modernWs.cell('A1').value).toBe('Name');
      expect(modernWs.cell('B2').value).toBe(30);

      // SheetJS
      const xlsxWs = XLSX.utils.aoa_to_sheet(aoa);
      expect(xlsxWs['A1']?.v).toBe('Name');
      expect(xlsxWs['B2']?.v).toBe(30);
    });

    it('modern-xlsx aoaToSheet with origin offset', () => {
      const ws = aoaToSheet(aoa, { origin: 'C3' });
      expect(ws.cell('C3').value).toBe('Name');
      expect(ws.cell('D4').value).toBe(30);
    });
  });

  describe('19c. Sheet to CSV', () => {
    it('both convert sheet to CSV', () => {
      // modern-xlsx
      const modernWs = aoaToSheet([
        ['A', 'B'],
        [1, 2],
      ]);
      const modernCsv = sheetToCsv(modernWs);
      expect(modernCsv).toContain('A,B');
      expect(modernCsv).toContain('1,2');

      // SheetJS
      const xlsxWs = XLSX.utils.aoa_to_sheet([
        ['A', 'B'],
        [1, 2],
      ]);
      const xlsxCsv = XLSX.utils.sheet_to_csv(xlsxWs);
      expect(xlsxCsv).toContain('A,B');
      expect(xlsxCsv).toContain('1,2');
    });

    it('modern-xlsx sheetToCsv with custom separator', () => {
      const ws = aoaToSheet([
        ['A', 'B'],
        [1, 2],
      ]);
      const tsv = sheetToCsv(ws, { separator: '\t' });
      expect(tsv).toContain('A\tB');
    });

    it('modern-xlsx sheetToCsv with forceQuote', () => {
      const ws = aoaToSheet([['hello', 'world']]);
      const csv = sheetToCsv(ws, { forceQuote: true });
      expect(csv).toContain('"hello"');
    });
  });

  describe('19d. Sheet to HTML (modern-xlsx only)', () => {
    it('modern-xlsx converts sheet to HTML table', () => {
      const ws = aoaToSheet([
        ['Name', 'Score'],
        ['Alice', 95],
      ]);
      const html = sheetToHtml(ws);
      expect(html).toContain('<table>');
      expect(html).toContain('<td>Name</td>');
      expect(html).toContain('<td>95</td>');
    });

    it('modern-xlsx sheetToHtml with header option', () => {
      const ws = aoaToSheet([
        ['Name', 'Score'],
        ['Alice', 95],
      ]);
      const html = sheetToHtml(ws, { header: true });
      expect(html).toContain('<thead>');
      expect(html).toContain('<th>Name</th>');
    });

    it('SheetJS: sheet_to_html also exists', () => {
      const xlsxWs = XLSX.utils.aoa_to_sheet([['Name'], ['Alice']]);
      const html = XLSX.utils.sheet_to_html(xlsxWs);
      expect(html).toContain('Name');
    });
  });

  describe('19e. Append utilities (modern-xlsx)', () => {
    it('sheetAddAoa appends data to existing sheet', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'Initial';

      sheetAddAoa(ws, [['Appended', 123]]);
      expect(ws.cell('A2').value).toBe('Appended');
      expect(ws.cell('B2').value).toBe(123);
    });

    it('sheetAddJson appends JSON objects', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'Header';

      sheetAddJson(ws, [{ x: 1, y: 2 }], { skipHeader: true });
      // Data should be appended after existing data
      expect(ws.rows.length).toBeGreaterThan(1);
    });
  });
});

// ==========================================================================
// 20. FORMAT CELL (modern-xlsx)
// ==========================================================================

describe('20. Number Formatting', () => {
  it('modern-xlsx formatCell applies number formats', () => {
    expect(formatCell(0.1234, '0.00%')).toBe('12.34%');
    expect(formatCell(1234567, '#,##0')).toBe('1,234,567');
    expect(formatCell(0.5, '0.00')).toBe('0.50');
  });

  it('SheetJS SSF also formats numbers', () => {
    expect(XLSX.SSF.format('0.00%', 0.1234)).toBe('12.34%');
    expect(XLSX.SSF.format('#,##0', 1234567)).toBe('1,234,567');
  });
});

// ==========================================================================
// 21. RICH TEXT (modern-xlsx)
// ==========================================================================

describe('21. Rich Text (modern-xlsx only)', () => {
  it('RichTextBuilder creates formatted text runs', () => {
    const builder = new RichTextBuilder()
      .bold('Important: ')
      .text('Normal text. ')
      .colored('Red text', 'FF0000')
      .styled('Custom', { bold: true, italic: true, fontSize: 14, fontName: 'Arial' });

    const runs = builder.build();
    expect(runs.length).toBe(4);

    const plain = builder.plainText();
    expect(plain).toBe('Important: Normal text. Red textCustom');

    // First run should be bold (fields are flat on RichTextRun, not nested under .font)
    expect(runs[0]?.bold).toBe(true);
    // Third run should have color
    expect(runs[2]?.color).toBe('FF0000');
  });
});

// ==========================================================================
// 22. BARCODE & QR CODE (modern-xlsx EXCLUSIVE)
// ==========================================================================

describe('22. Barcode & QR Code (modern-xlsx exclusive)', () => {
  it('encodeQR generates a barcode matrix', () => {
    const matrix = encodeQR('Hello World');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
    expect(matrix.modules.length).toBe(matrix.height);
    expect(matrix.modules[0]?.length).toBe(matrix.width);
  });

  it('renderBarcodePNG produces valid PNG bytes', () => {
    const matrix = encodeQR('Test', { ecLevel: 'M' });
    const png = renderBarcodePNG(matrix, { moduleSize: 4, quietZone: 2 });
    expect(png).toBeInstanceOf(Uint8Array);
    expect(png.length).toBeGreaterThan(0);
    // PNG magic: 0x89 0x50 0x4E 0x47
    expect(png[0]).toBe(0x89);
    expect(png[1]).toBe(0x50);
    expect(png[2]).toBe(0x4e);
    expect(png[3]).toBe(0x47);
  });

  it('addBarcode embeds QR code as image in XLSX', async () => {
    const wb = new Workbook();
    wb.addSheet('QR');
    wb.addBarcode('QR', { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 }, 'https://example.com', {
      type: 'qr',
      ecLevel: 'M',
    });

    const buf = await wb.toBuffer();
    expect(buf.length).toBeGreaterThan(0);

    // SheetJS can read the file (image won't be parsed but file is valid)
    const xlsxWb = XLSX.read(buf, { type: 'buffer' });
    expect(xlsxWb.SheetNames).toContain('QR');
  });

  it('SheetJS has NO barcode/QR code support', () => {
    // SheetJS community has no barcode generation capability
    // @ts-expect-error — proving the method doesn't exist
    expect(XLSX.utils.encode_qr).toBeUndefined();
  });
});

// ==========================================================================
// 23. TABLE LAYOUT ENGINE (modern-xlsx EXCLUSIVE)
// ==========================================================================

describe('23. Table Layout Engine (modern-xlsx exclusive)', () => {
  it('drawTableFromData creates a styled table with headers', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Table');

    const tableData = [
      { name: 'Alice', age: 30, city: 'NYC' },
      { name: 'Bob', age: 25, city: 'LA' },
    ];

    const result = drawTableFromData(wb, ws, tableData, {
      origin: 'A1',
      headerColor: '4472C4',
      freezeHeader: true,
      headerMap: { name: 'Name', age: 'Age', city: 'City' },
    });

    expect(result.range).toBeDefined();
    expect(ws.cell('A1').value).toBe('Name');
    expect(ws.cell('B1').value).toBe('Age');
    expect(ws.cell('A2').value).toBe('Alice');
    expect(ws.frozenPane?.rows).toBe(1);

    // Roundtrip
    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Table');
    expect(ws2?.cell('A1').value).toBe('Name');
    expect(Number(ws2?.cell('B2').value)).toBe(30);
  });

  it('SheetJS has NO built-in table layout engine', () => {
    // SheetJS requires manual cell-by-cell styling (Pro only) for table aesthetics
    // No equivalent to drawTableFromData
    // @ts-expect-error — proving the method doesn't exist
    expect(XLSX.utils.draw_table).toBeUndefined();
  });
});

// ==========================================================================
// 24. VALIDATION & REPAIR (modern-xlsx EXCLUSIVE)
// ==========================================================================

describe('24. Validation & Repair (modern-xlsx exclusive)', () => {
  it('validate checks workbook for OOXML compliance', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'test';
    const report = wb.validate();
    expect(report).toBeDefined();
    expect(report.issues).toBeDefined();
  });

  it('repair fixes repairable issues', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'test';
    const { workbook, report, repairCount } = wb.repair();
    expect(workbook).toBeDefined();
    expect(report).toBeDefined();
    expect(typeof repairCount).toBe('number');
  });
});
