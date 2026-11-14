import { describe, expect, it } from 'vitest';
import type { OutlinePropertiesData } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

// ---------------------------------------------------------------------------
// Outline Properties
// ---------------------------------------------------------------------------

describe('Outline Properties', () => {
  it('sets and gets outline properties', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const props: OutlinePropertiesData = {
      summaryBelow: false,
      summaryRight: false,
    };

    ws.outlineProperties = props;
    expect(ws.outlineProperties).toEqual(props);
  });

  it('defaults to null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.outlineProperties).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// Row / Column Grouping
// ---------------------------------------------------------------------------

describe('Row Grouping', () => {
  it('groups rows with outline level', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';

    ws.groupRows(2, 3, 1);

    const rows = ws.rows;
    const row2 = rows.find((r) => r.index === 2);
    const row3 = rows.find((r) => r.index === 3);
    expect(row2?.outlineLevel).toBe(1);
    expect(row3?.outlineLevel).toBe(1);
  });

  it('ungroups rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.groupRows(2, 3, 2);
    ws.ungroupRows(2, 3);

    const rows = ws.rows;
    const row2 = rows.find((r) => r.index === 2);
    expect(row2?.outlineLevel).toBeNull();
    expect(row2?.collapsed).toBe(false);
  });

  it('collapses and expands rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';
    ws.cell('A4').value = 'Summary';

    ws.groupRows(2, 3, 1);
    ws.collapseRows(2, 3);

    const rows = ws.rows;
    expect(rows.find((r) => r.index === 2)?.hidden).toBe(true);
    expect(rows.find((r) => r.index === 3)?.hidden).toBe(true);
    expect(rows.find((r) => r.index === 4)?.collapsed).toBe(true);

    ws.expandRows(2, 3);
    expect(rows.find((r) => r.index === 2)?.hidden).toBe(false);
    expect(rows.find((r) => r.index === 3)?.hidden).toBe(false);
    expect(rows.find((r) => r.index === 4)?.collapsed).toBe(false);
  });
});

describe('Column Grouping', () => {
  it('groups columns with outline level', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.groupColumns(2, 3, 1);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    const col3 = cols.find((c) => c.min === 3);
    expect(col2?.outlineLevel).toBe(1);
    expect(col3?.outlineLevel).toBe(1);
  });

  it('ungroups columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.groupColumns(2, 3, 2);
    ws.ungroupColumns(2, 3);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.outlineLevel).toBeNull();
    expect(col2?.collapsed).toBe(false);
  });

  it('updates existing column definitions', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.setColumnWidth(2, 20);
    ws.groupColumns(2, 2, 1);

    const cols = ws.columns;
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.width).toBe(20);
    expect(col2?.outlineLevel).toBe(1);
  });
});

describe('Roundtrip: Outline Properties', () => {
  it('roundtrips outline properties through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.outlineProperties = { summaryBelow: false, summaryRight: true };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.outlineProperties?.summaryBelow).toBe(false);
    expect(ws2?.outlineProperties?.summaryRight).toBe(true);
  });
});

describe('Roundtrip: Row/Column Grouping', () => {
  it('roundtrips a collapsed row group through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    for (const r of [1, 2, 3, 4]) ws.cell(`A${r}`).value = r;
    ws.groupRows(2, 3, 1);
    ws.collapseRows(2, 3);

    const buf = await wb.toBuffer();
    const rows = (await readBuffer(buf)).getSheet('Sheet1')?.rows ?? [];
    for (const index of [2, 3]) {
      const row = rows.find((r) => r.index === index);
      expect(row?.hidden).toBe(true);
      expect(row?.outlineLevel).toBe(1);
    }
    expect(rows.find((r) => r.index === 4)?.collapsed).toBe(true);

    const xml = readZipText(buf)['xl/worksheets/sheet1.xml'];
    expect(xml).toContain('<row r="2" hidden="1" outlineLevel="1">');
    expect(xml).toContain('<row r="4" collapsed="1">');
  });

  it('roundtrips row outline levels through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Header';
    ws.cell('A2').value = 'Detail 1';
    ws.cell('A3').value = 'Detail 2';

    ws.groupRows(2, 3, 2);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');
    const rows = ws2?.rows ?? [];
    const row2 = rows.find((r) => r.index === 2);
    expect(row2?.outlineLevel).toBe(2);
  });

  it('roundtrips column outline levels through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';

    ws.groupColumns(2, 3, 1);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');
    const cols = ws2?.columns ?? [];
    const col2 = cols.find((c) => c.min === 2);
    expect(col2?.outlineLevel).toBe(1);
  });
});
