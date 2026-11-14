import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const sheetXml = async (wb: Workbook): Promise<string> =>
  readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'] ?? '';

describe('row height', () => {
  it('writes ht with customHeight="1" for setRowHeight', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'tall';
    ws.setRowHeight(1, 40);
    expect(await sheetXml(wb)).toContain('<row r="1" ht="40" customHeight="1">');
  });

  it('reads customHeight back as true and keeps it through a second write', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'tall';
    ws.setRowHeight(1, 40);
    const back = await readBuffer(await wb.toBuffer());
    const row = back.getSheet('Sheet1')?.rows[0];
    expect(row?.height).toBe(40);
    expect(row?.customHeight).toBe(true);
    expect(await sheetXml(back)).toContain('<row r="1" ht="40" customHeight="1">');
  });

  it('keeps a row height read without customHeight an auto-fit height', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'x';
    ws.setRowHeight(1, 40);
    const row = ws.rows[0];
    expect(row).toBeDefined();
    if (row) delete row.customHeight;
    const xml = await sheetXml(wb);
    expect(xml).toContain('<row r="1" ht="40">');
    expect(xml).not.toContain('customHeight');
    const back = await readBuffer(await wb.toBuffer());
    const read = back.getSheet('Sheet1')?.rows[0];
    expect(read?.height).toBe(40);
    expect(read).not.toHaveProperty('customHeight');
  });

  it("writes Excel's default sheet view for a sheet without one, and reads it back as none", async () => {
    // Without any <sheetView>, Excel on a 150% display (144 DPI) showed setRowHeight(1, 40)
    // as 26.7 points; with this view it shows 40.
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'tall';
    ws.setRowHeight(1, 40);
    const xml = await sheetXml(wb);
    expect(xml).toContain(
      '<dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData>',
    );
    expect(ws.view).toBeNull();

    const back = await readBuffer(await wb.toBuffer());
    const read = back.getSheet('Sheet1');
    expect(read?.view).toBeNull();
    expect(read?.frozenPane ?? null).toBeNull();
    expect(read?.rows[0]?.height).toBe(40);
    expect(await sheetXml(back)).toBe(xml);
  });

  it('carries customHeight into a cloned sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'x';
    ws.setRowHeight(1, 30);
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.rows[0]?.customHeight).toBe(true);
  });
});
