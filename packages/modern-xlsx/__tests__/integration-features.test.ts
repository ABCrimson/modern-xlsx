import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Integration: Cross-Feature Roundtrips', () => {
  it('frozen pane + hidden gridlines + zoom survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.frozenPane = { rows: 2, cols: 1 };
    ws.view = { showGridLines: false, zoomScale: 150, view: 'pageLayout' };
    ws.pageSetup = { orientation: 'landscape', paperSize: 9 };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.frozenPane?.rows).toBe(2);
    expect(ws2?.frozenPane?.cols).toBe(1);
    expect(ws2?.view?.showGridLines).toBe(false);
    expect(ws2?.view?.zoomScale).toBe(150);
    expect(ws2?.view?.view).toBe('pageLayout');
    expect(ws2?.pageSetup?.orientation).toBe('landscape');
  });

  it('split pane on one sheet + frozen on another survives roundtrip', async () => {
    const wb = new Workbook();
    const ws1 = wb.addSheet('Split');
    ws1.cell('A1').value = 'split';
    ws1.splitPane = {
      horizontal: 2400,
      topLeftCell: 'A5',
      activePane: 'bottomLeft',
    };

    const ws2 = wb.addSheet('Frozen');
    ws2.cell('A1').value = 'frozen';
    ws2.frozenPane = { rows: 1, cols: 0 };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Split')?.splitPane?.horizontal).toBe(2400);
    expect(wb2.getSheet('Split')?.splitPane?.topLeftCell).toBe('A5');
    expect(wb2.getSheet('Split')?.splitPane?.activePane).toBe('bottomLeft');
    expect(wb2.getSheet('Split')?.frozenPane).toBeFalsy();
    expect(wb2.getSheet('Frozen')?.frozenPane?.rows).toBe(1);
    expect(wb2.getSheet('Frozen')?.splitPane).toBeFalsy();
  });

  it('workbook protection + hidden sheet + sheet protection survives roundtrip', async () => {
    const wb = new Workbook();
    const ws1 = wb.addSheet('Visible');
    ws1.cell('A1').value = 'public';
    ws1.sheetProtection = {
      sheet: true,
      objects: true,
      scenarios: true,
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

    const ws2 = wb.addSheet('Hidden');
    ws2.cell('A1').value = 'secret';
    ws2.state = 'hidden';

    wb.protection = { lockStructure: true };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.protection?.lockStructure).toBe(true);
    expect(wb2.getSheet('Hidden')?.state).toBe('hidden');
    expect(wb2.getSheet('Visible')?.sheetProtection?.sheet).toBe(true);
    expect(wb2.getSheet('Visible')?.sheetProtection?.scenarios).toBe(true);
  });

  it('clone sheet preserves all features through roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Original');
    ws.cell('A1').value = 'hello';
    ws.frozenPane = { rows: 1, cols: 0 };
    ws.view = { zoomScale: 120 };
    ws.tabColor = 'FF0000';

    wb.cloneSheet(0, 'Copy');
    wb.moveSheet(1, 0); // move Copy to front

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetNames).toEqual(['Copy', 'Original']);
    const copy = wb2.getSheet('Copy');
    expect(copy?.cell('A1').value).toBe('hello');
    expect(copy?.frozenPane?.rows).toBe(1);
    expect(copy?.view?.zoomScale).toBe(120);
    expect(copy?.tabColor).toBe('FF0000');
  });

  it('print-ready: margins + headers/footers + print titles survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('PrintReady');
    ws.cell('A1').value = 'Title Row';
    ws.cell('A2').value = 'Data';
    ws.viewMode = 'pageBreakPreview';
    ws.pageSetup = { orientation: 'portrait', paperSize: 9 };
    ws.headerFooter = {
      oddHeader: '&CPage &P of &N',
      oddFooter: '&L&D',
    };

    wb.setPrintTitles(0, 'PrintReady!$1:$1');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('PrintReady');
    expect(ws2?.viewMode).toBe('pageBreakPreview');
    expect(ws2?.pageSetup?.orientation).toBe('portrait');
    expect(ws2?.pageSetup?.paperSize).toBe(9);
    expect(ws2?.headerFooter?.oddHeader).toContain('&P');
    expect(ws2?.headerFooter?.oddFooter).toContain('&D');
    const titles = wb2.getPrintTitles(0);
    expect(titles).toContain('$1:$1');
  });
});
