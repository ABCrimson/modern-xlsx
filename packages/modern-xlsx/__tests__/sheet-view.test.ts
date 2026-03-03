import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Sheet View', () => {
  it('set and get view with hidden gridlines', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.view = { showGridLines: false };
    expect(ws.view?.showGridLines).toBe(false);
  });

  it('set and get view with zoom', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.view = { zoomScale: 150 };
    expect(ws.view?.zoomScale).toBe(150);
  });

  it('set and get RTL view', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.view = { rightToLeft: true, showRowColHeaders: false };
    expect(ws.view?.rightToLeft).toBe(true);
    expect(ws.view?.showRowColHeaders).toBe(false);
  });

  it('viewMode convenience getter/setter', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.viewMode).toBe('normal');
    ws.viewMode = 'pageBreakPreview';
    expect(ws.viewMode).toBe('pageBreakPreview');
    ws.viewMode = 'pageLayout';
    expect(ws.viewMode).toBe('pageLayout');
    ws.viewMode = 'normal';
    expect(ws.viewMode).toBe('normal');
  });

  it('hidden gridlines + zoom survive roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.view = { showGridLines: false, zoomScale: 150 };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.view?.showGridLines).toBe(false);
    expect(ws2.view?.zoomScale).toBe(150);
  });

  it('RTL view survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.view = { rightToLeft: true, showRowColHeaders: false };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.view?.rightToLeft).toBe(true);
    expect(ws2.view?.showRowColHeaders).toBe(false);
  });

  it('page break preview mode survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.view = {
      view: 'pageBreakPreview',
      zoomScale: 60,
      zoomScaleNormal: 100,
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.view?.view).toBe('pageBreakPreview');
    expect(ws2.view?.zoomScale).toBe(60);
    expect(ws2.view?.zoomScaleNormal).toBe(100);
  });

  it('page layout with custom zoom survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.view = {
      view: 'pageLayout',
      showRuler: false,
      showWhiteSpace: false,
      zoomScalePageLayoutView: 75,
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.view?.view).toBe('pageLayout');
    expect(ws2.view?.showRuler).toBe(false);
    expect(ws2.view?.showWhiteSpace).toBe(false);
    expect(ws2.view?.zoomScalePageLayoutView).toBe(75);
  });

  it('sheet view combined with frozen pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'header';
    ws.view = { showGridLines: false, zoomScale: 120 };
    ws.frozenPane = { rows: 1, cols: 0 };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.view?.showGridLines).toBe(false);
    expect(ws2.view?.zoomScale).toBe(120);
    expect(ws2.frozenPane).toEqual({ rows: 1, cols: 0 });
  });

  it('clear view resets to null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.view = { showGridLines: false };
    ws.view = null;
    expect(ws.view).toBeNull();
    expect(ws.viewMode).toBe('normal');
  });
});
