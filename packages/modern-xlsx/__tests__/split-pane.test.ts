import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Split Pane', () => {
  it('set and get horizontal split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = {
      horizontal: 2400,
      topLeftCell: 'A5',
      activePane: 'bottomLeft',
    };
    expect(ws.splitPane).toEqual({
      horizontal: 2400,
      topLeftCell: 'A5',
      activePane: 'bottomLeft',
    });
  });

  it('set and get vertical split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = {
      vertical: 3000,
      topLeftCell: 'D1',
      activePane: 'topRight',
    };
    expect(ws.splitPane?.vertical).toBe(3000);
  });

  it('set and get four-way split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = {
      horizontal: 2400,
      vertical: 3000,
      topLeftCell: 'D5',
      activePane: 'bottomRight',
    };
    ws.paneSelections = [
      { pane: 'topRight', activeCell: 'D1', sqref: 'D1' },
      { pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' },
      { pane: 'bottomRight', activeCell: 'D5', sqref: 'D5' },
    ];
    expect(ws.paneSelections).toHaveLength(3);
  });

  it('split pane clears frozen pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.frozenPane = { rows: 1, cols: 0 };
    expect(ws.frozenPane).toBeTruthy();
    ws.splitPane = { horizontal: 2400 };
    expect(ws.frozenPane).toBeNull();
    expect(ws.splitPane).toBeTruthy();
  });

  it('frozen pane clears split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400 };
    expect(ws.splitPane).toBeTruthy();
    ws.frozenPane = { rows: 1, cols: 0 };
    expect(ws.splitPane).toBeNull();
    expect(ws.frozenPane).toBeTruthy();
  });

  it('horizontal split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = {
      horizontal: 2400,
      topLeftCell: 'A5',
      activePane: 'bottomLeft',
    };
    ws.paneSelections = [
      { pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.horizontal).toBe(2400);
    expect(ws2.splitPane?.topLeftCell).toBe('A5');
    expect(ws2.paneSelections).toHaveLength(1);
  });

  it('vertical split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = {
      vertical: 3000,
      topLeftCell: 'D1',
      activePane: 'topRight',
    };

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.vertical).toBe(3000);
    expect(ws2.splitPane?.topLeftCell).toBe('D1');
  });

  it('four-way split pane survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.splitPane = {
      horizontal: 2400,
      vertical: 3000,
      topLeftCell: 'D5',
      activePane: 'bottomRight',
    };
    ws.paneSelections = [
      { pane: 'topRight', activeCell: 'D1', sqref: 'D1' },
      { pane: 'bottomLeft', activeCell: 'A5', sqref: 'A5' },
      { pane: 'bottomRight', activeCell: 'D5', sqref: 'D5' },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;
    expect(ws2.splitPane?.horizontal).toBe(2400);
    expect(ws2.splitPane?.vertical).toBe(3000);
    expect(ws2.paneSelections).toHaveLength(3);
  });

  it('clear split pane', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.splitPane = { horizontal: 2400 };
    ws.splitPane = null;
    expect(ws.splitPane).toBeNull();
  });

  it('default (no pane) - both null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.frozenPane).toBeNull();
    expect(ws.splitPane).toBeNull();
    expect(ws.paneSelections).toEqual([]);
  });
});
