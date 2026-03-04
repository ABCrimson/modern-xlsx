import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Sparklines', () => {
  it('add line sparkline group', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSparklineGroup({
      sparklines: [
        { formula: 'Sheet1!A1:A10', sqref: 'B1' },
        { formula: 'Sheet1!A1:A10', sqref: 'B2' },
      ],
    });
    expect(ws.sparklineGroups).toHaveLength(1);
    expect(ws.sparklineGroups[0].sparklines).toHaveLength(2);
  });

  it('add column sparkline group with colors', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSparklineGroup({
      sparklineType: 'column',
      sparklines: [{ formula: 'Sheet1!A1:A5', sqref: 'C1' }],
      colorSeries: 'FF376092',
      colorNegative: 'FFC00000',
      markers: true,
      high: true,
      low: true,
    });
    expect(ws.sparklineGroups[0].sparklineType).toBe('column');
    expect(ws.sparklineGroups[0].colorSeries).toBe('FF376092');
  });

  it('add win/loss sparkline with markers', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSparklineGroup({
      sparklineType: 'stacked',
      sparklines: [{ formula: 'Sheet1!D1:D10', sqref: 'E1' }],
      markers: true,
      negative: true,
      first: true,
      last: true,
    });
    expect(ws.sparklineGroups[0].sparklineType).toBe('stacked');
    expect(ws.sparklineGroups[0].markers).toBe(true);
  });

  it('line sparkline roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 1;
    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!A1:A10', sqref: 'B1' }],
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.sparklineGroups).toHaveLength(1);
    expect(ws2?.sparklineGroups[0].sparklines[0].formula).toBe('Sheet1!A1:A10');
    expect(ws2?.sparklineGroups[0].sparklines[0].sqref).toBe('B1');
  });

  it('column sparkline with colors roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 5;
    ws.addSparklineGroup({
      sparklineType: 'column',
      sparklines: [{ formula: 'Sheet1!A1:A5', sqref: 'C1' }],
      colorSeries: 'FF376092',
      colorNegative: 'FFC00000',
      markers: true,
      high: true,
      low: true,
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.sparklineGroups[0].sparklineType).toBe('column');
    expect(ws2?.sparklineGroups[0].colorSeries).toBe('FF376092');
    expect(ws2?.sparklineGroups[0].colorNegative).toBe('FFC00000');
    expect(ws2?.sparklineGroups[0].markers).toBe(true);
    expect(ws2?.sparklineGroups[0].high).toBe(true);
    expect(ws2?.sparklineGroups[0].low).toBe(true);
  });

  it('multiple sparkline groups roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 1;
    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!A1:A10', sqref: 'B1' }],
    });
    ws.addSparklineGroup({
      sparklineType: 'column',
      sparklines: [{ formula: 'Sheet1!C1:C10', sqref: 'D1' }],
      negative: true,
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.sparklineGroups).toHaveLength(2);
    expect(ws2?.sparklineGroups[1].sparklineType).toBe('column');
    expect(ws2?.sparklineGroups[1].negative).toBe(true);
  });

  it('clear sparkline groups', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!A1:A10', sqref: 'B1' }],
    });
    expect(ws.sparklineGroups).toHaveLength(1);
    ws.clearSparklineGroups();
    expect(ws.sparklineGroups).toHaveLength(0);
  });

  it('sparkline with manual min/max roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 1;
    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!A1:A10', sqref: 'B1' }],
      manualMin: 0,
      manualMax: 100,
      lineWeight: 1.25,
      displayEmptyCellsAs: 'gap',
      displayXAxis: true,
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2?.sparklineGroups[0].manualMin).toBe(0);
    expect(ws2?.sparklineGroups[0].manualMax).toBe(100);
    expect(ws2?.sparklineGroups[0].lineWeight).toBe(1.25);
    expect(ws2?.sparklineGroups[0].displayEmptyCellsAs).toBe('gap');
    expect(ws2?.sparklineGroups[0].displayXAxis).toBe(true);
  });
});
