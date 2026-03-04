import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Integration: Data Feature Roundtrips', () => {
  it('sparklines + data table in same sheet survives roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Add data
    ws.cell('A1').value = 10;
    ws.cell('A2').value = 20;
    ws.cell('A3').value = 30;

    // Add sparkline
    ws.addSparklineGroup({
      sparklineType: 'column',
      sparklines: [{ formula: 'Sheet1!A1:A3', sqref: 'B1' }],
      colorSeries: 'FF376092',
    });

    // Add data table formula via JSON manipulation
    const data = wb.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('missing sheet');
    // Add a data table cell to existing rows
    const row2 = sheet.worksheet.rows.find((r) => r.index === 2);
    if (row2) {
      row2.cells.push({
        reference: 'C2',
        cellType: 'formulaStr',
        styleIndex: null,
        value: '42',
        formula: '',
        formulaType: 'dataTable',
        formulaR1: 'A1',
      });
    }

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);
    const ws3 = wb3.getSheet('Sheet1');

    // Verify sparkline survived
    expect(ws3?.sparklineGroups).toHaveLength(1);
    expect(ws3?.sparklineGroups[0].sparklineType).toBe('column');
    expect(ws3?.sparklineGroups[0].colorSeries).toBe('FF376092');

    // Verify data table survived
    const readData = wb3.toJSON();
    const readRow2 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 2);
    const dtCell = readRow2?.cells.find((c) => c.reference === 'C2');
    expect(dtCell?.formulaType).toBe('dataTable');
    expect(dtCell?.formulaR1).toBe('A1');
  });

  it('sparkline groups on multiple sheets survive roundtrip', async () => {
    const wb = new Workbook();

    const ws1 = wb.addSheet('Sales');
    ws1.cell('A1').value = 100;
    ws1.cell('A2').value = 200;
    ws1.addSparklineGroup({
      sparklines: [{ formula: 'Sales!A1:A2', sqref: 'B1' }],
    });

    const ws2 = wb.addSheet('Profit');
    ws2.cell('A1').value = 50;
    ws2.cell('A2').value = 75;
    ws2.addSparklineGroup({
      sparklineType: 'stacked',
      sparklines: [{ formula: 'Profit!A1:A2', sqref: 'B1' }],
      negative: true,
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const sales = wb2.getSheet('Sales');
    expect(sales?.sparklineGroups).toHaveLength(1);
    expect(sales?.sparklineGroups[0].sparklines[0].formula).toBe('Sales!A1:A2');

    const profit = wb2.getSheet('Profit');
    expect(profit?.sparklineGroups).toHaveLength(1);
    expect(profit?.sparklineGroups[0].sparklineType).toBe('stacked');
    expect(profit?.sparklineGroups[0].negative).toBe(true);
  });

  it('data table + normal formula + sparkline coexist on same sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Normal formula
    ws.cell('A1').formula = 'SUM(B1:B5)';
    ws.cell('A1').value = '15';

    // Regular data
    for (let i = 1; i <= 5; i++) {
      ws.cell(`B${i}`).value = i;
    }

    // Sparkline on the data
    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!B1:B5', sqref: 'C1' }],
      lineWeight: 2.0,
    });

    // Data table formula via JSON
    const data = wb.toJSON();
    const sheet = data.sheets[0];
    if (!sheet) throw new Error('missing sheet');
    sheet.worksheet.rows.push({
      index: 7,
      cells: [
        {
          reference: 'D7',
          cellType: 'formulaStr',
          styleIndex: null,
          value: '99',
          formula: '',
          formulaType: 'dataTable',
          formulaR1: 'A1',
          formulaR2: 'B1',
          formulaDt2d: true,
        },
      ],
      height: null,
      hidden: false,
    });

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);
    const ws3 = wb3.getSheet('Sheet1');

    // Normal formula preserved
    expect(ws3?.cell('A1').formula).toBe('SUM(B1:B5)');

    // Sparkline preserved
    expect(ws3?.sparklineGroups).toHaveLength(1);
    expect(ws3?.sparklineGroups[0].lineWeight).toBe(2.0);

    // Data table preserved
    const readData = wb3.toJSON();
    const row7 = readData.sheets[0]?.worksheet.rows.find((r) => r.index === 7);
    const dtCell = row7?.cells.find((c) => c.reference === 'D7');
    expect(dtCell?.formulaType).toBe('dataTable');
    expect(dtCell?.formulaDt2d).toBe(true);
  });

  it('sparklines + preserved entries (external links) coexist', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 1;
    ws.cell('A2').value = 2;

    ws.addSparklineGroup({
      sparklines: [{ formula: 'Sheet1!A1:A2', sqref: 'B1' }],
    });

    // Add preserved external link
    const data = wb.toJSON();
    data.preservedEntries = {
      'xl/externalLinks/externalLink1.xml': Array.from(new TextEncoder().encode('<externalLink/>')),
    };

    const wb2 = new Workbook(data);
    const buffer = await wb2.toBuffer();
    const wb3 = await readBuffer(buffer);

    // Sparkline survived
    const ws3 = wb3.getSheet('Sheet1');
    expect(ws3?.sparklineGroups).toHaveLength(1);

    // External link survived
    const readData = wb3.toJSON();
    expect(readData.preservedEntries?.['xl/externalLinks/externalLink1.xml']).toBeDefined();
  });

  it('clone sheet preserves sparklines', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Original');
    ws.cell('A1').value = 10;
    ws.addSparklineGroup({
      sparklineType: 'column',
      sparklines: [{ formula: 'Original!A1:A5', sqref: 'B1' }],
      high: true,
      low: true,
    });

    wb.cloneSheet(0, 'Copy');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);

    const orig = wb2.getSheet('Original');
    expect(orig?.sparklineGroups).toHaveLength(1);
    expect(orig?.sparklineGroups[0].high).toBe(true);

    const copy = wb2.getSheet('Copy');
    expect(copy?.sparklineGroups).toHaveLength(1);
    expect(copy?.sparklineGroups[0].sparklineType).toBe('column');
  });

  it('performance: 50-sheet workbook with sparklines roundtrips in reasonable time', async () => {
    const wb = new Workbook();

    for (let i = 0; i < 50; i++) {
      const ws = wb.addSheet(`Sheet${i + 1}`);
      for (let r = 1; r <= 10; r++) {
        ws.cell(`A${r}`).value = r * (i + 1);
      }
      ws.addSparklineGroup({
        sparklines: [{ formula: `Sheet${i + 1}!A1:A10`, sqref: 'B1' }],
      });
    }

    const start = performance.now();
    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const elapsed = performance.now() - start;

    expect(wb2.sheetCount).toBe(50);
    expect(wb2.getSheet('Sheet1')?.sparklineGroups).toHaveLength(1);
    expect(wb2.getSheet('Sheet50')?.sparklineGroups).toHaveLength(1);
    // Should complete well within 10 seconds
    expect(elapsed).toBeLessThan(10000);
  });
});
