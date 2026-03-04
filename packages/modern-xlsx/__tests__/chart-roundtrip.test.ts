import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Chart roundtrip', () => {
  it('roundtrips a bar chart through WASM write/read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = '100';
    ws.cell('A3').value = 'Q2';
    ws.cell('B3').value = '200';

    ws.addChart('bar', (b) => {
      b.title('Sales')
        .addSeries({
          name: 'Revenue',
          catRef: 'Sheet1!$A$2:$A$3',
          valRef: 'Sheet1!$B$2:$B$3',
          fillColor: '4472C4',
        })
        .catAxis({ title: 'Quarter' })
        .valAxis({ title: 'Amount', majorGridlines: true })
        .legend('bottom')
        .grouping('clustered')
        .anchor({ col: 4, row: 0 }, { col: 12, row: 15 });
    });

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].chart.chartType).toBe('bar');

    // Write to buffer
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Read back
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    const charts = ws2!.charts;
    expect(charts).toHaveLength(1);
    expect(charts[0].chart.chartType).toBe('bar');
    expect(charts[0].chart.title?.text).toBe('Sales');
    expect(charts[0].chart.series).toHaveLength(1);
    expect(charts[0].chart.series[0].valRef).toBe('Sheet1!$B$2:$B$3');
    expect(charts[0].chart.series[0].fillColor).toBe('4472C4');
    expect(charts[0].chart.grouping).toBe('clustered');
    expect(charts[0].chart.legend?.position).toBe('bottom');
    expect(charts[0].anchor.fromCol).toBe(4);
    expect(charts[0].anchor.toCol).toBe(12);
  });

  it('roundtrips a pie chart', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Slice';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'A';
    ws.cell('B2').value = '30';
    ws.cell('A3').value = 'B';
    ws.cell('B3').value = '70';

    ws.addChart('pie', (b) => {
      b.title('Distribution')
        .addSeries({
          valRef: 'Sheet1!$B$2:$B$3',
          catRef: 'Sheet1!$A$2:$A$3',
          explosion: 10,
        })
        .dataLabels({ showPercent: true, showVal: true })
        .legend('right');
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;

    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0].chart.chartType).toBe('pie');
    expect(ws2.charts[0].chart.series[0].explosion).toBe(10);
    expect(ws2.charts[0].chart.dataLabels?.showPercent).toBe(true);
  });

  it('roundtrips a scatter chart', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');

    ws.addChart('scatter', (b) => {
      b.scatterStyle('lineMarker')
        .addSeries({
          xValRef: 'Data!$A$2:$A$5',
          valRef: 'Data!$B$2:$B$5',
          marker: 'circle',
          smooth: true,
        })
        .valAxis({ majorGridlines: true });
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data')!;

    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0].chart.chartType).toBe('scatter');
    expect(ws2.charts[0].chart.scatterStyle).toBe('lineMarker');
    expect(ws2.charts[0].chart.series[0].marker).toBe('circle');
  });

  it('roundtrips multiple charts on one sheet', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addChart('line', (b) => {
      b.addSeries({ valRef: 'Sheet1!$A$1:$A$5' }).anchor({ col: 0, row: 0 }, { col: 8, row: 15 });
    });
    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Sheet1!$B$1:$B$5' }).anchor({ col: 10, row: 0 }, { col: 18, row: 15 });
    });

    expect(ws.charts).toHaveLength(2);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;

    expect(ws2.charts).toHaveLength(2);
    expect(ws2.charts[0].chart.chartType).toBe('line');
    expect(ws2.charts[1].chart.chartType).toBe('bar');
  });

  it('roundtrips a line chart with style', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addChart('line', (b) => {
      b.title('Trend')
        .addSeries({ name: 'Sales', valRef: 'Sheet1!$A$1:$A$5', lineColor: 'FF0000' })
        .style(26);
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;

    expect(ws2.charts[0].chart.styleId).toBe(26);
    expect(ws2.charts[0].chart.series[0].lineColor).toBe('FF0000');
  });

  it('preserves chart data alongside cell data', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Hello';
    ws.cell('B1').value = '42';

    ws.addChart('column', (b) => {
      b.addSeries({ valRef: 'Sheet1!$B$1:$B$5' }).grouping('stacked');
    });

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1')!;

    // Cell data preserved
    expect(ws2.cell('A1').value).toBe('Hello');
    // Chart data preserved
    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0].chart.grouping).toBe('stacked');
  });
});
