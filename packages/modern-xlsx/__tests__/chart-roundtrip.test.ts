import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
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

    const charts = ws2?.charts;
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
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

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
    const ws2 = wb2.getSheet('Data') as Worksheet;

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
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

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
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

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
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

    // Cell data preserved
    expect(ws2.cell('A1').value).toBe('Hello');
    // Chart data preserved
    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0].chart.grouping).toBe('stacked');
  });

  it('preserves axis tick label font size on roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'X';
    ws.cell('B1').value = '10';

    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Data!$B$1:$B$1', catRef: 'Data!$A$1:$A$1' })
        .catAxis({ title: 'Category', fontSize: 1400 })
        .valAxis({ title: 'Value', fontSize: 1200 })
        .anchor({ col: 3, row: 0 }, { col: 10, row: 15 });
    });

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Data');
    const chart = ws2?.charts?.[0];

    expect(chart?.chart.catAxis?.fontSize).toBe(1400);
    expect(chart?.chart.valAxis?.fontSize).toBe(1200);
  });

  it('roundtrips a bubble chart with bubbleSizeRef and xValRef', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'X';
    ws.cell('B1').value = 'Y';
    ws.cell('C1').value = 'Size';
    ws.cell('A2').value = '1';
    ws.cell('B2').value = '10';
    ws.cell('C2').value = '5';
    ws.cell('A3').value = '2';
    ws.cell('B3').value = '20';
    ws.cell('C3').value = '8';

    ws.addChart('bubble', (b) => {
      b.title('Bubble Chart')
        .addSeries({
          name: 'Samples',
          xValRef: 'Data!$A$2:$A$3',
          valRef: 'Data!$B$2:$B$3',
          bubbleSizeRef: 'Data!$C$2:$C$3',
        })
        .anchor({ col: 4, row: 0 }, { col: 14, row: 18 });
    });

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].chart.chartType).toBe('bubble');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    const chart = ws2.charts[0];
    expect(chart.chart.chartType).toBe('bubble');
    expect(chart.chart.series).toHaveLength(1);
    expect(chart.chart.series[0].xValRef).toBe('Data!$A$2:$A$3');
    expect(chart.chart.series[0].valRef).toBe('Data!$B$2:$B$3');
    expect(chart.chart.series[0].bubbleSizeRef).toBe('Data!$C$2:$C$3');
  });

  it('roundtrips a stock chart with 3 HLC series', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Stock');
    ws.cell('A1').value = 'Date';
    ws.cell('B1').value = 'High';
    ws.cell('C1').value = 'Low';
    ws.cell('D1').value = 'Close';
    ws.cell('A2').value = '2026-01-01';
    ws.cell('B2').value = '150';
    ws.cell('C2').value = '140';
    ws.cell('D2').value = '145';
    ws.cell('A3').value = '2026-01-02';
    ws.cell('B3').value = '155';
    ws.cell('C3').value = '142';
    ws.cell('D3').value = '148';

    ws.addChart('stock', (b) => {
      b.title('HLC Stock Chart')
        .addSeries({
          name: 'High',
          catRef: 'Stock!$A$2:$A$3',
          valRef: 'Stock!$B$2:$B$3',
        })
        .addSeries({
          name: 'Low',
          catRef: 'Stock!$A$2:$A$3',
          valRef: 'Stock!$C$2:$C$3',
        })
        .addSeries({
          name: 'Close',
          catRef: 'Stock!$A$2:$A$3',
          valRef: 'Stock!$D$2:$D$3',
        })
        .anchor({ col: 5, row: 0 }, { col: 15, row: 20 });
    });

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].chart.chartType).toBe('stock');
    expect(ws.charts[0].chart.series).toHaveLength(3);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Stock') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    const chart = ws2.charts[0];
    expect(chart.chart.chartType).toBe('stock');
    expect(chart.chart.series).toHaveLength(3);
    expect(chart.chart.series[0].name).toBe('High');
    expect(chart.chart.series[1].name).toBe('Low');
    expect(chart.chart.series[2].name).toBe('Close');
  });

  it('roundtrips a oneCellAnchor chart through WASM write/read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = '100';

    // Use addChartData to set oneCellAnchor via extCx/extCy.
    ws.addChartData({
      chart: {
        chartType: 'bar',
        title: { text: 'OneCellTest', overlay: false },
        series: [
          {
            idx: 0,
            order: 0,
            name: 'Revenue',
            catRef: 'Sheet1!$A$2:$A$2',
            valRef: 'Sheet1!$B$2:$B$2',
          },
        ],
        showDataTable: false,
      },
      anchor: {
        fromCol: 3,
        fromRow: 5,
        fromColOff: 100,
        fromRowOff: 200,
        toCol: 0,
        toRow: 0,
        extCx: 5400000,
        extCy: 3240000,
      },
    });

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].anchor.extCx).toBe(5400000);
    expect(ws.charts[0].anchor.extCy).toBe(3240000);

    // Write to buffer.
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Read back.
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;
    expect(ws2).toBeDefined();

    const charts = ws2.charts;
    expect(charts).toHaveLength(1);
    expect(charts[0].chart.chartType).toBe('bar');
    expect(charts[0].chart.title?.text).toBe('OneCellTest');

    // Verify the oneCellAnchor fields survived the roundtrip.
    expect(charts[0].anchor.fromCol).toBe(3);
    expect(charts[0].anchor.fromRow).toBe(5);
    expect(charts[0].anchor.fromColOff).toBe(100);
    expect(charts[0].anchor.fromRowOff).toBe(200);
    expect(charts[0].anchor.extCx).toBe(5400000);
    expect(charts[0].anchor.extCy).toBe(3240000);
  });
});
