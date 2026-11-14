import { assert, describe, expect, it } from 'vitest';
import type { Worksheet } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';

/** Element `index` of `items`, narrowed with `assert.isDefined` (noUncheckedIndexedAccess). */
function at<T>(items: readonly T[] | undefined, index: number): T {
  const item = items?.[index];
  assert.isDefined(item);
  return item;
}

describe('Chart roundtrip', () => {
  it('roundtrips a bar chart through WASM write/read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = 100;
    ws.cell('A3').value = 'Q2';
    ws.cell('B3').value = 200;

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
    expect(at(ws.charts, 0).chart.chartType).toBe('bar');

    // Write to buffer
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Read back
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    const charts = ws2?.charts;
    expect(charts).toHaveLength(1);
    expect(at(charts, 0).chart.chartType).toBe('bar');
    expect(at(charts, 0).chart.title?.text).toBe('Sales');
    expect(at(charts, 0).chart.series).toHaveLength(1);
    expect(at(at(charts, 0).chart.series, 0).valRef).toBe('Sheet1!$B$2:$B$3');
    expect(at(at(charts, 0).chart.series, 0).fillColor).toBe('4472C4');
    expect(at(charts, 0).chart.grouping).toBe('clustered');
    expect(at(charts, 0).chart.legend?.position).toBe('bottom');
    expect(at(charts, 0).anchor.fromCol).toBe(4);
    expect(at(charts, 0).anchor.toCol).toBe(12);
    expect(at(charts, 0).chart.catAxis?.title?.text).toBe('Quarter');
    expect(at(charts, 0).chart.valAxis?.title?.text).toBe('Amount');
    expect(at(charts, 0).chart.valAxis?.majorGridlines).toBe(true);
  });

  it('roundtrips a pie chart', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Slice';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'A';
    ws.cell('B2').value = 30;
    ws.cell('A3').value = 'B';
    ws.cell('B3').value = 70;

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
    expect(at(ws2.charts, 0).chart.chartType).toBe('pie');
    expect(at(at(ws2.charts, 0).chart.series, 0).explosion).toBe(10);
    expect(at(ws2.charts, 0).chart.dataLabels?.showPercent).toBe(true);
    expect(at(ws2.charts, 0).chart.dataLabels?.showVal).toBe(true);
    expect(at(ws2.charts, 0).chart.legend?.position).toBe('right');
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
    expect(at(ws2.charts, 0).chart.chartType).toBe('scatter');
    expect(at(ws2.charts, 0).chart.scatterStyle).toBe('lineMarker');
    expect(at(at(ws2.charts, 0).chart.series, 0).marker).toBe('circle');
    expect(at(at(ws2.charts, 0).chart.series, 0).smooth).toBe(true);
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
    expect(at(ws2.charts, 0).chart.chartType).toBe('line');
    expect(at(ws2.charts, 1).chart.chartType).toBe('bar');
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

    expect(at(ws2.charts, 0).chart.styleId).toBe(26);
    expect(at(at(ws2.charts, 0).chart.series, 0).lineColor).toBe('FF0000');
  });

  it('preserves chart data alongside cell data', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Hello';
    ws.cell('B1').value = 42;

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
    expect(at(ws2.charts, 0).chart.grouping).toBe('stacked');
  });

  it('preserves axis tick label font size on roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'X';
    ws.cell('B1').value = 10;

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
    ws.cell('A2').value = 1;
    ws.cell('B2').value = 10;
    ws.cell('C2').value = 5;
    ws.cell('A3').value = 2;
    ws.cell('B3').value = 20;
    ws.cell('C3').value = 8;

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
    expect(at(ws.charts, 0).chart.chartType).toBe('bubble');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Data') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    const chart = at(ws2.charts, 0);
    expect(chart.chart.chartType).toBe('bubble');
    expect(chart.chart.series).toHaveLength(1);
    expect(at(chart.chart.series, 0).xValRef).toBe('Data!$A$2:$A$3');
    expect(at(chart.chart.series, 0).valRef).toBe('Data!$B$2:$B$3');
    expect(at(chart.chart.series, 0).bubbleSizeRef).toBe('Data!$C$2:$C$3');
  });

  it('roundtrips a stock chart with 3 HLC series', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Stock');
    ws.cell('A1').value = 'Date';
    ws.cell('B1').value = 'High';
    ws.cell('C1').value = 'Low';
    ws.cell('D1').value = 'Close';
    ws.cell('A2').value = '2026-01-01';
    ws.cell('B2').value = 150;
    ws.cell('C2').value = 140;
    ws.cell('D2').value = 145;
    ws.cell('A3').value = '2026-01-02';
    ws.cell('B3').value = 155;
    ws.cell('C3').value = 142;
    ws.cell('D3').value = 148;

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
    expect(at(ws.charts, 0).chart.chartType).toBe('stock');
    expect(at(ws.charts, 0).chart.series).toHaveLength(3);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Stock') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    const chart = at(ws2.charts, 0);
    expect(chart.chart.chartType).toBe('stock');
    expect(chart.chart.series).toHaveLength(3);
    // The model has no hiLowLines/upDownBars; only the series wiring is asserted.
    expect(chart.chart.series?.map((s) => [s.name, s.catRef, s.valRef])).toEqual([
      ['High', 'Stock!$A$2:$A$3', 'Stock!$B$2:$B$3'],
      ['Low', 'Stock!$A$2:$A$3', 'Stock!$C$2:$C$3'],
      ['Close', 'Stock!$A$2:$A$3', 'Stock!$D$2:$D$3'],
    ]);
  });

  it('roundtrips a oneCellAnchor chart through WASM write/read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = 100;

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
    expect(at(ws.charts, 0).anchor.extCx).toBe(5400000);
    expect(at(ws.charts, 0).anchor.extCy).toBe(3240000);

    // Write to buffer.
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Read back.
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;
    expect(ws2).toBeDefined();

    const charts = ws2.charts;
    expect(charts).toHaveLength(1);
    expect(at(charts, 0).chart.chartType).toBe('bar');
    expect(at(charts, 0).chart.title?.text).toBe('OneCellTest');

    // Verify the oneCellAnchor fields survived the roundtrip.
    expect(at(charts, 0).anchor.fromCol).toBe(3);
    expect(at(charts, 0).anchor.fromRow).toBe(5);
    expect(at(charts, 0).anchor.fromColOff).toBe(100);
    expect(at(charts, 0).anchor.fromRowOff).toBe(200);
    expect(at(charts, 0).anchor.extCx).toBe(5400000);
    expect(at(charts, 0).anchor.extCy).toBe(3240000);
  });

  it('roundtrips a combo chart (bar + line) with secondary axis', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Sales';
    ws.cell('C1').value = 'Trend';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = 100;
    ws.cell('C2').value = 10;
    ws.cell('A3').value = 'Q2';
    ws.cell('B3').value = 200;
    ws.cell('C3').value = 15;

    ws.addChartData({
      chart: {
        chartType: 'column',
        title: { text: 'Combo Chart', overlay: false },
        series: [
          {
            idx: 0,
            order: 0,
            name: 'Sales',
            catRef: 'Sheet1!$A$2:$A$3',
            valRef: 'Sheet1!$B$2:$B$3',
          },
        ],
        grouping: 'clustered',
        catAxis: { id: 100, crossAx: 200 },
        valAxis: { id: 200, crossAx: 100 },
        secondaryChart: {
          chartType: 'line',
          series: [
            {
              idx: 1,
              order: 1,
              name: 'Trend',
              catRef: 'Sheet1!$A$2:$A$3',
              valRef: 'Sheet1!$C$2:$C$3',
            },
          ],
          grouping: 'standard',
          showDataTable: false,
        },
        secondaryValAxis: {
          id: 300,
          crossAx: 100,
          position: 'right',
        },
        showDataTable: false,
      },
      anchor: {
        fromCol: 4,
        fromRow: 0,
        toCol: 14,
        toRow: 18,
      },
    });

    expect(ws.charts).toHaveLength(1);
    expect(at(ws.charts, 0).chart.secondaryChart?.chartType).toBe('line');

    // Write to buffer.
    const buffer = await wb.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Read back.
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;
    expect(ws2).toBeDefined();

    const charts = ws2.charts;
    expect(charts).toHaveLength(1);
    const chart = at(charts, 0).chart;

    // Primary chart.
    expect(chart.chartType).toBe('column');
    expect(chart.grouping).toBe('clustered');
    expect(chart.series).toHaveLength(1);
    expect(at(chart.series, 0).name).toBe('Sales');
    expect(at(chart.series, 0).valRef).toBe('Sheet1!$B$2:$B$3');

    // Secondary chart.
    expect(chart.secondaryChart).toBeDefined();
    expect(chart.secondaryChart?.chartType).toBe('line');
    expect(chart.secondaryChart?.grouping).toBe('standard');
    expect(chart.secondaryChart?.series).toHaveLength(1);
    expect(at(chart.secondaryChart?.series, 0).name).toBe('Trend');
    expect(at(chart.secondaryChart?.series, 0).valRef).toBe('Sheet1!$C$2:$C$3');

    // Primary value axis.
    expect(chart.valAxis?.id).toBe(200);
    expect(chart.valAxis?.crossAx).toBe(100);

    // Secondary value axis.
    expect(chart.secondaryValAxis).toBeDefined();
    expect(chart.secondaryValAxis?.id).toBe(300);
    expect(chart.secondaryValAxis?.crossAx).toBe(100);
    expect(chart.secondaryValAxis?.position).toBe('right');
  });

  /** Writes a one-sheet workbook with `build` applied to a chart, reads it back, returns the chart model. */
  async function roundtripChart(
    type: Parameters<Worksheet['addChart']>[0],
    build: Parameters<Worksheet['addChart']>[1],
  ) {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Category';
    ws.cell('B1').value = 'Value';
    ws.cell('A2').value = 'Q1';
    ws.cell('B2').value = 100;
    ws.cell('A3').value = 'Q2';
    ws.cell('B3').value = 200;
    ws.addChart(type, build);
    const wb2 = await readBuffer(await wb.toBuffer());
    const charts = wb2.getSheet('Sheet1')?.charts;
    expect(charts).toHaveLength(1);
    return at(charts, 0).chart;
  }

  const series = { catRef: 'Sheet1!$A$2:$A$3', valRef: 'Sheet1!$B$2:$B$3' };

  it('roundtrips an area chart', async () => {
    const chart = await roundtripChart('area', (b) => {
      b.addSeries(series).grouping('stacked');
    });
    expect(chart.chartType).toBe('area');
    expect(chart.grouping).toBe('stacked');
    expect(at(chart.series, 0).valRef).toBe(series.valRef);
  });

  it('roundtrips a doughnut chart with holeSize', async () => {
    const chart = await roundtripChart('doughnut', (b) => {
      b.addSeries(series).holeSize(50);
    });
    expect(chart.chartType).toBe('doughnut');
    expect(chart.holeSize).toBe(50);
  });

  it('roundtrips a radar chart with radarStyle', async () => {
    const chart = await roundtripChart('radar', (b) => {
      b.addSeries(series).radarStyle('filled');
    });
    expect(chart.chartType).toBe('radar');
    expect(chart.radarStyle).toBe('filled');
  });

  it('roundtrips a series trendline', async () => {
    const chart = await roundtripChart('line', (b) => {
      b.addSeries({
        ...series,
        trendline: { trendType: 'linear', forward: 1, displayEq: true, displayRSqr: true },
      });
    });
    expect(at(chart.series, 0).trendline).toEqual({
      trendType: 'linear',
      forward: 1,
      displayEq: true,
      displayRSqr: true,
    });
  });

  it('roundtrips series error bars', async () => {
    const chart = await roundtripChart('bar', (b) => {
      b.addSeries({ ...series, errorBars: { errType: 'percentage', value: 5 } });
    });
    expect(at(chart.series, 0).errorBars).toEqual({
      errType: 'percentage',
      direction: 'both',
      value: 5,
    });
  });

  it('roundtrips view3d', async () => {
    const chart = await roundtripChart('bar', (b) => {
      b.addSeries(series).view3d({ rotX: 15, rotY: 20, perspective: 30 });
    });
    expect(chart.view3d).toEqual({ rotX: 15, rotY: 20, perspective: 30 });
  });

  it('roundtrips showDataTable', async () => {
    const chart = await roundtripChart('bar', (b) => {
      b.addSeries(series).showDataTable(true);
    });
    expect(chart.showDataTable).toBe(true);
  });

  it('roundtrips showLegendKey and showBubbleSize in dataLabels', async () => {
    const chart = await roundtripChart('bar', (b) => {
      b.addSeries(series).dataLabels({ showVal: true, showLegendKey: true, showBubbleSize: false });
    });
    expect(chart.dataLabels?.showLegendKey).toBe(true);
    expect(chart.dataLabels?.showBubbleSize).toBe(false);
  });
});
