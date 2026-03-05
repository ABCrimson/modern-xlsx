import { describe, expect, it } from 'vitest';
import type { Worksheet, WorksheetChartData } from '../src/index.js';
import { ChartBuilder, readBuffer, Workbook } from '../src/index.js';

describe('addChartData', () => {
  it('adds a pre-built WorksheetChartData and it appears in charts', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const chartData: WorksheetChartData = new ChartBuilder('bar')
      .title('Revenue')
      .addSeries({ name: 'Q1', valRef: 'Sheet1!$B$2:$B$5' })
      .anchor({ col: 0, row: 0 }, { col: 10, row: 15 })
      .build();

    ws.addChartData(chartData);

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].chart.chartType).toBe('bar');
    expect(ws.charts[0].chart.title?.text).toBe('Revenue');
    expect(ws.charts[0].chart.series[0].name).toBe('Q1');
    expect(ws.charts[0].anchor.fromCol).toBe(0);
    expect(ws.charts[0].anchor.toCol).toBe(10);
  });

  it('roundtrips a chart added via addChartData', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'X';
    ws.cell('B1').value = '10';

    const chartData: WorksheetChartData = new ChartBuilder('line')
      .title('Trend')
      .addSeries({ name: 'Sales', valRef: 'Sheet1!$B$1:$B$1' })
      .legend('bottom')
      .anchor({ col: 3, row: 0 }, { col: 12, row: 18 })
      .build();

    ws.addChartData(chartData);

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1') as Worksheet;

    expect(ws2.charts).toHaveLength(1);
    expect(ws2.charts[0].chart.chartType).toBe('line');
    expect(ws2.charts[0].chart.title?.text).toBe('Trend');
    expect(ws2.charts[0].chart.series[0].name).toBe('Sales');
    expect(ws2.charts[0].chart.legend?.position).toBe('bottom');
    expect(ws2.charts[0].anchor.fromCol).toBe(3);
    expect(ws2.charts[0].anchor.toCol).toBe(12);
  });

  it('initializes charts array if absent', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Before adding anything, charts should be empty
    expect(ws.charts).toHaveLength(0);

    const chartData: WorksheetChartData = new ChartBuilder('pie')
      .addSeries({ valRef: 'Sheet1!$A$1:$A$3' })
      .build();

    ws.addChartData(chartData);

    expect(ws.charts).toHaveLength(1);
    expect(ws.charts[0].chart.chartType).toBe('pie');
  });
});

describe('removeChart', () => {
  it('removes chart by index and returns true', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addChart('bar', (b) => {
      b.title('First').addSeries({ valRef: 'Sheet1!$A$1:$A$5' });
    });
    ws.addChart('line', (b) => {
      b.title('Second').addSeries({ valRef: 'Sheet1!$B$1:$B$5' });
    });
    ws.addChart('pie', (b) => {
      b.title('Third').addSeries({ valRef: 'Sheet1!$C$1:$C$5' });
    });

    expect(ws.charts).toHaveLength(3);

    const removed = ws.removeChart(1);

    expect(removed).toBe(true);
    expect(ws.charts).toHaveLength(2);
    // Remaining charts shift correctly: First (index 0), Third (index 1)
    expect(ws.charts[0].chart.title?.text).toBe('First');
    expect(ws.charts[0].chart.chartType).toBe('bar');
    expect(ws.charts[1].chart.title?.text).toBe('Third');
    expect(ws.charts[1].chart.chartType).toBe('pie');
  });

  it('returns false for index 0 on empty sheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    expect(ws.removeChart(0)).toBe(false);
  });

  it('returns false for negative index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    expect(ws.removeChart(-1)).toBe(false);
  });

  it('returns false for out-of-range index', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Sheet1!$A$1:$A$5' });
    });

    expect(ws.removeChart(99)).toBe(false);
    expect(ws.charts).toHaveLength(1);
  });

  it('returns false for negative index when charts exist', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addChart('bar', (b) => {
      b.addSeries({ valRef: 'Sheet1!$A$1:$A$5' });
    });
    ws.addChart('line', (b) => {
      b.addSeries({ valRef: 'Sheet1!$B$1:$B$5' });
    });

    expect(ws.removeChart(-1)).toBe(false);
    // Charts are unchanged
    expect(ws.charts).toHaveLength(2);
  });
});

describe('charts getter', () => {
  it('returns empty array when no charts added', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // The getter returns readonly WorksheetChartData[] which defaults to []
    expect(ws.charts).toHaveLength(0);
    expect(ws.charts).toEqual([]);
  });
});
