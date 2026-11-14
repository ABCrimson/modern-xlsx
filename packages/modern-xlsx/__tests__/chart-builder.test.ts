import { assert, describe, expect, it } from 'vitest';
import { ChartBuilder } from '../src/authoring/chart-builder.js';

/** Element `index` of `items`, narrowed with `assert.isDefined` (noUncheckedIndexedAccess). */
function at<T>(items: readonly T[] | undefined, index: number): T {
  const item = items?.[index];
  assert.isDefined(item);
  return item;
}

describe('ChartBuilder', () => {
  it('creates a minimal bar chart', () => {
    const result = new ChartBuilder('bar').addSeries({ valRef: 'Sheet1!$B$2:$B$5' }).build();

    expect(result.chart.chartType).toBe('bar');
    expect(result.chart.series).toHaveLength(1);
    expect(at(result.chart.series, 0).valRef).toBe('Sheet1!$B$2:$B$5');
    expect(at(result.chart.series, 0).idx).toBe(0);
    expect(at(result.chart.series, 0).order).toBe(0);
  });

  it('sets title with formatting', () => {
    const result = new ChartBuilder('line')
      .title('Revenue', { bold: true, fontSize: 1400, color: '333333' })
      .build();

    expect(result.chart.title).toEqual({
      text: 'Revenue',
      overlay: false,
      bold: true,
      fontSize: 1400,
      color: '333333',
    });
  });

  it('adds multiple series with auto-indexing', () => {
    const result = new ChartBuilder('line')
      .addSeries({ name: 'Q1', valRef: 'Sheet1!$B$2:$B$5' })
      .addSeries({ name: 'Q2', valRef: 'Sheet1!$C$2:$C$5', fillColor: 'FF0000' })
      .addSeries({ name: 'Q3', valRef: 'Sheet1!$D$2:$D$5' })
      .build();

    expect(result.chart.series).toHaveLength(3);
    expect(at(result.chart.series, 0).idx).toBe(0);
    expect(at(result.chart.series, 1).idx).toBe(1);
    expect(at(result.chart.series, 2).idx).toBe(2);
    expect(at(result.chart.series, 1).fillColor).toBe('FF0000');
  });

  it('configures category axis', () => {
    const result = new ChartBuilder('bar')
      .catAxis({ title: 'Category', position: 'bottom', majorTickMark: 'out' })
      .build();

    expect(result.chart.catAxis).toBeDefined();
    expect(result.chart.catAxis?.id).toBe(0);
    expect(result.chart.catAxis?.crossAx).toBe(1);
    expect(result.chart.catAxis?.title?.text).toBe('Category');
    expect(result.chart.catAxis?.position).toBe('bottom');
    expect(result.chart.catAxis?.majorTickMark).toBe('out');
  });

  it('configures value axis with scale', () => {
    const result = new ChartBuilder('line')
      .valAxis({
        title: { text: 'Revenue ($)', bold: true, fontSize: 1200 },
        min: 0,
        max: 1000,
        majorUnit: 200,
        numFmt: '#,##0',
        majorGridlines: true,
      })
      .build();

    expect(result.chart.valAxis).toBeDefined();
    expect(result.chart.valAxis?.id).toBe(1);
    expect(result.chart.valAxis?.min).toBe(0);
    expect(result.chart.valAxis?.max).toBe(1000);
    expect(result.chart.valAxis?.majorUnit).toBe(200);
    expect(result.chart.valAxis?.numFmt).toBe('#,##0');
    expect(result.chart.valAxis?.majorGridlines).toBe(true);
    expect(result.chart.valAxis?.title?.bold).toBe(true);
  });

  it('sets legend', () => {
    const result = new ChartBuilder('pie').legend('right').build();

    expect(result.chart.legend).toEqual({ position: 'right', overlay: false });
  });

  it('configures data labels', () => {
    const result = new ChartBuilder('pie').dataLabels({ showVal: true, showPercent: true }).build();

    expect(result.chart.dataLabels).toBeDefined();
    expect(result.chart.dataLabels?.showVal).toBe(true);
    expect(result.chart.dataLabels?.showPercent).toBe(true);
    expect(result.chart.dataLabels?.showCatName).toBe(false);
  });

  it('sets grouping', () => {
    const result = new ChartBuilder('bar').grouping('stacked').build();
    expect(result.chart.grouping).toBe('stacked');
  });

  it('configures scatter chart', () => {
    const result = new ChartBuilder('scatter')
      .scatterStyle('lineMarker')
      .addSeries({
        xValRef: 'Sheet1!$A$2:$A$10',
        valRef: 'Sheet1!$B$2:$B$10',
        marker: 'diamond',
      })
      .build();

    expect(result.chart.chartType).toBe('scatter');
    expect(result.chart.scatterStyle).toBe('lineMarker');
    expect(at(result.chart.series, 0).xValRef).toBe('Sheet1!$A$2:$A$10');
    expect(at(result.chart.series, 0).marker).toBe('diamond');
  });

  it('configures doughnut chart', () => {
    const result = new ChartBuilder('doughnut').holeSize(50).build();

    expect(result.chart.chartType).toBe('doughnut');
    expect(result.chart.holeSize).toBe(50);
  });

  it('sets bar direction', () => {
    const result = new ChartBuilder('bar').barDirection(true).build();
    expect(result.chart.barDirHorizontal).toBe(true);
  });

  it('sets style ID', () => {
    const result = new ChartBuilder('line').style(26).build();
    expect(result.chart.styleId).toBe(26);
  });

  it('sets plot layout', () => {
    const result = new ChartBuilder('bar').plotLayout(0.1, 0.15, 0.8, 0.7).build();

    expect(result.chart.plotAreaLayout).toEqual({ x: 0.1, y: 0.15, w: 0.8, h: 0.7 });
  });

  it('sets custom anchor', () => {
    const result = new ChartBuilder('bar')
      .anchor({ col: 5, row: 10, colOff: 100 }, { col: 15, row: 30 })
      .build();

    expect(result.anchor).toEqual({
      fromCol: 5,
      fromRow: 10,
      fromColOff: 100,
      fromRowOff: 0,
      toCol: 15,
      toRow: 30,
      toColOff: 0,
      toRowOff: 0,
    });
  });

  it('uses default anchor when not specified', () => {
    const result = new ChartBuilder('bar').build();
    expect(result.anchor.fromCol).toBe(0);
    expect(result.anchor.toCol).toBe(10);
    expect(result.anchor.toRow).toBe(15);
  });

  it('chains all methods fluently', () => {
    const result = new ChartBuilder('bar')
      .title('Full Chart')
      .addSeries({
        name: 'S1',
        catRef: 'Sheet1!$A$2:$A$5',
        valRef: 'Sheet1!$B$2:$B$5',
      })
      .catAxis({ title: 'X', majorGridlines: true })
      .valAxis({ title: 'Y', majorGridlines: true })
      .legend('bottom')
      .dataLabels({ showVal: true })
      .grouping('clustered')
      .style(2)
      .anchor({ col: 0, row: 0 }, { col: 10, row: 20 })
      .build();

    expect(result.chart.chartType).toBe('bar');
    expect(result.chart.title?.text).toBe('Full Chart');
    expect(result.chart.series).toHaveLength(1);
    expect(result.chart.catAxis?.title?.text).toBe('X');
    expect(result.chart.valAxis?.title?.text).toBe('Y');
    expect(result.chart.legend?.position).toBe('bottom');
    expect(result.chart.dataLabels?.showVal).toBe(true);
    expect(result.chart.grouping).toBe('clustered');
    expect(result.chart.styleId).toBe(2);
  });

  it('supports series-level data labels', () => {
    const result = new ChartBuilder('pie')
      .addSeries({
        valRef: 'Sheet1!$B$2:$B$5',
        dataLabels: { showPercent: true, showLeaderLines: true },
      })
      .build();

    expect(at(result.chart.series, 0).dataLabels?.showPercent).toBe(true);
    expect(at(result.chart.series, 0).dataLabels?.showLeaderLines).toBe(true);
  });

  it('creates pie chart without axes', () => {
    const result = new ChartBuilder('pie')
      .addSeries({ valRef: 'Sheet1!$B$2:$B$5', explosion: 25 })
      .legend('right')
      .build();

    expect(result.chart.chartType).toBe('pie');
    expect(result.chart.catAxis).toBeNull();
    expect(result.chart.valAxis).toBeNull();
    expect(at(result.chart.series, 0).explosion).toBe(25);
  });

  it('sets radar style', () => {
    const result = new ChartBuilder('radar').radarStyle('filled').build();
    expect(result.chart.radarStyle).toBe('filled');
  });

  it('creates series with trendline', () => {
    const result = new ChartBuilder('scatter')
      .addSeries({
        valRef: 'Sheet1!$B$2:$B$5',
        trendline: { trendType: 'linear', displayEq: true, displayRSqr: true },
      })
      .build();

    expect(at(result.chart.series, 0).trendline?.trendType).toBe('linear');
    expect(at(result.chart.series, 0).trendline?.displayEq).toBe(true);
  });

  it('creates series with error bars', () => {
    const result = new ChartBuilder('bar')
      .addSeries({
        valRef: 'Sheet1!$B$2:$B$5',
        errorBars: { errType: 'percentage', value: 5 },
      })
      .build();

    expect(at(result.chart.series, 0).errorBars?.errType).toBe('percentage');
    expect(at(result.chart.series, 0).errorBars?.value).toBe(5);
  });

  it('sets 3D rotation', () => {
    const result = new ChartBuilder('bar').view3d({ rotX: 15, rotY: 20, perspective: 30 }).build();

    expect(result.chart.view3d).toEqual({ rotX: 15, rotY: 20, perspective: 30, rAngAx: null });
  });

  it('shows data table', () => {
    const result = new ChartBuilder('bar').showDataTable().build();
    expect(result.chart.showDataTable).toBe(true);
  });

  it('keeps showLegendKey and showBubbleSize in dataLabels', () => {
    const r = new ChartBuilder('bubble')
      .dataLabels({ showVal: true, showBubbleSize: true, showLegendKey: true })
      .build();
    expect(r.chart.dataLabels?.showBubbleSize).toBe(true);
    expect(r.chart.dataLabels?.showLegendKey).toBe(true);
    const s = new ChartBuilder('bar')
      .addSeries({ valRef: 'Sheet1!$B$2:$B$5', dataLabels: { showLegendKey: true } })
      .build();
    expect(at(s.chart.series, 0).dataLabels?.showLegendKey).toBe(true);
  });
});

describe('ChartBuilder combo charts', () => {
  /** `[idx, order]` of every series of `series`. */
  function ids(series: readonly { idx: number; order: number }[] | undefined): number[][] {
    return (series ?? []).map((s) => [s.idx, s.order]);
  }

  // Excel refuses a chart whose series share a c:order anywhere in its c:plotArea (excel-check.ps1,
  // 2026-11-12), so the secondary chart's series continue the primary chart's numbering.
  it('numbers the series 0..n-1 across the primary then the secondary chart', () => {
    const line = new ChartBuilder('line')
      .addSeries({ valRef: 'Sheet1!$C$2:$C$5' })
      .addSeries({ valRef: 'Sheet1!$D$2:$D$5' });
    const result = new ChartBuilder('column')
      .addSeries({ valRef: 'Sheet1!$A$2:$A$5' })
      .addSeries({ valRef: 'Sheet1!$B$2:$B$5' })
      .secondaryChart(line)
      .build();

    expect(ids(result.chart.series)).toEqual([
      [0, 0],
      [1, 1],
    ]);
    expect(result.chart.secondaryChart?.chartType).toBe('line');
    expect(ids(result.chart.secondaryChart?.series)).toEqual([
      [2, 2],
      [3, 3],
    ]);
    // The secondary builder on its own still numbers from 0.
    expect(ids(line.build().chart.series)).toEqual([
      [0, 0],
      [1, 1],
    ]);
  });

  it('keeps an idx and order the caller set and numbers the others around them', () => {
    const result = new ChartBuilder('column')
      .addSeries({ valRef: 'Sheet1!$A$2:$A$5', idx: 1, order: 5 })
      .addSeries({ valRef: 'Sheet1!$B$2:$B$5' })
      .secondaryChart(new ChartBuilder('line').addSeries({ valRef: 'Sheet1!$C$2:$C$5' }))
      .build();

    expect(ids(result.chart.series)).toEqual([
      [1, 5],
      [0, 0],
    ]);
    expect(ids(result.chart.secondaryChart?.series)).toEqual([[2, 1]]);
  });

  it('writes no axes of its own for the secondary chart and leaves a plain chart without one', () => {
    const result = new ChartBuilder('column')
      .addSeries({ valRef: 'Sheet1!$A$2:$A$5' })
      .secondaryChart(new ChartBuilder('line').addSeries({ valRef: 'Sheet1!$C$2:$C$5' }))
      .build();
    expect(result.chart.secondaryChart?.catAxis).toBeNull();
    expect(result.chart.secondaryChart?.valAxis).toBeNull();
    expect(result.chart.catAxis?.id).toBe(0);

    const plain = new ChartBuilder('column').addSeries({ valRef: 'Sheet1!$A$2:$A$5' }).build();
    expect('secondaryChart' in plain.chart).toBe(false);
  });

  it('refuses a secondary chart that has a secondary chart of its own', () => {
    const inner = new ChartBuilder('line')
      .addSeries({ valRef: 'Sheet1!$C$2:$C$5' })
      .secondaryChart(new ChartBuilder('area').addSeries({ valRef: 'Sheet1!$D$2:$D$5' }));
    const outer = new ChartBuilder('column')
      .addSeries({ valRef: 'Sheet1!$A$2:$A$5' })
      .secondaryChart(inner);
    expect(() => outer.build()).toThrow(RangeError);
  });
});
