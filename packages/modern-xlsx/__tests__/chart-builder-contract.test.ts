import { describe, expect, it } from 'vitest';
import { ChartBuilder } from '../src/authoring/chart-builder.js';
import { getChartStylePalette } from '../src/authoring/chart-styles.js';
import { Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const base = { valRef: 'Sheet1!$B$2:$B$5' };

describe('ChartBuilder contract', () => {
  it('converts lineWidth points to EMU', () => {
    const w = (lineWidth: number) =>
      new ChartBuilder('line').addSeries({ ...base, lineColor: 'FF0000', lineWidth }).build().chart
        .series[0]?.lineWidth;
    expect(w(2)).toBe(25400);
    expect(w(1.5)).toBe(19050);
    expect(w(0)).toBe(0);
    expect(w(1584)).toBe(1584 * 12700);
  });

  it('rejects an invalid lineWidth with RangeError', () => {
    for (const lineWidth of [-1, 1585, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => new ChartBuilder('line').addSeries({ ...base, lineWidth })).toThrow(RangeError);
    }
  });

  it('build() returns an independent deep copy', () => {
    const b = new ChartBuilder('bar').addSeries(base);
    const first = b.build();
    b.addSeries({ valRef: 'Sheet1!$C$2:$C$5' });
    expect(first.chart.series).toHaveLength(1);
    expect(b.build().chart.series).toHaveLength(2);
  });

  it('two build() calls do not share the anchor', () => {
    const b = new ChartBuilder('bar').addSeries(base);
    const a = b.build();
    const c = b.build();
    expect(a.anchor).not.toBe(c.anchor);
    a.anchor.toCol = 99;
    expect(c.anchor.toCol).toBe(10);
  });

  it('keeps builder state in #private fields', () => {
    expect(Object.keys(new ChartBuilder('bar'))).toEqual([]);
  });
});

describe('getChartStylePalette', () => {
  it('returns the bundled palette for style 1', () => {
    expect(getChartStylePalette(1)[0]).toBe('4472C4');
  });

  it('throws RangeError for styles without a bundled palette', () => {
    expect(() => getChartStylePalette(9)).toThrow(RangeError);
    expect(() => getChartStylePalette(9)).toThrow(/only 1-8/);
  });
});

describe('axes for axis-bearing chart types', () => {
  const axIds = async (type: 'scatter' | 'bar', configure: (b: ChartBuilder) => ChartBuilder) => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('B2').value = 4;
    ws.addChart(type, (b) => configure(b.addSeries({ ...base })));
    const xml = readZipText(await wb.toBuffer())['xl/charts/chart1.xml'] ?? '';
    const start = xml.indexOf(`<c:${type}Chart>`);
    const end = xml.indexOf(`</c:${type}Chart>`);
    const plot = start >= 0 && end > start ? xml.slice(start, end) : '';
    return plot.match(/<c:axId /g)?.length ?? 0;
  };

  it('fills the missing category axis of a scatter chart with only valAxis()', async () => {
    const { chart } = new ChartBuilder('scatter').addSeries(base).valAxis({ title: 'Y' }).build();
    expect(chart.catAxis).toMatchObject({ id: 0, crossAx: 1 });
    expect(chart.valAxis).toMatchObject({ id: 1, crossAx: 0 });
    expect(await axIds('scatter', (b) => b.valAxis({ title: 'Y' }))).toBe(2);
  });

  it('gives a bar chart with no axis call two axIds', async () => {
    const { chart } = new ChartBuilder('bar').addSeries(base).build();
    expect(chart.catAxis).toMatchObject({ id: 0, crossAx: 1 });
    expect(chart.valAxis).toMatchObject({ id: 1, crossAx: 0 });
    expect(await axIds('bar', (b) => b)).toBe(2);
  });

  it('gives a bar chart with only catAxis() two axIds', async () => {
    expect(await axIds('bar', (b) => b.catAxis({ title: 'X' }))).toBe(2);
  });

  it('keeps explicitly set axes and leaves pie and doughnut without axes', () => {
    const bar = new ChartBuilder('bar').addSeries(base).catAxis({ title: 'X' }).build();
    expect(bar.chart.catAxis?.title?.text).toBe('X');
    for (const t of ['pie', 'doughnut'] as const) {
      const { chart } = new ChartBuilder(t).addSeries(base).build();
      expect(chart.catAxis).toBeNull();
      expect(chart.valAxis).toBeNull();
    }
  });
});
