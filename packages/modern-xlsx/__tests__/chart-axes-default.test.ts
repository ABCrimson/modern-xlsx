import { assert, describe, expect, it } from 'vitest';
import type { ChartDataModel, ChartType } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

/**
 * A one-series chart model of `chartType` with no axis set, as hand-built JSON would carry it.
 * `idx` is the series' `c:idx` and `c:order`, unique across the chart-type elements of a combo.
 */
function chartWithoutAxes(chartType: ChartType, idx = 0): ChartDataModel {
  return {
    chartType,
    series: [{ idx, order: idx, catRef: 'Data!$A$2:$A$5', valRef: 'Data!$B$2:$B$5' }],
  };
}

/** A workbook whose sheet `Data` holds four rows of data and `chart`, set through the data model. */
function workbookWith(chart: ChartDataModel): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  const rows: readonly (readonly [string, number])[] = [
    ['Q1', 10],
    ['Q2', 20],
    ['Q3', 15],
    ['Q4', 30],
  ];
  for (const [i, [label, value]] of rows.entries()) {
    ws.cell(`A${i + 2}`).value = label;
    ws.cell(`B${i + 2}`).value = value;
  }
  ws.addChartData({ chart, anchor: { fromCol: 3, fromRow: 1, toCol: 10, toRow: 15 } });
  return wb;
}

/** The `c:axId` values inside every `<tag>` element of `xml`, one list per element. */
function axIdsIn(xml: string, tag: string): number[][] {
  return [...xml.matchAll(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`, 'g'))].map((m) =>
    [...(m[1] ?? '').matchAll(/<c:axId val="(\d+)"\/>/g)].map((id) => Number(id[1])),
  );
}

/** The ids of every `<tag>` axis element (`c:catAx`, `c:valAx`) of `xml`. */
function axisIds(xml: string, tag: string): number[] {
  return [...xml.matchAll(new RegExp(`<${tag}><c:axId val="(\\d+)"/>`, 'g'))].map((m) =>
    Number(m[1]),
  );
}

describe('an axis-bearing chart model without axes', () => {
  it('writes a line chart with two c:axId naming its two default axes, and reads them back', async () => {
    const chart = chartWithoutAxes('line');
    expect(chart.catAxis).toBeUndefined();
    expect(chart.valAxis).toBeUndefined();

    const buffer = await workbookWith(chart).toBuffer();
    const xml = readZipText(buffer)['xl/charts/chart1.xml'] ?? '';
    expect(axIdsIn(xml, 'c:lineChart')).toEqual([[0, 1]]);
    expect(axisIds(xml, 'c:catAx')).toEqual([0]);
    expect(axisIds(xml, 'c:valAx')).toEqual([1]);

    const read = (await readBuffer(buffer)).getSheet('Data')?.charts[0]?.chart;
    assert.isDefined(read);
    expect(read.chartType).toBe('line');
    expect(read.catAxis).toMatchObject({ id: 0, crossAx: 1, position: 'bottom' });
    expect(read.valAxis).toMatchObject({ id: 1, crossAx: 0, position: 'left' });
  });

  it('writes both elements of a combo chart without axes on one default pair', async () => {
    const chart: ChartDataModel = {
      ...chartWithoutAxes('column'),
      secondaryChart: chartWithoutAxes('line', 1),
    };
    const xml = readZipText(await workbookWith(chart).toBuffer())['xl/charts/chart1.xml'] ?? '';
    expect(axIdsIn(xml, 'c:barChart')).toEqual([[0, 1]]);
    expect(axIdsIn(xml, 'c:lineChart')).toEqual([[0, 1]]);
    expect(axisIds(xml, 'c:catAx')).toEqual([0]);
    expect(axisIds(xml, 'c:valAx')).toEqual([1]);
  });

  it('leaves a pie chart without axes', async () => {
    const xml =
      readZipText(await workbookWith(chartWithoutAxes('pie')).toBuffer())['xl/charts/chart1.xml'] ??
      '';
    expect(xml).toContain('<c:pieChart>');
    expect(xml).not.toContain('<c:axId');
  });
});
