import { assert, describe, expect, it } from 'vitest';
import { validateChartData } from '../src/authoring/validate-chart.js';
import type { Worksheet } from '../src/index.js';
import { ChartBuilder, INVALID_WORKBOOK, readBuffer, Workbook } from '../src/index.js';
import type { ChartSeriesData, WorksheetChartData } from '../src/model/types.js';

/** Helper: returns a minimal valid chart data object. */
function validChart(overrides?: Partial<WorksheetChartData>): WorksheetChartData {
  return {
    chart: {
      chartType: 'bar',
      series: [{ idx: 0, order: 0, valRef: 'Sheet1!$B$2:$B$5' }],
      ...overrides?.chart,
    },
    anchor: {
      fromCol: 0,
      fromRow: 0,
      toCol: 10,
      toRow: 15,
      ...overrides?.anchor,
    },
  };
}

describe('validateChartData', () => {
  it('accepts valid chart data without throwing', () => {
    expect(() => validateChartData(validChart())).not.toThrow();
  });

  it('accepts valid chart with all optional fields', () => {
    const data = validChart({
      chart: {
        chartType: 'doughnut',
        holeSize: 50,
        styleId: 2,
        view3d: { rotX: -45, rotY: 180, perspective: 30 },
        series: [
          {
            idx: 0,
            order: 0,
            valRef: 'Sheet1!$B$2:$B$5',
            lineWidth: 25400,
            explosion: 10,
          },
        ],
      },
    });
    expect(() => validateChartData(data)).not.toThrow();
  });

  // --- holeSize ---

  it('rejects negative holeSize', () => {
    const data = validChart({
      chart: {
        chartType: 'doughnut',
        holeSize: -1,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/holeSize/);
  });

  it('rejects holeSize > 90', () => {
    const data = validChart({
      chart: {
        chartType: 'doughnut',
        holeSize: 91,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/holeSize/);
  });

  it('rejects float holeSize', () => {
    const data = validChart({
      chart: {
        chartType: 'doughnut',
        holeSize: 50.5,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/holeSize/);
  });

  // --- lineWidth ---

  it('rejects float lineWidth', () => {
    const data = validChart({
      chart: {
        chartType: 'line',
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1', lineWidth: 1.5 }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/lineWidth/);
  });

  it('rejects negative lineWidth', () => {
    const data = validChart({
      chart: {
        chartType: 'line',
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1', lineWidth: -1 }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/lineWidth/);
  });

  // --- explosion ---

  it('rejects negative explosion', () => {
    const data = validChart({
      chart: {
        chartType: 'pie',
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1', explosion: -5 }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/explosion/);
  });

  // --- view3d ---

  it('rejects rotX > 90', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        view3d: { rotX: 91 },
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/rotX/);
  });

  it('rejects rotX < -90', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        view3d: { rotX: -91 },
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/rotX/);
  });

  it('rejects rotY > 360', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        view3d: { rotY: 361 },
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/rotY/);
  });

  it('rejects negative rotY', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        view3d: { rotY: -1 },
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/rotY/);
  });

  it('rejects perspective > 240', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        view3d: { perspective: 241 },
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/perspective/);
  });

  // --- series.valRef ---

  it('rejects empty valRef', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        series: [{ idx: 0, order: 0, valRef: '' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/valRef/);
  });

  // --- anchor ---

  it('rejects negative anchor.fromCol', () => {
    const data = validChart({ anchor: { fromCol: -1, fromRow: 0, toCol: 10, toRow: 15 } });
    expect(() => validateChartData(data)).toThrow(/fromCol/);
  });

  it('rejects float anchor.toRow', () => {
    const data = validChart({ anchor: { fromCol: 0, fromRow: 0, toCol: 10, toRow: 1.5 } });
    expect(() => validateChartData(data)).toThrow(/toRow/);
  });

  // --- styleId ---

  it('rejects negative styleId', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        styleId: -1,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/styleId/);
  });

  it('rejects float styleId', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        styleId: 2.5,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(/styleId/);
  });

  // --- error types ---

  it('throws RangeError for numeric violations', () => {
    const data = validChart({
      chart: {
        chartType: 'doughnut',
        holeSize: -1,
        series: [{ idx: 0, order: 0, valRef: 'Sheet1!$A$1' }],
      },
    });
    expect(() => validateChartData(data)).toThrow(RangeError);
  });

  it('throws RangeError for empty valRef', () => {
    const data = validChart({
      chart: { chartType: 'bar', series: [{ idx: 0, order: 0, valRef: '' }] },
    });
    expect(() => validateChartData(data)).toThrow(RangeError);
    expect(() => validateChartData(data)).toThrow(/series\.valRef/);
  });
});

describe('validateChartData boundaries', () => {
  const ser = { idx: 0, order: 0, valRef: 'Sheet1!$A$1' };
  const withChart = (extra: Record<string, unknown>) =>
    validChart({ chart: { chartType: 'doughnut', series: [ser], ...extra } });

  it('accepts the inclusive boundary values', () => {
    const accepted = [
      withChart({ holeSize: 1 }),
      withChart({ holeSize: 90 }),
      withChart({ view3d: { rotX: -90 } }),
      withChart({ view3d: { rotX: 90 } }),
      withChart({ view3d: { rotY: 0 } }),
      withChart({ view3d: { rotY: 360 } }),
      withChart({ view3d: { perspective: 0 } }),
      withChart({ view3d: { perspective: 240 } }),
      withChart({ styleId: 1 }),
      withChart({ styleId: 48 }),
      withChart({ series: [{ ...ser, lineWidth: 0 }] }),
    ];
    for (const d of accepted) expect(() => validateChartData(d)).not.toThrow();
  });

  it('rejects holeSize 0 with RangeError', () => {
    expect(() => validateChartData(withChart({ holeSize: 0 }))).toThrow(RangeError);
  });

  it('rejects NaN and Infinity on rotX and anchor.fromCol with RangeError', () => {
    for (const v of [Number.NaN, Number.POSITIVE_INFINITY]) {
      const rotX = () => validateChartData(withChart({ view3d: { rotX: v } }));
      expect(rotX).toThrow(RangeError);
      expect(rotX).toThrow(/view3d\.rotX/);
      const anchor = { fromCol: v, fromRow: 0, toCol: 10, toRow: 15 };
      const fromCol = () => validateChartData(validChart({ anchor }));
      expect(fromCol).toThrow(RangeError);
      expect(fromCol).toThrow(/anchor\.fromCol/);
    }
  });
});

describe('validateChartData spec ranges', () => {
  const ser = { idx: 0, order: 0, valRef: 'Sheet1!$A$1' };
  const withSeries = (extra: Record<string, unknown>) =>
    validChart({ chart: { chartType: 'bar', series: [{ ...ser, ...extra }] } });

  it('rejects trendline order 1, 1.5 and 7 and accepts 2 and 6', () => {
    for (const order of [1, 1.5, 7]) {
      const d = withSeries({ trendline: { trendType: 'polynomial', order } });
      expect(() => validateChartData(d)).toThrow(RangeError);
    }
    for (const order of [2, 6]) {
      expect(() =>
        validateChartData(withSeries({ trendline: { trendType: 'polynomial', order } })),
      ).not.toThrow();
    }
  });

  it('rejects trendline period 1', () => {
    const d = withSeries({ trendline: { trendType: 'movingAvg', period: 1 } });
    expect(() => validateChartData(d)).toThrow(RangeError);
    const ok = withSeries({ trendline: { trendType: 'movingAvg', period: 255 } });
    expect(() => validateChartData(ok)).not.toThrow();
  });

  it('checks title.fontSize against 100..400000', () => {
    const mk = (fontSize: number) =>
      validChart({
        chart: { chartType: 'bar', title: { text: 't', fontSize }, series: [ser] },
      });
    expect(() => validateChartData(mk(-1))).toThrow(RangeError);
    expect(() => validateChartData(mk(14))).toThrow(RangeError);
    expect(() => validateChartData(mk(1400))).not.toThrow();
  });

  it('rejects a non-u32 axis id and bad axis font sizes', () => {
    const mk = (axis: Record<string, unknown>) =>
      validChart({
        chart: {
          chartType: 'bar',
          series: [ser],
          catAxis: { id: 0, crossAx: 1, ...axis },
        },
      });
    expect(() => validateChartData(mk({ id: -1 }))).toThrow(RangeError);
    expect(() => validateChartData(mk({ id: 1.5 }))).toThrow(RangeError);
    expect(() => validateChartData(mk({ crossAx: 4294967296 }))).toThrow(RangeError);
    expect(() => validateChartData(mk({ fontSize: 14 }))).toThrow(RangeError);
    expect(() => validateChartData(mk({ title: { text: 'x', fontSize: 14 } }))).toThrow(RangeError);
    expect(() => validateChartData(mk({ fontSize: 1000 }))).not.toThrow();
  });

  it('checks styleId against 1..48', () => {
    const mk = (styleId: number) =>
      validChart({ chart: { chartType: 'bar', styleId, series: [ser] } });
    expect(() => validateChartData(mk(0))).toThrow(RangeError);
    expect(() => validateChartData(mk(49))).toThrow(RangeError);
    expect(() => validateChartData(mk(48))).not.toThrow();
  });

  it('accepts explosion 150', () => {
    expect(() => validateChartData(withSeries({ explosion: 150 }))).not.toThrow();
  });

  it('rejects empty or blank catRef, xValRef and bubbleSizeRef', () => {
    for (const key of ['catRef', 'xValRef', 'bubbleSizeRef']) {
      expect(() => validateChartData(withSeries({ [key]: '  ' }))).toThrow(RangeError);
    }
    expect(() => validateChartData(withSeries({ catRef: null }))).not.toThrow();
  });

  it('rejects a reversed two-cell anchor', () => {
    const cols = validChart({ anchor: { fromCol: 5, fromRow: 0, toCol: 2, toRow: 10 } });
    expect(() => validateChartData(cols)).toThrow(RangeError);
    const rows = validChart({ anchor: { fromCol: 0, fromRow: 9, toCol: 2, toRow: 3 } });
    expect(() => validateChartData(rows)).toThrow(RangeError);
    const offs = validChart({
      anchor: { fromCol: 2, fromRow: 0, fromColOff: 500, toCol: 2, toColOff: 100, toRow: 3 },
    });
    expect(() => validateChartData(offs)).toThrow(RangeError);
  });

  it('accepts a oneCellAnchor with toCol 0 and extCx set', () => {
    const d = validChart({
      anchor: { fromCol: 3, fromRow: 3, toCol: 0, toRow: 0, extCx: 5000000, extCy: 3000000 },
    });
    expect(() => validateChartData(d)).not.toThrow();
  });

  it('checks colours as six hex digits', () => {
    const withSeries = (extra: object) =>
      validChart({
        chart: { chartType: 'bar', series: [{ idx: 0, order: 0, valRef: 'S!$B$2', ...extra }] },
      });
    expect(() => validateChartData(withSeries({ fillColor: '#4472C4' }))).toThrow(RangeError);
    expect(() => validateChartData(withSeries({ lineColor: 'FF4472C4' }))).toThrow(RangeError);
    expect(() => validateChartData(withSeries({ fillColor: '4472c4' }))).not.toThrow();
    const titled = validChart({
      chart: { chartType: 'bar', series: [], title: { text: 'T', color: 'red' } },
    });
    expect(() => validateChartData(titled)).toThrow(RangeError);
  });

  it('checks logBase and axis units', () => {
    const withAxis = (valAxis: object) =>
      validChart({ chart: { chartType: 'bar', series: [], valAxis: valAxis as never } });
    expect(() => validateChartData(withAxis({ logBase: 1 }))).toThrow(RangeError);
    expect(() => validateChartData(withAxis({ logBase: 10 }))).not.toThrow();
    expect(() => validateChartData(withAxis({ majorUnit: 0 }))).toThrow(RangeError);
    expect(() => validateChartData(withAxis({ minorUnit: -1 }))).toThrow(RangeError);
  });
});

/** A column + line combo on explicit axes: one series each, the secondary one with `idx` and `order`. */
function combo(order: number, idx = 1): WorksheetChartData {
  return {
    chart: {
      chartType: 'column',
      series: [{ idx: 0, order: 0, catRef: 'Data!$A$2:$A$5', valRef: 'Data!$B$2:$B$5' }],
      catAxis: { id: 0, crossAx: 1, position: 'bottom' },
      valAxis: { id: 1, crossAx: 0, position: 'left' },
      secondaryChart: {
        chartType: 'line',
        series: [{ idx, order, catRef: 'Data!$A$2:$A$5', valRef: 'Data!$C$2:$C$5' }],
      },
    },
    anchor: { fromCol: 4, fromRow: 1, toCol: 12, toRow: 16 },
  };
}

/** A workbook whose sheet `Data` holds four rows in A:C. */
function dataWorkbook(): { wb: Workbook; ws: Worksheet } {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  for (let r = 2; r <= 5; r++) {
    ws.cell(`A${r}`).value = `Q${r - 1}`;
    ws.cell(`B${r}`).value = r * 10;
    ws.cell(`C${r}`).value = r * 3;
  }
  return { wb, ws };
}

// Excel (excel-check.ps1, 2026-11-12) refuses a chart whose series share a c:order, within one
// chart-type element or across the primary and secondary one, and opens one whose series share
// only a c:idx.
describe('validateChartData series order across a combo chart', () => {
  it('reports a c:order shared by the primary and secondary chart, naming both series', () => {
    expect(() => validateChartData(combo(0))).toThrow(RangeError);
    expect(() => validateChartData(combo(0))).toThrow(
      /series\[0\] and secondaryChart\.series\[0\] both have order 0/,
    );
  });

  it('reports a c:order shared by two series of one chart', () => {
    const data = validChart({
      chart: {
        chartType: 'bar',
        series: [
          { idx: 0, order: 3, valRef: 'Sheet1!$B$2:$B$5' },
          { idx: 1, order: 3, valRef: 'Sheet1!$C$2:$C$5' },
        ],
      },
    });
    expect(() => validateChartData(data)).toThrow(/series\[0\] and series\[1\] both have order 3/);
  });

  it('accepts unique orders, also when two series share an idx', () => {
    expect(() => validateChartData(combo(1))).not.toThrow();
    expect(() => validateChartData(combo(1, 0))).not.toThrow();
  });

  it('rejects a secondary series order or idx of -1 or 1.5 with RangeError', () => {
    for (const bad of [-1, 1.5]) {
      expect(() => validateChartData(combo(bad))).toThrow(RangeError);
      expect(() => validateChartData(combo(bad))).toThrow(/secondaryChart\.series\.order/);
      expect(() => validateChartData(combo(1, bad))).toThrow(/secondaryChart\.series\.idx/);
      const { ws } = dataWorkbook();
      expect(() =>
        ws.addChart('column', (b) =>
          b
            .addSeries({ valRef: 'Data!$B$2:$B$5' })
            .secondaryChart(
              new ChartBuilder('line').addSeries({ valRef: 'Data!$C$2:$C$5', order: bad }),
            ),
        ),
      ).toThrow(RangeError);
      expect(ws.charts).toHaveLength(0);
    }
  });
});

describe('writing a combo chart through the public API', () => {
  it('fails the write with INVALID_WORKBOOK when two series share a c:order', async () => {
    const { wb, ws } = dataWorkbook();
    ws.addChartData(combo(1));
    // Past validateChartData, as a model edited in place or built as JSON reaches the writer.
    const secondary = ws.charts[0]?.chart.secondaryChart?.series[0];
    assert.isDefined(secondary);
    secondary.order = 0;

    await expect(wb.toBuffer()).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: INVALID_WORKBOOK,
      message: expect.stringMatching(
        /sheet 'Data'.*series\[0\] and secondaryChart\.series\[0\] both have c:order 0/,
      ),
    });
    const issues = wb.validate().issues.filter((i) => i.message.includes('c:order'));
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({ severity: 'error', category: 'structure' });
  });

  it('writes a builder-made combo that reads back with the same series', async () => {
    const { wb, ws } = dataWorkbook();
    ws.addChart('column', (b) =>
      b.addSeries({ catRef: 'Data!$A$2:$A$5', valRef: 'Data!$B$2:$B$5' }).secondaryChart(
        new ChartBuilder('line').addSeries({
          catRef: 'Data!$A$2:$A$5',
          valRef: 'Data!$C$2:$C$5',
        }),
      ),
    );
    const built = ws.charts[0]?.chart;
    assert.isDefined(built);

    const read = (await readBuffer(await wb.toBuffer())).getSheet('Data')?.charts[0]?.chart;
    assert.isDefined(read);
    // The builder spells unset fields as null; the reader leaves them out.
    const set = (series: readonly ChartSeriesData[] | undefined) =>
      series?.map((s) => Object.fromEntries(Object.entries(s).filter(([, v]) => v != null)));
    expect(read.series).toEqual(set(built.series));
    expect(read.secondaryChart?.chartType).toBe('line');
    expect(read.secondaryChart?.series).toEqual(set(built.secondaryChart?.series));
    expect(read.secondaryChart?.series.map((s) => [s.idx, s.order])).toEqual([[1, 1]]);
  });
});
