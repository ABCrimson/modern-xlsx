import { describe, expect, it } from 'vitest';
import type { WorksheetChartData } from '../src/types.js';
import { validateChartData } from '../src/validate-chart.js';

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

  it('throws plain Error for empty valRef', () => {
    const data = validChart({
      chart: { chartType: 'bar', series: [{ idx: 0, order: 0, valRef: '' }] },
    });
    expect(() => validateChartData(data)).toThrow(Error);
  });
});
