import type {
  ChartAnchorData,
  ChartAxisData,
  ChartDataModel,
  ChartSeriesData,
  WorksheetChartData,
} from '../model/types.js';

function assertU32(name: string, value: number | null | undefined): void {
  if (value == null) return;
  if (!Number.isInteger(value) || value < 0 || value > 4294967295) {
    throw new RangeError(`${name} must be a non-negative integer (u32), got ${value}`);
  }
}

const HEX_RGB = /^[0-9A-Fa-f]{6}$/;

/** DrawingML `a:srgbClr` takes ST_HexColorRGB: six hex digits, no `#` and no alpha. */
function assertHexRgb(name: string, v: string | null | undefined): void {
  if (v != null && !HEX_RGB.test(v)) {
    throw new RangeError(`${name} must be six hex digits RRGGBB without #, got ${v}`);
  }
}

function assertPositive(name: string, v: number | null | undefined): void {
  if (v != null && !(Number.isFinite(v) && v > 0)) {
    throw new RangeError(`${name} must be a finite number greater than 0, got ${v}`);
  }
}

function assertRange(
  name: string,
  value: number | null | undefined,
  min: number,
  max: number,
): void {
  if (value == null) return;
  if (!Number.isInteger(value) || value < min || value > max) {
    throw new RangeError(`${name} must be ${min}\u2013${max}, got ${value}`);
  }
}

/** ECMA-376 ST_TextFontSize: hundredths of a point, 100..400000. */
function assertFontSize(name: string, value: number | null | undefined): void {
  assertRange(name, value, 100, 400000);
}

function assertNonEmptyRef(name: string, value: string | null | undefined): void {
  if (value == null) return;
  if (value.trim() === '') {
    throw new RangeError(`${name} must not be empty`);
  }
}

function validateAxis(name: string, axis: ChartAxisData | null | undefined): void {
  if (axis == null) return;
  assertU32(`${name}.id`, axis.id);
  assertU32(`${name}.crossAx`, axis.crossAx);
  assertFontSize(`${name}.fontSize`, axis.fontSize);
  assertFontSize(`${name}.title.fontSize`, axis.title?.fontSize);
  assertHexRgb(`${name}.title.color`, axis.title?.color);
  // ECMA-376 ST_LogBase: 2..1000
  if (
    axis.logBase != null &&
    !(Number.isFinite(axis.logBase) && axis.logBase >= 2 && axis.logBase <= 1000)
  ) {
    throw new RangeError(`${name}.logBase must be a number from 2 to 1000, got ${axis.logBase}`);
  }
  assertPositive(`${name}.majorUnit`, axis.majorUnit);
  assertPositive(`${name}.minorUnit`, axis.minorUnit);
}

/** Throws if a two-cell anchor ends before it starts (columns, then rows). */
function assertAnchorOrder(a: ChartAnchorData): void {
  if (a.extCx != null || a.extCy != null) return;
  if (a.toCol < a.fromCol || (a.toCol === a.fromCol && (a.toColOff ?? 0) < (a.fromColOff ?? 0))) {
    throw new RangeError('anchor must not end before it starts (column)');
  }
  if (a.toRow < a.fromRow || (a.toRow === a.fromRow && (a.toRowOff ?? 0) < (a.fromRowOff ?? 0))) {
    throw new RangeError('anchor must not end before it starts (row)');
  }
}

/**
 * Throws if two series of a chart share a `c:order` (within one chart-type element or across
 * its primary and secondary chart), which Excel refuses. Linear in the number of series.
 */
function assertUniqueOrders(chart: ChartDataModel): void {
  const seen = new Map<number, string>();
  const check = (series: readonly ChartSeriesData[], owner: string): void => {
    for (const [i, ser] of series.entries()) {
      const at = `${owner}series[${i}]`;
      const first = seen.get(ser.order);
      if (first !== undefined) {
        throw new RangeError(
          `series.order must be unique across the chart's series, including its secondary chart's: ${first} and ${at} both have order ${ser.order}`,
        );
      }
      seen.set(ser.order, at);
    }
  };
  check(chart.series, '');
  if (chart.secondaryChart) check(chart.secondaryChart.series, 'secondaryChart.');
}

/**
 * Validates chart data before it crosses the WASM boundary.
 *
 * TypeScript `number` fields that map to Rust `u32`/`i32` are checked
 * for integer-ness, sign, and ECMA-376 spec ranges where applicable, and
 * every series `order` must be unique across the chart and its `secondaryChart`.
 * Throws `RangeError` with a clear message on the first violation.
 */
export function validateChartData(data: WorksheetChartData): void {
  const c = data.chart;

  // holeSize: ECMA-376 ST_HoleSizeUByte, 1-90
  assertRange('holeSize', c.holeSize, 1, 90);

  // styleId: ECMA-376 ST_Style (§21.2.3.46), 1-48
  assertRange('styleId', c.styleId, 1, 48);

  assertFontSize('title.fontSize', c.title?.fontSize);
  assertHexRgb('title.color', c.title?.color);
  validateAxis('catAxis', c.catAxis);
  validateAxis('valAxis', c.valAxis);

  // View3D ranges (ECMA-376 spec)
  if (c.view3d) {
    assertRange('view3d.rotX', c.view3d.rotX, -90, 90);
    assertRange('view3d.rotY', c.view3d.rotY, 0, 360);
    assertRange('view3d.perspective', c.view3d.perspective, 0, 240);
  }

  // Series validation
  for (const ser of c.series) {
    assertU32('series.lineWidth', ser.lineWidth);
    assertU32('series.explosion', ser.explosion);
    assertHexRgb('series.fillColor', ser.fillColor);
    assertHexRgb('series.lineColor', ser.lineColor);
    assertU32('series.idx', ser.idx);
    assertU32('series.order', ser.order);
    assertNonEmptyRef('series.valRef', ser.valRef);
    assertNonEmptyRef('series.catRef', ser.catRef);
    assertNonEmptyRef('series.xValRef', ser.xValRef);
    assertNonEmptyRef('series.bubbleSizeRef', ser.bubbleSizeRef);
    assertRange('trendline.order', ser.trendline?.order, 2, 6);
    // ST_Period: unsignedInt >= 2
    assertRange('trendline.period', ser.trendline?.period, 2, 4294967295);
  }
  for (const ser of c.secondaryChart?.series ?? []) {
    assertU32('secondaryChart.series.idx', ser.idx);
    assertU32('secondaryChart.series.order', ser.order);
  }
  assertUniqueOrders(c);

  // Anchor validation
  const a = data.anchor;
  assertU32('anchor.fromCol', a.fromCol);
  assertU32('anchor.fromRow', a.fromRow);
  assertU32('anchor.toCol', a.toCol);
  assertU32('anchor.toRow', a.toRow);
  assertU32('anchor.fromColOff', a.fromColOff);
  assertU32('anchor.fromRowOff', a.fromRowOff);
  assertU32('anchor.toColOff', a.toColOff);
  assertU32('anchor.toRowOff', a.toRowOff);
  assertAnchorOrder(a);
}
