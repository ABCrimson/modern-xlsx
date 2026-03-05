import type { WorksheetChartData } from './types.js';

function assertU32(name: string, value: number | null | undefined): void {
  if (value == null) return;
  if (!Number.isInteger(value) || value < 0 || value > 4294967295) {
    throw new RangeError(`${name} must be a non-negative integer (u32), got ${value}`);
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

/**
 * Validates chart data before it crosses the WASM boundary.
 *
 * TypeScript `number` fields that map to Rust `u32`/`i32` are checked
 * for integer-ness, sign, and ECMA-376 spec ranges where applicable.
 * Throws `RangeError` with a clear message on the first violation.
 */
export function validateChartData(data: WorksheetChartData): void {
  const c = data.chart;

  // holeSize: u32, valid 0-90 for doughnut charts (ECMA-376 ST_HoleSize)
  if (c.holeSize != null) {
    assertRange('holeSize', c.holeSize, 0, 90);
  }

  // styleId: u32
  assertU32('styleId', c.styleId);

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
    assertU32('series.idx', ser.idx);
    assertU32('series.order', ser.order);
    if (ser.valRef === '') {
      throw new Error('series.valRef must not be empty');
    }
  }

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
}
