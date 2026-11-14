/**
 * Shared helpers for the linear (1D) barcode encoders.
 *
 * @module barcode/linear
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

/**
 * Expands alternating bar/space run widths (starting with a bar) into one boolean per module,
 * `true` for a bar.
 */
export function widths2mod(widths: number[]): boolean[] {
  const m: boolean[] = [];
  let black = true;
  for (const w of widths) {
    for (let i = 0; i < w; i++) m.push(black);
    black = !black;
  }
  return m;
}

/**
 * Stretches a module row (`true` for a bar) into a `height`-row matrix. Every row is its own array.
 *
 * @param bars - One entry per module.
 * @param height - Number of identical rows; defaults to 60.
 */
export function barsToMatrix(bars: readonly boolean[], height = 60): BarcodeMatrix {
  const w = bars.length;
  const modules = createGrid(height, w, false);
  for (let r = 0; r < height; r++) for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c));
  return { width: w, height, modules };
}
