/**
 * EAN-13 and UPC-A Encoders.
 *
 * @module barcode/ean13
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

// ---------------------------------------------------------------------------
// EAN-13 Encoder
// ---------------------------------------------------------------------------

const EAN_L: number[][] = [
  [0, 0, 0, 1, 1, 0, 1],
  [0, 0, 1, 1, 0, 0, 1],
  [0, 0, 1, 0, 0, 1, 1],
  [0, 1, 1, 1, 1, 0, 1],
  [0, 1, 0, 0, 0, 1, 1],
  [0, 1, 1, 0, 0, 0, 1],
  [0, 1, 0, 1, 1, 1, 1],
  [0, 1, 1, 1, 0, 1, 1],
  [0, 1, 1, 0, 1, 1, 1],
  [0, 0, 0, 1, 0, 1, 1],
];
const EAN_G: number[][] = [
  [0, 1, 0, 0, 1, 1, 1],
  [0, 1, 1, 0, 0, 1, 1],
  [0, 0, 1, 1, 0, 1, 1],
  [0, 1, 0, 0, 0, 0, 1],
  [0, 0, 1, 1, 1, 0, 1],
  [0, 1, 1, 1, 0, 0, 1],
  [0, 0, 0, 0, 1, 0, 1],
  [0, 0, 1, 0, 0, 0, 1],
  [0, 0, 0, 1, 0, 0, 1],
  [0, 0, 1, 0, 1, 1, 1],
];
const EAN_R: number[][] = EAN_L.map((p) => p.map((b) => b ^ 1));
const EAN_PAR: number[][] = [
  [0, 0, 0, 0, 0, 0],
  [0, 0, 1, 0, 1, 1],
  [0, 0, 1, 1, 0, 1],
  [0, 0, 1, 1, 1, 0],
  [0, 1, 0, 0, 1, 1],
  [0, 1, 1, 0, 0, 1],
  [0, 1, 1, 1, 0, 0],
  [0, 1, 0, 1, 0, 1],
  [0, 1, 0, 1, 1, 0],
  [0, 1, 1, 0, 1, 0],
];

function ean13Chk(d: number[]): number {
  let s = 0;
  for (let i = 0; i < 12; i++) s += at(d, i) * (i % 2 === 0 ? 1 : 3);
  return (10 - (s % 10)) % 10;
}

export function encodeEAN13(data: string): BarcodeMatrix {
  let digits: number[];
  if (data.length === 12) {
    digits = [...data].map(Number);
    digits.push(ean13Chk(digits));
  } else if (data.length === 13) {
    digits = [...data].map(Number);
    const exp = ean13Chk(digits.slice(0, 12));
    if (at(digits, 12) !== exp)
      throw new Error(`EAN-13: invalid check digit (expected ${exp}, got ${at(digits, 12)})`);
  } else {
    throw new Error('EAN-13: data must be 12 or 13 digits');
  }
  if (digits.some((d) => d < 0 || d > 9 || !Number.isFinite(d))) {
    throw new Error('EAN-13: data must contain only digits 0-9');
  }

  const bars: number[] = [1, 0, 1]; // start guard
  const par = at(EAN_PAR, at(digits, 0));
  for (let i = 0; i < 6; i++) {
    const pat = at(par, i) === 0 ? at(EAN_L, at(digits, i + 1)) : at(EAN_G, at(digits, i + 1));
    bars.push(...pat);
  }
  bars.push(0, 1, 0, 1, 0); // center guard
  for (let i = 7; i <= 12; i++) bars.push(...at(EAN_R, at(digits, i)));
  bars.push(1, 0, 1); // end guard

  const w = bars.length;
  const bh = 60;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++)
    for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c) === 1);
  return { width: w, height: bh, modules };
}

// ---------------------------------------------------------------------------
// UPC-A Encoder
// ---------------------------------------------------------------------------

export function encodeUPCA(data: string): BarcodeMatrix {
  if (data.length === 11) return encodeEAN13(`0${data}`);
  if (data.length === 12) return encodeEAN13(`0${data.slice(0, 11)}`);
  throw new Error('UPC-A: data must be 11 or 12 digits');
}
