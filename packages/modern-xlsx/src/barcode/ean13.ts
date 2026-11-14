/**
 * EAN-13 and UPC-A Encoders.
 *
 * @module barcode/ean13
 */

import type { BarcodeMatrix } from './common.js';
import { at } from './common.js';
import { barsToMatrix } from './linear.js';

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

/**
 * Encode an EAN-13 symbol from 12 digits (the check digit is computed) or 13 digits (the supplied
 * check digit is verified).
 *
 * @param data - 12 or 13 ASCII digits.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If `data` is not 12 or 13 ASCII digits, or the supplied check digit is wrong.
 */
export function encodeEAN13(data: string): BarcodeMatrix {
  if (!/^\d+$/.test(data)) throw new Error('EAN-13: data must contain only digits 0-9');
  const digits = [...data].map(Number);
  if (data.length === 12) {
    digits.push(ean13Chk(digits));
  } else if (data.length === 13) {
    const exp = ean13Chk(digits);
    if (at(digits, 12) !== exp)
      throw new Error(`EAN-13: invalid check digit (expected ${exp}, got ${at(digits, 12)})`);
  } else {
    throw new Error('EAN-13: data must be 12 or 13 digits');
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

  return barsToMatrix(bars.map((b) => b === 1));
}

// ---------------------------------------------------------------------------
// UPC-A Encoder
// ---------------------------------------------------------------------------

/**
 * Encode a UPC-A symbol from 11 digits (the check digit is computed) or 12 digits (the supplied
 * check digit is verified).
 *
 * @param data - 11 or 12 ASCII digits.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If `data` is not 11 or 12 ASCII digits, or the supplied check digit is wrong.
 */
export function encodeUPCA(data: string): BarcodeMatrix {
  if (!/^\d+$/.test(data)) throw new Error('UPC-A: data must contain only digits 0-9');
  if (data.length !== 11 && data.length !== 12) {
    throw new Error('UPC-A: data must be 11 or 12 digits');
  }
  return encodeEAN13(`0${data}`);
}
