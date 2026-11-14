/**
 * ITF-14 Encoder (Interleaved 2 of 5).
 *
 * @module barcode/itf14
 */

import type { BarcodeMatrix } from './common.js';
import { at } from './common.js';
import { barsToMatrix } from './linear.js';

// ---------------------------------------------------------------------------
// ITF-14 Encoder (Interleaved 2 of 5)
// ---------------------------------------------------------------------------

const ITF_PAT: number[][] = [
  [1, 1, 2, 2, 1],
  [2, 1, 1, 1, 2],
  [1, 2, 1, 1, 2],
  [2, 2, 1, 1, 1],
  [1, 1, 2, 1, 2],
  [2, 1, 2, 1, 1],
  [1, 2, 2, 1, 1],
  [1, 1, 1, 2, 2],
  [2, 1, 1, 2, 1],
  [1, 2, 1, 2, 1],
];

function itfChk(d: number[]): number {
  let s = 0;
  for (let i = 0; i < 13; i++) s += at(d, i) * (i % 2 === 0 ? 3 : 1);
  return (10 - (s % 10)) % 10;
}

/** The 14 digits of the symbol: the 13 given plus the computed check digit, or the 14 verified. */
function itf14Digits(data: string): number[] {
  if (!/^\d+$/.test(data)) throw new Error('ITF-14: data must contain only digits 0-9');
  const digits = [...data].map(Number);
  if (data.length === 13) {
    digits.push(itfChk(digits));
  } else if (data.length === 14) {
    const exp = itfChk(digits);
    if (at(digits, 13) !== exp)
      throw new Error(`ITF-14: invalid check digit (expected ${exp}, got ${at(digits, 13)})`);
  } else {
    throw new Error('ITF-14: data must be 13 or 14 digits');
  }
  return digits;
}

/**
 * Encode an ITF-14 symbol from 13 digits (the check digit is computed) or 14 digits (the supplied
 * check digit is verified).
 *
 * @param data - 13 or 14 ASCII digits.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If `data` is not 13 or 14 ASCII digits, or the supplied check digit is wrong.
 */
export function encodeITF14(data: string): BarcodeMatrix {
  const digits = itf14Digits(data);

  const bars: boolean[] = [true, false, true, false]; // start
  for (let i = 0; i < digits.length; i += 2) {
    const d1 = at(ITF_PAT, at(digits, i));
    const d2 = at(ITF_PAT, at(digits, i + 1));
    for (let j = 0; j < 5; j++) {
      for (let k = 0; k < at(d1, j); k++) bars.push(true);
      for (let k = 0; k < at(d2, j); k++) bars.push(false);
    }
  }
  bars.push(true, true, false, true); // stop

  return barsToMatrix(bars);
}
