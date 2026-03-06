/**
 * ITF-14 Encoder (Interleaved 2 of 5).
 *
 * @module barcode/itf14
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

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

export function encodeITF14(data: string): BarcodeMatrix {
  let digits: number[];
  if (data.length === 13) {
    digits = [...data].map(Number);
    digits.push(itfChk(digits));
  } else if (data.length === 14) {
    digits = [...data].map(Number);
    const exp = itfChk(digits.slice(0, 13));
    if (at(digits, 13) !== exp)
      throw new Error(`ITF-14: invalid check digit (expected ${exp}, got ${at(digits, 13)})`);
  } else {
    throw new Error('ITF-14: data must be 13 or 14 digits');
  }
  if (digits.some((d) => d < 0 || d > 9 || !Number.isFinite(d))) {
    throw new Error('ITF-14: data must contain only digits 0-9');
  }

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

  const w = bars.length;
  const bh = 60;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++) for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c));
  return { width: w, height: bh, modules };
}
