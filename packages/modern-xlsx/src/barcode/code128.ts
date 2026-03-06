/**
 * Code 128 Encoder.
 *
 * @module barcode/code128
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

// ---------------------------------------------------------------------------
// Code 128 Encoder
// ---------------------------------------------------------------------------

export const C128_W: number[][] = [
  [2, 1, 2, 2, 2, 2],
  [2, 2, 2, 1, 2, 2],
  [2, 2, 2, 2, 2, 1],
  [1, 2, 1, 2, 2, 3],
  [1, 2, 1, 3, 2, 2], // 0-4
  [1, 3, 1, 2, 2, 2],
  [1, 2, 2, 2, 1, 3],
  [1, 2, 2, 3, 1, 2],
  [1, 3, 2, 2, 1, 2],
  [2, 2, 1, 2, 1, 3], // 5-9
  [2, 2, 1, 3, 1, 2],
  [2, 3, 1, 2, 1, 2],
  [1, 1, 2, 2, 3, 2],
  [1, 2, 2, 1, 3, 2],
  [1, 2, 2, 2, 3, 1], // 10-14
  [1, 1, 3, 2, 2, 2],
  [1, 2, 3, 1, 2, 2],
  [1, 2, 3, 2, 2, 1],
  [2, 2, 3, 2, 1, 1],
  [2, 2, 1, 1, 3, 2], // 15-19
  [2, 2, 1, 2, 3, 1],
  [2, 1, 3, 2, 1, 2],
  [2, 2, 3, 1, 1, 2],
  [3, 1, 2, 1, 3, 1],
  [3, 1, 1, 2, 2, 2], // 20-24
  [3, 2, 1, 1, 2, 2],
  [3, 2, 1, 2, 2, 1],
  [3, 1, 2, 2, 1, 2],
  [3, 2, 2, 1, 1, 2],
  [3, 2, 2, 2, 1, 1], // 25-29
  [2, 1, 2, 1, 2, 3],
  [2, 1, 2, 3, 2, 1],
  [2, 3, 2, 1, 2, 1],
  [1, 1, 1, 3, 2, 3],
  [1, 3, 1, 1, 2, 3], // 30-34
  [1, 3, 1, 3, 2, 1],
  [1, 1, 2, 3, 1, 3],
  [1, 3, 2, 1, 1, 3],
  [1, 3, 2, 3, 1, 1],
  [2, 1, 1, 3, 1, 3], // 35-39
  [2, 3, 1, 1, 1, 3],
  [2, 3, 1, 3, 1, 1],
  [1, 1, 2, 1, 3, 3],
  [1, 1, 2, 3, 3, 1],
  [1, 3, 2, 1, 3, 1], // 40-44
  [1, 1, 3, 1, 2, 3],
  [1, 1, 3, 3, 2, 1],
  [1, 3, 3, 1, 2, 1],
  [3, 1, 3, 1, 2, 1],
  [2, 1, 1, 3, 3, 1], // 45-49
  [2, 3, 1, 1, 3, 1],
  [2, 1, 3, 1, 1, 3],
  [2, 1, 3, 3, 1, 1],
  [2, 1, 3, 1, 3, 1],
  [3, 1, 1, 1, 2, 3], // 50-54
  [3, 1, 1, 3, 2, 1],
  [3, 3, 1, 1, 2, 1],
  [3, 1, 2, 1, 1, 3],
  [3, 1, 2, 3, 1, 1],
  [3, 3, 2, 1, 1, 1], // 55-59
  [3, 1, 4, 1, 1, 1],
  [2, 2, 1, 4, 1, 1],
  [4, 3, 1, 1, 1, 1],
  [1, 1, 1, 2, 2, 4],
  [1, 1, 1, 4, 2, 2], // 60-64
  [1, 2, 1, 1, 2, 4],
  [1, 2, 1, 4, 2, 1],
  [1, 4, 1, 1, 2, 2],
  [1, 4, 1, 2, 2, 1],
  [1, 1, 2, 2, 1, 4], // 65-69
  [1, 1, 2, 4, 1, 2],
  [1, 2, 2, 1, 1, 4],
  [1, 2, 2, 4, 1, 1],
  [1, 4, 2, 1, 1, 2],
  [1, 4, 2, 2, 1, 1], // 70-74
  [2, 4, 1, 2, 1, 1],
  [2, 2, 1, 1, 1, 4],
  [4, 1, 3, 1, 1, 1],
  [2, 4, 1, 1, 1, 2],
  [1, 3, 4, 1, 1, 1], // 75-79
  [1, 1, 1, 2, 4, 2],
  [1, 2, 1, 1, 4, 2],
  [1, 2, 1, 2, 4, 1],
  [1, 1, 4, 2, 1, 2],
  [1, 2, 4, 1, 1, 2], // 80-84
  [1, 2, 4, 2, 1, 1],
  [4, 1, 1, 2, 1, 2],
  [4, 2, 1, 1, 1, 2],
  [4, 2, 1, 2, 1, 1],
  [2, 1, 2, 1, 4, 1], // 85-89
  [2, 1, 4, 1, 2, 1],
  [4, 1, 2, 1, 2, 1],
  [1, 1, 1, 1, 4, 3],
  [1, 1, 1, 3, 4, 1],
  [1, 3, 1, 1, 4, 1], // 90-94
  [1, 1, 4, 1, 1, 3],
  [1, 1, 4, 3, 1, 1],
  [4, 1, 1, 1, 1, 3],
  [4, 1, 1, 3, 1, 1],
  [1, 1, 3, 1, 4, 1], // 95-99
  [1, 1, 4, 1, 3, 1],
  [3, 1, 1, 1, 4, 1],
  [4, 1, 1, 1, 3, 1],
  [2, 1, 1, 4, 1, 2],
  [2, 1, 1, 2, 1, 4], // 100-104
  [2, 1, 1, 2, 3, 2], // 105: START C
];

export const C128_STOP: number[] = [2, 3, 3, 1, 1, 1, 2];

export function widths2mod(widths: number[]): boolean[] {
  const m: boolean[] = [];
  let black = true;
  for (const w of widths) {
    for (let i = 0; i < w; i++) m.push(black);
    black = !black;
  }
  return m;
}

export function encodeCode128(data: string): BarcodeMatrix {
  if (data.length === 0) throw new Error('Code 128: data must not be empty');

  const vals: number[] = [];
  let chk: number;
  const allDig = /^\d+$/.test(data) && data.length % 2 === 0;

  if (allDig) {
    vals.push(105);
    chk = 105;
    for (let i = 0; i < data.length; i += 2) {
      const v = (data.charCodeAt(i) - 48) * 10 + (data.charCodeAt(i + 1) - 48);
      vals.push(v);
      chk += v * (i / 2 + 1);
    }
  } else {
    vals.push(104);
    chk = 104;
    let i = 0;
    while (i < data.length) {
      let dRun = 0;
      let j = i;
      while (j < data.length && data.charCodeAt(j) >= 48 && data.charCodeAt(j) <= 57) {
        dRun++;
        j++;
      }
      if (dRun >= 4 && dRun % 2 === 0) {
        vals.push(99);
        chk += 99 * vals.length;
        for (let k = i; k < i + dRun; k += 2) {
          const v = (data.charCodeAt(k) - 48) * 10 + (data.charCodeAt(k + 1) - 48);
          vals.push(v);
          chk += v * vals.length;
        }
        i += dRun;
        if (i < data.length) {
          vals.push(100);
          chk += 100 * vals.length;
        }
      } else {
        const ch = data.charCodeAt(i);
        if (ch < 32 || ch > 127) throw new Error(`Code 128: unsupported character code ${ch}`);
        const v = ch - 32;
        vals.push(v);
        chk += v * (vals.length - 1);
        i++;
      }
    }
  }
  chk %= 103;
  vals.push(chk);

  const bars: boolean[] = [];
  for (const v of vals) bars.push(...widths2mod(at(C128_W, v)));
  bars.push(...widths2mod(C128_STOP));

  const bh = 60;
  const w = bars.length;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++) for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c));
  return { width: w, height: bh, modules };
}
