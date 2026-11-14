/**
 * Code 128 Encoder.
 *
 * @module barcode/code128
 */

import type { BarcodeMatrix } from './common.js';
import { at } from './common.js';
import { barsToMatrix, widths2mod } from './linear.js';

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

const C128_STOP: number[] = [2, 3, 3, 1, 1, 1, 2];

/**
 * Module row of a Code 128 symbol: each symbol value's bar widths (checksum value included), then
 * the stop pattern.
 */
export function code128Bars(vals: readonly number[]): boolean[] {
  const bars: boolean[] = [];
  for (const v of vals) bars.push(...widths2mod(at(C128_W, v)));
  bars.push(...widths2mod(C128_STOP));
  return bars;
}

/** Length of the run of ASCII digits starting at `from`. */
function digitRunLength(data: string, from: number): number {
  let end = from;
  while (end < data.length && data.charCodeAt(end) >= 48 && data.charCodeAt(end) <= 57) end++;
  return end - from;
}

/**
 * Appends to `out` the Code C symbol values for the digit pairs in `data[from, from + count)`
 * (`count` is even). Pushes one value at a time: spreading a long run into a single `push` call
 * would exceed the engine's argument limit.
 */
function pushDigitPairs(out: number[], data: string, from: number, count: number): void {
  for (let k = from; k < from + count; k += 2) {
    out.push((data.charCodeAt(k) - 48) * 10 + (data.charCodeAt(k + 1) - 48));
  }
}

/** Symbol values (start code included, checksum excluded) of a Code 128 B symbol with Code C runs. */
function mixedSymbolValues(data: string): number[] {
  const vals: number[] = [104];
  let i = 0;
  while (i < data.length) {
    const run = digitRunLength(data, i);
    if (run >= 4 && run % 2 === 0) {
      vals.push(99);
      pushDigitPairs(vals, data, i, run);
      i += run;
      if (i < data.length) vals.push(100);
    } else {
      const ch = data.charCodeAt(i);
      if (ch < 32 || ch > 127) throw new Error(`Code 128: unsupported character code ${ch}`);
      vals.push(ch - 32);
      i++;
    }
  }
  return vals;
}

/** Symbol values (start code included, checksum excluded) for `data`. */
function symbolValues(data: string): number[] {
  const allDigitPairs = /^\d+$/.test(data) && data.length % 2 === 0;
  if (!allDigitPairs) return mixedSymbolValues(data);
  const vals = [105];
  pushDigitPairs(vals, data, 0, data.length);
  return vals;
}

/** Checksum: the start value plus each following symbol value times its symbol index, modulo 103. */
export function code128Checksum(vals: readonly number[]): number {
  let chk = at(vals, 0);
  for (let i = 1; i < vals.length; i++) chk += at(vals, i) * i;
  return chk % 103;
}

/**
 * Encodes a Code 128 symbol from ASCII characters 32-127. A run of four or more digits switches
 * to Code C (an odd-length run starts with one Code B digit); all other characters use Code B. All-digit data of even length is one Code C
 * symbol. The check character is added.
 *
 * @param data - Text to encode; must not be empty.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If `data` is empty or contains a character outside ASCII 32-127.
 */
export function encodeCode128(data: string): BarcodeMatrix {
  if (data.length === 0) throw new Error('Code 128: data must not be empty');

  const vals = symbolValues(data);
  vals.push(code128Checksum(vals));

  return barsToMatrix(code128Bars(vals));
}
