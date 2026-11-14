/**
 * Code 39 Encoder.
 *
 * @module barcode/code39
 */

import type { BarcodeMatrix } from './common.js';
import { at } from './common.js';
import { barsToMatrix } from './linear.js';

// ---------------------------------------------------------------------------
// Code 39 Encoder
// ---------------------------------------------------------------------------

const C39_CHARS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*';
const C39_PAT: number[][] = [
  [1, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, 1],
  [1, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 1],
  [1, 0, 1, 1, 0, 0, 1, 0, 1, 0, 1, 1],
  [1, 1, 0, 1, 1, 0, 0, 1, 0, 1, 0, 1],
  [1, 0, 1, 0, 0, 1, 1, 0, 1, 0, 1, 1],
  [1, 1, 0, 1, 0, 0, 1, 1, 0, 1, 0, 1],
  [1, 0, 1, 1, 0, 0, 1, 1, 0, 1, 0, 1],
  [1, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1, 1],
  [1, 1, 0, 1, 0, 0, 1, 0, 1, 1, 0, 1],
  [1, 0, 1, 1, 0, 0, 1, 0, 1, 1, 0, 1],
  [1, 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 1],
  [1, 0, 1, 1, 0, 1, 0, 0, 1, 0, 1, 1],
  [1, 1, 0, 1, 1, 0, 1, 0, 0, 1, 0, 1],
  [1, 0, 1, 0, 1, 1, 0, 0, 1, 0, 1, 1],
  [1, 1, 0, 1, 0, 1, 1, 0, 0, 1, 0, 1],
  [1, 0, 1, 1, 0, 1, 1, 0, 0, 1, 0, 1],
  [1, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1, 1],
  [1, 1, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1],
  [1, 0, 1, 1, 0, 1, 0, 0, 1, 1, 0, 1],
  [1, 0, 1, 0, 1, 1, 0, 0, 1, 1, 0, 1],
  [1, 1, 0, 1, 0, 1, 0, 1, 0, 0, 1, 1],
  [1, 0, 1, 1, 0, 1, 0, 1, 0, 0, 1, 1],
  [1, 1, 0, 1, 1, 0, 1, 0, 1, 0, 0, 1],
  [1, 0, 1, 0, 1, 1, 0, 1, 0, 0, 1, 1],
  [1, 1, 0, 1, 0, 1, 1, 0, 1, 0, 0, 1],
  [1, 0, 1, 1, 0, 1, 1, 0, 1, 0, 0, 1],
  [1, 0, 1, 0, 1, 0, 1, 1, 0, 0, 1, 1],
  [1, 1, 0, 1, 0, 1, 0, 1, 1, 0, 0, 1],
  [1, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 1],
  [1, 0, 1, 0, 1, 1, 0, 1, 1, 0, 0, 1],
  [1, 1, 0, 0, 1, 0, 1, 0, 1, 0, 1, 1],
  [1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1, 1],
  [1, 1, 0, 0, 1, 1, 0, 1, 0, 1, 0, 1],
  [1, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1, 1],
  [1, 1, 0, 0, 1, 0, 1, 1, 0, 1, 0, 1],
  [1, 0, 0, 1, 1, 0, 1, 1, 0, 1, 0, 1],
  [1, 0, 0, 1, 0, 1, 0, 1, 1, 0, 1, 1],
  [1, 1, 0, 0, 1, 0, 1, 0, 1, 1, 0, 1],
  [1, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 1],
  [1, 0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 1],
  [1, 0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 1],
  [1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 1],
  [1, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1],
  [1, 0, 0, 1, 0, 1, 1, 0, 1, 1, 0, 1],
];

/**
 * Encodes a Code 39 symbol. Lower-case letters are upper-cased. The characters are `0-9`, `A-Z`,
 * `- . $ / + %` and space; the `*` start/stop characters are added, with no check character.
 *
 * @param data - Text to encode; must not be empty and must not contain `*`.
 * @returns The symbol matrix (60 rows), without quiet zone.
 * @throws {Error} If `data` is empty, contains `*`, or contains an unsupported character.
 */
export function encodeCode39(data: string): BarcodeMatrix {
  if (data.length === 0) throw new Error('Code 39: data must not be empty');
  const u = data.toUpperCase();
  for (const ch of u) {
    if (ch === '*') throw new Error('Code 39: data must not contain *');
    if (!C39_CHARS.includes(ch)) throw new Error(`Code 39: unsupported character '${ch}'`);
  }
  const starI = C39_CHARS.indexOf('*');
  const bars: number[] = [...at(C39_PAT, starI), 0];
  for (const ch of u) {
    bars.push(...at(C39_PAT, C39_CHARS.indexOf(ch)), 0);
  }
  bars.push(...at(C39_PAT, starI));

  return barsToMatrix(bars.map((b) => b === 1));
}
