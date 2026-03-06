/**
 * GS1-128 Encoder.
 *
 * @module barcode/gs1128
 */

import { C128_STOP, C128_W, widths2mod } from './code128.js';
import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

// ---------------------------------------------------------------------------
// GS1-128 Encoder
// ---------------------------------------------------------------------------

export function encodeGS1128(data: string): BarcodeMatrix {
  const stripped = data.replace(/[()]/g, '');
  if (stripped.length === 0) throw new Error('GS1-128: data must not be empty');

  const allDig = /^\d+$/.test(stripped) && stripped.length % 2 === 0;
  const vals: number[] = [];
  let chk: number;

  if (allDig) {
    vals.push(105);
    chk = 105;
    vals.push(102);
    chk += 102;
    let pos = 2;
    for (let i = 0; i < stripped.length; i += 2) {
      const v = (stripped.charCodeAt(i) - 48) * 10 + (stripped.charCodeAt(i + 1) - 48);
      vals.push(v);
      chk += v * pos;
      pos++;
    }
  } else {
    vals.push(104);
    chk = 104;
    vals.push(102);
    chk += 102;
    let pos = 2;
    for (let i = 0; i < stripped.length; i++) {
      const ch = stripped.charCodeAt(i);
      if (ch < 32 || ch > 127) throw new Error(`GS1-128: unsupported character code ${ch}`);
      const v = ch - 32;
      vals.push(v);
      chk += v * pos;
      pos++;
    }
  }
  chk %= 103;
  vals.push(chk);

  const bars: boolean[] = [];
  for (const v of vals) bars.push(...widths2mod(at(C128_W, v)));
  bars.push(...widths2mod(C128_STOP));

  const w = bars.length;
  const bh = 60;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++) for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c));
  return { width: w, height: bh, modules };
}
