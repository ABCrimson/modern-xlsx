/**
 * Data Matrix Encoder (ECC 200, ASCII mode).
 *
 * @module barcode/datamatrix
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, DM_FIELD, rsEncode, setCell } from './common.js';

// ---------------------------------------------------------------------------
// Data Matrix Encoder (ECC 200, ASCII mode)
// ---------------------------------------------------------------------------

const DM_SIZES: [number, number, number, number][] = [
  [10, 10, 3, 5],
  [12, 12, 5, 7],
  [14, 14, 8, 10],
  [16, 16, 12, 12],
  [18, 18, 18, 14],
  [20, 20, 22, 18],
  [22, 22, 30, 20],
  [24, 24, 36, 24],
  [26, 26, 44, 28],
  [32, 32, 62, 36],
  [36, 36, 86, 42],
  [40, 40, 114, 48],
  [44, 44, 144, 56],
  [48, 48, 174, 68],
];

/** @internal */
export function dmAsciiEnc(data: string): number[] {
  const cps = Array.from(data, (ch) => ch.codePointAt(0) as number);
  const cws: number[] = [];
  // Latin-1 text keeps the default interpretation; anything wider switches to UTF-8 via
  // ECI 000026 (codeword 241, then value + 1 = 27) and encodes the UTF-8 bytes.
  const latin1 = cps.every((cp) => cp <= 255);
  if (!latin1) cws.push(241, 27);
  const units = latin1 ? cps : Array.from(new TextEncoder().encode(data));
  let i = 0;
  while (i < units.length) {
    const ch = units[i] as number;
    const ch2 = units[i + 1];
    if (ch >= 48 && ch <= 57 && ch2 !== undefined && ch2 >= 48 && ch2 <= 57) {
      cws.push((ch - 48) * 10 + (ch2 - 48) + 130);
      i += 2;
      continue;
    }
    if (ch <= 127) cws.push(ch + 1);
    else cws.push(235, ch - 127);
    i++;
  }
  return cws;
}

function dmSelSize(len: number): [number, number, number, number] {
  for (const sz of DM_SIZES) if (len <= sz[2]) return sz;
  throw new Error('Data Matrix: data too long');
}

/** @internal */
export function dmPad(cws: number[], cap: number): number[] {
  const p = [...cws];
  if (p.length < cap) p.push(129);
  while (p.length < cap) {
    const r = ((149 * (p.length + 1)) % 253) + 1;
    const t = 129 + r;
    p.push(t <= 254 ? t : t - 254);
  }
  return p;
}

/**
 * Encodes a Data Matrix (ECC 200) symbol.
 *
 * @experimental Module placement is not ISO/IEC 16022 conformant and symbols may not scan.
 *
 * @param data - Text to encode. Latin-1 text is encoded as is; any other text is encoded as UTF-8 behind an ECI marker.
 * @returns The symbol matrix of the smallest square size that fits (10x10 to 48x48), without quiet zone.
 * @throws {Error} `Data Matrix: data too long` if the data exceeds the 48x48 capacity.
 */
export function encodeDataMatrix(data: string): BarcodeMatrix {
  const cws = dmAsciiEnc(data);
  const [rows, cols, dataCap, ecCount] = dmSelSize(cws.length);
  const padded = dmPad(cws, dataCap);
  const ecCws = Array.from(rsEncode(Uint8Array.from(padded), ecCount, DM_FIELD, 1));
  const allCws = [...padded, ...ecCws];

  const modules = createGrid(rows, cols, false);

  // L-shaped finder pattern
  for (let c = 0; c < cols; c++) {
    setCell(modules, rows - 1, c, true);
    setCell(modules, 0, c, c % 2 === 0);
  }
  for (let r = 0; r < rows; r++) {
    setCell(modules, r, 0, true);
    setCell(modules, r, cols - 1, r % 2 === 0);
  }

  // Data region
  const dR = rows - 2;
  const dC = cols - 2;
  const bits: boolean[] = [];
  for (const cw of allCws) {
    for (let b = 7; b >= 0; b--) bits.push(((cw >> b) & 1) === 1);
  }

  // Place bits sequentially in data region (simplified placement)
  let bIdx = 0;
  for (let r = 0; r < dR; r++) {
    for (let c = 0; c < dC; c++) {
      const val = bIdx < bits.length ? at(bits, bIdx) : false;
      setCell(modules, r + 1, c + 1, val);
      bIdx++;
    }
  }

  return { width: cols, height: rows, modules };
}
