/**
 * Data Matrix Encoder (ECC 200, ASCII mode).
 *
 * @module barcode/datamatrix
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

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

function dmAsciiEnc(data: string): number[] {
  const cws: number[] = [];
  let i = 0;
  while (i < data.length) {
    const ch = data.charCodeAt(i);
    if (i + 1 < data.length) {
      const ch2 = data.charCodeAt(i + 1);
      if (ch >= 48 && ch <= 57 && ch2 >= 48 && ch2 <= 57) {
        cws.push((ch - 48) * 10 + (ch2 - 48) + 130);
        i += 2;
        continue;
      }
    }
    if (ch >= 0 && ch <= 127) cws.push(ch + 1);
    else {
      cws.push(235);
      cws.push(ch - 127);
    }
    i++;
  }
  return cws;
}

function dmSelSize(len: number): [number, number, number, number] {
  for (const sz of DM_SIZES) if (len <= sz[2]) return sz;
  throw new Error('Data Matrix: data too long');
}

function dmPad(cws: number[], cap: number): number[] {
  const p = [...cws];
  if (p.length < cap) p.push(129);
  while (p.length < cap) {
    const r = ((149 * (p.length + 1)) % 253) + 1;
    p.push((129 + r) % 254);
  }
  return p;
}

const DM_GF_EXP = new Uint8Array(512);
const DM_GF_LOG = new Uint8Array(256);
{
  let v = 1;
  for (let i = 0; i < 255; i++) {
    DM_GF_EXP[i] = v;
    DM_GF_LOG[v] = i;
    v <<= 1;
    if (v >= 256) v ^= 0x12d;
  }
  for (let i = 255; i < 512; i++) DM_GF_EXP[i] = DM_GF_EXP[i - 255] ?? 0;
}

function dmEC(data: number[], numEc: number): number[] {
  const mul = (a: number, b: number): number => {
    if (a === 0 || b === 0) return 0;
    return DM_GF_EXP[(DM_GF_LOG[a] ?? 0) + (DM_GF_LOG[b] ?? 0)] ?? 0;
  };

  let g = new Uint8Array([1]);
  for (let i = 0; i < numEc; i++) {
    const next = new Uint8Array(g.length + 1);
    const factor = DM_GF_EXP[i + 1] ?? 0;
    for (let j = 0; j < g.length; j++) {
      next[j] = (next[j] ?? 0) ^ (g[j] ?? 0);
      next[j + 1] = (next[j + 1] ?? 0) ^ mul(g[j] ?? 0, factor);
    }
    g = next;
  }

  const result = new Uint8Array(numEc);
  for (const cw of data) {
    const coef = cw ^ (result[0] ?? 0);
    for (let j = 0; j < numEc - 1; j++) result[j] = (result[j + 1] ?? 0) ^ mul(g[j + 1] ?? 0, coef);
    result[numEc - 1] = mul(g[numEc] ?? 0, coef);
  }
  return Array.from(result);
}

export function encodeDataMatrix(data: string): BarcodeMatrix {
  const cws = dmAsciiEnc(data);
  const [rows, cols, dataCap, ecCount] = dmSelSize(cws.length);
  const padded = dmPad(cws, dataCap);
  const ecCws = dmEC(padded, ecCount);
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
