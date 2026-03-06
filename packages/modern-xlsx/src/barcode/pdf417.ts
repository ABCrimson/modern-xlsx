/**
 * PDF417 Encoder (text compaction, simplified).
 *
 * @module barcode/pdf417
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';

// ---------------------------------------------------------------------------
// PDF417 Encoder (text compaction, simplified)
// ---------------------------------------------------------------------------

const PDF417_EC_COUNTS = [2, 4, 8, 16, 32, 64, 128, 256, 512];
const PDF417_START = [8, 1, 1, 1, 1, 1, 1, 3];
const PDF417_STOP_PAT = [7, 1, 1, 1, 1, 1, 1, 3];

const PDF417_TEXT_UP = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ';

function gf929Pow(base: number, exp: number): number {
  let r = 1;
  let b = base % 929;
  let e = exp;
  while (e > 0) {
    if (e & 1) r = (r * b) % 929;
    b = (b * b) % 929;
    e >>= 1;
  }
  return r;
}

function pdf417TextCompact(data: string): number[] {
  const cws: number[] = [900]; // text compaction latch
  const vals: number[] = [];
  for (const ch of data) {
    const ui = PDF417_TEXT_UP.indexOf(ch);
    if (ui >= 0) {
      vals.push(ui);
      continue;
    }
    const li = PDF417_TEXT_UP.indexOf(ch.toUpperCase());
    if (li >= 0) {
      vals.push(27);
      vals.push(li);
      continue;
    }
    const code = ch.charCodeAt(0);
    if (code >= 48 && code <= 57) {
      vals.push(28);
      vals.push(code - 48);
      continue;
    }
    vals.push(29);
    vals.push(Math.max(0, Math.min(29, code - 32)));
  }
  for (let i = 0; i < vals.length; i += 2) {
    const a = at(vals, i);
    const b = i + 1 < vals.length ? at(vals, i + 1) : 29;
    cws.push(a * 30 + b);
  }
  return cws;
}

function pdf417EC(dataCws: number[], ecLevel: number): number[] {
  const numEc = at(PDF417_EC_COUNTS, ecLevel);
  const MOD = 929;
  let g = [1];
  for (let i = 0; i < numEc; i++) {
    const newG = new Array<number>(g.length + 1).fill(0);
    const factor = gf929Pow(3, i);
    for (let j = 0; j < g.length; j++) {
      newG[j] = ((newG[j] ?? 0) + at(g, j)) % MOD;
      newG[j + 1] = ((newG[j + 1] ?? 0) + at(g, j) * factor) % MOD;
    }
    g = newG;
  }
  const ec = new Array<number>(numEc).fill(0);
  for (const cw of dataCws) {
    const t = (cw + at(ec, numEc - 1)) % MOD;
    for (let j = numEc - 1; j > 0; j--) {
      ec[j] = (at(ec, j - 1) + MOD - ((t * at(g, j)) % MOD)) % MOD;
    }
    ec[0] = (MOD - ((t * at(g, 0)) % MOD)) % MOD;
  }
  return ec.map((v) => (MOD - v) % MOD).reverse();
}

function widths2mod(widths: number[]): boolean[] {
  const m: boolean[] = [];
  let black = true;
  for (const w of widths) {
    for (let i = 0; i < w; i++) m.push(black);
    black = !black;
  }
  return m;
}

function pdf417CwPat(cw: number, cluster: number): boolean[] {
  const base = (((cw * 3 + cluster) % 929) + 929) % 929;
  const w = [1, 1, 1, 1, 1, 1, 1, 1];
  const remaining = 9;
  let seed = base;
  for (let i = 0; i < remaining; i++) {
    const idx = seed % 8;
    seed = (seed * 7 + 13) % 929;
    if (at(w, idx) < 6) {
      w[idx] = at(w, idx) + 1;
    } else {
      for (let j = 0; j < 8; j++) {
        const k = (idx + j) % 8;
        if (at(w, k) < 6) {
          w[k] = at(w, k) + 1;
          break;
        }
      }
    }
  }
  const m: boolean[] = [];
  for (let i = 0; i < 8; i++) {
    const isBar = i % 2 === 0;
    for (let n = 0; n < at(w, i); n++) m.push(isBar);
  }
  return m;
}

export function encodePDF417(
  data: string,
  options?: { ecLevel?: number; columns?: number },
): BarcodeMatrix {
  const ecLevel = Math.min(8, Math.max(0, options?.ecLevel ?? 2));
  const columns = Math.min(30, Math.max(1, options?.columns ?? 4));
  const dataCws = pdf417TextCompact(data);
  const ecCount = at(PDF417_EC_COUNTS, ecLevel);
  const allData = [dataCws.length + 1, ...dataCws];
  const minRows = Math.max(3, Math.ceil((allData.length + ecCount) / columns));
  const rows = Math.min(90, minRows);
  const totalSlots = rows * columns;
  while (allData.length < totalSlots - ecCount) allData.push(900);
  allData[0] = allData.length + ecCount;
  const ecCws = pdf417EC(allData, ecLevel);
  const fullCws = [...allData, ...ecCws];

  const startMods = widths2mod(PDF417_START);
  const stopMods = widths2mod(PDF417_STOP_PAT);
  const moduleRows: boolean[][] = [];

  for (let r = 0; r < rows; r++) {
    const cluster = r % 3;
    const rowMods: boolean[] = [...startMods];
    let li: number;
    if (cluster === 0) li = ((r / 3) | 0) * 30 + (((rows - 1) / 3) | 0);
    else if (cluster === 1) li = ((r / 3) | 0) * 30 + ecLevel * 3 + ((rows - 1) % 3);
    else li = ((r / 3) | 0) * 30 + (columns - 1);
    rowMods.push(...pdf417CwPat(li % 929, cluster));
    for (let c = 0; c < columns; c++) {
      const cwI = r * columns + c;
      const cw = cwI < fullCws.length ? at(fullCws, cwI) : 900;
      rowMods.push(...pdf417CwPat(cw, cluster));
    }
    let ri: number;
    if (cluster === 0) ri = ((r / 3) | 0) * 30 + (columns - 1);
    else if (cluster === 1) ri = ((r / 3) | 0) * 30 + (((rows - 1) / 3) | 0);
    else ri = ((r / 3) | 0) * 30 + ecLevel * 3 + ((rows - 1) % 3);
    rowMods.push(...pdf417CwPat(ri % 929, cluster));
    rowMods.push(...stopMods);
    moduleRows.push(rowMods);
  }

  const rowH = 3;
  const w = moduleRows[0]?.length ?? 0;
  const h = rows * rowH;
  const modules = createGrid(h, w, false);
  for (let r = 0; r < rows; r++) {
    const mr = at(moduleRows, r);
    for (let hh = 0; hh < rowH; hh++) {
      for (let c = 0; c < w; c++) setCell(modules, r * rowH + hh, c, at(mr, c));
    }
  }
  return { width: w, height: h, modules };
}
