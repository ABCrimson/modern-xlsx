/**
 * PDF417 Encoder (text compaction, simplified).
 *
 * @module barcode/pdf417
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, setCell } from './common.js';
import { widths2mod } from './linear.js';

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

/** Validates and clamps the options: `ecLevel` to 0-8 and `columns` to 1-30. */
function resolvePDF417Options(options?: { ecLevel?: number; columns?: number }): {
  ecLevel: number;
  columns: number;
} {
  const rawEc = options?.ecLevel ?? 2;
  const rawCols = options?.columns ?? 4;
  if (!Number.isFinite(rawEc)) throw new RangeError('PDF417: ecLevel must be a finite number');
  if (!Number.isFinite(rawCols)) throw new RangeError('PDF417: columns must be a finite number');
  return {
    ecLevel: Math.min(8, Math.max(0, Math.trunc(rawEc))),
    columns: Math.min(30, Math.max(1, Math.trunc(rawCols))),
  };
}

/**
 * Builds the data codewords (length descriptor first), checks the size limits
 * and pads with 900 to fill the `rows` x `columns` grid left after the
 * error-correction codewords.
 */
function pdf417DataCodewords(
  data: string,
  ecCount: number,
  columns: number,
): { allData: number[]; rows: number } {
  const dataCws = pdf417TextCompact(data);
  const allData = [dataCws.length + 1, ...dataCws];
  if (allData.length + ecCount > 928 || Math.ceil((allData.length + ecCount) / columns) > 90) {
    throw new Error('PDF417: data too long');
  }
  const minRows = Math.max(3, Math.ceil((allData.length + ecCount) / columns));
  const rows = Math.min(90, minRows);
  const totalSlots = rows * columns;
  while (allData.length < totalSlots - ecCount) allData.push(900);
  allData[0] = allData.length + ecCount;
  return { allData, rows };
}

/** The left and right row-indicator codeword values for row `r` (before `% 929`). */
function pdf417RowIndicators(
  r: number,
  rows: number,
  columns: number,
  ecLevel: number,
): { left: number; right: number } {
  const cluster = r % 3;
  const base = ((r / 3) | 0) * 30;
  const rowsPart = ((rows - 1) / 3) | 0;
  const ecPart = ecLevel * 3 + ((rows - 1) % 3);
  const colsPart = columns - 1;
  if (cluster === 0) return { left: base + rowsPart, right: base + colsPart };
  if (cluster === 1) return { left: base + ecPart, right: base + rowsPart };
  return { left: base + colsPart, right: base + ecPart };
}

/** Lays the codewords out as start pattern, indicators, data columns and stop pattern per row. */
function pdf417ModuleRows(
  fullCws: number[],
  rows: number,
  columns: number,
  ecLevel: number,
): boolean[][] {
  const startMods = widths2mod(PDF417_START);
  const stopMods = widths2mod(PDF417_STOP_PAT);
  const moduleRows: boolean[][] = [];
  for (let r = 0; r < rows; r++) {
    const cluster = r % 3;
    const { left, right } = pdf417RowIndicators(r, rows, columns, ecLevel);
    const rowMods: boolean[] = [...startMods];
    rowMods.push(...pdf417CwPat(left % 929, cluster));
    for (let c = 0; c < columns; c++) {
      const cwI = r * columns + c;
      const cw = cwI < fullCws.length ? at(fullCws, cwI) : 900;
      rowMods.push(...pdf417CwPat(cw, cluster));
    }
    rowMods.push(...pdf417CwPat(right % 929, cluster));
    rowMods.push(...stopMods);
    moduleRows.push(rowMods);
  }
  return moduleRows;
}

/** Renders the module rows into a grid, each row repeated three times. */
function renderPDF417(moduleRows: boolean[][], rows: number): BarcodeMatrix {
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

/**
 * Encodes a PDF417 symbol.
 *
 * @experimental Output is not ISO/IEC 15438 conformant and may not scan.
 *
 * @param data - Text to encode.
 * @param options - `ecLevel`: error-correction level 0-8 (default `2`); `columns`: data columns 1-30 (default `4`). Out-of-range numbers are truncated and clamped.
 * @returns The symbol matrix, each module row repeated three times, without quiet zone.
 * @throws {RangeError} If `ecLevel` or `columns` is not a finite number.
 * @throws {Error} `PDF417: data too long` if the codewords exceed 928 or the symbol would need more than 90 rows.
 */
export function encodePDF417(
  data: string,
  options?: { ecLevel?: number; columns?: number },
): BarcodeMatrix {
  const { ecLevel, columns } = resolvePDF417Options(options);
  const ecCount = at(PDF417_EC_COUNTS, ecLevel);
  const { allData, rows } = pdf417DataCodewords(data, ecCount, columns);
  const fullCws = [...allData, ...pdf417EC(allData, ecLevel)];
  return renderPDF417(pdf417ModuleRows(fullCws, rows, columns, ecLevel), rows);
}
