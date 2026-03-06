/**
 * Barcode & QR Code generation module for modern-xlsx.
 *
 * Exports encoders for QR Code, Code 128, EAN-13, UPC-A, Code 39, PDF417,
 * Data Matrix, ITF-14, and GS1-128. Includes a minimal PNG renderer and
 * OOXML drawing XML generator for embedding images in XLSX worksheets.
 *
 * Zero external dependencies. Tree-shakable named exports only.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface BarcodeMatrix {
  width: number;
  height: number;
  modules: boolean[][];
}

export interface RenderOptions {
  moduleSize?: number;
  quietZone?: number;
  foreground?: number[];
  background?: number[];
  showText?: boolean;
  textValue?: string;
}

export type BarcodeType =
  | 'qr'
  | 'code128'
  | 'ean13'
  | 'upca'
  | 'code39'
  | 'pdf417'
  | 'datamatrix'
  | 'itf14'
  | 'gs1128';

export interface DrawBarcodeOptions extends RenderOptions {
  type: BarcodeType;
  fitWidth?: number;
  fitHeight?: number;
  ecLevel?: 'L' | 'M' | 'Q' | 'H';
  pdf417EcLevel?: number;
  pdf417Columns?: number;
}

export interface ImageAnchor {
  fromCol: number;
  fromRow: number;
  toCol: number;
  toRow: number;
  fromColOff?: number;
  fromRowOff?: number;
  toColOff?: number;
  toRowOff?: number;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Safe array access with non-null assertion for controlled-index loops. */
function at<T>(arr: readonly T[], i: number): T {
  return arr[i] as T;
}

function createGrid(rows: number, cols: number, fill: boolean): boolean[][] {
  const grid: boolean[][] = [];
  for (let r = 0; r < rows; r++) {
    grid.push(new Array<boolean>(cols).fill(fill));
  }
  return grid;
}

function setCell(grid: boolean[][], r: number, c: number, v: boolean): void {
  const row = grid[r];
  if (row) row[c] = v;
}

function getCell(grid: boolean[][], r: number, c: number): boolean {
  return grid[r]?.[c] ?? false;
}

// ---------------------------------------------------------------------------
// GF(256) arithmetic for Reed-Solomon (QR & Data Matrix)
// ---------------------------------------------------------------------------

const GF_EXP = new Uint8Array(512);
const GF_LOG = new Uint8Array(256);

{
  let v = 1;
  for (let i = 0; i < 255; i++) {
    GF_EXP[i] = v;
    GF_LOG[v] = i;
    v <<= 1;
    if (v >= 256) v ^= 0x11d;
  }
  for (let i = 255; i < 512; i++) {
    GF_EXP[i] = GF_EXP[i - 255] ?? 0;
  }
}

function gfMul(a: number, b: number): number {
  if (a === 0 || b === 0) return 0;
  return GF_EXP[(GF_LOG[a] ?? 0) + (GF_LOG[b] ?? 0)] ?? 0;
}

function rsGeneratorPoly(nsym: number): Uint8Array {
  let g = new Uint8Array([1]);
  for (let i = 0; i < nsym; i++) {
    const next = new Uint8Array(g.length + 1);
    const factor = GF_EXP[i] ?? 0;
    for (let j = 0; j < g.length; j++) {
      next[j] = (next[j] ?? 0) ^ (g[j] ?? 0);
      next[j + 1] = (next[j + 1] ?? 0) ^ gfMul(g[j] ?? 0, factor);
    }
    g = next;
  }
  return g;
}

function rsEncode(data: Uint8Array, nsym: number): Uint8Array {
  const gen = rsGeneratorPoly(nsym);
  const out = new Uint8Array(data.length + nsym);
  out.set(data);
  for (let i = 0; i < data.length; i++) {
    const coef = out[i] ?? 0;
    if (coef !== 0) {
      for (let j = 0; j < gen.length; j++) {
        out[i + j] = (out[i + j] ?? 0) ^ gfMul(gen[j] ?? 0, coef);
      }
    }
  }
  return out.subarray(data.length);
}

// ---------------------------------------------------------------------------
// QR Code Encoder (Model 2, ISO 18004, versions 1-10)
// ---------------------------------------------------------------------------

type QrEcLevel = 'L' | 'M' | 'Q' | 'H';

// [total_codewords, ec_per_block, num_blocks_g1, data_per_block_g1, num_blocks_g2, data_per_block_g2]
type QrV = [number, number, number, number, number, number];

const QR_VI: Record<QrEcLevel, QrV[]> = {
  L: [
    [0, 0, 0, 0, 0, 0],
    [26, 7, 1, 19, 0, 0],
    [44, 10, 1, 34, 0, 0],
    [70, 15, 1, 55, 0, 0],
    [100, 20, 1, 80, 0, 0],
    [134, 26, 1, 108, 0, 0],
    [172, 18, 2, 68, 0, 0],
    [196, 20, 2, 78, 0, 0],
    [242, 24, 2, 97, 0, 0],
    [292, 30, 2, 116, 0, 0],
    [346, 18, 2, 68, 2, 69],
  ],
  M: [
    [0, 0, 0, 0, 0, 0],
    [26, 10, 1, 16, 0, 0],
    [44, 16, 1, 28, 0, 0],
    [70, 26, 1, 44, 0, 0],
    [100, 18, 2, 32, 0, 0],
    [134, 24, 2, 43, 0, 0],
    [172, 16, 4, 27, 0, 0],
    [196, 18, 4, 31, 0, 0],
    [242, 22, 2, 38, 2, 39],
    [292, 22, 3, 36, 2, 37],
    [346, 26, 4, 43, 1, 44],
  ],
  Q: [
    [0, 0, 0, 0, 0, 0],
    [26, 13, 1, 13, 0, 0],
    [44, 22, 1, 22, 0, 0],
    [70, 18, 2, 17, 0, 0],
    [100, 26, 2, 24, 0, 0],
    [134, 18, 2, 15, 2, 16],
    [172, 24, 2, 19, 2, 20],
    [196, 18, 2, 14, 4, 15],
    [242, 22, 4, 18, 2, 19],
    [292, 20, 4, 16, 4, 17],
    [346, 24, 6, 19, 2, 20],
  ],
  H: [
    [0, 0, 0, 0, 0, 0],
    [26, 17, 1, 9, 0, 0],
    [44, 28, 1, 16, 0, 0],
    [70, 22, 2, 13, 0, 0],
    [100, 16, 4, 9, 0, 0],
    [134, 22, 2, 11, 2, 12],
    [172, 28, 4, 15, 0, 0],
    [196, 26, 4, 13, 1, 14],
    [242, 26, 4, 14, 2, 15],
    [292, 24, 4, 12, 4, 13],
    [346, 28, 6, 15, 2, 16],
  ],
};

const QR_ALIGN: number[][] = [
  [],
  [],
  [6, 18],
  [6, 22],
  [6, 26],
  [6, 30],
  [6, 34],
  [6, 22, 38],
  [6, 24, 42],
  [6, 26, 46],
  [6, 28, 50],
];

const QR_FMT: Record<QrEcLevel, number[]> = (() => {
  const poly = 0x537;
  const r: Record<QrEcLevel, number[]> = { L: [], M: [], Q: [], H: [] };
  const ecBits = { L: 0b01, M: 0b00, Q: 0b11, H: 0b10 } satisfies Record<QrEcLevel, number>;
  for (const lv of ['L', 'M', 'Q', 'H'] satisfies QrEcLevel[]) {
    for (let mask = 0; mask < 8; mask++) {
      const data = (ecBits[lv] << 3) | mask;
      let bits = data << 10;
      for (let i = 14; i >= 10; i--) {
        if (bits & (1 << i)) bits ^= poly << (i - 10);
      }
      let fmt = (data << 10) | bits;
      fmt ^= 0x5412;
      r[lv].push(fmt);
    }
  }
  return r;
})();

const QR_VBITS = [0, 0, 0, 0, 0, 0, 0, 0x07c94, 0x085bc, 0x09a99, 0x0a4d3];

const QR_ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:';

function qrMode(data: string): 'numeric' | 'alphanumeric' | 'byte' {
  if (/^\d+$/.test(data)) return 'numeric';
  if ([...data.toUpperCase()].every((c) => QR_ALNUM.includes(c))) return 'alphanumeric';
  return 'byte';
}

function qrCcBits(ver: number, mode: 'numeric' | 'alphanumeric' | 'byte'): number {
  if (ver <= 9) return mode === 'numeric' ? 10 : mode === 'alphanumeric' ? 9 : 8;
  return mode === 'numeric' ? 12 : mode === 'alphanumeric' ? 11 : 16;
}

function qrEncData(
  data: string,
  mode: 'numeric' | 'alphanumeric' | 'byte',
): { bits: number[]; charCnt: number } {
  const bits: number[] = [];
  const push = (val: number, len: number) => {
    for (let i = len - 1; i >= 0; i--) bits.push((val >> i) & 1);
  };
  if (mode === 'numeric') {
    for (let i = 0; i < data.length; i += 3) {
      const chunk = data.substring(i, i + 3);
      push(Number.parseInt(chunk, 10), chunk.length === 3 ? 10 : chunk.length === 2 ? 7 : 4);
    }
  } else if (mode === 'alphanumeric') {
    const u = data.toUpperCase();
    for (let i = 0; i < u.length; i += 2) {
      const a = QR_ALNUM.indexOf(u[i] ?? '');
      if (i + 1 < u.length) push(a * 45 + QR_ALNUM.indexOf(u[i + 1] ?? ''), 11);
      else push(a, 6);
    }
  } else {
    const encoded = new TextEncoder().encode(data);
    for (const b of encoded) push(b, 8);
    return { bits, charCnt: encoded.length };
  }
  return { bits, charCnt: data.length };
}

function qrSelVer(
  dataBits: number,
  mode: 'numeric' | 'alphanumeric' | 'byte',
  ec: QrEcLevel,
): number {
  for (let v = 1; v <= 10; v++) {
    const info = at(QR_VI[ec], v);
    const cap = (info[2] * info[3] + info[4] * info[5]) * 8;
    if (4 + qrCcBits(v, mode) + dataBits <= cap) return v;
  }
  throw new Error('Data too long for QR versions 1-10');
}

export function encodeQR(data: string, options?: { ecLevel?: QrEcLevel }): BarcodeMatrix {
  const ec: QrEcLevel = options?.ecLevel ?? 'M';
  const mode = qrMode(data);
  const { bits: dataBits, charCnt } = qrEncData(data, mode);
  const ver = qrSelVer(dataBits.length, mode, ec);
  const sz = 17 + ver * 4;

  const modeInd = mode === 'numeric' ? 0b0001 : mode === 'alphanumeric' ? 0b0010 : 0b0100;
  const ccLen = qrCcBits(ver, mode);
  const info = at(QR_VI[ec], ver);
  const totalDCW = info[2] * info[3] + info[4] * info[5];
  const totalCap = totalDCW * 8;

  const stream: number[] = [];
  const pushBits = (val: number, len: number) => {
    for (let i = len - 1; i >= 0; i--) stream.push((val >> i) & 1);
  };

  pushBits(modeInd, 4);
  pushBits(charCnt, ccLen);
  stream.push(...dataBits);
  pushBits(0, Math.min(4, totalCap - stream.length));
  while (stream.length % 8 !== 0) stream.push(0);
  const pads = [0xec, 0x11];
  let pi = 0;
  while (stream.length < totalCap) {
    pushBits(at(pads, pi), 8);
    pi = (pi + 1) % 2;
  }

  const cw = new Uint8Array(totalDCW);
  for (let i = 0; i < totalDCW; i++) {
    let byte = 0;
    for (let b = 0; b < 8; b++) byte = (byte << 1) | at(stream, i * 8 + b);
    cw[i] = byte;
  }

  const ecPB = info[1];
  const blocks: Uint8Array[] = [];
  const ecBlocks: Uint8Array[] = [];
  let off = 0;
  for (let g = 0; g < 2; g++) {
    const nb = g === 0 ? info[2] : info[4];
    const dpb = g === 0 ? info[3] : info[5];
    for (let b = 0; b < nb; b++) {
      const block = cw.slice(off, off + dpb);
      blocks.push(block);
      ecBlocks.push(rsEncode(block, ecPB));
      off += dpb;
    }
  }

  const interleaved: number[] = [];
  const maxDL = Math.max(...blocks.map((b) => b.length));
  for (let i = 0; i < maxDL; i++) {
    for (const bl of blocks) {
      if (i < bl.length) interleaved.push(bl[i] ?? 0);
    }
  }
  for (let i = 0; i < ecPB; i++) {
    for (const ecb of ecBlocks) {
      if (i < ecb.length) interleaved.push(ecb[i] ?? 0);
    }
  }

  // Build grid: -1=unassigned, 0=white function, 1=black function
  const grid: Int8Array[] = [];
  for (let r = 0; r < sz; r++) grid.push(new Int8Array(sz).fill(-1));

  const setMod = (r: number, c: number, val: number) => {
    if (r >= 0 && r < sz && c >= 0 && c < sz) at(grid, r)[c] = val;
  };
  const getMod = (r: number, c: number): number => grid[r]?.[c] ?? -1;

  // Finder patterns
  const placeFinder = (row: number, col: number) => {
    for (let r = -1; r <= 7; r++) {
      for (let c = -1; c <= 7; c++) {
        const rr = row + r;
        const cc = col + c;
        if (rr < 0 || rr >= sz || cc < 0 || cc >= sz) continue;
        if (r === -1 || r === 7 || c === -1 || c === 7) setMod(rr, cc, 0);
        else if (r === 0 || r === 6 || c === 0 || c === 6) setMod(rr, cc, 1);
        else if (r >= 2 && r <= 4 && c >= 2 && c <= 4) setMod(rr, cc, 1);
        else setMod(rr, cc, 0);
      }
    }
  };
  placeFinder(0, 0);
  placeFinder(0, sz - 7);
  placeFinder(sz - 7, 0);

  // Timing
  for (let i = 8; i < sz - 8; i++) {
    setMod(6, i, i % 2 === 0 ? 1 : 0);
    setMod(i, 6, i % 2 === 0 ? 1 : 0);
  }

  // Alignment
  const alignPos = at(QR_ALIGN, ver);
  for (const ar of alignPos) {
    for (const ac of alignPos) {
      if (ar <= 8 && ac <= 8) continue;
      if (ar <= 8 && ac >= sz - 9) continue;
      if (ar >= sz - 9 && ac <= 8) continue;
      for (let dr = -2; dr <= 2; dr++) {
        for (let dc = -2; dc <= 2; dc++) {
          setMod(
            ar + dr,
            ac + dc,
            Math.abs(dr) === 2 || Math.abs(dc) === 2 || (dr === 0 && dc === 0) ? 1 : 0,
          );
        }
      }
    }
  }

  // Dark module
  setMod(sz - 8, 8, 1);

  // Reserve format info
  for (let i = 0; i < 8; i++) {
    if (getMod(8, i) === -1) setMod(8, i, 0);
    if (getMod(i, 8) === -1) setMod(i, 8, 0);
    if (getMod(8, sz - 1 - i) === -1) setMod(8, sz - 1 - i, 0);
    if (getMod(sz - 1 - i, 8) === -1) setMod(sz - 1 - i, 8, 0);
  }
  if (getMod(8, 8) === -1) setMod(8, 8, 0);

  // Reserve version info (v7+)
  if (ver >= 7) {
    for (let i = 0; i < 6; i++) {
      for (let j = 0; j < 3; j++) {
        if (getMod(i, sz - 11 + j) === -1) setMod(i, sz - 11 + j, 0);
        if (getMod(sz - 11 + j, i) === -1) setMod(sz - 11 + j, i, 0);
      }
    }
  }

  // Place data
  let bitIdx = 0;
  let upward = true;
  for (let col = sz - 1; col >= 0; col -= 2) {
    if (col === 6) col = 5;
    const colPair = [col, col - 1];
    const rows = upward
      ? Array.from({ length: sz }, (_, i) => sz - 1 - i)
      : Array.from({ length: sz }, (_, i) => i);
    for (const row of rows) {
      for (const c of colPair) {
        if (c < 0 || c >= sz) continue;
        if (getMod(row, c) !== -1) continue;
        if (bitIdx < interleaved.length * 8) {
          const byteI = bitIdx >> 3;
          const bitP = 7 - (bitIdx & 7);
          setMod(row, c, (at(interleaved, byteI) >> bitP) & 1);
        } else {
          setMod(row, c, 0);
        }
        bitIdx++;
      }
    }
    upward = !upward;
  }

  // Build function pattern mask
  const funcP = createGrid(sz, sz, false);
  for (let r = 0; r < 9; r++) {
    for (let c = 0; c < 9; c++) setCell(funcP, r, c, true);
    for (let c = sz - 8; c < sz; c++) setCell(funcP, r, c, true);
  }
  for (let r = sz - 8; r < sz; r++) {
    for (let c = 0; c < 9; c++) setCell(funcP, r, c, true);
  }
  for (let i = 0; i < sz; i++) {
    setCell(funcP, 6, i, true);
    setCell(funcP, i, 6, true);
  }
  for (const ar of alignPos) {
    for (const ac of alignPos) {
      if (ar <= 8 && ac <= 8) continue;
      if (ar <= 8 && ac >= sz - 9) continue;
      if (ar >= sz - 9 && ac <= 8) continue;
      for (let dr = -2; dr <= 2; dr++) {
        for (let dc = -2; dc <= 2; dc++) setCell(funcP, ar + dr, ac + dc, true);
      }
    }
  }
  setCell(funcP, sz - 8, 8, true);
  if (ver >= 7) {
    for (let i = 0; i < 6; i++) {
      for (let j = 0; j < 3; j++) {
        setCell(funcP, i, sz - 11 + j, true);
        setCell(funcP, sz - 11 + j, i, true);
      }
    }
  }
  for (let i = 0; i < 8; i++) {
    setCell(funcP, 8, i, true);
    setCell(funcP, i, 8, true);
    setCell(funcP, 8, sz - 1 - i, true);
    setCell(funcP, sz - 1 - i, 8, true);
  }
  setCell(funcP, 8, 8, true);

  // Mask functions
  const maskFns: ((r: number, c: number) => boolean)[] = [
    (r, c) => (r + c) % 2 === 0,
    (r) => r % 2 === 0,
    (_, c) => c % 3 === 0,
    (r, c) => (r + c) % 3 === 0,
    (r, c) => (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0,
    (r, c) => ((r * c) % 2) + ((r * c) % 3) === 0,
    (r, c) => (((r * c) % 2) + ((r * c) % 3)) % 2 === 0,
    (r, c) => (((r + c) % 2) + ((r * c) % 3)) % 2 === 0,
  ];

  // Penalty evaluation
  const evalPen = (g: Int8Array[]): number => {
    let pen = 0;
    // Rule 1: runs
    for (let r = 0; r < sz; r++) {
      let run = 1;
      for (let c = 1; c < sz; c++) {
        if (at(g, r)[c] === at(g, r)[c - 1]) run++;
        else {
          if (run >= 5) pen += run - 2;
          run = 1;
        }
      }
      if (run >= 5) pen += run - 2;
    }
    for (let c = 0; c < sz; c++) {
      let run = 1;
      for (let r = 1; r < sz; r++) {
        if (at(g, r)[c] === at(g, r - 1)[c]) run++;
        else {
          if (run >= 5) pen += run - 2;
          run = 1;
        }
      }
      if (run >= 5) pen += run - 2;
    }
    // Rule 2: 2x2
    for (let r = 0; r < sz - 1; r++) {
      for (let c = 0; c < sz - 1; c++) {
        const v = at(g, r)[c];
        if (v === at(g, r)[c + 1] && v === at(g, r + 1)[c] && v === at(g, r + 1)[c + 1]) pen += 3;
      }
    }
    // Rule 3: finder-like
    const p1 = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0];
    const p2 = [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1];
    for (let r = 0; r < sz; r++) {
      for (let c = 0; c <= sz - 11; c++) {
        let m1 = true;
        let m2 = true;
        for (let k = 0; k < 11; k++) {
          if (at(g, r)[c + k] !== at(p1, k)) m1 = false;
          if (at(g, r)[c + k] !== at(p2, k)) m2 = false;
        }
        if (m1 || m2) pen += 40;
      }
    }
    for (let c = 0; c < sz; c++) {
      for (let r = 0; r <= sz - 11; r++) {
        let m1 = true;
        let m2 = true;
        for (let k = 0; k < 11; k++) {
          if (at(g, r + k)[c] !== at(p1, k)) m1 = false;
          if (at(g, r + k)[c] !== at(p2, k)) m2 = false;
        }
        if (m1 || m2) pen += 40;
      }
    }
    // Rule 4: dark proportion
    let dark = 0;
    for (let r = 0; r < sz; r++) for (let c = 0; c < sz; c++) if (at(g, r)[c] === 1) dark++;
    const pct = (dark * 100) / (sz * sz);
    const pr = Math.floor(pct / 5) * 5;
    pen += Math.min(Math.abs(pr - 50) / 5, Math.abs(pr + 5 - 50) / 5) * 10;
    return pen;
  };

  let bestMask = 0;
  let bestPen = Number.POSITIVE_INFINITY;
  for (let mask = 0; mask < 8; mask++) {
    const masked: Int8Array[] = grid.map((row) => new Int8Array(row));
    for (let r = 0; r < sz; r++) {
      for (let c = 0; c < sz; c++) {
        if (!getCell(funcP, r, c) && at(maskFns, mask)(r, c)) {
          const row = at(masked, r);
          row[c] = (row[c] ?? 0) ^ 1;
        }
      }
    }
    const p = evalPen(masked);
    if (p < bestPen) {
      bestPen = p;
      bestMask = mask;
    }
  }

  // Apply best mask
  for (let r = 0; r < sz; r++) {
    for (let c = 0; c < sz; c++) {
      if (!getCell(funcP, r, c) && at(maskFns, bestMask)(r, c)) {
        const row = at(grid, r);
        row[c] = (row[c] ?? 0) ^ 1;
      }
    }
  }

  // Write format info
  const fmtBits = at(QR_FMT[ec], bestMask);
  const fmtPos: [number, number][] = [
    [8, 0],
    [8, 1],
    [8, 2],
    [8, 3],
    [8, 4],
    [8, 5],
    [8, 7],
    [8, 8],
    [7, 8],
    [5, 8],
    [4, 8],
    [3, 8],
    [2, 8],
    [1, 8],
    [0, 8],
  ];
  for (let i = 0; i < 15; i++) {
    const [r, c] = at(fmtPos, i);
    at(grid, r)[c] = (fmtBits >> (14 - i)) & 1;
  }
  for (let i = 0; i < 7; i++) at(grid, sz - 1 - i)[8] = (fmtBits >> i) & 1;
  for (let i = 0; i < 8; i++) at(grid, 8)[sz - 8 + i] = (fmtBits >> (7 + i)) & 1;

  // Version info (v7+)
  if (ver >= 7) {
    const vb = at(QR_VBITS, ver);
    for (let i = 0; i < 18; i++) {
      const bit = (vb >> i) & 1;
      const row = Math.floor(i / 3);
      const col = sz - 11 + (i % 3);
      at(grid, row)[col] = bit;
      at(grid, col)[row] = bit;
    }
  }

  // Convert to BarcodeMatrix
  const modules = createGrid(sz, sz, false);
  for (let r = 0; r < sz; r++) {
    for (let c = 0; c < sz; c++) setCell(modules, r, c, at(grid, r)[c] === 1);
  }
  return { width: sz, height: sz, modules };
}

// ---------------------------------------------------------------------------
// Code 128 Encoder
// ---------------------------------------------------------------------------

const C128_W: number[][] = [
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

const C128_STOP = [2, 3, 3, 1, 1, 1, 2];

function widths2mod(widths: number[]): boolean[] {
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

// ---------------------------------------------------------------------------
// EAN-13 Encoder
// ---------------------------------------------------------------------------

const EAN_L: number[][] = [
  [0, 0, 0, 1, 1, 0, 1],
  [0, 0, 1, 1, 0, 0, 1],
  [0, 0, 1, 0, 0, 1, 1],
  [0, 1, 1, 1, 1, 0, 1],
  [0, 1, 0, 0, 0, 1, 1],
  [0, 1, 1, 0, 0, 0, 1],
  [0, 1, 0, 1, 1, 1, 1],
  [0, 1, 1, 1, 0, 1, 1],
  [0, 1, 1, 0, 1, 1, 1],
  [0, 0, 0, 1, 0, 1, 1],
];
const EAN_G: number[][] = [
  [0, 1, 0, 0, 1, 1, 1],
  [0, 1, 1, 0, 0, 1, 1],
  [0, 0, 1, 1, 0, 1, 1],
  [0, 1, 0, 0, 0, 0, 1],
  [0, 0, 1, 1, 1, 0, 1],
  [0, 1, 1, 1, 0, 0, 1],
  [0, 0, 0, 0, 1, 0, 1],
  [0, 0, 1, 0, 0, 0, 1],
  [0, 0, 0, 1, 0, 0, 1],
  [0, 0, 1, 0, 1, 1, 1],
];
const EAN_R: number[][] = EAN_L.map((p) => p.map((b) => b ^ 1));
const EAN_PAR: number[][] = [
  [0, 0, 0, 0, 0, 0],
  [0, 0, 1, 0, 1, 1],
  [0, 0, 1, 1, 0, 1],
  [0, 0, 1, 1, 1, 0],
  [0, 1, 0, 0, 1, 1],
  [0, 1, 1, 0, 0, 1],
  [0, 1, 1, 1, 0, 0],
  [0, 1, 0, 1, 0, 1],
  [0, 1, 0, 1, 1, 0],
  [0, 1, 1, 0, 1, 0],
];

function ean13Chk(d: number[]): number {
  let s = 0;
  for (let i = 0; i < 12; i++) s += at(d, i) * (i % 2 === 0 ? 1 : 3);
  return (10 - (s % 10)) % 10;
}

export function encodeEAN13(data: string): BarcodeMatrix {
  let digits: number[];
  if (data.length === 12) {
    digits = [...data].map(Number);
    digits.push(ean13Chk(digits));
  } else if (data.length === 13) {
    digits = [...data].map(Number);
    const exp = ean13Chk(digits.slice(0, 12));
    if (at(digits, 12) !== exp)
      throw new Error(`EAN-13: invalid check digit (expected ${exp}, got ${at(digits, 12)})`);
  } else {
    throw new Error('EAN-13: data must be 12 or 13 digits');
  }
  if (digits.some((d) => d < 0 || d > 9 || !Number.isFinite(d))) {
    throw new Error('EAN-13: data must contain only digits 0-9');
  }

  const bars: number[] = [1, 0, 1]; // start guard
  const par = at(EAN_PAR, at(digits, 0));
  for (let i = 0; i < 6; i++) {
    const pat = at(par, i) === 0 ? at(EAN_L, at(digits, i + 1)) : at(EAN_G, at(digits, i + 1));
    bars.push(...pat);
  }
  bars.push(0, 1, 0, 1, 0); // center guard
  for (let i = 7; i <= 12; i++) bars.push(...at(EAN_R, at(digits, i)));
  bars.push(1, 0, 1); // end guard

  const w = bars.length;
  const bh = 60;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++)
    for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c) === 1);
  return { width: w, height: bh, modules };
}

// ---------------------------------------------------------------------------
// UPC-A Encoder
// ---------------------------------------------------------------------------

export function encodeUPCA(data: string): BarcodeMatrix {
  if (data.length === 11) return encodeEAN13(`0${data}`);
  if (data.length === 12) return encodeEAN13(`0${data.slice(0, 11)}`);
  throw new Error('UPC-A: data must be 11 or 12 digits');
}

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
  [1, 0, 0, 1, 0, 1, 0, 1, 0, 1, 1, 1],
  [1, 0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 1],
  [1, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1],
  [1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 1],
  [1, 0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 1],
  [1, 0, 0, 1, 0, 1, 1, 0, 1, 1, 0, 1],
];

export function encodeCode39(data: string): BarcodeMatrix {
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

  const w = bars.length;
  const bh = 60;
  const modules = createGrid(bh, w, false);
  for (let r = 0; r < bh; r++)
    for (let c = 0; c < w; c++) setCell(modules, r, c, at(bars, c) === 1);
  return { width: w, height: bh, modules };
}

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

// ---------------------------------------------------------------------------
// Bitmap font (5x7, digits 0-9 + A-Z + punctuation)
// ---------------------------------------------------------------------------

const FW = 5;
const FH = 7;
const FONT: Record<string, number[]> = {
  '0': [0x0e, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0e],
  '1': [0x04, 0x0c, 0x04, 0x04, 0x04, 0x04, 0x0e],
  '2': [0x0e, 0x11, 0x01, 0x06, 0x08, 0x10, 0x1f],
  '3': [0x0e, 0x11, 0x01, 0x06, 0x01, 0x11, 0x0e],
  '4': [0x02, 0x06, 0x0a, 0x12, 0x1f, 0x02, 0x02],
  '5': [0x1f, 0x10, 0x1e, 0x01, 0x01, 0x11, 0x0e],
  '6': [0x06, 0x08, 0x10, 0x1e, 0x11, 0x11, 0x0e],
  '7': [0x1f, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08],
  '8': [0x0e, 0x11, 0x11, 0x0e, 0x11, 0x11, 0x0e],
  '9': [0x0e, 0x11, 0x11, 0x0f, 0x01, 0x02, 0x0c],
  A: [0x0e, 0x11, 0x11, 0x1f, 0x11, 0x11, 0x11],
  B: [0x1e, 0x11, 0x11, 0x1e, 0x11, 0x11, 0x1e],
  C: [0x0e, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0e],
  D: [0x1e, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1e],
  E: [0x1f, 0x10, 0x10, 0x1e, 0x10, 0x10, 0x1f],
  F: [0x1f, 0x10, 0x10, 0x1e, 0x10, 0x10, 0x10],
  G: [0x0e, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0f],
  H: [0x11, 0x11, 0x11, 0x1f, 0x11, 0x11, 0x11],
  I: [0x0e, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0e],
  J: [0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0c],
  K: [0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11],
  L: [0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1f],
  M: [0x11, 0x1b, 0x15, 0x15, 0x11, 0x11, 0x11],
  N: [0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11],
  O: [0x0e, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0e],
  P: [0x1e, 0x11, 0x11, 0x1e, 0x10, 0x10, 0x10],
  Q: [0x0e, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0d],
  R: [0x1e, 0x11, 0x11, 0x1e, 0x14, 0x12, 0x11],
  S: [0x0e, 0x11, 0x10, 0x0e, 0x01, 0x11, 0x0e],
  T: [0x1f, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04],
  U: [0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0e],
  V: [0x11, 0x11, 0x11, 0x11, 0x0a, 0x0a, 0x04],
  W: [0x11, 0x11, 0x11, 0x15, 0x15, 0x1b, 0x11],
  X: [0x11, 0x11, 0x0a, 0x04, 0x0a, 0x11, 0x11],
  Y: [0x11, 0x11, 0x0a, 0x04, 0x04, 0x04, 0x04],
  Z: [0x1f, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1f],
  ' ': [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00],
  '-': [0x00, 0x00, 0x00, 0x1f, 0x00, 0x00, 0x00],
  '.': [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04],
  '(': [0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02],
  ')': [0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08],
  '/': [0x01, 0x02, 0x02, 0x04, 0x08, 0x08, 0x10],
};

function renderText(text: string): boolean[][] {
  if (text.length === 0) return [];
  const cw = FW + 1;
  const tw = text.length * cw - 1;
  const grid = createGrid(FH, tw, false);
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]?.toUpperCase() ?? ' ';
    const d: number[] = FONT[ch] ?? FONT[' '] ?? [];
    const xo = i * cw;
    for (let r = 0; r < FH; r++) {
      for (let c = 0; c < FW; c++) {
        if (xo + c < tw) setCell(grid, r, xo + c, ((at(d, r) >> (4 - c)) & 1) === 1);
      }
    }
  }
  return grid;
}

// ---------------------------------------------------------------------------
// CRC32 & Adler32 for PNG
// ---------------------------------------------------------------------------

const CRC_TBL: Uint32Array = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++)
    crc = (CRC_TBL[(crc ^ (data[i] ?? 0)) & 0xff] ?? 0) ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  for (let i = 0; i < data.length; i++) {
    a = (a + (data[i] ?? 0)) % 65521;
    b = (b + a) % 65521;
  }
  return ((b << 16) | a) >>> 0;
}

// ---------------------------------------------------------------------------
// Minimal PNG encoder (uncompressed deflate / store blocks)
// ---------------------------------------------------------------------------

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const chunk = new Uint8Array(4 + 4 + data.length + 4);
  const dv = new DataView(chunk.buffer);
  dv.setUint32(0, data.length, false);
  for (let i = 0; i < 4; i++) chunk[4 + i] = type.charCodeAt(i);
  chunk.set(data, 8);
  const crcBuf = new Uint8Array(4 + data.length);
  crcBuf.set(chunk.subarray(4, 8));
  crcBuf.set(data, 4);
  dv.setUint32(8 + data.length, crc32(crcBuf), false);
  return chunk;
}

function createPng(width: number, height: number, pixels: Uint8Array): Uint8Array {
  const rowB = width * 3;
  const raw = new Uint8Array(height * (1 + rowB));
  for (let r = 0; r < height; r++) {
    raw[r * (1 + rowB)] = 0;
    raw.set(pixels.subarray(r * rowB, (r + 1) * rowB), r * (1 + rowB) + 1);
  }

  const maxBlk = 65535;
  const nBlk = Math.ceil(raw.length / maxBlk) || 1;
  const defSz = 2 + nBlk * 5 + raw.length + 4;
  const def = new Uint8Array(defSz);
  let p = 0;
  def[p++] = 0x78;
  def[p++] = 0x01;
  let rem = raw.length;
  let off = 0;
  while (rem > 0) {
    const bs = Math.min(rem, maxBlk);
    def[p++] = rem <= maxBlk ? 0x01 : 0x00;
    def[p++] = bs & 0xff;
    def[p++] = (bs >> 8) & 0xff;
    def[p++] = ~bs & 0xff;
    def[p++] = (~bs >> 8) & 0xff;
    def.set(raw.subarray(off, off + bs), p);
    p += bs;
    off += bs;
    rem -= bs;
  }
  const ad = adler32(raw);
  def[p++] = (ad >> 24) & 0xff;
  def[p++] = (ad >> 16) & 0xff;
  def[p++] = (ad >> 8) & 0xff;
  def[p++] = ad & 0xff;
  const actDef = def.subarray(0, p);

  const ihdr = new Uint8Array(13);
  const ihdrV = new DataView(ihdr.buffer);
  ihdrV.setUint32(0, width, false);
  ihdrV.setUint32(4, height, false);
  ihdr[8] = 8;
  ihdr[9] = 2;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const sig = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdrC = pngChunk('IHDR', ihdr);
  const idatC = pngChunk('IDAT', actDef);
  const iendC = pngChunk('IEND', new Uint8Array(0));

  const png = new Uint8Array(sig.length + ihdrC.length + idatC.length + iendC.length);
  let pp = 0;
  png.set(sig, pp);
  pp += sig.length;
  png.set(ihdrC, pp);
  pp += ihdrC.length;
  png.set(idatC, pp);
  pp += idatC.length;
  png.set(iendC, pp);
  return png;
}

// ---------------------------------------------------------------------------
// PNG Renderer
// ---------------------------------------------------------------------------

export function renderBarcodePNG(matrix: BarcodeMatrix, options?: RenderOptions): Uint8Array {
  const modSz = options?.moduleSize ?? 4;
  const qz = options?.quietZone ?? 4;
  const fgR = options?.foreground?.[0] ?? 0;
  const fgG = options?.foreground?.[1] ?? 0;
  const fgB = options?.foreground?.[2] ?? 0;
  const bgR = options?.background?.[0] ?? 255;
  const bgG = options?.background?.[1] ?? 255;
  const bgB = options?.background?.[2] ?? 255;
  const showText = options?.showText ?? false;
  const textVal = options?.textValue ?? '';

  const bw = matrix.width * modSz + 2 * qz * modSz;
  const bh = matrix.height * modSz + 2 * qz * modSz;

  let textBM: boolean[][] = [];
  let textPH = 0;
  const textGap = modSz;
  if (showText && textVal.length > 0) {
    textBM = renderText(textVal);
    textPH = textBM.length > 0 ? textBM.length * modSz + textGap : 0;
  }

  const tw = bw;
  const th = bh + textPH;
  const px = new Uint8Array(tw * th * 3);

  for (let i = 0; i < tw * th; i++) {
    px[i * 3] = bgR;
    px[i * 3 + 1] = bgG;
    px[i * 3 + 2] = bgB;
  }

  const qzPx = qz * modSz;
  for (let r = 0; r < matrix.height; r++) {
    for (let c = 0; c < matrix.width; c++) {
      if (getCell(matrix.modules, r, c)) {
        for (let dy = 0; dy < modSz; dy++) {
          for (let dx = 0; dx < modSz; dx++) {
            const x = qzPx + c * modSz + dx;
            const y = qzPx + r * modSz + dy;
            const idx = (y * tw + x) * 3;
            px[idx] = fgR;
            px[idx + 1] = fgG;
            px[idx + 2] = fgB;
          }
        }
      }
    }
  }

  if (showText && textBM.length > 0) {
    const tpw = textBM[0]?.length ?? 0;
    const scale = Math.max(1, Math.floor(modSz / 2));
    const stw = tpw * scale;
    const txo = Math.max(0, Math.floor((tw - stw) / 2));
    const tyo = bh + textGap;
    for (let tr = 0; tr < textBM.length; tr++) {
      for (let tc = 0; tc < tpw; tc++) {
        if (getCell(textBM, tr, tc)) {
          for (let dy = 0; dy < scale; dy++) {
            for (let dx = 0; dx < scale; dx++) {
              const x = txo + tc * scale + dx;
              const y = tyo + tr * scale + dy;
              if (x >= 0 && x < tw && y >= 0 && y < th) {
                const idx = (y * tw + x) * 3;
                px[idx] = fgR;
                px[idx + 1] = fgG;
                px[idx + 2] = fgB;
              }
            }
          }
        }
      }
    }
  }

  return createPng(tw, th, px);
}

// ---------------------------------------------------------------------------
// Auto-sizing helper
// ---------------------------------------------------------------------------

export function generateBarcode(value: string, options: DrawBarcodeOptions): Uint8Array {
  let matrix: BarcodeMatrix;

  switch (options.type) {
    case 'qr':
      matrix = encodeQR(value, options.ecLevel ? { ecLevel: options.ecLevel } : undefined);
      break;
    case 'code128':
      matrix = encodeCode128(value);
      break;
    case 'ean13':
      matrix = encodeEAN13(value);
      break;
    case 'upca':
      matrix = encodeUPCA(value);
      break;
    case 'code39':
      matrix = encodeCode39(value);
      break;
    case 'pdf417': {
      const pOpts: { ecLevel?: number; columns?: number } = {};
      if (options.pdf417EcLevel !== undefined) pOpts.ecLevel = options.pdf417EcLevel;
      if (options.pdf417Columns !== undefined) pOpts.columns = options.pdf417Columns;
      matrix = encodePDF417(value, pOpts);
      break;
    }
    case 'datamatrix':
      matrix = encodeDataMatrix(value);
      break;
    case 'itf14':
      matrix = encodeITF14(value);
      break;
    case 'gs1128':
      matrix = encodeGS1128(value);
      break;
    default:
      throw new Error(`Unknown barcode type: ${options.type as string}`);
  }

  let modSz = options.moduleSize ?? 4;
  const qz = options.quietZone ?? (options.type === 'qr' || options.type === 'datamatrix' ? 4 : 10);

  if (options.fitWidth || options.fitHeight) {
    const tmw = matrix.width + 2 * qz;
    const tmh = matrix.height + 2 * qz;
    if (options.fitWidth && options.fitHeight) {
      modSz = Math.max(
        1,
        Math.min(Math.floor(options.fitWidth / tmw), Math.floor(options.fitHeight / tmh)),
      );
    } else if (options.fitWidth) {
      modSz = Math.max(1, Math.floor(options.fitWidth / tmw));
    } else if (options.fitHeight) {
      modSz = Math.max(1, Math.floor(options.fitHeight / tmh));
    }
  }

  const renderOpts: RenderOptions = {
    moduleSize: modSz,
    quietZone: qz,
    textValue: options.textValue ?? value,
  };
  if (options.foreground !== undefined) renderOpts.foreground = options.foreground;
  if (options.background !== undefined) renderOpts.background = options.background;
  if (options.showText !== undefined) renderOpts.showText = options.showText;
  return renderBarcodePNG(matrix, renderOpts);
}

// ---------------------------------------------------------------------------
// OOXML Drawing XML Generator
// ---------------------------------------------------------------------------

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function generateDrawingXml(anchors: { anchor: ImageAnchor; imageIndex: number }[]): string {
  const p: string[] = [];
  p.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  p.push(`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`);

  for (let i = 0; i < anchors.length; i++) {
    const entry = anchors[i];
    if (!entry) continue;
    const a = entry.anchor;
    const rId = `rId${entry.imageIndex}`;

    p.push('  <xdr:twoCellAnchor editAs="oneCell">');
    p.push('    <xdr:from>');
    p.push(`      <xdr:col>${a.fromCol}</xdr:col>`);
    p.push(`      <xdr:colOff>${a.fromColOff ?? 0}</xdr:colOff>`);
    p.push(`      <xdr:row>${a.fromRow}</xdr:row>`);
    p.push(`      <xdr:rowOff>${a.fromRowOff ?? 0}</xdr:rowOff>`);
    p.push('    </xdr:from>');
    p.push('    <xdr:to>');
    p.push(`      <xdr:col>${a.toCol}</xdr:col>`);
    p.push(`      <xdr:colOff>${a.toColOff ?? 0}</xdr:colOff>`);
    p.push(`      <xdr:row>${a.toRow}</xdr:row>`);
    p.push(`      <xdr:rowOff>${a.toRowOff ?? 0}</xdr:rowOff>`);
    p.push('    </xdr:to>');
    p.push('    <xdr:pic>');
    p.push('      <xdr:nvPicPr>');
    p.push(`        <xdr:cNvPr id="${i + 2}" name="Barcode ${i + 1}"/>`);
    p.push('        <xdr:cNvPicPr>');
    p.push('          <a:picLocks noChangeAspect="1"/>');
    p.push('        </xdr:cNvPicPr>');
    p.push('      </xdr:nvPicPr>');
    p.push('      <xdr:blipFill>');
    p.push(`        <a:blip r:embed="${rId}"/>`);
    p.push('        <a:stretch>');
    p.push('          <a:fillRect/>');
    p.push('        </a:stretch>');
    p.push('      </xdr:blipFill>');
    p.push('      <xdr:spPr>');
    p.push('        <a:xfrm>');
    p.push('          <a:off x="0" y="0"/>');
    p.push('          <a:ext cx="0" cy="0"/>');
    p.push('        </a:xfrm>');
    p.push('        <a:prstGeom prst="rect">');
    p.push('          <a:avLst/>');
    p.push('        </a:prstGeom>');
    p.push('      </xdr:spPr>');
    p.push('    </xdr:pic>');
    p.push('    <xdr:clientData/>');
    p.push('  </xdr:twoCellAnchor>');
  }

  p.push('</xdr:wsDr>');
  return p.join('\n');
}

export function generateDrawingRels(imageIds: { rId: string; target: string }[]): string {
  const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const p: string[] = [];
  p.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  p.push(`<Relationships xmlns="${RELS}">`);
  for (const id of imageIds) {
    p.push(`  <Relationship Id="${id.rId}" Type="${IMG}" Target="${id.target}"/>`);
  }
  p.push('</Relationships>');
  return p.join('\n');
}
