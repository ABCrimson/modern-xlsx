/**
 * QR Code Encoder (Model 2, ISO 18004, versions 1-10).
 *
 * @module barcode/qr
 */

import type { BarcodeMatrix } from './common.js';
import { at, createGrid, getCell, rsEncode, setCell } from './common.js';

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
