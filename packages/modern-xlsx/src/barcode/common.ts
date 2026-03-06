/**
 * Shared types, constants, and helper functions for barcode encoders.
 *
 * @module barcode/common
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
export function at<T>(arr: readonly T[], i: number): T {
  return arr[i] as T;
}

export function createGrid(rows: number, cols: number, fill: boolean): boolean[][] {
  const grid: boolean[][] = [];
  for (let r = 0; r < rows; r++) {
    grid.push(new Array<boolean>(cols).fill(fill));
  }
  return grid;
}

export function setCell(grid: boolean[][], r: number, c: number, v: boolean): void {
  const row = grid[r];
  if (row) row[c] = v;
}

export function getCell(grid: boolean[][], r: number, c: number): boolean {
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

export function gfMul(a: number, b: number): number {
  if (a === 0 || b === 0) return 0;
  return GF_EXP[(GF_LOG[a] ?? 0) + (GF_LOG[b] ?? 0)] ?? 0;
}

export function rsGeneratorPoly(nsym: number): Uint8Array {
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

export function rsEncode(data: Uint8Array, nsym: number): Uint8Array {
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
// Bitmap font (5x7, digits 0-9 + A-Z + punctuation)
// ---------------------------------------------------------------------------

export const FW = 5;
export const FH = 7;
export const FONT: Record<string, number[]> = {
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

export function renderText(text: string): boolean[][] {
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
