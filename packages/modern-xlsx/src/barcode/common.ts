/**
 * Shared types, constants, and helper functions for barcode encoders.
 *
 * @module barcode/common
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** A barcode as a grid of modules, as returned by the `encode*` functions. */
export interface BarcodeMatrix {
  /** Number of module columns. */
  width: number;
  /** Number of module rows. */
  height: number;
  /** Row-major grid: `modules[row][col]` is `true` for a dark module. */
  modules: boolean[][];
}

/** Options for {@link renderBarcodePNG}. */
export interface RenderOptions {
  /** Pixels per module, a positive integer. Default `4`. */
  moduleSize?: number;
  /** Quiet zone around the symbol, in modules, a non-negative integer. Default `4`. */
  quietZone?: number;
  /** Foreground colour as `[r, g, b]`, each channel 0-255. Missing channels default to `0`, so the default colour is black. */
  foreground?: number[];
  /** Background colour as `[r, g, b]`, each channel 0-255. Missing channels default to `255`, so the default colour is white. */
  background?: number[];
  /** Draw `textValue` as a caption below the symbol. Default `false`. The caption is skipped when `textValue` is empty. */
  showText?: boolean;
  /** Caption text, drawn in a built-in 5x7 pixel font (letters are upper-cased; characters the font lacks render as a space). Default: empty, so no caption. */
  textValue?: string;
}

/** The symbology that {@link generateBarcode} encodes. */
export type BarcodeType =
  | 'qr'
  | 'code128'
  | 'ean13'
  | 'upca'
  | 'code39'
  /** Experimental: output is not ISO/IEC 15438 conformant and may not scan. */
  | 'pdf417'
  /** Experimental: module placement is not ISO/IEC 16022 conformant and symbols may not scan. */
  | 'datamatrix'
  | 'itf14'
  | 'gs1128';

/**
 * Options for {@link generateBarcode}: the {@link RenderOptions} plus the symbology and sizing.
 * `generateBarcode` differs from {@link renderBarcodePNG} in two defaults: `quietZone` is `4` modules
 * for `'qr'` and `'datamatrix'` and `10` for every other type, and `textValue` is the encoded value.
 */
export interface DrawBarcodeOptions extends RenderOptions {
  /** The symbology to encode. */
  type: BarcodeType;
  /** Target width in pixels for the symbol plus quiet zone, a positive finite number. The module size becomes the largest whole number at which they fit, never below 1, so a too-small bound is exceeded; it replaces `moduleSize`. */
  fitWidth?: number;
  /** Target height in pixels for the symbol plus quiet zone; any `showText` caption is drawn below it. Applied like `fitWidth`; when both are set the smaller module size wins. */
  fitHeight?: number;
  /** Error-correction level, used by `'qr'` only (default `'M'`). Other types ignore it. */
  ecLevel?: 'L' | 'M' | 'Q' | 'H';
  /** @experimental PDF417 output is not ISO/IEC 15438 conformant. Error-correction level 0-8 for `'pdf417'` (default `2`; values are truncated and clamped to the range; a non-finite value throws a RangeError). */
  pdf417EcLevel?: number;
  /** @experimental PDF417 output is not ISO/IEC 15438 conformant. Data columns 1-30 for `'pdf417'` (default `4`; values are truncated and clamped to the range; a non-finite value throws a RangeError). */
  pdf417Columns?: number;
}

/**
 * A cell range that an image is anchored to, as an OOXML `xdr:twoCellAnchor`. Columns and rows are
 * 0-based indexes (`0` is column A and row 1).
 */
export interface ImageAnchor {
  /** 0-based column of the top-left corner. */
  fromCol: number;
  /** 0-based row of the top-left corner. */
  fromRow: number;
  /** 0-based column of the bottom-right corner. */
  toCol: number;
  /** 0-based row of the bottom-right corner. */
  toRow: number;
  /** Horizontal offset of the top-left corner inside its cell, in EMUs (English Metric Units, 9525 per pixel). Default `0`. */
  fromColOff?: number;
  /** Vertical offset of the top-left corner inside its cell, in EMUs. Default `0`. */
  fromRowOff?: number;
  /** Horizontal offset of the bottom-right corner inside its cell, in EMUs. Default `0`. */
  toColOff?: number;
  /** Vertical offset of the bottom-right corner inside its cell, in EMUs. Default `0`. */
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
// GF(256) arithmetic for Reed-Solomon: QR uses the field polynomial 0x11d with
// generator roots starting at alpha^0; Data Matrix uses 0x12d with roots
// starting at alpha^1.
// ---------------------------------------------------------------------------

/** Exp/log tables of GF(256) for one field polynomial. */
export interface GaloisField {
  readonly exp: Uint8Array;
  readonly log: Uint8Array;
}

/** Builds the exp (doubled, so sums of two logs index directly) and log tables for `poly`. */
function makeGf(poly: number): GaloisField {
  const exp = new Uint8Array(512);
  const log = new Uint8Array(256);
  let v = 1;
  for (let i = 0; i < 255; i++) {
    exp[i] = v;
    log[v] = i;
    v <<= 1;
    if (v >= 256) v ^= poly;
  }
  for (let i = 255; i < 512; i++) {
    exp[i] = exp[i - 255] ?? 0;
  }
  return { exp, log };
}

const QR_FIELD = makeGf(0x11d);
/** @internal Data Matrix field (polynomial 0x12d), for `rsEncode(data, nsym, DM_FIELD, 1)`. */
export const DM_FIELD: GaloisField = makeGf(0x12d);

function fieldMul(f: GaloisField, a: number, b: number): number {
  if (a === 0 || b === 0) return 0;
  return f.exp[(f.log[a] ?? 0) + (f.log[b] ?? 0)] ?? 0;
}

/**
 * Reed-Solomon generator polynomial with roots alpha^firstRoot .. alpha^(firstRoot + nsym - 1).
 * Defaults to the QR field and first root 0.
 */
function rsGeneratorPoly(nsym: number, field: GaloisField = QR_FIELD, firstRoot = 0): Uint8Array {
  let g = new Uint8Array([1]);
  for (let i = 0; i < nsym; i++) {
    const next = new Uint8Array(g.length + 1);
    const factor = field.exp[i + firstRoot] ?? 0;
    for (let j = 0; j < g.length; j++) {
      next[j] = (next[j] ?? 0) ^ (g[j] ?? 0);
      next[j + 1] = (next[j + 1] ?? 0) ^ fieldMul(field, g[j] ?? 0, factor);
    }
    g = next;
  }
  return g;
}

/** Returns the `nsym` Reed-Solomon check bytes of `data`. Defaults to the QR field and first root 0. */
export function rsEncode(
  data: Uint8Array,
  nsym: number,
  field: GaloisField = QR_FIELD,
  firstRoot = 0,
): Uint8Array {
  const gen = rsGeneratorPoly(nsym, field, firstRoot);
  const out = new Uint8Array(data.length + nsym);
  out.set(data);
  for (let i = 0; i < data.length; i++) {
    const coef = out[i] ?? 0;
    if (coef !== 0) {
      for (let j = 0; j < gen.length; j++) {
        out[i + j] = (out[i + j] ?? 0) ^ fieldMul(field, gen[j] ?? 0, coef);
      }
    }
  }
  return out.subarray(data.length);
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
  $: [0x04, 0x0f, 0x14, 0x0e, 0x05, 0x1e, 0x04],
  '+': [0x00, 0x04, 0x04, 0x1f, 0x04, 0x04, 0x00],
  '%': [0x18, 0x19, 0x02, 0x04, 0x08, 0x13, 0x03],
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
// Minimal PNG encoder (1-bit palette, uncompressed deflate / store blocks)
// ---------------------------------------------------------------------------

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const chunk = new Uint8Array(4 + 4 + data.length + 4);
  const dv = new DataView(chunk.buffer);
  dv.setUint32(0, data.length, false);
  for (let i = 0; i < 4; i++) chunk[4 + i] = type.charCodeAt(i);
  chunk.set(data, 8);
  // The CRC covers the type and data bytes, which are contiguous in the chunk.
  dv.setUint32(8 + data.length, crc32(chunk.subarray(4, 8 + data.length)), false);
  return chunk;
}

/**
 * Encode a 1-bit palette PNG. `pixels` holds one palette index per pixel (0 or 1,
 * row-major); palette entry 0 is `bg`, entry 1 is `fg` (RGB triples).
 */
function createPng(
  width: number,
  height: number,
  pixels: Uint8Array,
  bg: readonly [number, number, number],
  fg: readonly [number, number, number],
): Uint8Array {
  const rowB = Math.ceil(width / 8);
  const raw = new Uint8Array(height * (1 + rowB));
  for (let r = 0; r < height; r++) {
    const rowOff = r * (1 + rowB);
    raw[rowOff] = 0; // filter type: None
    for (let x = 0; x < width; x++) {
      if (pixels[r * width + x]) {
        const at8 = rowOff + 1 + (x >> 3);
        raw[at8] = (raw[at8] ?? 0) | (0x80 >> (x & 7));
      }
    }
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
  ihdr[8] = 1; // bit depth
  ihdr[9] = 3; // colour type: palette
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const plte = new Uint8Array([
    bg[0] & 0xff,
    bg[1] & 0xff,
    bg[2] & 0xff,
    fg[0] & 0xff,
    fg[1] & 0xff,
    fg[2] & 0xff,
  ]);

  const sig = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdrC = pngChunk('IHDR', ihdr);
  const plteC = pngChunk('PLTE', plte);
  const idatC = pngChunk('IDAT', actDef);
  const iendC = pngChunk('IEND', new Uint8Array(0));

  const png = new Uint8Array(
    sig.length + ihdrC.length + plteC.length + idatC.length + iendC.length,
  );
  let pp = 0;
  png.set(sig, pp);
  pp += sig.length;
  png.set(ihdrC, pp);
  pp += ihdrC.length;
  png.set(plteC, pp);
  pp += plteC.length;
  png.set(idatC, pp);
  pp += idatC.length;
  png.set(iendC, pp);
  return png;
}

// ---------------------------------------------------------------------------
// PNG Renderer
// ---------------------------------------------------------------------------

/** Set palette index 1 in the `size` x `size` block at (x0, y0), clipped to the raster. */
function paintBlock(
  px: Uint8Array,
  tw: number,
  th: number,
  x0: number,
  y0: number,
  size: number,
): void {
  for (let y = Math.max(0, y0); y < Math.min(th, y0 + size); y++) {
    for (let x = Math.max(0, x0); x < Math.min(tw, x0 + size); x++) px[y * tw + x] = 1;
  }
}

/**
 * Paint every true cell of `bitmap` (`rows` x `cols`) as a `scale` x `scale` block
 * with the bitmap's top-left corner at (x0, y0).
 */
function paintBitmap(
  px: Uint8Array,
  tw: number,
  th: number,
  bitmap: boolean[][],
  rows: number,
  cols: number,
  x0: number,
  y0: number,
  scale: number,
): void {
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (getCell(bitmap, r, c)) paintBlock(px, tw, th, x0 + c * scale, y0 + r * scale, scale);
    }
  }
}

type Rgb = [number, number, number];

function pickColour(c: number[] | undefined, fallback: number): Rgb {
  return [c?.[0] ?? fallback, c?.[1] ?? fallback, c?.[2] ?? fallback];
}

/**
 * Renders a barcode matrix as a PNG image: dark modules in `foreground`, the rest and the quiet
 * zone in `background`, with an optional caption below.
 *
 * @param matrix - The symbol, from an `encode*` function.
 * @param options - Module size, quiet zone, colours and caption; see {@link RenderOptions} for defaults.
 * @returns The PNG file bytes.
 * @throws {RangeError} If `moduleSize` is not a positive safe integer or `quietZone` is not a non-negative safe integer, or a foreground or background channel is not an integer 0-255.
 */
export function renderBarcodePNG(matrix: BarcodeMatrix, options?: RenderOptions): Uint8Array {
  const modSz = options?.moduleSize ?? 4;
  const qz = options?.quietZone ?? 4;
  if (!Number.isSafeInteger(modSz) || modSz < 1) {
    throw new RangeError('moduleSize must be a positive integer');
  }
  if (!Number.isSafeInteger(qz) || qz < 0) {
    throw new RangeError('quietZone must be a non-negative integer');
  }
  const fg = pickColour(options?.foreground, 0);
  const bg = pickColour(options?.background, 255);
  if (![...fg, ...bg].every((v) => Number.isInteger(v) && v >= 0 && v <= 255)) {
    throw new RangeError('foreground and background channels must be integers 0-255');
  }
  const textVal = options?.textValue ?? '';
  const textBM = options?.showText && textVal.length > 0 ? renderText(textVal) : [];

  const bw = matrix.width * modSz + 2 * qz * modSz;
  const bh = matrix.height * modSz + 2 * qz * modSz;
  const textGap = modSz;
  const textPH = textBM.length > 0 ? textBM.length * modSz + textGap : 0;

  const tw = bw;
  const th = bh + textPH;
  // One palette index per pixel: 0 = background (zero-initialised), 1 = foreground.
  const px = new Uint8Array(tw * th);

  const qzPx = qz * modSz;
  paintBitmap(px, tw, th, matrix.modules, matrix.height, matrix.width, qzPx, qzPx, modSz);

  if (textBM.length > 0) {
    const tpw = textBM[0]?.length ?? 0;
    const scale = Math.max(1, Math.floor(modSz / 2));
    const txo = Math.max(0, Math.floor((tw - tpw * scale) / 2));
    paintBitmap(px, tw, th, textBM, textBM.length, tpw, txo, bh + textGap, scale);
  }

  return createPng(tw, th, px, bg, fg);
}

// ---------------------------------------------------------------------------
// OOXML Drawing XML Generator
// ---------------------------------------------------------------------------

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Last 0-based column index of a worksheet (column XFD). */
const MAX_COL_INDEX = 16383;
/** Last 0-based row index of a worksheet (row 1048576). */
const MAX_ROW_INDEX = 1048575;

/** Throws a RangeError naming `name` unless `value` is an integer in `0..max`. */
function assertIndex(name: string, value: number, max: number): void {
  if (!Number.isInteger(value) || value < 0 || value > max) {
    throw new RangeError(`anchor.${name} must be an integer 0-${max}, got ${value}`);
  }
}

/** Throws a RangeError naming `name` unless `value` is absent or a non-negative safe integer. */
function assertOffset(name: string, value: number | undefined): void {
  if (value != null && !(Number.isSafeInteger(value) && value >= 0)) {
    throw new RangeError(`anchor.${name} must be a non-negative integer, got ${value}`);
  }
}

/**
 * Checks that an image anchor addresses real worksheet cells: columns are integers 0-16383, rows
 * are integers 0-1048575, offsets (when set) are non-negative safe integers, and the anchor does
 * not end before it starts (columns, then rows, comparing offsets within the same cell).
 *
 * @param a - The anchor to check.
 * @throws {RangeError} Naming the first field that is out of range, or the axis on which the
 *   anchor ends before it starts.
 */
export function assertImageAnchor(a: ImageAnchor): void {
  assertIndex('fromCol', a.fromCol, MAX_COL_INDEX);
  assertIndex('fromRow', a.fromRow, MAX_ROW_INDEX);
  assertIndex('toCol', a.toCol, MAX_COL_INDEX);
  assertIndex('toRow', a.toRow, MAX_ROW_INDEX);
  assertOffset('fromColOff', a.fromColOff);
  assertOffset('fromRowOff', a.fromRowOff);
  assertOffset('toColOff', a.toColOff);
  assertOffset('toRowOff', a.toRowOff);
  if (a.toCol < a.fromCol || (a.toCol === a.fromCol && (a.toColOff ?? 0) < (a.fromColOff ?? 0))) {
    throw new RangeError('anchor must not end before it starts (column)');
  }
  if (a.toRow < a.fromRow || (a.toRow === a.fromRow && (a.toRowOff ?? 0) < (a.fromRowOff ?? 0))) {
    throw new RangeError('anchor must not end before it starts (row)');
  }
}

/** Escapes `&`, `<`, `>` and `"` for an XML attribute value. */
function escapeAttr(s: string): string {
  return s
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

/**
 * Builds the drawing part (`xl/drawings/drawingN.xml`) that places images on a worksheet, one
 * `xdr:twoCellAnchor` per entry. Each picture is named `Barcode 1`, `Barcode 2`, ... in order.
 *
 * @param anchors - Images to place. `anchor` is the cell range; `imageIndex` selects the
 *   relationship the picture embeds, `rId<imageIndex>`, so it must match an `rId` passed to {@link generateDrawingRels}.
 * @returns The drawing XML as a string.
 * @throws {RangeError} If an anchor column is not an integer 0-16383, a row is not an integer
 *   0-1048575, an offset is not a non-negative safe integer, or an anchor ends before it starts.
 */
export function generateDrawingXml(anchors: { anchor: ImageAnchor; imageIndex: number }[]): string {
  const p: string[] = [];
  p.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  p.push(`<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`);

  for (let i = 0; i < anchors.length; i++) {
    const entry = anchors[i];
    if (!entry) continue;
    const a = entry.anchor;
    assertImageAnchor(a);
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

/**
 * Builds the relationships part (`xl/drawings/_rels/drawingN.xml.rels`) that links a drawing to its image files.
 *
 * @param imageIds - One entry per image: `rId` is the relationship id and `target` the path of the
 *   image part; both are XML-escaped.
 * @returns The relationships XML as a string.
 */
export function generateDrawingRels(imageIds: { rId: string; target: string }[]): string {
  const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const p: string[] = [];
  p.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  p.push(`<Relationships xmlns="${RELS}">`);
  for (const id of imageIds) {
    p.push(
      `  <Relationship Id="${escapeAttr(id.rId)}" Type="${IMG}" Target="${escapeAttr(id.target)}"/>`,
    );
  }
  p.push('</Relationships>');
  return p.join('\n');
}
