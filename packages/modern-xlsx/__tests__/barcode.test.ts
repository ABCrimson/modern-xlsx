import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { C128_W } from '../src/barcode/code128.js';
import { at } from '../src/barcode/common.js';
import { widths2mod } from '../src/barcode/linear.js';
import { QR_VI } from '../src/barcode/qr.js';
import {
  encodeCode39,
  encodeCode128,
  encodeDataMatrix,
  encodeEAN13,
  encodeGS1128,
  encodeITF14,
  encodePDF417,
  encodeQR,
  encodeUPCA,
  generateBarcode,
  generateDrawingRels,
  generateDrawingXml,
  renderBarcodePNG,
  Workbook,
} from '../src/index.js';

describe('Barcode Encoders', () => {
  it('encodes QR code with valid matrix', () => {
    const matrix = encodeQR('Hello');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBe(matrix.width); // QR codes are square
    expect(matrix.modules.length).toBe(matrix.height);
    expect(matrix.modules[0]?.length).toBe(matrix.width);
  });

  it('encodes QR with different EC levels', () => {
    const l = encodeQR('Test', { ecLevel: 'L' });
    const h = encodeQR('Test', { ecLevel: 'H' });
    // Higher EC level = larger matrix
    expect(h.width).toBeGreaterThanOrEqual(l.width);
  });

  it('encodes Code 128', () => {
    const matrix = encodeCode128('ABC-123');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
    // Code 128 is 1D — height should be uniform
    expect(matrix.modules.length).toBe(matrix.height);
  });

  it('encodes EAN-13', () => {
    const matrix = encodeEAN13('5901234123457');
    expect(matrix.width).toBe(95); // EAN-13 is always 95 modules wide
    expect(matrix.height).toBeGreaterThan(0);
  });

  it('encodes EAN-13 with auto check digit', () => {
    // 12 digits — check digit should be auto-calculated
    const matrix = encodeEAN13('590123412345');
    expect(matrix.modules).toEqual(encodeEAN13('5901234123457').modules);
    // GS1 mod-10 check digit is 1 (20 + 3*23 = 89); swapped weights would give 7.
    expect(encodeEAN13('400638133393').modules).toEqual(encodeEAN13('4006381333931').modules);
  });

  it('encodes UPC-A', () => {
    const matrix = encodeUPCA('012345678905');
    expect(matrix.width).toBe(95); // UPC-A wraps EAN-13
  });

  it('encodes Code 39', () => {
    const matrix = encodeCode39('HELLO');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
  });

  it('encodes every Code 39 character with its standard pattern', () => {
    // Standard Code 39 encodings: 9 elements, most significant bit first, bars and spaces
    // alternating from a bar; a set bit is a wide (2-module) element.
    const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%';
    const codes = [
      0x034, 0x121, 0x061, 0x160, 0x031, 0x130, 0x070, 0x025, 0x124, 0x064, 0x109, 0x049, 0x148,
      0x019, 0x118, 0x058, 0x00d, 0x10c, 0x04c, 0x01c, 0x103, 0x043, 0x142, 0x013, 0x112, 0x052,
      0x007, 0x106, 0x046, 0x016, 0x181, 0x0c1, 0x1c0, 0x091, 0x190, 0x0d0, 0x085, 0x184, 0x0c4,
      0x0a8, 0x0a2, 0x08a, 0x02a,
    ];
    const star = 0x094;
    const expand = (code: number): boolean[] => {
      const out: boolean[] = [];
      for (let k = 0; k < 9; k++) {
        const wide = ((code >> (8 - k)) & 1) === 1;
        const bar = k % 2 === 0;
        for (let w = 0; w < (wide ? 2 : 1); w++) out.push(bar);
      }
      return out;
    };
    expect(codes.length).toBe(chars.length);
    for (let i = 0; i < chars.length; i++) {
      const ch = chars[i] as string;
      const expected = [...expand(star), false, ...expand(at(codes, i)), false, ...expand(star)];
      expect(encodeCode39(ch).modules[0], `Code 39 '${ch}'`).toEqual(expected);
    }
  });

  it('keeps every QR block structure consistent with the total codeword count', () => {
    for (const level of ['L', 'M', 'Q', 'H'] as const) {
      for (let version = 1; version <= 10; version++) {
        const [total, ec, g1, d1, g2, d2] = at(QR_VI[level], version);
        expect(g1 * (d1 + ec) + g2 * (d2 + ec), `${version}-${level}`).toBe(total);
      }
    }
  });

  it('encodes 70 bytes at level Q as version 6 (41 modules)', () => {
    expect(encodeQR('x'.repeat(70), { ecLevel: 'Q' }).width).toBe(41);
  });

  it('encodes ITF-14', () => {
    const matrix = encodeITF14('00012345600012');
    expect(matrix.width).toBeGreaterThan(0);
    // 13 digits: weights 3,1 give sum 48, so the check digit is 2.
    expect(encodeITF14('0001234560001').modules).toEqual(matrix.modules);
  });

  it('encodes GS1-128', () => {
    const matrix = encodeGS1128('(01)09501101530003');
    expect(matrix.width).toBeGreaterThan(0);
  });

  it('encodes PDF417', () => {
    const matrix = encodePDF417('Hello World');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBeGreaterThan(0);
    // PDF417 is 2D stacked
    expect(matrix.height).toBeGreaterThan(1);
  });

  it('encodes Data Matrix', () => {
    const matrix = encodeDataMatrix('Hello');
    expect(matrix.width).toBeGreaterThan(0);
    expect(matrix.height).toBe(matrix.width); // Data Matrix is square
  });
});

describe('PNG Renderer', () => {
  it('renders barcode to valid PNG bytes', () => {
    const matrix = encodeQR('Test');
    const png = renderBarcodePNG(matrix);
    expect(png).toBeInstanceOf(Uint8Array);
    expect(png.length).toBeGreaterThan(0);
    // Check PNG magic bytes
    expect(png[0]).toBe(0x89);
    expect(png[1]).toBe(0x50); // P
    expect(png[2]).toBe(0x4e); // N
    expect(png[3]).toBe(0x47); // G
  });

  it('respects moduleSize option', () => {
    const matrix = encodeQR('A');
    const small = renderBarcodePNG(matrix, { moduleSize: 2 });
    const large = renderBarcodePNG(matrix, { moduleSize: 8 });
    expect(large.length).toBeGreaterThan(small.length);
  });

  it('renders with showText option', () => {
    const matrix = encodeCode128('123');
    const withText = renderBarcodePNG(matrix, { showText: true, textValue: '123' });
    const withoutText = renderBarcodePNG(matrix, { showText: false });
    expect(withText.length).toBeGreaterThan(withoutText.length);
  });
});

describe('generateBarcode helper', () => {
  it('generates QR code PNG', () => {
    const png = generateBarcode('https://example.com', { type: 'qr' });
    expect(png[0]).toBe(0x89); // PNG magic
  });

  it('generates Code 128 PNG', () => {
    const png = generateBarcode('ABC-123', { type: 'code128' });
    expect(png[0]).toBe(0x89);
  });

  it('generates with fit dimensions', () => {
    const png = generateBarcode('Test', {
      type: 'qr',
      fitWidth: 200,
      fitHeight: 200,
    });
    expect(png.length).toBeGreaterThan(0);
  });
});

describe('Drawing XML Generation', () => {
  it('generates valid drawing XML', () => {
    const xml = generateDrawingXml([
      {
        anchor: { fromCol: 0, fromRow: 0, toCol: 3, toRow: 5 },
        imageIndex: 1,
      },
    ]);
    expect(xml).toContain('xdr:wsDr');
    expect(xml).toContain('xdr:twoCellAnchor');
    expect(xml).toContain('xdr:from');
    expect(xml).toContain('xdr:to');
  });

  it('generates valid drawing rels', () => {
    const xml = generateDrawingRels([{ rId: 'rId1', target: '../media/image1.png' }]);
    expect(xml).toContain('Relationships');
    expect(xml).toContain('rId1');
    expect(xml).toContain('../media/image1.png');
  });
});

describe('Workbook barcode integration', () => {
  it('adds barcode image to workbook preservedEntries', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addBarcode('Sheet1', { fromCol: 0, fromRow: 0, toCol: 4, toRow: 4 }, 'Hello', {
      type: 'qr',
    });

    const data = wb.toJSON();
    expect(data.preservedEntries).toBeDefined();

    // Should have media, drawing, rels entries
    const entries = Object.keys(data.preservedEntries ?? {});
    expect(entries.some((e) => e.startsWith('xl/media/'))).toBe(true);
    expect(entries.some((e) => e.startsWith('xl/drawings/'))).toBe(true);
    expect(entries.some((e) => e.includes('_rels'))).toBe(true);
  });

  it('writes workbook with barcode to valid XLSX', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Barcodes');
    ws.cell('A1').value = 'QR Code:';
    wb.addBarcode(
      'Barcodes',
      { fromCol: 1, fromRow: 0, toCol: 5, toRow: 4 },
      'https://example.com',
      {
        type: 'qr',
      },
    );

    const buffer = await wb.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('addImage embeds raw PNG bytes', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    // Render a QR code to PNG bytes to embed as the image
    const qr = encodeQR('Test');
    const png = renderBarcodePNG(qr, { moduleSize: 1 });

    wb.addImage('Sheet1', { fromCol: 0, fromRow: 0, toCol: 2, toRow: 2 }, png);

    const data = wb.toJSON();
    const entries = Object.keys(data.preservedEntries ?? {});
    expect(entries.some((e) => e.includes('image1.png'))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Known vectors: expectations computed from the standards, not from the encoder under test, except the QR golden below
// ---------------------------------------------------------------------------

/**
 * Derived golden from repro qrmask.mjs build(0): the encoder's mask-0 data modules with the
 * ISO 18004 format information in both copies; the codeword assertion (independent Reed-Solomon)
 * and the mask-0 penalty check back it. Regenerate if the mask choice changes.
 */
const QR_01234567_M_MASK0 = [
  '111111100011101111111',
  '100000101110001000001',
  '101110100110001011101',
  '101110100101101011101',
  '101110101101101011101',
  '100000100001001000001',
  '111111101010101111111',
  '000000000000000000000',
  '101010100010100010010',
  '110100001011010100010',
  '000110111011011101110',
  '110011010101110110010',
  '001001110111011100001',
  '000000001010001000010',
  '111111100000100010001',
  '100000100010001001011',
  '101110101110101011101',
  '101110100101010101110',
  '101110101101011100101',
  '100000100001110111000',
  '111111101001011100101',
];

const QR_MASKS: ((r: number, c: number) => boolean)[] = [
  (r, c) => (r + c) % 2 === 0,
  (r) => r % 2 === 0,
  (_r, c) => c % 3 === 0,
  (r, c) => (r + c) % 3 === 0,
  (r, c) => (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0,
  (r, c) => ((r * c) % 2) + ((r * c) % 3) === 0,
  (r, c) => (((r * c) % 2) + ((r * c) % 3)) % 2 === 0,
  (r, c) => (((r + c) % 2) + ((r * c) % 3)) % 2 === 0,
];

/** True for the version-1 function-pattern cells (finders, format areas, timing lines). */
function qrV1IsFunction(r: number, c: number, n: number): boolean {
  return (r < 9 && c < 9) || (r < 9 && c >= n - 8) || (r >= n - 8 && c < 9) || r === 6 || c === 6;
}

/** Data-module cells of a version-1 symbol in zigzag codeword order. */
function qrV1DataCells(n: number): [number, number][] {
  const cells: [number, number][] = [];
  const ascending = Array.from({ length: n }, (_, k) => n - 1 - k);
  let up = true;
  for (let right = n - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    const rows = up ? ascending : [...ascending].reverse();
    for (const r of rows) {
      for (const c of [right, right - 1]) {
        if (!qrV1IsFunction(r, c, n)) cells.push([r, c]);
      }
    }
    up = !up;
  }
  return cells;
}

/** Pack a bit list into whole bytes, most significant bit first; a trailing partial byte is dropped. */
function bitsToBytes(bits: number[]): number[] {
  const bytes: number[] = [];
  for (let i = 0; i + 8 <= bits.length; i += 8) {
    bytes.push(Number.parseInt(bits.slice(i, i + 8).join(''), 2));
  }
  return bytes;
}

/** Version-1 zigzag codeword reader; undoes mask 0 and skips the function patterns. */
function qrCodewords(modules: boolean[][], n: number): number[] {
  const bits = qrV1DataCells(n).map(
    ([r, c]) => (modules[r]?.[c] ? 1 : 0) ^ ((r + c) % 2 === 0 ? 1 : 0),
  );
  return bitsToBytes(bits);
}

/** The 4-bit mode indicator of a version-1 symbol: bottom-right 2x2 cells, unmasked. */
function qrModeBits(modules: boolean[][], mask: (r: number, c: number) => boolean): string {
  const cells: [number, number][] = [
    [20, 20],
    [20, 19],
    [19, 20],
    [19, 19],
  ];
  return cells
    .map(([r, c]) => ((modules[r]?.[c] ? 1 : 0) ^ (mask(r, c) ? 1 : 0)).toString())
    .join('');
}

const bitString = (m: boolean[][], cells: [number, number][]): string =>
  cells.map(([r, c]) => (m[r]?.[c] ? '1' : '0')).join('');

/** Format-information copy 1 cells in ISO 18004 order (bit 14 first). */
const copy1: [number, number][] = [
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

/** The mask pattern index (0-7) a symbol declares in format copy 1: unmask with 0x5412, take bits 12-10. */
const qrMask = (m: boolean[][]): number =>
  ((Number.parseInt(bitString(m, copy1), 2) ^ 0x5412) >> 10) & 7;

const PNG_CHANNELS: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

function paeth(a: number, b: number, c: number): number {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

/** Apply one PNG scanline filter (types 0-4) to a filtered byte. */
function unfilterByte(ft: number, cur: number, a: number, b: number, c: number): number {
  switch (ft) {
    case 1:
      return cur + a;
    case 2:
      return cur + b;
    case 3:
      return cur + ((a + b) >> 1);
    case 4:
      return cur + paeth(a, b, c);
    default:
      return cur;
  }
}

/** Undo the per-scanline filters of inflated PNG data; returns the raw scanline bytes. */
function unfilterScanlines(
  raw: Uint8Array,
  height: number,
  stride: number,
  bpp: number,
): Uint8Array {
  const out = new Uint8Array(height * stride);
  for (let y = 0; y < height; y++) {
    const ft = raw[y * (stride + 1)] as number;
    for (let x = 0; x < stride; x++) {
      const cur = raw[y * (stride + 1) + 1 + x] as number;
      const a = x >= bpp ? (out[y * stride + x - bpp] as number) : 0;
      const b = y > 0 ? (out[(y - 1) * stride + x] as number) : 0;
      const c = x >= bpp && y > 0 ? (out[(y - 1) * stride + x - bpp] as number) : 0;
      out[y * stride + x] = unfilterByte(ft, cur, a, b, c) & 0xff;
    }
  }
  return out;
}

interface PngLumContext {
  depth: number;
  colourType: number;
  channels: number;
  bpp: number;
  bitsPerPixel: number;
  stride: number;
  plte: Uint8Array;
}

const paletteLum = (plte: Uint8Array, i: number): number =>
  ((plte[i * 3] ?? 0) + (plte[i * 3 + 1] ?? 0) + (plte[i * 3 + 2] ?? 0)) / 3;

/** Luminance (0-255) of pixel (x, y) in unfiltered scanline bytes. */
function sampleLum(out: Uint8Array, x: number, y: number, ctx: PngLumContext): number {
  const { depth, colourType, channels, bpp, bitsPerPixel, stride, plte } = ctx;
  if (depth < 8) {
    const bit = x * bitsPerPixel;
    const sample =
      ((out[y * stride + (bit >> 3)] as number) >> (8 - depth - (bit & 7))) & ((1 << depth) - 1);
    return colourType === 3 ? paletteLum(plte, sample) : (sample * 255) / ((1 << depth) - 1);
  }
  if (colourType === 3) return paletteLum(plte, out[y * stride + x] as number);
  const o = y * stride + x * bpp;
  if (channels < 3) return out[o] as number;
  return ((out[o] as number) + (out[o + 1] as number) + (out[o + 2] as number)) / 3;
}

/** Decode a PNG by its IHDR bit depth and colour type; returns dark = true per pixel. */
function decodePngDark(png: Uint8Array): { width: number; height: number; dark: boolean[][] } {
  const dv = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const width = dv.getUint32(16);
  const height = dv.getUint32(20);
  const depth = png[24] as number;
  const colourType = png[25] as number;
  const idat: Uint8Array[] = [];
  let plte: Uint8Array = new Uint8Array(0);
  for (let off = 8; off < png.length; ) {
    const len = dv.getUint32(off);
    const type = String.fromCharCode(...png.subarray(off + 4, off + 8));
    const body = png.subarray(off + 8, off + 8 + len);
    if (type === 'IDAT') idat.push(body);
    else if (type === 'PLTE') plte = body;
    off += 12 + len;
  }
  const channels = PNG_CHANNELS[colourType];
  if (channels === undefined) throw new Error(`unexpected PNG colour type ${colourType}`);
  const bitsPerPixel = channels * depth;
  const bpp = Math.max(1, bitsPerPixel >> 3);
  const stride = Math.ceil((width * bitsPerPixel) / 8);
  const out = unfilterScanlines(inflateSync(Buffer.concat(idat)), height, stride, bpp);
  const ctx: PngLumContext = { depth, colourType, channels, bpp, bitsPerPixel, stride, plte };
  const dark = Array.from({ length: height }, (_, y) =>
    Array.from({ length: width }, (_, x) => sampleLum(out, x, y, ctx) < 128),
  );
  return { width, height, dark };
}

describe('known vectors', () => {
  it('QR: format information (both copies), codewords and golden matrix for 01234567 / M', () => {
    const q = encodeQR('01234567', { ecLevel: 'M' });
    expect(q.width).toBe(21);
    expect(q.height).toBe(21);
    // Copy 2 in ISO 18004 order: (20..14, 8) then (8, 13..20).
    const copy2: [number, number][] = [];
    for (let r = 20; r >= 14; r--) copy2.push([r, 8]);
    for (let c = 13; c <= 20; c++) copy2.push([8, c]);
    expect(bitString(q.modules, copy1)).toBe('101010000010010');
    expect(bitString(q.modules, copy2)).toBe('101010000010010');
    expect(qrCodewords(q.modules, 21)).toEqual([
      16, 32, 12, 86, 97, 128, 236, 17, 236, 17, 236, 17, 236, 17, 236, 17, 165, 36, 212, 193, 237,
      54, 199, 135, 44, 85,
    ]);
    expect(q.modules.map((row) => row.map((b) => (b ? '1' : '0')).join(''))).toEqual(
      QR_01234567_M_MASK0,
    );
  });

  it('QR: lowercase text is byte mode, not upper-cased alphanumeric', () => {
    const lower = encodeQR('hello');
    const upper = encodeQR('HELLO');
    expect(lower.modules).not.toEqual(upper.modules);
    // Unmask with the mask each symbol declares in its own format information.
    expect(qrModeBits(lower.modules, at(QR_MASKS, qrMask(lower.modules)))).toBe('0100');
    expect(qrModeBits(upper.modules, at(QR_MASKS, qrMask(upper.modules)))).toBe('0010');
  });

  it('Code 128: mixed-path checksum weights follow the symbol index', () => {
    const c = encodeCode128('AB1234');
    expect(c.width).toBe(90);
    // start B (104) + A*1 + B*2 + 99*3 + 12*4 + 34*5 = 720; 720 mod 103 = 102
    expect(c.modules[0]?.slice(66, 77)).toEqual(widths2mod(at(C128_W, 102)));
    // 'Hello123456': the check symbol sits just before the 13-module stop pattern.
    const h = encodeCode128('Hello123456');
    expect(h.modules[0]?.slice(h.width - 24, h.width - 13)).toEqual(widths2mod(at(C128_W, 88)));
  });

  it('EAN-13: 5901234123457 matches the published L/G/R pattern', () => {
    const L = [
      '0001101',
      '0011001',
      '0010011',
      '0111101',
      '0100011',
      '0110001',
      '0101111',
      '0111011',
      '0110111',
      '0001011',
    ];
    const R = L.map((p) => [...p].map((b) => (b === '1' ? '0' : '1')).join(''));
    const G = R.map((p) => [...p].reverse().join(''));
    const parity = [
      'LLLLLL',
      'LLGLGG',
      'LLGGLG',
      'LLGGGL',
      'LGLLGG',
      'LGGLLG',
      'LGGGLL',
      'LGLGLG',
      'LGLGGL',
      'LGGLGL',
    ];
    const d = [...'5901234123457'].map(Number);
    let expected = '101';
    for (let i = 1; i <= 6; i++) {
      const t = parity[d[0] as number]?.[i - 1] === 'L' ? L : G;
      expected += t[d[i] as number];
    }
    expected += '01010';
    for (let i = 7; i <= 12; i++) expected += R[d[i] as number];
    expected += '101';
    expect(expected).toHaveLength(95);
    const row = encodeEAN13('5901234123457').modules[0] ?? [];
    expect(row.map((b) => (b ? '1' : '0')).join('')).toBe(expected);
  });

  it('UPC-A, EAN-13, ITF-14 and Code 39 validate their input', () => {
    expect(() => encodeUPCA('012345678900')).toThrow(/check digit/);
    expect(() => encodeITF14('00012345600013')).toThrow(
      'ITF-14: invalid check digit (expected 2, got 3)',
    );
    expect(encodeUPCA('01234567890').modules).toEqual(encodeUPCA('012345678905').modules);
    expect(encodeUPCA('012345678905').modules).toEqual(encodeEAN13('0012345678905').modules);
    expect(() => encodeEAN13('12345678901 ')).toThrow(/only digits/);
    expect(() => encodeEAN13('123456789012A')).toThrow(/only digits/);
    expect(() => encodeITF14('123456789012 ')).toThrow(/only digits/);
    expect(() => encodeCode39('')).toThrow('Code 39: data must not be empty');
  });

  it('PNG: fitted QR is 174x174 and a 1:1 render decodes to the module grid', () => {
    const fitted = generateBarcode('01234567', { type: 'qr', fitWidth: 200, fitHeight: 200 });
    const dv = new DataView(fitted.buffer, fitted.byteOffset, fitted.byteLength);
    expect([dv.getUint32(16), dv.getUint32(20)]).toEqual([174, 174]);

    const q = encodeQR('01234567', { ecLevel: 'M' });
    const img = decodePngDark(renderBarcodePNG(q, { moduleSize: 1, quietZone: 0 }));
    expect([img.width, img.height]).toEqual([21, 21]);
    expect(img.dark).toEqual(q.modules);
  });

  // The PDF417 and Data Matrix encoders are @experimental (not yet ISO/IEC 15438 / 16022 conformant), so their known-answer vectors stay todo.
  it.todo('PDF417: ISO/IEC 15438 known-answer vector');
  it.todo('Data Matrix: ISO/IEC 16022 known-answer vector');
});
