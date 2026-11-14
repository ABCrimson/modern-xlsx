import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { type BarcodeMatrix, getCell, renderText } from '../src/barcode/common.js';
import type { ImageAnchor } from '../src/index.js';
import {
  encodeQR,
  generateDrawingRels,
  generateDrawingXml,
  renderBarcodePNG,
  Workbook,
} from '../src/index.js';

describe('encodeQR early length reject', () => {
  it('accepts the version 10-L numeric capacity', () => {
    expect(() => encodeQR('1'.repeat(652), { ecLevel: 'L' })).not.toThrow();
  });

  it('rejects one digit more', () => {
    expect(() => encodeQR('1'.repeat(653), { ecLevel: 'L' })).toThrow(/too long/);
  });

  it('rejects a 5,000,000-character input as too long', () => {
    const big = 'a'.repeat(5e6);
    // No wall-clock bound: without the early reject, qrSelVer throws the same error after a linear encode, so a time bound could not tell the paths apart; the early reject is a bench candidate.
    expect(() => encodeQR(big)).toThrow(/too long/);
  });
});

interface DecodedPng {
  width: number;
  height: number;
  depth: number;
  colourType: number;
  palette: number[][];
  /** RGB triples, row-major. */
  rgb: Uint8Array;
}

interface PngChunks {
  ihdr: Uint8Array;
  plte: Uint8Array;
  idat: Uint8Array;
}

/** Split a PNG into its IHDR, PLTE and (concatenated) IDAT payloads. */
function readChunks(png: Uint8Array): PngChunks {
  const dv = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const found = new Map<string, Uint8Array[]>();
  for (let pos = 8; pos < png.length; ) {
    const len = dv.getUint32(pos, false);
    const type = String.fromCharCode(...png.subarray(pos + 4, pos + 8));
    found.set(type, [...(found.get(type) ?? []), png.subarray(pos + 8, pos + 8 + len)]);
    pos += 12 + len;
  }
  const join = (type: string) => new Uint8Array(Buffer.concat(found.get(type) ?? []));
  return { ihdr: join('IHDR'), plte: join('PLTE'), idat: join('IDAT') };
}

function decodePng(png: Uint8Array): DecodedPng {
  const { ihdr, plte, idat } = readChunks(png);
  const dv = new DataView(ihdr.buffer, ihdr.byteOffset, ihdr.byteLength);
  const width = dv.getUint32(0, false);
  const height = dv.getUint32(4, false);
  const palette: number[][] = [];
  for (let i = 0; i < plte.length; i += 3) palette.push([...plte.subarray(i, i + 3)]);

  const raw = new Uint8Array(inflateSync(idat));
  const rowBytes = Math.ceil(width / 8);
  const rgb = new Uint8Array(width * height * 3);
  for (let y = 0; y < height; y++) {
    const base = y * (1 + rowBytes);
    expect(raw[base]).toBe(0);
    for (let x = 0; x < width; x++) {
      const bit = ((raw[base + 1 + (x >> 3)] ?? 0) >> (7 - (x & 7))) & 1;
      rgb.set(palette[bit] ?? [0, 0, 0], (y * width + x) * 3);
    }
  }
  return { width, height, depth: ihdr[8] ?? 0, colourType: ihdr[9] ?? -1, palette, rgb };
}

/** Colour the scale x scale block at (x0, y0), clipped to the raster. */
function stampBlock(
  rgb: Uint8Array,
  tw: number,
  th: number,
  x0: number,
  y0: number,
  scale: number,
  fg: number[],
): void {
  for (let y = Math.max(0, y0); y < Math.min(th, y0 + scale); y++) {
    for (let x = Math.max(0, x0); x < Math.min(tw, x0 + scale); x++) rgb.set(fg, (y * tw + x) * 3);
  }
}

/** Stamp every true cell of `bitmap` as a scale x scale block of `fg`. */
function stamp(
  rgb: Uint8Array,
  tw: number,
  th: number,
  bitmap: boolean[][],
  [x0, y0, scale]: [number, number, number],
  fg: number[],
): void {
  for (let r = 0; r < bitmap.length; r++) {
    for (let c = 0; c < (bitmap[r]?.length ?? 0); c++) {
      if (getCell(bitmap, r, c)) stampBlock(rgb, tw, th, x0 + c * scale, y0 + r * scale, scale, fg);
    }
  }
}

/** RGB reference built straight from the matrix (and the text raster). */
function referenceRgb(
  matrix: BarcodeMatrix,
  fg: number[],
  bg: number[],
  text: string,
): { width: number; height: number; rgb: Uint8Array } {
  const modSz = 4;
  const qz = 4;
  const bw = matrix.width * modSz + 2 * qz * modSz;
  const bh = matrix.height * modSz + 2 * qz * modSz;
  const textBM = renderText(text);
  const tw = bw;
  const th = bh + textBM.length * modSz + modSz;
  const rgb = new Uint8Array(tw * th * 3);
  for (let i = 0; i < tw * th; i++) rgb.set(bg, i * 3);
  stamp(rgb, tw, th, matrix.modules, [qz * modSz, qz * modSz, modSz], fg);
  const scale = Math.max(1, Math.floor(modSz / 2));
  const txo = Math.max(0, Math.floor((tw - (textBM[0]?.length ?? 0) * scale) / 2));
  stamp(rgb, tw, th, textBM, [txo, bh + modSz, scale], fg);
  return { width: tw, height: th, rgb };
}

describe('renderBarcodePNG 1-bit palette output', () => {
  it('decodes to the RGB reference for a QR code with text', () => {
    const matrix = encodeQR('Test');
    const png = renderBarcodePNG(matrix, { showText: true, textValue: 'Test' });
    const decoded = decodePng(png);
    expect(decoded.depth).toBe(1);
    expect(decoded.colourType).toBe(3);
    const ref = referenceRgb(matrix, [0, 0, 0], [255, 255, 255], 'Test');
    expect(decoded.width).toBe(ref.width);
    expect(decoded.height).toBe(ref.height);
    expect(Buffer.from(decoded.rgb).equals(Buffer.from(ref.rgb))).toBe(true);
  });

  it('keeps custom foreground and background colours', () => {
    const matrix = encodeQR('Test');
    const png = renderBarcodePNG(matrix, {
      foreground: [200, 0, 0],
      background: [0, 0, 255],
      showText: true,
      textValue: 'Test',
    });
    const decoded = decodePng(png);
    expect(decoded.depth).toBe(1);
    expect(decoded.colourType).toBe(3);
    expect(decoded.palette).toEqual([
      [0, 0, 255],
      [200, 0, 0],
    ]);
    const ref = referenceRgb(matrix, [200, 0, 0], [0, 0, 255], 'Test');
    expect(Buffer.from(decoded.rgb).equals(Buffer.from(ref.rgb))).toBe(true);
  });

  it.each([
    ['foreground 256', { foreground: [256, 0, 0] }],
    ['background -1', { background: [-1, 0, 0] }],
    ['foreground 1.5', { foreground: [1.5, 0, 0] }],
  ])('rejects a %s colour channel with a RangeError', (_label, options) => {
    const matrix = encodeQR('Test');
    expect(() => renderBarcodePNG(matrix, options)).toThrow(RangeError);
  });
});

describe('image anchor validation', () => {
  const valid: ImageAnchor = { fromCol: 1, fromRow: 1, toCol: 4, toRow: 9 };
  const bad: [string, ImageAnchor][] = [
    ['fromCol 1.5', { ...valid, fromCol: 1.5 }],
    ['fromRow -1', { ...valid, fromRow: -1 }],
    ['toRow NaN', { ...valid, toRow: Number.NaN }],
    ['toCol < fromCol', { ...valid, fromCol: 4, toCol: 1 }],
  ];
  const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
  const mediaKeys = (wb: Workbook): string[] =>
    Object.keys(wb.toJSON().preservedEntries ?? {}).filter((k) => k.startsWith('xl/media/'));

  it.each(bad)('addImage rejects %s and stores nothing', (_label, anchor) => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => wb.addImage('Sheet1', anchor, png)).toThrow(RangeError);
    expect(mediaKeys(wb)).toEqual([]);
  });

  it.each(bad)('generateDrawingXml rejects %s', (_label, anchor) => {
    expect(() => generateDrawingXml([{ anchor, imageIndex: 1 }])).toThrow(RangeError);
  });

  it('accepts a valid addImage after a rejected one', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => wb.addImage('Sheet1', { ...valid, fromCol: 1.5 }, png)).toThrow(RangeError);
    expect(() => wb.addImage('Sheet1', valid, png)).not.toThrow();
    expect(mediaKeys(wb)).toEqual(['xl/media/image1.png']);
  });

  it('XML-escapes the relationship target', () => {
    const xml = generateDrawingRels([{ rId: 'rId1', target: 'a&b.png' }]);
    expect(xml).toContain('Target="a&amp;b.png"');
  });
});
