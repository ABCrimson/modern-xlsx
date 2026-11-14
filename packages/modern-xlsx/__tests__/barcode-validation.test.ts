import { describe, expect, it } from 'vitest';
import { C128_W } from '../src/barcode/code128.js';
import { dmAsciiEnc, dmPad } from '../src/barcode/datamatrix.js';
import {
  encodeCode128,
  encodeDataMatrix,
  encodeGS1128,
  encodePDF417,
  generateBarcode,
  renderBarcodePNG,
} from '../src/barcode/index.js';
import { widths2mod } from '../src/barcode/linear.js';

describe('GS1-128 AI parsing', () => {
  it('inserts FNC1 after a variable-length AI', () => {
    const g = encodeGS1128('(10)ABC(21)123');
    expect(g.width).toBe(167);
    expect(g.modules[0]?.slice(77, 88)).toEqual(widths2mod(C128_W[102] as number[]));
  });

  it('encodes a fixed-length GTIN', () => {
    expect(() => encodeGS1128('(01)09501101530003')).not.toThrow();
  });

  it('removes a stray ")" from input without "(" as 1.2 did', () => {
    expect(encodeGS1128('10)ABC').modules).toEqual(encodeGS1128('10ABC').modules);
  });

  it('rejects a wrong fixed length and malformed syntax', () => {
    expect(() => encodeGS1128('(01)123')).toThrow(/must be 14 characters/);
    expect(() => encodeGS1128('(10')).toThrow(/malformed AI syntax/);
  });
});

describe('Data Matrix padding and ECI', () => {
  it('wraps the pad randomisation at 254 and never emits 0', () => {
    expect(dmPad(Array(23).fill(66), 30)[27]).toBe(254);
    for (let n = 1; n <= 40; n++) {
      const cws = dmAsciiEnc('A'.repeat(n));
      const cap = [3, 5, 8, 12, 18, 22, 30, 36, 44, 62].find((c) => c >= cws.length) ?? 62;
      expect(dmPad(cws, cap)).not.toContain(0);
    }
  });

  it('encodes Latin-1 and non-Latin-1 text', () => {
    expect(dmAsciiEnc('é')).toEqual([235, 106]);
    expect(dmAsciiEnc('中')).toEqual([241, 27, 235, 101, 235, 57, 235, 46]);
    expect(encodeDataMatrix('Ж').modules).not.toEqual(encodeDataMatrix('Ԗ').modules);
  });
});

describe('PDF417 options and capacity', () => {
  it('rejects data beyond capacity', () => {
    expect(() => encodePDF417('A'.repeat(3000))).toThrow(/data too long/);
  });

  it('truncates fractional options and rejects non-finite ones', () => {
    expect(encodePDF417('ABC', { ecLevel: 2.5 })).toEqual(encodePDF417('ABC', { ecLevel: 2 }));
    expect(() => encodePDF417('ABC', { ecLevel: Number.NaN })).toThrow(RangeError);
    expect(() => encodePDF417('ABC', { columns: Number.POSITIVE_INFINITY })).toThrow(RangeError);
  });
});

describe('render option validation', () => {
  const gen = (o: object) => generateBarcode('ABC', { type: 'code128', ...o });

  it('rejects bad moduleSize and quietZone', () => {
    for (const v of [0, -1, Number.NaN, 2.5]) {
      expect(() => gen({ moduleSize: v })).toThrow(RangeError);
    }
    expect(() => gen({ quietZone: -1 })).toThrow(RangeError);
    expect(() => gen({ quietZone: 0 })).not.toThrow();
  });

  it('rejects bad fit sizes', () => {
    expect(() => gen({ fitWidth: 0 })).toThrow(RangeError);
    expect(() => gen({ fitWidth: -50 })).toThrow(RangeError);
    expect(() => gen({ fitHeight: Number.NaN })).toThrow(RangeError);
  });

  it('fits the module size to fitWidth, fitHeight or the smaller of both', () => {
    // Code 128 of 'ABC' is 68 modules wide and 60 high; with the default quiet zone of 10 the
    // total is 88 x 80.
    const m = encodeCode128('ABC');
    const tw = m.width + 20;
    const th = m.height + 20;
    const png = (o: object) => generateBarcode('ABC', { type: 'code128', ...o }).join(',');
    expect(png({ fitWidth: tw * 3 + 2 })).toBe(png({ moduleSize: 3 }));
    expect(png({ fitHeight: th * 2 + 1 })).toBe(png({ moduleSize: 2 }));
    expect(png({ fitWidth: tw * 5, fitHeight: th * 2 })).toBe(png({ moduleSize: 2 }));
    expect(png({ fitWidth: 1 })).toBe(png({ moduleSize: 1 }));
  });

  it('rejects an unknown barcode type, including Object.prototype names', () => {
    for (const type of ['bogus', 'constructor', 'toString']) {
      expect(() => generateBarcode('A', { type: type as never })).toThrow(
        `Unknown barcode type: ${type}`,
      );
    }
  });

  it('rejects bad moduleSize in renderBarcodePNG', () => {
    expect(() => renderBarcodePNG(encodeCode128('A'), { moduleSize: 0 })).toThrow(RangeError);
  });
});

describe('text glyphs', () => {
  const png = (t: string) =>
    generateBarcode('A', { type: 'code39', showText: true, textValue: t }).join(',');

  it('draws $ + % distinctly from space', () => {
    expect(png('A$B')).not.toBe(png('A B'));
    expect(png('A+B')).not.toBe(png('A B'));
    expect(png('A%B')).not.toBe(png('A B'));
  });
});
