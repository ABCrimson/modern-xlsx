import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { encodePDF417 } from '../src/barcode/pdf417.js';

/** SHA-256 of `width x height` plus the module bits, row-major, one char per module. */
function digest(data: string, options?: { ecLevel?: number; columns?: number }): string {
  const { width, height, modules } = encodePDF417(data, options);
  const bits = modules.map((row) => row.map((m) => (m ? '1' : '0')).join('')).join('\n');
  return `${width}x${height}:${createHash('sha256').update(bits).digest('hex')}`;
}

describe('encodePDF417 output snapshot (experimental encoder, behaviour frozen)', () => {
  it('default options', () => {
    expect(digest('Hello, PDF417!')).toBe(
      '135x18:04c30c9ae350d6bba122341e29b36a5af2c4d6050155ab3b19ca8f85e8d35368',
    );
  });
  it('ecLevel 5, 6 columns', () => {
    expect(
      digest('The quick brown fox jumps over the lazy dog 0123456789', { ecLevel: 5, columns: 6 }),
    ).toBe('169x60:c09edacd3977bd871a019fc8a40d398ac209eba9291cbddf611b863848afd2da');
  });
  it('clamped options (ecLevel 99 -> 8, columns 99 -> 30)', () => {
    expect(digest('x'.repeat(40), { ecLevel: 99, columns: 99 })).toBe(
      '577x57:6886219b9e82dd3dada4dd10ea4f3ee9ef73d8a40aab5965da2ee81a67bba218',
    );
  });
  it('keeps the validation errors and their order', () => {
    expect(() => encodePDF417('a', { ecLevel: Number.NaN, columns: Number.NaN })).toThrow(
      'PDF417: ecLevel must be a finite number',
    );
    expect(() => encodePDF417('a', { columns: Number.POSITIVE_INFINITY })).toThrow(
      'PDF417: columns must be a finite number',
    );
    expect(() => encodePDF417('x'.repeat(5000))).toThrow('PDF417: data too long');
  });
});
