import { describe, expect, it } from 'vitest';
import { decodeRow, encodeRow, splitCellRef } from '../src/index.js';

describe('extended cell reference utilities', () => {
  describe('encodeRow', () => {
    it('converts 0-based index to 1-based string', () => {
      expect(encodeRow(0)).toBe('1');
      expect(encodeRow(9)).toBe('10');
      expect(encodeRow(1048575)).toBe('1048576');
    });
  });

  describe('decodeRow', () => {
    it('converts 1-based string to 0-based index', () => {
      expect(decodeRow('1')).toBe(0);
      expect(decodeRow('10')).toBe(9);
      expect(decodeRow('1048576')).toBe(1048575);
    });

    it('throws on invalid input', () => {
      expect(() => decodeRow('')).toThrow();
      expect(() => decodeRow('0')).toThrow();
      expect(() => decodeRow('abc')).toThrow();
    });
  });

  describe('splitCellRef', () => {
    it('splits simple reference', () => {
      expect(splitCellRef('A1')).toEqual({
        col: 'A', row: '1', absCol: false, absRow: false,
      });
    });

    it('splits fully absolute reference', () => {
      expect(splitCellRef('$A$1')).toEqual({
        col: 'A', row: '1', absCol: true, absRow: true,
      });
    });

    it('splits mixed references', () => {
      expect(splitCellRef('$A1')).toEqual({
        col: 'A', row: '1', absCol: true, absRow: false,
      });
      expect(splitCellRef('A$1')).toEqual({
        col: 'A', row: '1', absCol: false, absRow: true,
      });
    });

    it('handles multi-letter columns', () => {
      expect(splitCellRef('$XFD$1048576')).toEqual({
        col: 'XFD', row: '1048576', absCol: true, absRow: true,
      });
    });

    it('throws on invalid ref', () => {
      expect(() => splitCellRef('')).toThrow();
      expect(() => splitCellRef('123')).toThrow();
    });
  });
});
