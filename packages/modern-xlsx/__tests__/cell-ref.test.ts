import { describe, expect, it } from 'vitest';
import {
  columnToLetter,
  decodeCellRef,
  decodeRange,
  encodeCellRef,
  encodeRange,
  letterToColumn,
} from '../src/cell-ref.js';

describe('columnToLetter', () => {
  it('converts single-letter columns', () => {
    expect(columnToLetter(0)).toBe('A');
    expect(columnToLetter(25)).toBe('Z');
  });

  it('converts double-letter columns', () => {
    expect(columnToLetter(26)).toBe('AA');
    expect(columnToLetter(27)).toBe('AB');
    expect(columnToLetter(51)).toBe('AZ');
    expect(columnToLetter(52)).toBe('BA');
    expect(columnToLetter(701)).toBe('ZZ');
  });

  it('converts triple-letter columns', () => {
    expect(columnToLetter(702)).toBe('AAA');
  });
});

describe('letterToColumn', () => {
  it('converts single letters', () => {
    expect(letterToColumn('A')).toBe(0);
    expect(letterToColumn('Z')).toBe(25);
  });

  it('converts double letters', () => {
    expect(letterToColumn('AA')).toBe(26);
    expect(letterToColumn('AB')).toBe(27);
    expect(letterToColumn('AZ')).toBe(51);
    expect(letterToColumn('BA')).toBe(52);
    expect(letterToColumn('ZZ')).toBe(701);
  });

  it('converts triple letters', () => {
    expect(letterToColumn('AAA')).toBe(702);
  });

  it('is inverse of columnToLetter', () => {
    for (let i = 0; i < 100; i++) {
      expect(letterToColumn(columnToLetter(i))).toBe(i);
    }
  });
});

describe('encodeCellRef', () => {
  it('encodes basic references', () => {
    expect(encodeCellRef(0, 0)).toBe('A1');
    expect(encodeCellRef(0, 25)).toBe('Z1');
    expect(encodeCellRef(9, 2)).toBe('C10');
    expect(encodeCellRef(99, 26)).toBe('AA100');
  });
});

describe('decodeCellRef', () => {
  it('decodes basic references', () => {
    expect(decodeCellRef('A1')).toEqual({ row: 0, col: 0 });
    expect(decodeCellRef('Z1')).toEqual({ row: 0, col: 25 });
    expect(decodeCellRef('C10')).toEqual({ row: 9, col: 2 });
    expect(decodeCellRef('AA100')).toEqual({ row: 99, col: 26 });
  });

  it('handles absolute references ($)', () => {
    expect(decodeCellRef('$A$1')).toEqual({ row: 0, col: 0 });
    expect(decodeCellRef('$B$3')).toEqual({ row: 2, col: 1 });
  });

  it('throws on invalid reference', () => {
    expect(() => decodeCellRef('1A')).toThrow('Invalid cell reference');
    expect(() => decodeCellRef('')).toThrow('Invalid cell reference');
  });

  it('is inverse of encodeCellRef', () => {
    for (let r = 0; r < 10; r++) {
      for (let c = 0; c < 30; c++) {
        const ref = encodeCellRef(r, c);
        const decoded = decodeCellRef(ref);
        expect(decoded).toEqual({ row: r, col: c });
      }
    }
  });
});

describe('encodeRange / decodeRange', () => {
  it('encodes a range', () => {
    expect(encodeRange({ row: 0, col: 0 }, { row: 9, col: 2 })).toBe('A1:C10');
  });

  it('decodes a range', () => {
    expect(decodeRange('A1:C10')).toEqual({
      start: { row: 0, col: 0 },
      end: { row: 9, col: 2 },
    });
  });

  it('throws on invalid range', () => {
    expect(() => decodeRange('A1')).toThrow('Invalid range');
  });

  it('roundtrips', () => {
    const s = { row: 5, col: 3 };
    const e = { row: 20, col: 10 };
    expect(decodeRange(encodeRange(s, e))).toEqual({ start: s, end: e });
  });
});
