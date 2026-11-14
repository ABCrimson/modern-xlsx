import { describe, expect, it } from 'vitest';
import {
  columnToLetter,
  decodeCellRef,
  decodeRange,
  decodeRow,
  encodeCellRef,
  encodeRange,
  encodeRow,
  letterToColumn,
  ModernXlsxError,
  splitCellRef,
} from '../src/index.js';

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
    expect(() => decodeRange('A1:')).toThrow('Invalid range');
  });

  it('throws on a range with more than one colon', () => {
    expect(() => decodeRange('A1:B2:C3')).toThrow('Invalid range');
  });

  it('roundtrips', () => {
    const s = { row: 5, col: 3 };
    const e = { row: 20, col: 10 };
    expect(decodeRange(encodeRange(s, e))).toEqual({ start: s, end: e });
  });
});

describe('bounds (XFD / 1048576)', () => {
  function codeOf(fn: () => unknown): { isModern: boolean; code: unknown; message: string } {
    try {
      fn();
    } catch (e) {
      return {
        isModern: e instanceof ModernXlsxError,
        code: e instanceof ModernXlsxError ? e.code : undefined,
        message: e instanceof Error ? e.message : '',
      };
    }
    throw new Error('expected a throw');
  }

  it.each([
    ['decodeCellRef A0', () => decodeCellRef('A0')],
    ['decodeCellRef XFE1', () => decodeCellRef('XFE1')],
    ['decodeCellRef A1048577', () => decodeCellRef('A1048577')],
    ['decodeCellRef empty', () => decodeCellRef('')],
    ['letterToColumn empty', () => letterToColumn('')],
    ['letterToColumn a1', () => letterToColumn('a1')],
    ['letterToColumn XFE', () => letterToColumn('XFE')],
    ['columnToLetter -1', () => columnToLetter(-1)],
    ['columnToLetter NaN', () => columnToLetter(Number.NaN)],
    ['columnToLetter 1.5', () => columnToLetter(1.5)],
    ['columnToLetter 16384', () => columnToLetter(16384)],
    ['encodeCellRef -1,0', () => encodeCellRef(-1, 0)],
    ['encodeCellRef 1048576,0', () => encodeCellRef(1048576, 0)],
  ])('%s throws ModernXlsxError INVALID_CELL_REF', (_name, fn) => {
    const r = codeOf(fn);
    expect(r.isModern).toBe(true);
    expect(r.code).toBe('INVALID_CELL_REF');
  });

  it('decodeCellRef messages start with "Invalid cell reference"', () => {
    for (const ref of ['A0', 'XFE1', 'A1048577', '']) {
      expect(codeOf(() => decodeCellRef(ref)).message).toMatch(/^Invalid cell reference/);
    }
  });

  it('accepts the boundaries', () => {
    expect(decodeCellRef('XFD1048576')).toEqual({ col: 16383, row: 1048575 });
    expect(columnToLetter(16383)).toBe('XFD');
    expect(letterToColumn('xfd')).toBe(16383);
    expect(letterToColumn('XFD')).toBe(16383);
    expect(decodeCellRef('$B$2')).toEqual({ col: 1, row: 1 });
  });
});

describe('extended', () => {
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
        col: 'A',
        row: '1',
        absCol: false,
        absRow: false,
      });
    });

    it('splits fully absolute reference', () => {
      expect(splitCellRef('$A$1')).toEqual({
        col: 'A',
        row: '1',
        absCol: true,
        absRow: true,
      });
    });

    it('splits mixed references', () => {
      expect(splitCellRef('$A1')).toEqual({
        col: 'A',
        row: '1',
        absCol: true,
        absRow: false,
      });
      expect(splitCellRef('A$1')).toEqual({
        col: 'A',
        row: '1',
        absCol: false,
        absRow: true,
      });
    });

    it('handles multi-letter columns', () => {
      expect(splitCellRef('$XFD$1048576')).toEqual({
        col: 'XFD',
        row: '1048576',
        absCol: true,
        absRow: true,
      });
    });

    it('throws on invalid ref', () => {
      expect(() => splitCellRef('')).toThrow();
      expect(() => splitCellRef('123')).toThrow();
    });
  });
});
