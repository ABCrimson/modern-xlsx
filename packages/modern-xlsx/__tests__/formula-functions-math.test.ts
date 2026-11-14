import { describe, expect, it } from 'vitest';
import type { CellValue } from '../src/index.js';
import { ev } from './helpers/formula-context.js';

// ---------------------------------------------------------------------------
// SUMIF / AVERAGEIF: sum_range is sized like range from its top-left cell
// ---------------------------------------------------------------------------

describe('SUMIF / AVERAGEIF sum_range sizing', () => {
  const cells: Record<string, CellValue> = {
    'Sheet1!A1': 1,
    'Sheet1!A2': 2,
    'Sheet1!A3': 3,
    'Sheet1!B1': 10,
    'Sheet1!B2': 20,
    'Sheet1!B3': 30,
  };

  it('SUMIF reads a single-cell sum_range at the size of range', () => {
    expect(ev('SUMIF(A1:A3,">0",B1)', cells)).toBe(60);
  });

  it('AVERAGEIF reads a single-cell average_range at the size of range', () => {
    expect(ev('AVERAGEIF(A1:A3,">0",B1)', cells)).toBe(20);
  });

  it('SUMIF extends a shorter sum_range to the size of range', () => {
    expect(ev('SUMIF(A1:A3,">1",B1:B2)', cells)).toBe(50);
  });
});

// ---------------------------------------------------------------------------
// SUMPRODUCT: non-numeric entries, booleans included, count as 0
// ---------------------------------------------------------------------------

describe('SUMPRODUCT non-numeric entries', () => {
  it('treats booleans in an array constant as 0', () => {
    expect(ev('SUMPRODUCT({TRUE,FALSE},{1,1})')).toBe(0);
  });

  it('treats a single boolean array as 0', () => {
    expect(ev('SUMPRODUCT({TRUE,FALSE})')).toBe(0);
  });

  it('treats booleans in a range as 0', () => {
    expect(
      ev('SUMPRODUCT(C1:C2,D1:D2)', {
        'Sheet1!C1': true,
        'Sheet1!C2': 2,
        'Sheet1!D1': 5,
        'Sheet1!D2': 1,
      }),
    ).toBe(2);
  });

  it('treats numeric text in an array constant as 0', () => {
    expect(ev('SUMPRODUCT({"3",2},{1,1})')).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// FLOOR / CEILING: quotients normalised to 15 significant digits
// ---------------------------------------------------------------------------

describe('FLOOR / CEILING decimal significance', () => {
  it('FLOOR(0.3,0.1) is 0.3', () => {
    expect(ev('FLOOR(0.3,0.1)')).toBe(0.3);
  });

  it('FLOOR(0.6,0.2) is 0.6', () => {
    expect(ev('FLOOR(0.6,0.2)')).toBe(0.6);
  });

  it('CEILING(0.6,0.2) is 0.6', () => {
    expect(ev('CEILING(0.6,0.2)')).toBe(0.6);
  });

  it('CEILING(0.7,0.1) is 0.7', () => {
    expect(ev('CEILING(0.7,0.1)')).toBe(0.7);
  });

  it('FLOOR(-2.5,-2) is -2', () => {
    expect(ev('FLOOR(-2.5,-2)')).toBe(-2);
  });

  it('FLOOR(-2.5,2) is -4', () => {
    expect(ev('FLOOR(-2.5,2)')).toBe(-4);
  });
});
