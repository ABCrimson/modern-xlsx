import { describe, expect, it } from 'vitest';
import {
  dateToSerial,
  isDateFormatCode,
  isDateFormatId,
  serialToDate,
} from '../src/format/dates.js';
import { getBuiltinFormat } from '../src/format/format-cell.js';

describe('dateToSerial', () => {
  it('converts Jan 1, 1900 (serial 1)', () => {
    const date = new Date(Date.UTC(1900, 0, 1));
    expect(dateToSerial(date, 'date1900')).toBe(1);
  });

  it('converts Jan 1, 2000', () => {
    const date = new Date(Date.UTC(2000, 0, 1));
    expect(dateToSerial(date, 'date1900')).toBe(36526);
  });

  it('handles Lotus 1-2-3 leap year bug', () => {
    // Feb 28, 1900 = serial 59
    const feb28 = new Date(Date.UTC(1900, 1, 28));
    expect(dateToSerial(feb28, 'date1900')).toBe(59);

    // Mar 1, 1900 = serial 61 (60 is the phantom Feb 29)
    const mar1 = new Date(Date.UTC(1900, 2, 1));
    expect(dateToSerial(mar1, 'date1900')).toBe(61);
  });

  it('works with date1904 system', () => {
    const date = new Date(Date.UTC(1904, 0, 1));
    expect(dateToSerial(date, 'date1904')).toBe(0);

    const date2 = new Date(Date.UTC(1904, 0, 2));
    expect(dateToSerial(date2, 'date1904')).toBe(1);
  });
});

describe('serialToDate', () => {
  it('converts serial 1 to Jan 1, 1900', () => {
    const date = serialToDate(1, 'date1900');
    expect(date.getUTCFullYear()).toBe(1900);
    expect(date.getUTCMonth()).toBe(0);
    expect(date.getUTCDate()).toBe(1);
  });

  it('converts serial 36526 to Jan 1, 2000', () => {
    const date = serialToDate(36526, 'date1900');
    expect(date.getUTCFullYear()).toBe(2000);
    expect(date.getUTCMonth()).toBe(0);
    expect(date.getUTCDate()).toBe(1);
  });

  it('handles phantom Feb 29, 1900 (serial 60)', () => {
    // Serial 60 is the phantom Feb 29, 1900 — we return Feb 28
    const date = serialToDate(60, 'date1900');
    expect(date.getUTCMonth()).toBe(1); // February
    expect(date.getUTCDate()).toBe(28);
  });

  it('roundtrips with dateToSerial for date1900', () => {
    const original = new Date(Date.UTC(2024, 5, 15)); // June 15, 2024
    const serial = dateToSerial(original, 'date1900');
    const result = serialToDate(serial, 'date1900');
    expect(result.getUTCFullYear()).toBe(2024);
    expect(result.getUTCMonth()).toBe(5);
    expect(result.getUTCDate()).toBe(15);
  });

  it('roundtrips with dateToSerial for date1904', () => {
    const original = new Date(Date.UTC(2024, 5, 15));
    const serial = dateToSerial(original, 'date1904');
    const result = serialToDate(serial, 'date1904');
    expect(result.getUTCFullYear()).toBe(2024);
    expect(result.getUTCMonth()).toBe(5);
    expect(result.getUTCDate()).toBe(15);
  });
});

describe('isDateFormatId', () => {
  it('returns true for known date format IDs', () => {
    expect(isDateFormatId(14)).toBe(true);
    expect(isDateFormatId(22)).toBe(true);
    expect(isDateFormatId(45)).toBe(true);
  });

  it('returns false for non-date format IDs', () => {
    expect(isDateFormatId(0)).toBe(false); // General
    expect(isDateFormatId(1)).toBe(false); // 0
    expect(isDateFormatId(164)).toBe(false); // Custom
  });
});

describe('isDateFormatCode', () => {
  it('detects date formats', () => {
    expect(isDateFormatCode('yyyy-mm-dd')).toBe(true);
    expect(isDateFormatCode('dd/mm/yyyy')).toBe(true);
    expect(isDateFormatCode('hh:mm:ss')).toBe(true);
    expect(isDateFormatCode('yyyy-mm-dd hh:mm')).toBe(true);
  });

  it('detects AM/PM format', () => {
    expect(isDateFormatCode('h:mm AM/PM')).toBe(true);
  });

  it('rejects non-date formats', () => {
    expect(isDateFormatCode('#,##0.00')).toBe(false);
    expect(isDateFormatCode('0%')).toBe(false);
    expect(isDateFormatCode('General')).toBe(false);
  });

  it('ignores date chars inside quoted strings', () => {
    expect(isDateFormatCode('"Today is" yyyy')).toBe(true);
    // A format that is ONLY quoted text
    expect(isDateFormatCode('"yyyy"')).toBe(false);
  });

  it('ignores escaped characters', () => {
    expect(isDateFormatCode('\\y#,##0')).toBe(false);
  });
});

describe('isDateFormatCode bracket directives', () => {
  it.each([
    '[Red]#,##0',
    '#,##0.00;[Red](#,##0.00)',
    '[Cyan]0',
    '[White]0',
    '[Red]0.00',
    '0.0E+00',
    '[Color10]0',
    '[<-5]0',
    '[Red',
  ])('is not a date: %s', (code) => {
    expect(isDateFormatCode(code)).toBe(false);
  });

  it('is not a date for built-ins 38 and 40', () => {
    expect(isDateFormatCode(getBuiltinFormat(38) as string)).toBe(false);
    expect(isDateFormatCode(getBuiltinFormat(40) as string)).toBe(false);
  });

  it.each(['[h]:mm', '[hh]:mm', '[mm]:ss', '[ss]', '[$-409]m/d/yyyy'])('is a date: %s', (code) => {
    expect(isDateFormatCode(code)).toBe(true);
  });

  it('skips pad and fill characters', () => {
    expect(isDateFormatCode('0_d')).toBe(false);
    expect(isDateFormatCode('0*d')).toBe(false);
  });
});

describe('serialToDate millisecond rounding', () => {
  it('rounds float error to whole seconds', () => {
    expect(serialToDate(45000 + 7 / 86400).toISOString()).toBe('2023-03-15T00:00:07.000Z');
    expect(serialToDate(1000 + 7 / 86400, 'date1904').getUTCMilliseconds()).toBe(0);
    expect(serialToDate(45000 + 7 / 86400, 'date1904').toISOString()).toBe(
      '2027-03-16T00:00:07.000Z',
    );
  });

  it('has zero milliseconds for every second of a day', () => {
    for (let s = 0; s < 86400; s++) {
      expect(serialToDate(45000 + s / 86400).getUTCMilliseconds()).toBe(0);
    }
  });
});

describe('dates edge cases', () => {
  it('converts small and fractional serials', () => {
    expect(serialToDate(0.5).toISOString()).toBe('1899-12-31T12:00:00.000Z');
    expect(serialToDate(0).toISOString()).toBe('1899-12-31T00:00:00.000Z');
    expect(serialToDate(45000.5).toISOString()).toBe('2023-03-15T12:00:00.000Z');
  });

  it('returns NaN for an invalid Date', () => {
    expect(dateToSerial(new Date(Number.NaN))).toBeNaN();
  });

  it('matches the built-in date format id set', () => {
    const ids = new Set([
      ...Array.from({ length: 9 }, (_, i) => 14 + i),
      ...Array.from({ length: 10 }, (_, i) => 27 + i),
      45,
      46,
      47,
      ...Array.from({ length: 9 }, (_, i) => 50 + i),
    ]);
    for (let n = 0; n <= 200; n++) {
      expect(isDateFormatId(n)).toBe(ids.has(n));
    }
  });

  it('does not remap years 0-99 to 19xx', () => {
    const expected = dateToSerial(new Date('0050-01-01T00:00:00Z'));
    expect(dateToSerial({ year: 50, month: 1, day: 1 })).toBe(expected);
    expect(expected).not.toBe(dateToSerial(new Date('1950-01-01T00:00:00Z')));
    expect(expected).not.toBe(18264);
  });
});
