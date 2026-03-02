import { describe, expect, it } from 'vitest';
import { dateToSerial, isDateFormatCode, isDateFormatId, serialToDate } from '../src/dates.js';

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
