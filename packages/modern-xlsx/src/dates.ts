/**
 * Date serial number utilities for Excel date systems.
 *
 * Excel stores dates as serial numbers: the number of days since
 * the epoch (1899-12-30 for date1900, 1904-01-01 for date1904).
 *
 * The date1900 system has a known bug (inherited from Lotus 1-2-3)
 * that treats 1900 as a leap year. Serial number 60 is Feb 29, 1900
 * (which doesn't exist). We replicate this behavior for compatibility.
 */

import type { DateSystem } from './types.js';

/** Epoch for the 1900 date system: Dec 31, 1899 (serial 1 = Jan 1, 1900). */
const EPOCH_1900 = Date.UTC(1899, 11, 31); // 1899-12-31

/** Epoch for the 1904 date system: Jan 1, 1904 (serial 0). */
const EPOCH_1904 = Date.UTC(1904, 0, 1); // 1904-01-01

/** Milliseconds per day. */
const MS_PER_DAY = 86_400_000;

/** Known date number format IDs (built-in Excel formats). */
const DATE_FORMAT_IDS = new Set([
  14,
  15,
  16,
  17,
  18,
  19,
  20,
  21,
  22, // Standard date/time
  27,
  28,
  29,
  30,
  31,
  32,
  33,
  34,
  35,
  36, // CJK date formats
  45,
  46,
  47, // Time-only formats
  50,
  51,
  52,
  53,
  54,
  55,
  56,
  57,
  58, // More CJK
]);

/** Duck-typed Temporal.PlainDate / PlainDateTime. */
interface TemporalLike {
  year: number;
  month: number;
  day: number;
  hour?: number;
  minute?: number;
  second?: number;
  millisecond?: number;
}

/** Type guard for duck-typed Temporal.PlainDate / PlainDateTime objects. */
function isTemporalLike(value: unknown): value is TemporalLike {
  return (
    typeof value === 'object' &&
    value !== null &&
    'year' in value &&
    'month' in value &&
    'day' in value
  );
}

function toUtcMs(input: Date | TemporalLike): number {
  if (input instanceof Date) {
    return Date.UTC(
      input.getUTCFullYear(),
      input.getUTCMonth(),
      input.getUTCDate(),
      input.getUTCHours(),
      input.getUTCMinutes(),
      input.getUTCSeconds(),
      input.getUTCMilliseconds(),
    );
  }
  return Date.UTC(
    input.year,
    input.month - 1,
    input.day,
    input.hour ?? 0,
    input.minute ?? 0,
    input.second ?? 0,
    input.millisecond ?? 0,
  );
}

/**
 * Convert a Date or Temporal-like object to an Excel serial number.
 *
 * Accepts a JavaScript Date, Temporal.PlainDate, Temporal.PlainDateTime,
 * or any object with `year`, `month`, `day` properties.
 *
 * @param date - Date or Temporal-like object
 * @param system - Date system to use (default: date1900)
 * @returns Serial number (fractional part represents time of day)
 */
export function dateToSerial(date: Date | TemporalLike, system: DateSystem = 'date1900'): number {
  const utcMs = toUtcMs(date);

  if (system === 'date1904') {
    return (utcMs - EPOCH_1904) / MS_PER_DAY;
  }

  let serial = (utcMs - EPOCH_1900) / MS_PER_DAY;
  // Lotus 1-2-3 bug: dates after Feb 28, 1900 are off by 1
  if (serial >= 60) {
    serial += 1;
  }
  return serial;
}

export { isTemporalLike };

/**
 * Convert an Excel serial number to a Date (UTC).
 *
 * @param serial - Excel serial number
 * @param system - Date system to use (default: date1900)
 * @returns JavaScript Date (UTC)
 */
export function serialToDate(serial: number, system: DateSystem = 'date1900'): Date {
  if (system === 'date1904') {
    return new Date(EPOCH_1904 + serial * MS_PER_DAY);
  }

  let adjusted = serial;
  // Lotus 1-2-3 bug compensation
  if (adjusted > 60) {
    adjusted -= 1;
  } else if (adjusted === 60) {
    // Feb 29, 1900 doesn't exist — return Feb 28
    adjusted = 59;
  }
  return new Date(EPOCH_1900 + adjusted * MS_PER_DAY);
}

/**
 * Check whether a given number format ID represents a date format.
 *
 * Only checks against built-in date format IDs. For custom formats,
 * use `isDateFormatCode()` which analyzes the format string.
 */
export function isDateFormatId(numFmtId: number): boolean {
  return DATE_FORMAT_IDS.has(numFmtId);
}

/**
 * Heuristic check whether a format code string represents a date/time format.
 * Looks for date/time tokens (y, m, d, h, s, AM/PM) while ignoring
 * quoted strings and escaped characters.
 */
export function isDateFormatCode(formatCode: string): boolean {
  let inQuote = false;
  let i = 0;
  while (i < formatCode.length) {
    const ch = formatCode.charAt(i);

    if (ch === '"') {
      inQuote = !inQuote;
      i++;
      continue;
    }
    if (inQuote) {
      i++;
      continue;
    }
    if (ch === '\\') {
      i += 2; // skip escaped char
      continue;
    }

    // Date/time tokens (case-insensitive)
    const lower = ch.toLowerCase();
    if (lower === 'y' || lower === 'd' || lower === 'h' || lower === 's') {
      return true;
    }
    // 'm' is ambiguous (month vs minutes) — but if present, likely date
    if (lower === 'm') {
      return true;
    }
    // AM/PM marker
    if (formatCode.slice(i, i + 5).toUpperCase() === 'AM/PM') {
      return true;
    }

    i++;
  }
  return false;
}
