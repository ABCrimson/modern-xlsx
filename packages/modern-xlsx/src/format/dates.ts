/**
 * Date serial number utilities for Excel date systems.
 *
 * Excel stores dates as serial numbers: the number of days since an epoch.
 *
 * date1900: serial 1 is 1900-01-01. Excel inherits a Lotus 1-2-3 bug that
 * treats 1900 as a leap year, so serial 60 is the phantom 1900-02-29. No
 * `Date` maps to it: `dateToSerial` never returns 60 and `serialToDate(60)`
 * returns 1900-02-28. From serial 61 (1900-03-01) onwards the effective
 * epoch is 1899-12-30.
 *
 * date1904: serial 0 is 1904-01-01.
 */

import type { DateSystem } from '../model/types.js';

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

/** Inner text of an elapsed-time directive: `[h]`, `[hh]`, `[mm]`, `[ss]`, ... */
const ELAPSED_DIRECTIVE = /^(?:h+|m+|s+)$/i;

/**
 * A calendar date with an optional time of day, typed structurally.
 *
 * Native `Temporal.PlainDate` and `Temporal.PlainDateTime` satisfy it, and so
 * does any plain object with numeric `year`, `month` (1-12) and `day` fields.
 * Omitted time fields count as 0.
 */
export interface TemporalLike {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour?: number;
  readonly minute?: number;
  readonly second?: number;
  readonly millisecond?: number;
}

/**
 * Type guard for {@link TemporalLike} values: native `Temporal.PlainDate` /
 * `Temporal.PlainDateTime` or any object with `year`, `month` and `day`.
 */
export function isTemporalLike(value: unknown): value is TemporalLike {
  return (
    typeof value === 'object' &&
    value !== null &&
    'year' in value &&
    'month' in value &&
    'day' in value
  );
}

/**
 * Milliseconds since the Unix epoch (UTC) of a Date or Temporal-like object.
 * The Temporal branch avoids `Date.UTC`, which maps years 0-99 to 1900-1999.
 */
function toUtcMs(input: Date | TemporalLike): number {
  if (input instanceof Date) {
    return input.getTime();
  }
  const d = new Date(0);
  d.setUTCFullYear(input.year, input.month - 1, input.day);
  d.setUTCHours(input.hour ?? 0, input.minute ?? 0, input.second ?? 0, input.millisecond ?? 0);
  return d.getTime();
}

/**
 * Convert a Date or Temporal-like object to an Excel serial number.
 *
 * Accepts a JavaScript Date or a {@link TemporalLike}: a native
 * `Temporal.PlainDate` or `Temporal.PlainDateTime`, or any object with
 * `year`, `month`, `day` properties.
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

/**
 * Convert an Excel serial number to a Date (UTC).
 *
 * @param serial - Excel serial number
 * @param system - Date system to use (default: date1900)
 * @returns JavaScript Date (UTC), rounded to the nearest millisecond so
 *   whole-second serials do not land on `.999`
 */
export function serialToDate(serial: number, system: DateSystem = 'date1900'): Date {
  if (system === 'date1904') {
    return new Date(EPOCH_1904 + Math.round(serial * MS_PER_DAY));
  }

  let adjusted = serial;
  // Lotus 1-2-3 bug compensation
  if (adjusted > 60) {
    adjusted -= 1;
  } else if (adjusted === 60) {
    // Feb 29, 1900 doesn't exist — return Feb 28
    adjusted = 59;
  }
  return new Date(EPOCH_1900 + Math.round(adjusted * MS_PER_DAY));
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

/** Date/time token letters; `m` is ambiguous (month vs minutes) and counts as a date. */
const DATE_LETTER = /[dhmsy]/i;

/** True when a date/time letter or the `AM/PM` or `A/P` marker starts at `code[i]`. */
function isDateTokenAt(code: string, i: number): boolean {
  return (
    DATE_LETTER.test(code.charAt(i)) ||
    code.slice(i, i + 5).toUpperCase() === 'AM/PM' ||
    code.slice(i, i + 3).toUpperCase() === 'A/P'
  );
}

/**
 * Scan a bracketed directive opening at `code[open]`. Returns `false` when
 * there is no closing `]` (ends the scan), `true` for an elapsed-time form
 * (`[h]`, `[mm]`, `[ss]`), else the index just past the `]`.
 */
function scanBracket(code: string, open: number): number | boolean {
  const close = code.indexOf(']', open + 1);
  if (close === -1) {
    return false;
  }
  return ELAPSED_DIRECTIVE.test(code.slice(open + 1, close)) ? true : close + 1;
}

/**
 * Heuristic check whether a format code string represents a date/time format.
 * Looks for date/time tokens (y, m, d, h, s, AM/PM, A/P) while ignoring
 * quoted strings, escaped characters, pad/fill characters (`_x`, `*x`) and
 * bracketed directives such as `[Red]`, `[Color10]`, `[$-409]` or `[<-5]`.
 * Only the elapsed-time forms `[h]`, `[mm]`, `[ss]` (any repeat count) count
 * as date/time. An unterminated `[` ends the scan with `false`.
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
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2; // skip escaped char, or the pad/fill character
      continue;
    }
    if (ch === '[') {
      const next = scanBracket(formatCode, i);
      if (typeof next === 'boolean') {
        return next;
      }
      i = next;
      continue;
    }

    // Date/time tokens (case-insensitive) and the AM/PM marker; 'm' is
    // ambiguous (month vs minutes) but if present, likely date
    if (isDateTokenAt(formatCode, i)) {
      return true;
    }

    i++;
  }
  return false;
}
