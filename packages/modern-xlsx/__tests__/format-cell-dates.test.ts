import { describe, expect, it } from 'vitest';
import { type FormatCellOptions, formatCell, formatCellRich } from '../src/index.js';

/**
 * Date and time display. Every expected string is what Excel 16.0 (build
 * 20527) shows, both as cell display and as `TEXT()`, checked on 2026-11-11.
 * Where SheetJS SSF 0.20.3 agrees (the spec rows) it gave the same value.
 */
type Case = readonly [
  value: number,
  format: string | number,
  expected: string,
  opts?: FormatCellOptions,
];

/** `m`/`mm` is minutes after an hour or before a seconds token; `mmm`+ stay month names. */
const MINUTE_CASES: readonly Case[] = [
  [0.5 + 1 / 1440, 'h:mm', '12:01'],
  [0.5, 'h:mm', '12:00'],
  [45000.5 + 30 / 1440, 22, '3/15/23 12:30'],
  [45000.5 + 61 / 86400, 45, '01:01'],
  [45000 + 7 / 86400, 'h:mm:ss', '0:00:07'],
  [45000 + 7 / 86400, 'ss', '07'],
  [45000.5 + 61 / 86400, 18, '12:01 PM'],
  [45000.5 + 61 / 86400, 19, '12:01:01 PM'],
  [45000.5 + 61 / 86400, 20, '12:01'],
  [45000.5 + 61 / 86400, 21, '12:01:01'],
  [45000.5 + 61 / 86400, 'h:mmm', '12:Mar'],
  [45000.5 + 61 / 86400, 'd mm', '15 03'],
  [45000.5 + 61 / 86400, 'h "x" mm', '12 x 01'],
  [45000.5 + 61 / 86400, 'mm "x" ss', '01 x 01'],
  [45000.5 + 61 / 86400, 'h m', '12 1'],
  [45000.5 + 61 / 86400, 'm s', '1 1'],
  [45000.5 + 61 / 86400, 'h AM/PM mm', '12 PM 01'],
  [45000.5 + 61 / 86400, '[h] mm', '1080012 01'],
  [45000.5 + 61 / 86400, '[m] mm', '64800721 03'],
  // An hour or a seconds token (before any minutes) leaves the next m/mm as
  // minutes across other tokens; once minutes are shown, later m/mm are months.
  [45000.5 + 61 / 86400, 'h yyyy mm', '12 2023 01'],
  [45000.5 + 61 / 86400, 'h d mm', '12 15 01'],
  [45000.5 + 61 / 86400, 'h mm dd mm', '12 01 15 03'],
  [45000.5 + 61 / 86400, 'h mmm mm', '12 Mar 01'],
  [45000.5 + 61 / 86400, 'h:mm mm', '12:01 03'],
  [45000.5 + 61 / 86400, 's mm', '1 01'],
  [45000.5 + 61 / 86400, 's d mm', '1 15 01'],
  [45000.5 + 61 / 86400, 'h:mm:ss mm', '12:01:01 03'],
  [45000.5 + 61 / 86400, 'mm s mm', '01 1 03'],
  [45000.5 + 61 / 86400, 's mm h mm', '1 01 12 01'],
  [45000.5 + 61 / 86400, '[m]:ss mm', '64800721:01 03'],
  [45000.5 + 61 / 86400, 'ss.0 mm', '01.0 01'],
  [45000.5 + 61 / 86400, 's.0 h mm', '1.0 12 01'],
  [45000.5 + 61 / 86400, 'mm:ss.0 mm', '01:01.0 03'],
  [45000.5 + 61 / 86400, '"a"mm"b"ss', 'a01b01'],
  [45000.5 + 61 / 86400, 'yy h mm ss mm', '23 12 01 01 03'],
  [45000.5 + 61 / 86400, 'AM/PM mm ss', 'PM 01 01'],
  [45000.5 + 61 / 86400, 'hh mm AM/PM', '12 01 PM'],
  [45000.5 + 61 / 86400, 'mm ss.000', '01 01.000'],
  [45000.5 + 61 / 86400, 'm:s', '1:1'],
  // Before a seconds token only literals and AM/PM may sit in between.
  [45000.5 + 61 / 86400, 'mm d ss', '03 15 01'],
  [45000.5 + 61 / 86400, 'mm yyyy ss', '03 2023 01'],
  [45000.5 + 61 / 86400, 'mm mm:ss', '03 01:01'],
  [45000.5 + 61 / 86400, 'mm AM/PM ss', '01 PM 01'],
  [45000.5 + 61 / 86400, 'mm h', '03 12'],
  [45000.5 + 61 / 86400, 'mm', '03'],
];

/** Day names, the month initial, and Excel's 1900 weekday (serials below 60 run a day behind). */
const NAME_CASES: readonly Case[] = [
  [45000, 'dddd', 'Wednesday'],
  [45000, 'ddd', 'Wed'],
  [45000, 'mmmmm', 'M'],
  [45000, 'mmmmmm', 'March'],
  [45296, 'dddd', 'Friday'],
  [45296, 'ddd', 'Fri'],
  [45296, 'mmmmm', 'J'],
  [1, 'dddd', 'Sunday'],
  [59, 'dddd', 'Tuesday'],
  [60, 'dddd', 'Wednesday'],
  [61, 'dddd', 'Thursday'],
  [45427, 'dddd, mmmm d, yyyy', 'Wednesday, May 15, 2024'],
  [45000.5, 'DDDD MMMMM', 'Wednesday M'],
  [45000, 'ddddddd', 'Wednesday'],
  [45000, 'mmmmmmmm', 'March'],
  [0, 'dddd', 'Friday', { dateSystem: 'date1904' }],
];

/** Elapsed [h]/[m]/[s], sub-second '.0', rounding to the shown precision, and the AM/PM marker against a quoted "AM/PM" literal. */
const ELAPSED_CASES: readonly Case[] = [
  [1.5, '[h]:mm:ss', '36:00:00'],
  [1.5, 46, '36:00:00'],
  [1.9999999, '[h]:mm:ss', '48:00:00'],
  [1.5, '[mm]:ss', '2160:00'],
  [0.5, '[m]', '720'],
  [1.5, '[s]', '129600'],
  [0.5000058, 47, '0000.5'],
  [61.7 / 86400, 47, '0101.7'],
  [1.5 + 0.7 / 86400, '[h]:mm:ss.0', '36:00:00.7'],
  [1.5 + 0.7 / 86400, '[h]:mm:ss', '36:00:01'],
  [0.5 + 0.7 / 86400, 'h:mm:ss', '12:00:01'],
  [45000.5 + 0.7 / 86400, 'h:mm', '12:00'],
  [1.5, '[hh]:mm', '36:00'],
  [0.04, '[hhh]:mm', '000:57'],
  [0.5, '[mmm]', '720'],
  [0.5, '[sss]', '43200'],
  [45000.5, '[h]', '1080012'],
  [1.5, '[H]:MM:SS', '36:00:00'],
  [1.123456, 'ss.000', '46.598'],
  [1.123456, 'ss.00', '46.60'],
  [1.123456, 's.0', '46.6'],
  [1.123456, '[s].00', '97066.60'],
  [1.123456, '[m]:ss.00', '1617:46.60'],
  [1.123456, '[h]:mm:ss.000', '26:57:46.598'],
  [1.123456, 'mm:ss.0', '57:46.6'],
  [1.123456, 'h:mm .0', '2:57 .6'],
  [1.123456, 'mm.0', '57.6'],
  [1.123456, 'ss .0', '46 .6'],
  [1.123456, 'h:mm:ss.0;@', '2:57:46.6'],
  [1.123456, 's', '47'],
  [1.123456, 'sss', '47'],
  [1.123456, 'hhh', '02'],
  [1 / 86400, 'ss.000', '01.000'],
  [0.999999, 'h:mm:ss.00', '23:59:59.91'],
  [45000.9999999, 'hh:mm:ss.000', '23:59:59.991'],
  [0.99999999, 'hh:mm:ss.00', '00:00:00.00'],
  [0.99999999, '[h]:mm', '24:00'],
  [0.99999999, '[m]', '1440'],
  // Rounding to the second carries into the date.
  [45000.99999, 'yyyy-mm-dd hh:mm:ss', '2023-03-15 23:59:59'],
  [45000.99999, 'yyyy-mm-dd hh:mm', '2023-03-15 23:59'],
  [45000.99999, 'yyyy-mm-dd', '2023-03-15'],
  [45000.999999, 'yyyy-mm-dd', '2023-03-16'],
  [45000.999999, 'dddd hh:mm:ss', 'Thursday 00:00:00'],
  [45000.999999, 'hh:mm', '00:00'],
  [45000.75, 'yyyy-mm-dd hh:mm:ss AM/PM', '2023-03-15 06:00:00 PM'],
  // A quoted run is literal text, never the AM/PM marker; the unquoted last row is the control.
  [45000.75, 'h "AM/PM"', '18 AM/PM'],
  [45000.75, 'hh:mm "AM/PM"', '18:00 AM/PM'],
  [45000.75, 'h AM/PM', '6 PM'],
];

function check([value, format, expected, opts]: Case): void {
  expect(formatCell(value, format, opts)).toBe(expected);
  expect(formatCellRich(value, format, opts).text).toBe(expected);
}

describe('formatCell date and time codes', () => {
  it.for(MINUTE_CASES)('m as minutes: %s with %s gives %s', (row) => check(row));
  it.for(NAME_CASES)('day and month names: %s with %s gives %s', (row) => check(row));
  it.for(ELAPSED_CASES)('elapsed and sub-seconds: %s with %s gives %s', (row) => check(row));
});
