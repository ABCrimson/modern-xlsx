/**
 * Number format renderer — converts cell values to formatted strings
 * using Excel-compatible format codes.
 *
 * Equivalent to SheetJS SSF.format(). Handles built-in formats
 * (General, 0, 0.00, #,##0, dates, percentages, fractions, scientific)
 * and custom format strings.
 */

import { serialToDate } from './dates.js';
import type { DateSystem } from './types.js';

/** Built-in Excel format codes (ECMA-376 18.8.30). */
const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

export interface FormatCellOptions {
  /** Date system for serial-to-date conversion. */
  dateSystem?: DateSystem;
}

/**
 * Format a cell value using an Excel number format string.
 *
 * @param value - The raw cell value (number or string)
 * @param format - The format code string or built-in format ID
 * @param opts - Optional settings
 * @returns Formatted string
 */
export function formatCell(
  value: string | number | boolean | null,
  format: string | number,
  opts?: FormatCellOptions,
): string {
  if (value === null || value === undefined) return '';

  // Excel always renders booleans as uppercase TRUE/FALSE.
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';

  const formatCode = typeof format === 'number' ? (BUILTIN_FORMATS[format] ?? 'General') : format;

  if (formatCode === 'General' || formatCode === '' || formatCode === '@') {
    return String(value);
  }

  const numVal = typeof value === 'number' ? value : Number.parseFloat(String(value));
  if (Number.isNaN(numVal)) return String(value);

  return dispatchFormat(numVal, formatCode, opts?.dateSystem ?? 'date1900');
}

/** Get the format code string for a built-in format ID. */
export function getBuiltinFormat(id: number): string | undefined {
  return BUILTIN_FORMATS[id];
}

// ---------------------------------------------------------------------------
// Dispatch
// ---------------------------------------------------------------------------

function dispatchFormat(numVal: number, code: string, system: DateSystem): string {
  if (isDateFormat(code)) return formatDate(numVal, code, system);
  if (code.includes('%')) return formatPercentage(numVal, code);
  if (code.includes('E+') || code.includes('E-') || code.includes('e+')) {
    return formatScientific(numVal, code);
  }
  if (code.includes('?/') || code.includes('#/')) return formatFraction(numVal);
  return formatNumber(numVal, code);
}

// ---------------------------------------------------------------------------
// Date format detection
// ---------------------------------------------------------------------------

function isDateFormat(code: string): boolean {
  let inQuote = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code.charAt(i);
    if (ch === '"') {
      inQuote = !inQuote;
      continue;
    }
    if (inQuote) continue;
    if (ch === '\\') {
      i++;
      continue;
    }
    const lower = ch.toLowerCase();
    if (lower === 'y' || lower === 'd' || lower === 'h' || lower === 's' || lower === 'm') {
      return true;
    }
    if (code.slice(i, i + 5).toUpperCase() === 'AM/PM') return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Date formatting
// ---------------------------------------------------------------------------

const MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

const MONTH_SHORT = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];

function formatDate(serial: number, code: string, system: DateSystem): string {
  const date = serialToDate(serial, system);
  const parts = extractDateParts(date, code);
  return applyDateTokens(code, parts);
}

interface DateParts {
  year: number;
  month: number;
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
  ampm: string;
}

function extractDateParts(date: Date, code: string): DateParts {
  const hours24 = date.getUTCHours();
  const isAmPm = /AM\/PM/i.test(code);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hours: isAmPm ? hours24 % 12 || 12 : hours24,
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
    ampm: hours24 < 12 ? 'AM' : 'PM',
  };
}

/**
 * Single-pass tokenizer for date format strings. Avoids the sequential
 * regex-replace approach which can corrupt output when a substituted value
 * contains characters matching later patterns (e.g. year "1900" → yy matches "00").
 */
function applyDateTokens(code: string, p: DateParts): string {
  let result = '';
  let i = 0;
  const s = code;

  while (i < s.length) {
    // Skip color brackets [Red] etc.
    if (s[i] === '[') {
      const close = s.indexOf(']', i);
      if (close !== -1) {
        i = close + 1;
        continue;
      }
    }

    // Quoted literal strings
    if (s[i] === '"') {
      const end = s.indexOf('"', i + 1);
      if (end !== -1) {
        result += s.slice(i + 1, end);
        i = end + 1;
        continue;
      }
    }

    // Escape sequence
    if (s[i] === '\\' && i + 1 < s.length) {
      result += s[i + 1];
      i += 2;
      continue;
    }

    const lower = s.slice(i).toLowerCase();

    // AM/PM
    if (lower.startsWith('am/pm')) {
      result += p.ampm;
      i += 5;
      continue;
    }

    // Year tokens
    if (lower.startsWith('yyyy')) {
      result += String(p.year);
      i += 4;
      continue;
    }
    if (lower.startsWith('yy')) {
      result += String(p.year).slice(-2);
      i += 2;
      continue;
    }

    // Month tokens (mmmm, mmm, mm, m)
    const ch = s[i]?.toLowerCase();
    if (ch === 'm') {
      const run = countRun(s, i, 'm');
      if (run >= 4) {
        result += MONTH_NAMES[p.month - 1] ?? '';
        i += run;
        continue;
      }
      if (run === 3) {
        result += MONTH_SHORT[p.month - 1] ?? '';
        i += 3;
        continue;
      }
      if (run === 2) {
        result += String(p.month).padStart(2, '0');
        i += 2;
        continue;
      }
      // single m
      result += String(p.month);
      i += 1;
      continue;
    }

    // Day tokens
    if (ch === 'd') {
      const run = countRun(s, i, 'd');
      if (run >= 2) {
        result += String(p.day).padStart(2, '0');
        i += run;
        continue;
      }
      result += String(p.day);
      i += 1;
      continue;
    }

    // Hour tokens
    if (ch === 'h') {
      const run = countRun(s, i, 'h');
      if (run >= 2) {
        result += String(p.hours).padStart(2, '0');
        i += run;
        continue;
      }
      result += String(p.hours);
      i += 1;
      continue;
    }

    // Second tokens
    if (ch === 's') {
      const run = countRun(s, i, 's');
      if (run >= 2) {
        result += String(p.seconds).padStart(2, '0');
        i += run;
        continue;
      }
      result += String(p.seconds);
      i += 1;
      continue;
    }

    // Pass through everything else (colons, slashes, spaces, etc.)
    result += s[i];
    i += 1;
  }

  return result;
}

/** Count consecutive occurrences of a character (case-insensitive). */
function countRun(s: string, start: number, ch: string): number {
  const lower = ch.toLowerCase();
  let count = 0;
  while (start + count < s.length && s[start + count]?.toLowerCase() === lower) {
    count++;
  }
  return count;
}

// ---------------------------------------------------------------------------
// Percentage, scientific, fraction
// ---------------------------------------------------------------------------

function formatPercentage(value: number, code: string): string {
  const pctValue = value * 100;
  const decimals = countDecimals(code, /\.(\d+|0+)%/);
  return `${pctValue.toFixed(decimals)}%`;
}

function formatScientific(value: number, code: string): string {
  const decimals = countDecimals(code, /\.(\d+|0+)[Ee]/) || 2;
  return value.toExponential(decimals).toUpperCase();
}

function formatFraction(value: number): string {
  const whole = Math.trunc(value);
  const frac = Math.abs(value - whole);
  if (frac === 0) return String(whole);

  const { num, denom } = approximateFraction(frac, 99);
  if (num === 0) return String(whole);
  if (whole === 0) return `${num}/${denom}`;
  return `${whole} ${num}/${denom}`;
}

function approximateFraction(frac: number, maxDenom: number): { num: number; denom: number } {
  let bestNum = 0;
  let bestDenom = 1;
  let bestError = frac;

  for (let d = 1; d <= maxDenom; d++) {
    const n = Math.round(frac * d);
    const error = Math.abs(frac - n / d);
    if (error < bestError) {
      bestError = error;
      bestNum = n;
      bestDenom = d;
    }
    if (bestError === 0) break;
  }

  return { num: bestNum, denom: bestDenom };
}

// ---------------------------------------------------------------------------
// Number formatting
// ---------------------------------------------------------------------------

function formatNumber(value: number, code: string): string {
  const { section, value: resolvedValue } = resolveSection(code, value);
  if (section === 'General') return String(resolvedValue);

  const cleaned = section.replace(/\[(?:Red|Blue|Green|Yellow|Magenta|Cyan|White|Black)\]/gi, '');
  return applyNumberFormat(resolvedValue, cleaned);
}

function resolveSection(code: string, value: number): { section: string; value: number } {
  const sections = splitSections(code);

  if (sections.length >= 3 && value === 0) {
    return { section: sections[2] ?? sections[0] ?? 'General', value };
  }
  if (sections.length >= 2 && value < 0) {
    return { section: sections[1] ?? sections[0] ?? 'General', value: Math.abs(value) };
  }
  return { section: sections[0] ?? 'General', value };
}

function applyNumberFormat(value: number, section: string): string {
  // Strip quoted literals, replace with placeholders
  const literals: string[] = [];
  const stripped = section.replace(/"([^"]*)"/g, (_, text: string) => {
    literals.push(text);
    return `<<${literals.length - 1}>>`;
  });

  const hasComma = stripped.includes('#,') || stripped.includes('0,');
  const decimals = countDecimals(stripped, /\.(0+|#+)/);

  let formatted: string;
  if (hasComma) {
    formatted = value.toLocaleString('en-US', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping: true,
    });
  } else {
    formatted = value.toFixed(decimals);
  }

  // Reinsert literals
  return formatted.replace(/<<(\d+)>>/g, (_, idx: string) => literals[Number(idx)] ?? '');
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function countDecimals(code: string, pattern: RegExp): number {
  const match = code.match(pattern);
  return match ? (match[1]?.length ?? 0) : 0;
}

function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  let inQuote = false;

  for (const ch of code) {
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
    } else if (ch === ';' && !inQuote) {
      sections.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  sections.push(current);
  return sections;
}
