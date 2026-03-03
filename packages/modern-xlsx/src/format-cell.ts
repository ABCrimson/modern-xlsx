/**
 * Number format renderer — converts cell values to formatted strings
 * using Excel-compatible format codes.
 *
 * Equivalent to SheetJS SSF.format(). Handles built-in formats
 * (General, 0, 0.00, #,##0, dates, percentages, fractions, scientific)
 * and custom format strings.
 */

import { isDateFormatCode, serialToDate } from './dates.js';
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

/** Options for the {@link formatCell} function. */
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
  if (isDateFormatCode(code)) return formatDate(numVal, code, system);
  if (code.includes('%')) return formatPercentage(numVal, code);
  if (code.includes('E+') || code.includes('E-') || code.includes('e+')) {
    return formatScientific(numVal, code);
  }
  if (code.includes('?/') || code.includes('#/')) return formatFraction(numVal);
  return formatNumber(numVal, code);
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

// ---------------------------------------------------------------------------
// Token table — data-driven date format tokenizer
// ---------------------------------------------------------------------------

type TokenFormatter = (p: DateParts) => string;

interface DateToken {
  /** Characters to match (case-insensitive). Longest match wins. */
  readonly pattern: string;
  readonly render: TokenFormatter;
}

/**
 * Tokens for run-length date characters (m, d, h, s).
 * Keyed by lowercase character; value maps run length to formatter.
 */
const RUN_TOKENS: Record<string, readonly [number, TokenFormatter][]> = {
  m: [
    [4, (p) => MONTH_NAMES[p.month - 1] ?? ''],
    [3, (p) => MONTH_SHORT[p.month - 1] ?? ''],
    [2, (p) => String(p.month).padStart(2, '0')],
    [1, (p) => String(p.month)],
  ],
  d: [
    [2, (p) => String(p.day).padStart(2, '0')],
    [1, (p) => String(p.day)],
  ],
  h: [
    [2, (p) => String(p.hours).padStart(2, '0')],
    [1, (p) => String(p.hours)],
  ],
  s: [
    [2, (p) => String(p.seconds).padStart(2, '0')],
    [1, (p) => String(p.seconds)],
  ],
};

/** Fixed-string tokens matched before run tokens. Ordered longest-first. */
const FIXED_TOKENS: readonly DateToken[] = [
  { pattern: 'am/pm', render: (p) => p.ampm },
  { pattern: 'yyyy', render: (p) => String(p.year) },
  { pattern: 'yy', render: (p) => String(p.year).slice(-2) },
];

/** Count consecutive case-insensitive occurrences of `ch` starting at `start`. */
function countRun(s: string, start: number, ch: string): number {
  let count = 0;
  while (start + count < s.length && (s[start + count] ?? '').toLowerCase() === ch) {
    count++;
  }
  return count;
}

/** Try to match a fixed-string token at position `i`. */
function matchFixedToken(lower: string): DateToken | undefined {
  for (const token of FIXED_TOKENS) {
    if (lower.startsWith(token.pattern)) return token;
  }
  return undefined;
}

/** Try to match a run-length token at position `i`. Returns [consumed, output] or undefined. */
function matchRunToken(s: string, i: number): [number, TokenFormatter] | undefined {
  const ch = (s[i] ?? '').toLowerCase();
  const entries = RUN_TOKENS[ch];
  if (!entries) return undefined;
  const run = countRun(s, i, ch);
  for (const [minLen, render] of entries) {
    if (run >= minLen) return [Math.max(run, minLen), render];
  }
  return undefined;
}

/** Skip a bracketed sequence like [Red], [Color1], etc. Returns new index or -1. */
function skipBracket(s: string, i: number): number {
  const close = s.indexOf(']', i);
  return close !== -1 ? close + 1 : -1;
}

/** Extract a quoted literal and return [content, newIndex] or undefined. */
function extractQuoted(s: string, i: number): [string, number] | undefined {
  const end = s.indexOf('"', i + 1);
  return end !== -1 ? [s.slice(i + 1, end), end + 1] : undefined;
}

/**
 * Process a single character/token at position `i` in the format string.
 * Returns [charsConsumed, outputText].
 */
function processDateChar(code: string, i: number, p: DateParts): [number, string] {
  const ch = code[i] ?? '';

  // Bracketed directives: [Red], [Color1], etc.
  if (ch === '[') {
    const next = skipBracket(code, i);
    if (next !== -1) return [next - i, ''];
  }

  // Quoted literal strings
  if (ch === '"') {
    const quoted = extractQuoted(code, i);
    if (quoted) return [quoted[1] - i, quoted[0]];
  }

  // Escape sequence
  if (ch === '\\' && i + 1 < code.length) {
    return [2, code[i + 1] ?? ''];
  }

  // Fixed-string tokens (am/pm, yyyy, yy)
  const fixed = matchFixedToken(code.slice(i).toLowerCase());
  if (fixed) return [fixed.pattern.length, fixed.render(p)];

  // Run-length tokens (m, d, h, s)
  const run = matchRunToken(code, i);
  if (run) return [run[0], run[1](p)];

  // Pass through everything else
  return [1, ch];
}

/**
 * Single-pass tokenizer for date format strings.
 *
 * Uses a token table for all date pattern matching — no branching per token type.
 * Avoids the sequential regex-replace approach which can corrupt output when a
 * substituted value contains characters matching later patterns.
 */
function applyDateTokens(code: string, p: DateParts): string {
  let result = '';
  let i = 0;
  while (i < code.length) {
    const [consumed, output] = processDateChar(code, i, p);
    result += output;
    i += consumed;
  }
  return result;
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
