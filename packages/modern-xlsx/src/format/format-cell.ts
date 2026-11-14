/**
 * Number format renderer: turns a cell value into the text Excel displays
 * for a number format code.
 *
 * Follows Excel's number-format semantics (ECMA-376 Part 1 §18.8.30 and
 * Microsoft's "Review guidelines for customizing a number format"). The
 * number and date/time paths were validated against Excel 16.0 (build 20527,
 * `TEXT()` and cell display, 2026-11-11) and against SheetJS SSF 0.20.3
 * (`XLSX.SSF.format`) where SSF matches Excel. Covers the built-in formats,
 * up to four sections with conditions and colours, digit templates with
 * literals, grouping and scaling, percent, scientific and engineering
 * notation, fractions, `General`, and date/time codes with day and month
 * names, `AM/PM` and `A/P`, elapsed `[h]`/`[m]`/`[s]` and sub-second
 * `.0`/`.00`/`.000`.
 *
 * Known divergences from Excel: a negative value that rounds to zero keeps
 * its '-' (`0.00` shows -0.001 as `-0.00`, Excel `0.00`); `*x` fill is
 * dropped because there is no column width; a negative date or time, which
 * Excel shows as `####`, renders the time of its absolute value; the 1900
 * system's day 0 and day 60 render as {@link serialToDate} maps them
 * (1899-12-31 and 1900-02-28), where Excel shows 1900-01-00 and 1900-02-29.
 */

import type { DateSystem } from '../model/types.js';
import { isDateFormatCode, isDateFormatId, serialToDate } from './dates.js';
import { formatGeneralNumber } from './number-text.js';

/**
 * Built-in format codes: ECMA-376 Part 1 §18.8.30 (numFmt), with the ids
 * SheetJS SSF 0.20.3 `get_table()` also defines (56 is the zh-TW time format).
 */
const BUILTIN_FORMATS = {
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
  56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
} as const satisfies Record<number, string>;

/** Type-safe lookup for BUILTIN_FORMATS by arbitrary numeric key. */
const builtinFormatLookup = BUILTIN_FORMATS as Record<number, string | undefined>;

/** Runtime-registered format codes; a registration wins over the built-in code for its id. */
const CUSTOM_FORMATS = new Map<number, string>();

/**
 * Register a number format code for `id` at runtime (the SheetJS `SSF.load`
 * equivalent). The registration takes precedence over a built-in code with
 * the same id.
 */
export function loadFormat(formatCode: string, id: number): void {
  CUSTOM_FORMATS.set(id, formatCode);
}

/**
 * Bulk-register format codes at runtime (the SheetJS `SSF.load_table`
 * equivalent); see {@link loadFormat}.
 */
export function loadFormatTable(table: Record<number, string>): void {
  for (const [id, code] of Object.entries(table)) {
    CUSTOM_FORMATS.set(Number(id), code);
  }
}

/** Options for the {@link formatCell} function. */
export interface FormatCellOptions {
  /** Date system for serial-to-date conversion. */
  dateSystem?: DateSystem;
}

/**
 * Format a cell value using an Excel number format string.
 *
 * Same as {@link formatCellRich} without the colour.
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
  return formatCellRich(value, format, opts).text;
}

/** Get the format code string for a built-in format ID. */
export function getBuiltinFormat(id: number): string | undefined {
  return builtinFormatLookup[id];
}

/** Result of rich cell formatting, including optional color metadata. */
export interface FormatCellResult {
  /** The formatted text string. */
  text: string;
  /** Color name from bracket directive (e.g., "Red", "Color3"). */
  color?: string;
}

/**
 * Format a cell value and return rich result with color metadata.
 * Supports conditional sections like `[>100]#,##0;[<=100]0.00`
 * and bracket color codes like `[Red]`, `[Color3]`.
 *
 * A string value is formatted as a number only when its trimmed text is a
 * plain decimal number (`' 12 '`, `'1e3'`); any other text, and every string
 * under `General` or `@`, is returned unchanged. Numbers under `General` or
 * `@` use Excel's General display (`1e21` gives `'1E+21'`).
 */
export function formatCellRich(
  value: string | number | boolean | null,
  format: string | number,
  opts?: FormatCellOptions,
): FormatCellResult {
  if (value === null || value === undefined) return { text: '' };
  // Excel always renders booleans as uppercase TRUE/FALSE.
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };

  const formatCode = resolveFormatCode(format);
  if (isGeneralCode(formatCode)) {
    return { text: typeof value === 'number' ? formatGeneralNumber(value) : value };
  }

  const numVal = typeof value === 'number' ? value : parseNumericText(value);
  if (!Number.isFinite(numVal)) return { text: String(value) };

  return dispatchFormatRich(numVal, formatCode, opts?.dateSystem ?? 'date1900');
}

/**
 * The code for a format id (a {@link loadFormat} registration first), or the
 * code itself. A built-in date id the table has no code for (27-36, 50-55,
 * 57 and 58: locale-specific CJK dates) renders with the code of id 14, as
 * `Cell.numberFormat` resolves it.
 */
function resolveFormatCode(format: string | number): string {
  if (typeof format === 'string') return format;
  return (
    CUSTOM_FORMATS.get(format) ??
    builtinFormatLookup[format] ??
    (isDateFormatId(format) ? BUILTIN_FORMATS[14] : 'General')
  );
}

/** A whole format code that displays numbers as General: '', `@` or `General`. */
function isGeneralCode(code: string): boolean {
  return code === '' || code === '@' || /^general$/i.test(code);
}

/**
 * A plain decimal number: no hex, binary, octal, `Infinity` or trailing text.
 * The point and its fraction digits share one optional group, so the digit
 * runs never overlap and a long digit run with a rejecting tail cannot make
 * the match backtrack quadratically.
 */
const NUMERIC_RE = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Parse a string cell value as a number, or NaN when it is not a plain decimal number. */
function parseNumericText(text: string): number {
  const trimmed = text.trim();
  return NUMERIC_RE.test(trimmed) ? Number(trimmed) : Number.NaN;
}

// ---------------------------------------------------------------------------
// Dispatch: section first (conditions, colours), then the section's kind
// ---------------------------------------------------------------------------

/** `[Red]`, `[color3]`: a colour directive. */
const COLOR_RE = /\[(Red|Blue|Green|Yellow|Magenta|Cyan|White|Black|Color\d{1,2})\]/i;
/**
 * `[<-5]`, `[>=1e3]`, `[<>0]`: an operator and a decimal threshold, with the
 * threshold shaped like {@link NUMERIC_RE} so it cannot backtrack quadratically.
 */
const CONDITION_RE = /\[([<>=!]+)([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\]/;

function buildRichResult(text: string, color: string | undefined): FormatCellResult {
  return color ? { text, color } : { text };
}

function dispatchFormatRich(numVal: number, code: string, system: DateSystem): FormatCellResult {
  const sections = splitSections(code);
  const { section, value } = resolveSectionConditional(sections, numVal);
  const { color, cleaned } = scanDirectives(section);
  return buildRichResult(renderSection(value, cleaned, sections.length, system), color);
}

/** A section's unquoted colour and condition directives. */
interface SectionDirectives {
  /** Name of the first colour directive (`Red`, `Color3`). */
  readonly color: string | undefined;
  /** The first condition directive. */
  readonly condition: Condition | undefined;
  /** The section with every unquoted colour and condition bracket removed. */
  readonly cleaned: string;
}

/** The match of `re` when it spans all of `text`. */
function wholeMatch(re: RegExp, text: string): RegExpExecArray | undefined {
  const match = re.exec(text);
  return match?.index === 0 && match[0].length === text.length ? match : undefined;
}

/**
 * Walk a section as {@link isDateFormatCode} does: a quoted run and the
 * pairs `\x`, `_x` and `*x` are text, and each other `[` opens a directive
 * that runs to the next `]`. Only those brackets are classified as colours
 * ({@link COLOR_RE}) and conditions ({@link CONDITION_RE}), so `"[Red]"0`
 * shows 5 as `[Red]5` in no colour. A `[` without a `]` ends the scan.
 */
function scanDirectives(section: string): SectionDirectives {
  let color: string | undefined;
  let condition: Condition | undefined;
  let cleaned = '';
  let copied = 0;
  let i = 0;
  while (i < section.length) {
    if (section.charAt(i) !== '[') {
      i = skipTextUnit(section, i);
      continue;
    }
    const close = section.indexOf(']', i + 1);
    if (close === -1) break;
    const found = classifyBracket(section.slice(i, close + 1));
    if (found) {
      color ??= found.color;
      condition ??= found.condition;
      cleaned += section.slice(copied, i);
      copied = close + 1;
    }
    i = close + 1;
  }
  return { color, condition, cleaned: cleaned + section.slice(copied) };
}

/**
 * Index just past the text unit at `section[i]`: a quoted run (to the end of
 * the section when unterminated), a `\x`, `_x` or `*x` pair, or one character.
 */
function skipTextUnit(section: string, i: number): number {
  const ch = section.charAt(i);
  if (ch === '"') {
    const close = section.indexOf('"', i + 1);
    return close === -1 ? section.length : close + 1;
  }
  return ch === '\\' || ch === '_' || ch === '*' ? i + 2 : i + 1;
}

/** The colour or condition a whole `[...]` bracket states; `undefined` for any other directive. */
function classifyBracket(bracket: string): Omit<SectionDirectives, 'cleaned'> | undefined {
  const color = wholeMatch(COLOR_RE, bracket)?.[1];
  if (color) return { color, condition: undefined };
  const match = wholeMatch(CONDITION_RE, bracket);
  return match?.[1] && match[2]
    ? { color: undefined, condition: { op: match[1], threshold: Number(match[2]) } }
    : undefined;
}

/**
 * Render one cleaned section. `value` is already signed for the section (see
 * {@link resolveSectionConditional}), so only a value that is still negative
 * prints a '-', in front of everything else (`(0)` shows -3 as `-(3)`).
 */
function renderSection(
  value: number,
  section: string,
  sectionCount: number,
  system: DateSystem,
): string {
  // An empty section of a multi-section code hides the value ('0;;' and ';;;').
  if (section === '') return sectionCount > 1 ? '' : formatGeneralNumber(value);
  if (isDateFormatCode(section)) return formatDate(value, section, system);

  const { tokens, grouping, scale } = classifyCommas(tokenize(section));
  if (tokens.some((t) => t.kind === 'general')) return renderGeneralSection(value, tokens);
  // Without placeholders only the literals show, after the sign ('"x"' shows -5 as '-x').
  if (!tokens.some((t) => t.kind === 'digit')) {
    return (value < 0 ? '-' : '') + renderLiterals(tokens);
  }

  const layout: NumberLayout = { grouping, percent: countPercent(tokens), scale };
  const expAt = tokens.findIndex((t) => t.kind === 'exp');
  if (expAt !== -1) return renderScientific(value, tokens, expAt, layout);
  const slashAt = findFractionSlash(tokens);
  if (slashAt !== -1) return renderFractionSection(value, tokens, slashAt, layout);
  return renderNumber(value, tokens, layout);
}

/** A section's `[op threshold]` condition. */
interface Condition {
  readonly op: string;
  readonly threshold: number;
}

/** The first unquoted condition directive of `section` (see {@link scanDirectives}). */
function parseCondition(section: string): Condition | undefined {
  return scanDirectives(section).condition;
}

/** A condition only negative numbers meet (`[<0]`, `[<-5]`, `[<=-1]`, `[=-5]`). */
function isNegativeOnly({ op, threshold }: Condition): boolean {
  return (op === '<' && threshold <= 0) || ((op === '<=' || op === '=') && threshold < 0);
}

/** `cond;default` whose condition is an ordering against a threshold <= 0. */
function isAbsoluteDefault(conditions: readonly (Condition | undefined)[]): boolean {
  const [first, second] = conditions;
  return (
    conditions.length === 2 &&
    second === undefined &&
    first !== undefined &&
    ['<', '<=', '>', '>='].includes(first.op) &&
    first.threshold <= 0
  );
}

/**
 * Pick the section that applies to `value`, and the value it shows. A plain
 * `pos;neg[;zero[;text]]` code gives its negative section the absolute value,
 * because that section supplies its own sign or brackets. With conditions,
 * Excel 16.0 (build 20527, observed 2026-11-11) shows the absolute value in
 * a matched section whose condition only negative numbers meet
 * (`[<-5]0;0` shows -10 as `10`, `[<5]0;0` as `-10`), and in the second
 * section of a two-section code whose first condition has a threshold
 * <= 0 (`[>=0]0;0` shows -5 as `5`, `[>100]0;0` as `-5`).
 */
function resolveSectionConditional(
  sections: readonly string[],
  value: number,
): { section: string; value: number } {
  const conditions = sections.map(parseCondition);
  // Check for explicit conditions [>100], [<=50], etc.
  for (const [k, condition] of conditions.entries()) {
    if (condition && evaluateCondition(value, condition.op, condition.threshold)) {
      const shown = isNegativeOnly(condition) ? Math.abs(value) : value;
      return { section: sections[k] ?? 'General', value: shown };
    }
  }

  // No condition matched: the last section is the default of a conditional code.
  if (conditions.some((c) => c !== undefined)) {
    const shown = isAbsoluteDefault(conditions) ? Math.abs(value) : value;
    return { section: sections.at(-1) ?? 'General', value: shown };
  }
  return resolvePlainSection(sections, value);
}

/** The `pos;neg;zero;text` section for `value` in a code without conditions. */
function resolvePlainSection(
  sections: readonly string[],
  value: number,
): { section: string; value: number } {
  if (sections.length >= 3 && value === 0) {
    return { section: sections[2] ?? sections[0] ?? 'General', value };
  }
  if (sections.length >= 2 && value < 0) {
    return { section: sections[1] ?? sections[0] ?? 'General', value: Math.abs(value) };
  }
  return { section: sections[0] ?? 'General', value };
}

function evaluateCondition(value: number, op: string, threshold: number): boolean {
  switch (op) {
    case '>':
      return value > threshold;
    case '<':
      return value < threshold;
    case '>=':
      return value >= threshold;
    case '<=':
      return value <= threshold;
    case '=':
      return value === threshold;
    case '<>':
    case '!=':
      return value !== threshold;
    default:
      return false;
  }
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
] as const;

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
] as const;

/** Weekday names, Sunday first (the order of `Date.prototype.getUTCDay`). */
const DAY_NAMES = [
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
] as const;

/** Abbreviated weekday names, Sunday first. */
const DAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] as const;

const SECONDS_PER_DAY = 86_400;

/** Most sub-second digits a code shows (`ss.000`); Excel rejects a fourth. */
const MAX_SUBSECOND_DIGITS = 3;

/** Seconds in one unit of an elapsed-time token. */
const ELAPSED_UNIT_SECONDS = { h: 3600, m: 60, s: 1 } as const;

/**
 * Render `serial` with a date/time section: the code is split into tokens,
 * the `m` runs that mean minutes are marked, and each token prints its part.
 */
function formatDate(serial: number, code: string, system: DateSystem): string {
  const tokens = markMinutes(lexDateCode(code));
  const parts = extractDateParts(serial, system, tokens);
  return tokens.map((token) => renderDateToken(token, parts)).join('');
}

/**
 * The fields a date code can show. The time is rounded to whole seconds, or to
 * the digits of the code's longest `.0` token, before it is split, so rounding
 * carries into minutes, hours and the date (`h:mm:ss` shows 0.5 + 0.7 s as
 * `12:00:01`).
 */
interface DateParts {
  /** Calendar year. */
  readonly year: number;
  /** Month, 1 = January. */
  readonly month: number;
  /** Day of the month, 1-31. */
  readonly day: number;
  /**
   * Day of the week, 0 = Sunday. Serials below 60 in the 1900 system run a
   * day behind the calendar, as in Excel (serial 1 is a Sunday), because
   * Excel counts the nonexistent 1900-02-29.
   */
  readonly weekday: number;
  /** Hour of the day: 0-23, or 1-12 when the code has `AM/PM` or `A/P`. */
  readonly hours: number;
  /** Minute of the hour, 0-59. */
  readonly minutes: number;
  /** Second of the minute, 0-59. */
  readonly seconds: number;
  /** Whole seconds in the absolute serial, the source of `[h]`, `[m]` and `[s]`. */
  readonly totalSeconds: number;
  /** Sub-second digits, zero-padded to the longest `.0` token ('' without one). */
  readonly fraction: string;
  /** `AM` before noon, else `PM`; `A/P` shows the matching letter of the code. */
  readonly ampm: string;
}

/** Characters that form run-length date tokens (`m`, `dd`, `hhh`, `s`, `yyyy`, `ee`). */
type RunChar = 'm' | 'd' | 'h' | 's' | 'y' | 'e';

type TokenFormatter = (p: DateParts) => string;

/**
 * Renderers for run-length date tokens, keyed by lowercase character. A run
 * uses the first entry whose minimum length it reaches, so longer entries
 * come first: `mmmmmm` and longer print the full month name, `mmmmm` its
 * first letter, `mmmm` the full name, `mmm` the short name, `mm` and `m` the
 * month number; `dddd`+ the weekday, `ddd` its short name; `yyy`+ the
 * four-digit year, `y` and `yy` its last two digits; an `e` run of any
 * length prints the era year once, which is the year in the Gregorian
 * calendar. An `m` or `mm` run that means minutes never reaches this table
 * (see {@link markMinutes}).
 */
const RUN_TOKENS: Readonly<Record<RunChar, readonly (readonly [number, TokenFormatter])[]>> = {
  m: [
    [6, (p) => MONTH_NAMES[p.month - 1] ?? ''],
    [5, (p) => MONTH_NAMES[p.month - 1]?.[0] ?? ''],
    [4, (p) => MONTH_NAMES[p.month - 1] ?? ''],
    [3, (p) => MONTH_SHORT[p.month - 1] ?? ''],
    [2, (p) => String(p.month).padStart(2, '0')],
    [1, (p) => String(p.month)],
  ],
  d: [
    [4, (p) => DAY_NAMES[p.weekday] ?? ''],
    [3, (p) => DAY_SHORT[p.weekday] ?? ''],
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
  y: [
    [3, (p) => String(p.year).padStart(4, '0')],
    [1, (p) => String(p.year % 100).padStart(2, '0')],
  ],
  e: [[1, (p) => String(p.year)]],
};

function isRunChar(ch: string): ch is RunChar {
  return Object.hasOwn(RUN_TOKENS, ch);
}

/** A token matched by its exact (case-insensitive) text. */
interface FixedToken {
  readonly pattern: string;
  readonly render: TokenFormatter;
}

/**
 * Fixed-string tokens, matched before run tokens; keep any addition
 * longest-first. Today the only one is `AM/PM`.
 */
const FIXED_TOKENS = [
  { pattern: 'am/pm', render: (p: DateParts) => p.ampm },
] as const satisfies readonly FixedToken[];

/** One unit of a date code. */
type DateToken =
  /** Literal output: quoted text, an escape, any other character; '' for `[$-409]`. */
  | { readonly kind: 'text'; readonly text: string }
  | { readonly kind: 'fixed'; readonly token: FixedToken }
  /** `A/P`: the letter before (`am`) or after (`pm`) the slash, each in its source case. */
  | { readonly kind: 'ap'; readonly am: string; readonly pm: string }
  | { readonly kind: 'run'; readonly ch: RunChar; readonly len: number }
  /** An `m` or `mm` run that shows the minute (`len` 1 or 2). */
  | { readonly kind: 'minutes'; readonly len: number }
  /** `[h]`, `[mm]`, `[sss]`: the total elapsed units, zero-padded to `len`. */
  | {
      readonly kind: 'elapsed';
      readonly unit: keyof typeof ELAPSED_UNIT_SECONDS;
      readonly len: number;
    }
  /** `.0`, `.00`, `.000`: the fraction of the second. */
  | { readonly kind: 'subsecond'; readonly digits: number };

/** `[h]`, `[mm]`, `[sss]` at `lastIndex`, on the lower-cased code. */
const ELAPSED_RE = /\[(h+|m+|s+)\]/y;

/** Count consecutive occurrences of `ch` in `lower` starting at `start`. */
function countRun(lower: string, start: number, ch: string): number {
  let count = 0;
  while (lower.charAt(start + count) === ch) count++;
  return count;
}

/**
 * Try to match a fixed-string token at position `i` of `lower`, the ASCII
 * lower-cased format code (same length as the original, so indices line up).
 */
function matchFixedToken(lower: string, i: number): FixedToken | undefined {
  return FIXED_TOKENS.find((token) => lower.startsWith(token.pattern, i));
}

/** Skip a bracketed directive such as `[$-409]`. Returns the index after `]`, or -1. */
function skipBracket(s: string, i: number): number {
  const close = s.indexOf(']', i);
  return close !== -1 ? close + 1 : -1;
}

/** Extract a quoted literal and return [content, newIndex] or undefined. */
function extractQuoted(s: string, i: number): [string, number] | undefined {
  const end = s.indexOf('"', i + 1);
  return end !== -1 ? [s.slice(i + 1, end), end + 1] : undefined;
}

/** Read a bracketed directive at `lower[i]`: an elapsed-time token, or nothing to print. */
function lexBracket(lower: string, i: number): [DateToken, number] | undefined {
  ELAPSED_RE.lastIndex = i;
  const elapsed = ELAPSED_RE.exec(lower);
  const units = elapsed?.[1];
  if (units) {
    const unit = units.charAt(0) as keyof typeof ELAPSED_UNIT_SECONDS;
    return [{ kind: 'elapsed', unit, len: units.length }, units.length + 2];
  }
  const next = skipBracket(lower, i);
  return next === -1 ? undefined : [{ kind: 'text', text: '' }, next - i];
}

/**
 * Read the token at `code[i]` and the number of characters it spans. `lower`
 * is `code` with ASCII letters lower-cased (same length, same indices).
 * A `.` followed by `0`s is a sub-second token anywhere in a date code, as in
 * Excel (`mm.0`, `h:mm .0`).
 */
function lexDateToken(code: string, lower: string, i: number): [DateToken, number] {
  const ch = code.charAt(i);
  if (ch === '[') {
    const bracket = lexBracket(lower, i);
    if (bracket) return bracket;
  }
  if (ch === '"') {
    const quoted = extractQuoted(code, i);
    if (quoted) return [{ kind: 'text', text: quoted[0] }, quoted[1] - i];
  }
  if (ch === '\\' && i + 1 < code.length) {
    return [{ kind: 'text', text: code.charAt(i + 1) }, 2];
  }
  if (ch === '.' && code.charAt(i + 1) === '0') {
    const digits = Math.min(countRun(lower, i + 1, '0'), MAX_SUBSECOND_DIGITS);
    return [{ kind: 'subsecond', digits }, digits + 1];
  }
  const fixed = matchFixedToken(lower, i);
  if (fixed) return [{ kind: 'fixed', token: fixed }, fixed.pattern.length];
  if (lower.startsWith('a/p', i)) {
    return [{ kind: 'ap', am: code.charAt(i), pm: code.charAt(i + 2) }, 3];
  }
  const lc = lower.charAt(i);
  if (isRunChar(lc)) {
    const len = countRun(lower, i, lc);
    return [{ kind: 'run', ch: lc, len }, len];
  }
  return [{ kind: 'text', text: ch }, 1];
}

/** Split a date code into tokens. */
function lexDateCode(code: string): DateToken[] {
  // ASCII-only lower-casing: String.prototype.toLowerCase can change the length
  // (U+0130 becomes two code units), which would misalign indices into `code`.
  const lower = code.replace(/[A-Z]/g, (c) => String.fromCharCode(c.charCodeAt(0) | 0x20));
  const tokens: DateToken[] = [];
  let i = 0;
  while (i < code.length) {
    const [token, consumed] = lexDateToken(code, lower, i);
    tokens.push(token);
    i += consumed;
  }
  return tokens;
}

/** A meridiem token, `AM/PM` or `A/P`, which switches the hour to the 12-hour clock. */
function isAmPm(token: DateToken): boolean {
  return token.kind === 'ap' || (token.kind === 'fixed' && token.token.pattern === 'am/pm');
}

/**
 * A seconds token (`s`, `ss`, `.0`) is the next token after `k`, past
 * literals, `AM/PM` and `A/P`.
 */
function secondsFollow(tokens: readonly DateToken[], k: number): boolean {
  for (let j = k + 1; j < tokens.length; j++) {
    const token = tokens[j];
    if (token === undefined || token.kind === 'text' || isAmPm(token)) continue;
    return token.kind === 'subsecond' || (token.kind === 'run' && token.ch === 's');
  }
  return false;
}

/**
 * Mark the `m` and `mm` runs that show minutes. Microsoft documents `m`
 * right after `h` or right before `ss` as minutes; Excel 16.0 (build 20527,
 * observed 2026-11-11) applies it more widely. An `m`/`mm` run is minutes
 * when a seconds token (`s`, `ss`, `.0`) follows it with only literals,
 * `AM/PM` and `A/P` between, or when an hour token (`h`, `[h]`) came
 * earlier with no minutes run since (`h yyyy mm` shows the minute), or a
 * seconds run did before any minutes or `[m]` (`s mm` shows the minute,
 * `h:mm:ss mm` and `[m]:ss mm` the month). Longer runs (`mmm` and up) are
 * always month names.
 */
function markMinutes(tokens: readonly DateToken[]): DateToken[] {
  let pending = false;
  let minutesShown = false;
  return tokens.map((token, k) => {
    if (
      (token.kind === 'run' && token.ch === 'h') ||
      (token.kind === 'elapsed' && token.unit === 'h')
    ) {
      pending = true;
    } else if (token.kind === 'elapsed' && token.unit === 'm') {
      pending = false;
      minutesShown = true;
    } else if (token.kind === 'run' && token.ch === 's') {
      pending ||= !minutesShown;
    } else if (
      token.kind === 'run' &&
      token.ch === 'm' &&
      token.len <= 2 &&
      (pending || secondsFollow(tokens, k))
    ) {
      pending = false;
      minutesShown = true;
      return { kind: 'minutes', len: token.len };
    }
    return token;
  });
}

/**
 * Round `serial` to whole seconds, or to the digits of the longest `.0` token,
 * and split it into {@link DateParts}. The time comes from |serial| (Excel shows `####` for a
 * negative date or time); the date from the rounded whole days.
 */
function extractDateParts(
  serial: number,
  system: DateSystem,
  tokens: readonly DateToken[],
): DateParts {
  // A loop, not Math.max(...spread): a long code would pass one argument per token.
  let digits = 0;
  for (const t of tokens) if (t.kind === 'subsecond' && t.digits > digits) digits = t.digits;
  const scale = 10 ** digits;
  const units = Math.round(Math.abs(serial) * SECONDS_PER_DAY * scale);
  const totalSeconds = Math.floor(units / scale);
  const wholeDays = Math.floor(totalSeconds / SECONDS_PER_DAY);
  const daySerial = serial < 0 ? -wholeDays : wholeDays;
  const date = serialToDate(daySerial, system);
  let weekday = date.getUTCDay();
  if (system === 'date1900' && daySerial < 60) weekday = (weekday + 6) % 7;

  const hours24 = Math.floor(totalSeconds / 3600) % 24;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekday,
    hours: tokens.some(isAmPm) ? hours24 % 12 || 12 : hours24,
    minutes: Math.floor(totalSeconds / 60) % 60,
    seconds: totalSeconds % 60,
    totalSeconds,
    fraction: digits > 0 ? String(units % scale).padStart(digits, '0') : '',
    ampm: hours24 < 12 ? 'AM' : 'PM',
  };
}

/** The text one date token shows for `p`. */
function renderDateToken(token: DateToken, p: DateParts): string {
  switch (token.kind) {
    case 'text':
      return token.text;
    case 'fixed':
      return token.token.render(p);
    case 'ap':
      return p.ampm === 'AM' ? token.am : token.pm;
    case 'run': {
      const entry = RUN_TOKENS[token.ch].find(([minLen]) => token.len >= minLen);
      return entry ? entry[1](p) : '';
    }
    case 'minutes':
      return String(p.minutes).padStart(token.len, '0');
    case 'elapsed': {
      const total = Math.floor(p.totalSeconds / ELAPSED_UNIT_SECONDS[token.unit]);
      return String(total).padStart(token.len, '0');
    }
    case 'subsecond':
      return `.${p.fraction.slice(0, token.digits)}`;
  }
}

// ---------------------------------------------------------------------------
// Number sections: tokens
// ---------------------------------------------------------------------------

/** A digit placeholder: `0` pads with zeros, `#` shows significant digits only, `?` pads with spaces. */
type Placeholder = '0' | '#' | '?';

/** One unit of a number section. */
type NumToken =
  | { readonly kind: 'digit'; readonly ch: Placeholder }
  | { readonly kind: 'point' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'percent' }
  | { readonly kind: 'exp'; readonly text: string; readonly sign: '+' | '-' }
  | { readonly kind: 'slash' }
  | { readonly kind: 'general' }
  /** An unquoted digit 1-9: literal text, or part of a fixed fraction denominator. */
  | { readonly kind: 'numeral'; readonly text: string }
  | { readonly kind: 'literal'; readonly text: string };

/** How a section scales and groups its value. */
interface NumberLayout {
  /** Thousands separators between integer placeholders. */
  readonly grouping: boolean;
  /** Number of `%` signs: each multiplies by 100 (not in scientific sections). */
  readonly percent: number;
  /** Number of scaling commas: each divides by 1000 (not in scientific sections). */
  readonly scale: number;
}

function literal(text: string): NumToken {
  return { kind: 'literal', text };
}

/** Text a bracket directive shows: the symbol of `[$€-407]`; nothing for `[$-409]` and others. */
function bracketText(inner: string): string {
  return /^\$([^-]*)/.exec(inner)?.[1] ?? '';
}

/** Tokens that are exactly one character of the section. */
const CHAR_TOKENS: Readonly<Record<string, NumToken>> = {
  '0': { kind: 'digit', ch: '0' },
  '#': { kind: 'digit', ch: '#' },
  '?': { kind: 'digit', ch: '?' },
  '.': { kind: 'point' },
  ',': { kind: 'comma' },
  '%': { kind: 'percent' },
  '/': { kind: 'slash' },
};

/** Index of the first `closer` after `section[i]`, or the section length when there is none. */
function closingIndex(section: string, i: number, closer: string): number {
  const end = section.indexOf(closer, i + 1);
  return end === -1 ? section.length : end;
}

/**
 * Read a construct that spans more than one character at `section[i]`:
 * quoted text, an escape, pad or fill, a bracket, an exponent or `General`.
 * Returns the token (none for fill) and the characters consumed, or
 * `undefined` when `section[i]` is a single-character token.
 */
function readCompound(section: string, i: number): [NumToken | undefined, number] | undefined {
  const ch = section.charAt(i);
  switch (ch) {
    case '"': {
      const stop = closingIndex(section, i, '"');
      return [literal(section.slice(i + 1, stop)), stop + 1 - i];
    }
    case '\\':
      return [literal(section.charAt(i + 1)), 2];
    case '_':
      return [literal(' '), 2];
    case '*':
      return [undefined, 2];
    case '[': {
      const stop = closingIndex(section, i, ']');
      return [literal(bracketText(section.slice(i + 1, stop))), stop + 1 - i];
    }
    case 'E':
    case 'e': {
      const sign = section.charAt(i + 1);
      return sign === '+' || sign === '-' ? [{ kind: 'exp', text: ch, sign }, 2] : undefined;
    }
    case 'G':
    case 'g':
      return /^general$/i.test(section.slice(i, i + 7)) ? [{ kind: 'general' }, 7] : undefined;
    default:
      return undefined;
  }
}

/**
 * Split a cleaned number section into tokens. Quoted text and `\x` are
 * literals, `_x` is one space, `*x` is dropped (there is no column width to
 * fill), a bracket shows only a currency symbol (`[$€-407]`), the first
 * unquoted `.` is the decimal point, `E+`/`E-`/`e+`/`e-` start an exponent,
 * and every other character is a literal.
 */
function tokenize(section: string): NumToken[] {
  const tokens: NumToken[] = [];
  let i = 0;
  while (i < section.length) {
    const [token, consumed] = readCompound(section, i) ?? [charToken(section.charAt(i)), 1];
    if (token) tokens.push(token);
    i += consumed;
  }
  // Only the first point is the decimal point; later ones print as dots.
  const pointAt = tokens.findIndex((t) => t.kind === 'point');
  return tokens.map((t, j) => (t.kind === 'point' && j !== pointAt ? literal('.') : t));
}

/** The token for a single character: a placeholder, `.`, `,`, `%`, `/`, a numeral or a literal. */
function charToken(ch: string): NumToken {
  return CHAR_TOKENS[ch] ?? (/[1-9]/.test(ch) ? { kind: 'numeral', text: ch } : literal(ch));
}

type NumTokenKind = NumToken['kind'];

/**
 * The kind of the nearest non-comma token before (`prevKind`) and after
 * (`nextKind`) each index, `undefined` at the ends. Two linear passes, so a
 * run of n commas costs O(n), not O(n²).
 */
function commaNeighbourKinds(tokens: readonly NumToken[]): {
  prevKind: (NumTokenKind | undefined)[];
  nextKind: (NumTokenKind | undefined)[];
} {
  const prevKind: (NumTokenKind | undefined)[] = [];
  const nextKind = new Array<NumTokenKind | undefined>(tokens.length).fill(undefined);
  let last: NumTokenKind | undefined;
  for (const token of tokens) {
    prevKind.push(last);
    if (token.kind !== 'comma') last = token.kind;
  }
  last = undefined;
  for (let j = tokens.length - 1; j >= 0; j--) {
    nextKind[j] = last;
    const kind = tokens[j]?.kind;
    if (kind !== 'comma') last = kind;
  }
  return { prevKind, nextKind };
}

/**
 * Resolve commas: one between integer placeholders turns on grouping; one
 * after the last placeholder of a run divides by 1000 (`#,##0,` shows
 * thousands); a comma at or after the exponent is dropped (`0.0E+0,0` shows
 * 1234567 as `1.2E+06`); any other comma is a literal.
 */
function classifyCommas(tokens: readonly NumToken[]): {
  tokens: NumToken[];
  grouping: boolean;
  scale: number;
} {
  const pointAt = tokens.findIndex((t) => t.kind === 'point');
  const intEnd = pointAt === -1 ? tokens.length : pointAt;
  const expAt = tokens.findIndex((t) => t.kind === 'exp');
  const { prevKind, nextKind } = commaNeighbourKinds(tokens);
  const out: NumToken[] = [];
  let grouping = false;
  let scale = 0;
  tokens.forEach((token, j) => {
    if (token.kind !== 'comma') {
      out.push(token);
      return;
    }
    if (expAt !== -1 && j >= expAt) return;
    const before = prevKind[j] === 'digit';
    const after = nextKind[j] === 'digit';
    if (before && after && j < intEnd) grouping = true;
    else if (before && !after) scale++;
    else out.push(literal(','));
  });
  return { tokens: out, grouping, scale };
}

function countPercent(tokens: readonly NumToken[]): number {
  return tokens.filter((t) => t.kind === 'percent').length;
}

/** Text of a token outside a digit template. */
function tokenText(token: NumToken): string {
  switch (token.kind) {
    case 'digit':
      return token.ch;
    case 'point':
      return '.';
    case 'comma':
      return ',';
    case 'percent':
      return '%';
    case 'exp':
      return token.text + token.sign;
    case 'slash':
      return '/';
    case 'general':
      return 'General';
    case 'numeral':
    case 'literal':
      return token.text;
  }
}

/** The tokens as plain text, without a number. */
function renderLiterals(tokens: readonly NumToken[]): string {
  return tokens.map(tokenText).join('');
}

/** `General` with literals around it (`0;-0;General`, `"Total: "General`). */
function renderGeneralSection(value: number, tokens: readonly NumToken[]): string {
  return tokens
    .map((t) => (t.kind === 'general' ? formatGeneralNumber(value) : tokenText(t)))
    .join('');
}

// ---------------------------------------------------------------------------
// Decimal-safe rounding
// ---------------------------------------------------------------------------

/** Significant digits Excel keeps in a number. */
const EXCEL_PRECISION = 15;

/** A non-negative number as `d.ddd x 10^exp`, rounded to 15 significant digits. */
interface Decimal {
  /** The significant digits (`'0'` for zero). */
  readonly digits: string;
  /** Power of ten of the first digit. */
  readonly exp: number;
}

/** Excel's 15-digit decimal view of `abs`, shifted by `shift` powers of ten. */
function toDecimal(abs: number, shift: number): Decimal {
  if (abs === 0) return { digits: '0', exp: 0 };
  const [mantissa = '0', exp = '0'] = abs.toExponential(EXCEL_PRECISION - 1).split('e');
  return { digits: mantissa.replace('.', ''), exp: Number(exp) + shift };
}

/**
 * Round `dec` half away from zero to `decimals` places, on its decimal
 * digits (so 1.005 rounds to 1.01), and split it. `int` has no leading
 * zeros (`''` for a zero integer part); `frac` has exactly `decimals` digits.
 */
function roundDecimal(dec: Decimal, decimals: number): { int: string; frac: string } {
  const keep = dec.exp + 1 + decimals;
  let scaled = '0';
  if (keep >= 0) {
    const head = dec.digits.slice(0, keep);
    const roundUp = dec.digits.charAt(keep) >= '5';
    scaled = (BigInt(head === '' ? '0' : head) + (roundUp ? 1n : 0n)).toString();
    if (keep > dec.digits.length) scaled += '0'.repeat(keep - dec.digits.length);
  }
  const padded = scaled.padStart(decimals + 1, '0');
  const cut = padded.length - decimals;
  return { int: padded.slice(0, cut).replace(/^0+/, ''), frac: padded.slice(cut) };
}

// ---------------------------------------------------------------------------
// Number sections: template walker
// ---------------------------------------------------------------------------

/** What an integer placeholder shows when it has no digit. */
function emptySlot(ch: Placeholder): string {
  return ch === '0' ? '0' : ch === '?' ? ' ' : '';
}

/**
 * Fill integer placeholders right to left with the digits of `int`; extra
 * leading digits go into the leftmost placeholder.
 */
function fillIntegerSlots(int: string, slots: readonly Placeholder[]): string[] {
  const out = slots.map(emptySlot);
  let d = int.length - 1;
  for (let k = slots.length - 1; k >= 0 && d >= 0; k--, d--) out[k] = int.charAt(d);
  if (d >= 0 && out.length > 0) out[0] = int.slice(0, d + 1) + (out[0] ?? '');
  return out;
}

/** Fill decimal placeholders; trailing zeros vanish under `#` and become spaces under `?`. */
function fillFractionSlots(frac: string, slots: readonly Placeholder[]): string[] {
  let significant = frac.length;
  while (significant > 0 && frac.charAt(significant - 1) === '0') significant--;
  return slots.map((ch, k) => (k < significant || ch === '0' ? frac.charAt(k) : emptySlot(ch)));
}

/**
 * Insert thousands separators into the filled integer placeholders: ',' left
 * of every third digit, or a space where a `?` padded the position
 * (`??,??0` shows 12 as `    12`).
 */
function groupSlots(slots: readonly string[]): string[] {
  let count = 0;
  const out = [...slots];
  for (let k = out.length - 1; k >= 0; k--) {
    let text = '';
    for (const ch of [...(out[k] ?? '')].reverse()) {
      const separator = ch === ' ' ? ' ' : ',';
      text = count > 0 && count % 3 === 0 ? `${ch}${separator}${text}` : ch + text;
      count++;
    }
    out[k] = text;
  }
  return out;
}

function placeholders(tokens: readonly NumToken[]): Placeholder[] {
  return tokens.flatMap((t) => (t.kind === 'digit' ? [t.ch] : []));
}

/**
 * Walk a digit template that holds the rounded `int` and `frac` parts:
 * placeholders take digits, every other token prints its literal text.
 */
function renderTemplate(
  tokens: readonly NumToken[],
  parts: { int: string; frac: string },
  grouping: boolean,
): string {
  const pointAt = tokens.findIndex((t) => t.kind === 'point');
  const intEnd = pointAt === -1 ? tokens.length : pointAt;
  const intSlots = placeholders(tokens.slice(0, intEnd));
  const filledInt = fillIntegerSlots(parts.int, intSlots);
  const intText = grouping ? groupSlots(filledInt) : filledInt;
  const fracText = fillFractionSlots(parts.frac, placeholders(tokens.slice(intEnd)));
  // With no integer placeholder ('.00'), the integer digits print before the point.
  const bareInt =
    intSlots.length > 0 ? '' : grouping ? (groupSlots([parts.int])[0] ?? '') : parts.int;

  let out = '';
  let intK = 0;
  let fracK = 0;
  tokens.forEach((token, j) => {
    if (token.kind === 'digit') {
      out += j < intEnd ? (intText[intK++] ?? '') : (fracText[fracK++] ?? '');
    } else if (token.kind === 'point') {
      out += `${bareInt}.`;
    } else {
      out += tokenText(token);
    }
  });
  return out;
}

/** A plain number section: digits, literals, grouping, scaling and percent. */
function renderNumber(value: number, tokens: readonly NumToken[], layout: NumberLayout): string {
  const pointAt = tokens.findIndex((t) => t.kind === 'point');
  const decimals = pointAt === -1 ? 0 : placeholders(tokens.slice(pointAt)).length;
  const parts = roundDecimal(
    toDecimal(Math.abs(value), 2 * layout.percent - 3 * layout.scale),
    decimals,
  );
  return (value < 0 ? '-' : '') + renderTemplate(tokens, parts, layout.grouping);
}

/**
 * Scientific section `mantissa E±exponent`. `E+`/`e+` always print the
 * exponent sign, `E-`/`e-` only a '-'; the exponent is zero-padded to its
 * placeholder count and keeps the format's `E`/`e`. The exponent is a
 * multiple of the k integer placeholders, so k > 1 is engineering notation
 * (`##0.0E+0` shows 12300 as `12.3E+3`; `00.00E+00` shows 12345 as
 * `01.23E+04` and 0.5 as `50.00E-02`), and with no integer placeholder the
 * mantissa is below 1. `%` is printed but does not scale the value
 * (`0.00E+00%` shows 0.5 as `5.00E-01%`), and scaling commas are ignored
 * (`0.0,E+0` shows 1234567 as `1.2E+6`), as in Excel 16.0 (build 20527,
 * observed 2026-11-11).
 */
function renderScientific(
  value: number,
  tokens: readonly NumToken[],
  expAt: number,
  layout: NumberLayout,
): string {
  const expToken = tokens[expAt];
  if (expToken?.kind !== 'exp') return renderLiterals(tokens);
  const mantissa = tokens.slice(0, expAt);
  const pointAt = mantissa.findIndex((t) => t.kind === 'point');
  const intSlots = placeholders(pointAt === -1 ? mantissa : mantissa.slice(0, pointAt));
  const decimals = pointAt === -1 ? 0 : placeholders(mantissa.slice(pointAt)).length;
  const k = intSlots.length;
  const step = Math.max(k, 1);

  const dec = toDecimal(Math.abs(value), 0);
  let exponent = 0;
  // Zero shows a '0' even under '#' ('#.##E+00' shows 0 as '0.E+00').
  let parts = { int: k > 0 ? '0' : '', frac: '0'.repeat(decimals) };
  if (dec.digits !== '0') {
    exponent = k === 0 ? dec.exp + 1 : Math.floor(dec.exp / step) * step;
    parts = roundDecimal({ digits: dec.digits, exp: dec.exp - exponent }, decimals);
    // Rounding can carry the mantissa past its integer width (9.99 -> 10.0).
    if (parts.int.length > k) {
      exponent += step;
      parts = roundDecimal({ digits: dec.digits, exp: dec.exp - exponent }, decimals);
    }
  }

  const tail = tokens.slice(expAt + 1);
  let width = 0;
  while (tail[width]?.kind === 'digit') width++;
  const expSign = exponent < 0 ? '-' : expToken.sign === '+' ? '+' : '';
  const expDigits = String(Math.abs(exponent)).padStart(width, '0');

  return (
    (value < 0 ? '-' : '') +
    renderTemplate(mantissa, parts, layout.grouping) +
    expToken.text +
    expSign +
    expDigits +
    renderLiterals(tail.slice(width))
  );
}

// ---------------------------------------------------------------------------
// Fraction sections
// ---------------------------------------------------------------------------

/** Index of a `/` between a numerator placeholder and a denominator, or -1. */
function findFractionSlash(tokens: readonly NumToken[]): number {
  return tokens.findIndex((t, j) => {
    if (t.kind !== 'slash' || tokens[j - 1]?.kind !== 'digit') return false;
    const next = tokens[j + 1]?.kind;
    return next === 'digit' || next === 'numeral';
  });
}

/** Remainder below which the continued-fraction expansion stops. */
const FRACTION_EPSILON = 5e-8;

/**
 * The last continued-fraction convergent p/q of `x` with q <= maxDen, expanded
 * in floating point. This is how Excel picks fractions: Excel 16.0 (build
 * 20527, observed 2026-11-11) shows 3.14159 under `# ??/??` as `3  1/7 `, not
 * the closer 3 14/99, and 0.3 under `# ?/?` as ` 2/7` (the floating-point
 * expansion of 0.3 is [0; 3, 2, 1]). SheetJS SSF 0.20.3 uses the same method.
 */
function convergentFraction(x: number, maxDen: number): { num: number; den: number } {
  let [p2, p1, q2, q1] = [0, 1, 1, 0];
  let p = 0;
  let q = 1;
  let rest = x;
  while (q1 < maxDen) {
    const a = Math.floor(rest);
    p = a * p1 + p2;
    q = a * q1 + q2;
    if (rest - a < FRACTION_EPSILON) break;
    rest = 1 / (rest - a);
    [p2, p1, q2, q1] = [p1, p, q1, q];
  }
  if (q > maxDen) {
    [p, q] = q1 > maxDen ? [p2, q2] : [p1, q1];
  }
  return { num: p, den: q };
}

/**
 * Split `abs` into whole part, numerator and denominator. A fixed
 * denominator rounds the numerator; otherwise the convergent rule applies.
 * Without a whole part the fraction is improper.
 */
function splitFraction(
  abs: number,
  hasWhole: boolean,
  fixedDen: number,
  maxDen: number,
): { whole: number; num: number; den: number } {
  const total =
    fixedDen > 0
      ? { num: Math.round(abs * fixedDen), den: fixedDen }
      : convergentFraction(abs, maxDen);
  const whole = hasWhole ? Math.floor(total.num / total.den) : 0;
  return { whole, num: total.num - whole * total.den, den: total.den };
}

/**
 * A fixed denominator written with numerals (`8`, `16`, `10`: a `0` inside it
 * is a digit), or 0 when the denominator is a placeholder width.
 */
function fixedDenominator(denTokens: readonly NumToken[]): number {
  if (!denTokens.some((t) => t.kind === 'numeral')) return 0;
  return Number(
    denTokens.map((t) => (t.kind === 'digit' ? t.ch.replace(/[#?]/, '') : tokenText(t))).join(''),
  );
}

/** Denominator text: digits left-aligned, then the unused placeholders. */
function fillDenominator(den: number, slots: readonly Placeholder[]): string {
  const digits = String(den);
  return digits + slots.slice(digits.length).map(emptySlot).join('');
}

/**
 * Fraction section `[whole ]numerator/denominator`. The denominator is a
 * width (n placeholders allow up to 10^n - 1) or a fixed integer (`?/8`).
 * Without a whole-part template the fraction is improper (`?/?` shows 1.5
 * as `3/2`). A negative value prints '-' first; a numerator that reaches the
 * denominator carries into the whole part, and a zero fraction beside a
 * whole part becomes spaces of the same width (`# ?/?` shows 0.999 as `1    `).
 * Each `%` multiplies the value by 100 (`# ?/?%` shows 1.5 as `150    %`).
 */
function renderFractionSection(
  value: number,
  tokens: readonly NumToken[],
  slashAt: number,
  layout: NumberLayout,
): string {
  const left = tokens.slice(0, slashAt);
  const right = tokens.slice(slashAt + 1);
  let numStart = left.length;
  while (left[numStart - 1]?.kind === 'digit') numStart--;
  let denEnd = 0;
  while (right[denEnd]?.kind === 'digit' || right[denEnd]?.kind === 'numeral') denEnd++;

  const head = left.slice(0, numStart);
  const lastWholeDigit = head.findLastIndex((t) => t.kind === 'digit');
  const wholeTemplate = head.slice(0, lastWholeDigit + 1);
  const gap = head.slice(lastWholeDigit + 1);
  const hasWhole = lastWholeDigit !== -1;
  const numSlots = placeholders(left.slice(numStart));
  const denTokens = right.slice(0, denEnd);
  const denSlots = placeholders(denTokens);
  const fixedDen = fixedDenominator(denTokens);

  const abs = Number((Math.abs(value) * 100 ** layout.percent).toPrecision(EXCEL_PRECISION));
  const maxDen = Math.max(10 ** denSlots.length - 1, 1);
  const { whole, num, den } = splitFraction(abs, hasWhole, fixedDen, maxDen);

  const denText = fixedDen > 0 ? String(fixedDen) : fillDenominator(den, denSlots);
  const fractionText = `${renderLiterals(gap)}${fillIntegerSlots(String(num), numSlots).join('')}/${denText}`;
  const zeroFraction = hasWhole && num === 0;
  const wholeDigits = whole === 0 ? (zeroFraction ? '0' : '') : String(whole);
  const wholeText = hasWhole
    ? renderTemplate(wholeTemplate, { int: wholeDigits, frac: '' }, layout.grouping)
    : '';
  // Excel keeps the sign of a value that shows as a whole 0 ('-0    ') but not of an improper '0/1'.
  const negative = value < 0 && (hasWhole || num !== 0);

  return (
    (negative ? '-' : '') +
    wholeText +
    (zeroFraction ? ' '.repeat(fractionText.length) : fractionText) +
    renderLiterals(right.slice(denEnd))
  );
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Split a format code at each `;` outside quotes. Outside quotes `\x`, `_x`
 * and `*x` are two-character units, so their second character never splits
 * (`0\;0` and `0_;0` are one section), as in SheetJS SSF 0.20.3.
 */
function splitSections(code: string): string[] {
  const sections: string[] = [];
  let current = '';
  let inQuote = false;
  let i = 0;
  while (i < code.length) {
    const ch = code.charAt(i);
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
    } else if (!inQuote && (ch === '\\' || ch === '_' || ch === '*')) {
      current += code.slice(i, i + 2);
      i++;
    } else if (ch === ';' && !inQuote) {
      sections.push(current);
      current = '';
    } else {
      current += ch;
    }
    i++;
  }
  sections.push(current);
  return sections;
}
