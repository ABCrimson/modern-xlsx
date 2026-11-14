import { afterEach, describe, expect, it } from 'vitest';
import { formatCell, formatCellRich, getBuiltinFormat, loadFormat } from '../src/index.js';

/**
 * Expected strings follow Excel's number-format semantics: every code Excel
 * 16.0 (build 20527) accepts was checked against its cell display and
 * `TEXT()` on 2026-11-11. They equal SheetJS SSF 0.20.3 (the devDependency
 * `xlsx`) except where SSF diverges from Excel ('0E+00' keeps two decimals,
 * lowercase 'e' is mangled, '##0.0E+0' does not carry). Excel refuses the
 * lowercase code '0.00e-00'; its expected value applies the 'e-' rule with
 * the format's case kept.
 */
type Case = readonly [value: number | string, format: string | number, expected: string];

const NUMBER_CASES: readonly Case[] = [
  [-1234.5, 40, '(1,234.50)'],
  [1234.5, 40, '1,234.50'],
  [-5, 37, '(5)'],
  [5, 37, '5 '],
  [1234, 38, '1,234 '],
  [5, '000', '005'],
  [-5, '#,##0;(#,##0)', '(5)'],
  [1234, '$#,##0', '$1,234'],
  [1.25, '0.0#', '1.25'],
  [0.5, '#.##', '.5'],
  [1234, '[Red]#,##0', '1,234'],
];

const PERCENT_SECTION_CASES: readonly Case[] = [
  [-0.5, '0.0%;(0.0%)', '(50.0%)'],
  [12.345, '#,##0.0%', '1,234.5%'],
  [0.5, '"Rate: "0.0%', 'Rate: 50.0%'],
  [-12345, '0.00E+00;-0.00E+00', '-1.23E+04'],
];

const SCIENTIFIC_CASES: readonly Case[] = [
  [12345, '0.00E+00', '1.23E+04'],
  [12300, '0E+00', '1E+04'],
  [0.000123, '0.00E+00', '1.23E-04'],
  [12300, '0.00E-00', '1.23E04'],
  [0.000123, '0.00e-00', '1.23e-04'],
  [12300, 48, '12.3E+3'],
];

const FRACTION_CASES: readonly Case[] = [
  [-0.5, '# ?/?', '- 1/2'],
  [0.33, '# ?/?', ' 1/3'],
  [0.999, '# ?/?', '1    '],
  [0.3, '# ?/8', ' 2/8'],
  [1.5, '?/?', '3/2'],
];

const STRICT_PARSE_CASES: readonly Case[] = [
  ['12abc', '0.00', '12abc'],
  ['0x1A', '0.00', '0x1A'],
  [' 12 ', '0.00', '12.00'],
  ['1e3', '0', '1000'],
];

const GENERAL_CASES: readonly Case[] = [
  [1e21, 'General', '1E+21'],
  [0.1 + 0.2, 'General', '0.3'],
  [1 / 3, 'General', '0.333333333'],
];

/** Cell text Excel 16.0 (build 20527) displayed for these codes on 2026-11-11. */
const EXCEL_OBSERVED_CASES: readonly Case[] = [
  [-3, '(0)', '-(3)'],
  [123, '0-0', '12-3'],
  [123456789, '000-00-0000', '123-45-6789'],
  [-5, '$#,##0_);($#,##0)', '($5)'],
  [1234567, '#,##0,', '1,235'],
  [1234567, '0.0,,', '1.2'],
  [5, '0\\%', '5%'],
  [0.01, '0%%', '100%%'],
  [1.5, '???.??', '  1.5 '],
  [1.5, '00.00?', '01.50 '],
  [12, '??,??0', '    12'],
  [5, '0,000', '0,005'],
  [1e21, '0.00', '1000000000000000000000.00'],
  [123456789012345680, '0', '123456789012346000'],
  [1.005, '0.00', '1.01'],
  [1234.5, '[$€-407]#,##0.00', '€1,234.50'],
  [0, '0;;', ''],
  [5, ';;;', ''],
  [-5, '"x"', '-x'],
  [12.5, '"Total: "General', 'Total: 12.5'],
  [0, '#.##E+00', '0.E+00'],
  [9.99, '0.0E+0', '1.0E+1'],
  [12345, '00.00E+00', '01.23E+04'],
  [0.5, '00.00E+00', '50.00E-02'],
  [0.00012345, '##0.0E+0', '123.5E-6'],
  [999999, '##0.0E+0', '1.0E+6'],
  [0.5, '0.00E+00%', '5.00E-01%'],
  [1234567, '0.0,E+0', '1.2E+6'],
  [1234567, '0.0E+0,', '1.2E+6'],
  [1234567, '0.0E+00,,', '1.2E+06'],
  [1234567, '0.0E+0,0', '1.2E+06'],
  [-1234, '0.0E+00;(0.0E+00)', '(1.2E+03)'],
  [Math.PI, '# ??/??', '3  1/7 '],
  [Math.PI, '# ???/???', '3  16/113'],
  [-0.3, '# ?/?', '- 2/7'],
  [0, '# ?/?', '0    '],
  [-0.01, '# ?/?', '-0    '],
  [-0.01, '?/?', '0/1'],
  [0.4, '# ??/16', '  6/16'],
  [0.25, '# ?/10', ' 3/10'],
  [1.5, '# ?/?%', '150    %'],
  [0.25, '?/?%', '25/1%'],
  [0.25, '# ?/?%', '25    %'],
  [2.5, '# ?/?%%', '25000    %%'],
  [-10, '[<-5]0;0', '10'],
  [-10, '[<5]0;0', '-10'],
  [-5, '[>=0]0;0', '5'],
  [-5, '[>100]0;0', '-5'],
  [-5, '[<=0]0;0', '-5'],
  [-5, '[=1]"one";"x"', '-x'],
];

const ALL_CASES: readonly Case[] = [
  ...NUMBER_CASES,
  ...PERCENT_SECTION_CASES,
  ...SCIENTIFIC_CASES,
  ...FRACTION_CASES,
  ...STRICT_PARSE_CASES,
  ...GENERAL_CASES,
  ...EXCEL_OBSERVED_CASES,
];

describe('formatCell number template walker', () => {
  it.each(NUMBER_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('formatCell percent and sections', () => {
  it.each(PERCENT_SECTION_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('formatCell scientific', () => {
  it.each(SCIENTIFIC_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('formatCell fractions', () => {
  it.each(FRACTION_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('formatCellRich conditions', () => {
  it('matches a negative threshold', () => {
    expect(formatCellRich(-10, '[<-5]"low";"other"').text).toBe('low');
    expect(formatCellRich(-3, '[<-5]"low";"other"').text).toBe('other');
  });
});

describe('formatCell strict numeric strings', () => {
  it.each(STRICT_PARSE_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('loadFormat precedence', () => {
  const builtin14 = getBuiltinFormat(14) ?? 'General';

  afterEach(() => {
    loadFormat(builtin14, 14);
  });

  it('a registration overrides the built-in code for the same id', () => {
    loadFormat('0.000', 14);
    expect(formatCell(1, 14)).toBe('1.000');
    expect(formatCellRich(1, 14).text).toBe('1.000');
  });
});

describe('formatCell General', () => {
  it.each(GENERAL_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

describe('formatCell matches Excel 16.0', () => {
  it.each(EXCEL_OBSERVED_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });
});

/**
 * Section splitting honours `\x`, `_x` and `*x` as two-character units, and
 * colour and condition directives inside quotes are literal text. SheetJS
 * SSF 0.20.3 gives the same strings, except that it rejects `0*;0`, where the
 * `*;` fill is dropped as in every other code.
 */
const ESCAPE_CASES: readonly Case[] = [
  [5, '0\\;0', '0;5'],
  [-5, '0\\;0', '-0;5'],
  [-5, '0\\;0;"neg"', 'neg'],
  [5, '0\\"0;0', '0"5'],
  [5, '0_;0', '0 5'],
  [-5, '0_;0', '-0 5'],
  [5, '0*;0', '05'],
  [-5, '#,##0_);(#,##0)', '(5)'],
  [5, '#,##0_);(#,##0)', '5 '],
  [10, '"[>5]"0;0', '[>5]10'],
  [5, '"[Red] "0', '[Red] 5'],
];

describe('formatCell escapes and quoted directives', () => {
  it.each(ESCAPE_CASES)('formatCell(%j, %j) is %j', (value, format, expected) => {
    expect(formatCell(value, format)).toBe(expected);
  });

  it('takes no colour from a quoted [Red]', () => {
    expect(formatCellRich(5, '"[Red]"0')).toEqual({ text: '[Red]5' });
  });

  it('keeps unquoted colour and condition directives', () => {
    expect(formatCellRich(5, '[Red]0;[Blue]-0')).toEqual({ text: '5', color: 'Red' });
    expect(formatCellRich(-5, '[Red]0;[Blue]-0')).toEqual({ text: '-5', color: 'Blue' });
    expect(formatCell(200, '[>100]#,##0;[<=100]0.00')).toBe('200');
    expect(formatCell(50, '[>100]#,##0;[<=100]0.00')).toBe('50.00');
  });
});

describe('formatCell and formatCellRich agree', () => {
  it.each(ALL_CASES)('formatCell(%j, %j) equals formatCellRich().text', (value, format) => {
    expect(formatCell(value, format)).toBe(formatCellRich(value, format).text);
  });
});
