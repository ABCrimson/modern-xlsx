import { describe, expect, it } from 'vitest';
import { isDateFormatCode } from '../src/format/dates.js';
import {
  formatCell,
  formatCellRich,
  getBuiltinFormat,
  loadFormat,
  loadFormatTable,
} from '../src/index.js';

describe('formatCell', () => {
  it('returns the value as-is for General format', () => {
    expect(formatCell(123, 'General')).toBe('123');
    expect(formatCell('hello', 'General')).toBe('hello');
  });

  it('formats numbers with #,##0.00', () => {
    expect(formatCell(1234567.891, '#,##0.00')).toBe('1,234,567.89');
    expect(formatCell(0, '#,##0.00')).toBe('0.00');
    expect(formatCell(42, '#,##0.00')).toBe('42.00');
  });

  it('formats percentages with 0.00%', () => {
    expect(formatCell(0.1234, '0.00%')).toBe('12.34%');
    expect(formatCell(1, '0.00%')).toBe('100.00%');
    expect(formatCell(0, '0.00%')).toBe('0.00%');
  });

  it('formats dates with yyyy-mm-dd (serial 45292 = 2024-01-01)', () => {
    // Serial 45292 corresponds to 2024-01-01 in the date1900 system
    const result = formatCell(45292, 'yyyy-mm-dd');
    expect(result).toBe('2024-01-01');
  });

  it('formats dates with mm/dd/yy', () => {
    const result = formatCell(45292, 'mm/dd/yy');
    expect(result).toBe('01/01/24');
  });

  it('formats dates using built-in format ID 14 (mm-dd-yy)', () => {
    // Built-in format 14 is 'mm-dd-yy'
    const result = formatCell(45292, 14);
    expect(result).toBe('01-01-24');
  });

  it('formats scientific notation with 0.00E+00', () => {
    const result = formatCell(12345, '0.00E+00');
    // Excel pads the exponent to the two '0' placeholders after 'E+'.
    expect(result).toBe('1.23E+04');
  });

  it('returns correct built-in format codes via getBuiltinFormat', () => {
    expect(getBuiltinFormat(0)).toBe('General');
    expect(getBuiltinFormat(1)).toBe('0');
    expect(getBuiltinFormat(2)).toBe('0.00');
    expect(getBuiltinFormat(9)).toBe('0%');
    expect(getBuiltinFormat(10)).toBe('0.00%');
    expect(getBuiltinFormat(14)).toBe('mm-dd-yy');
    expect(getBuiltinFormat(49)).toBe('@');
    expect(getBuiltinFormat(999)).toBeUndefined();
  });

  it('returns empty string for null and undefined', () => {
    expect(formatCell(null, 'General')).toBe('');
    expect(formatCell(undefined as unknown as null, '#,##0')).toBe('');
  });

  it('returns string as-is for text format @', () => {
    expect(formatCell('hello world', '@')).toBe('hello world');
    expect(formatCell(42, '@')).toBe('42');
  });

  it('formats integer percentages with 0%', () => {
    expect(formatCell(0.75, '0%')).toBe('75%');
  });

  it('handles built-in format ID 0 as General', () => {
    expect(formatCell(42, 0)).toBe('42');
  });
});

describe('extended number formatting', () => {
  describe('conditional sections', () => {
    it('applies [>100] condition', () => {
      const fmt = '[>100]"High";[<=100]"Low"';
      expect(formatCellRich(200, fmt).text).toBe('High');
      expect(formatCellRich(50, fmt).text).toBe('Low');
    });

    it('applies [Red] condition with number', () => {
      const fmt = '[Red][>100]#,##0;[Blue]#,##0';
      const r1 = formatCellRich(200, fmt);
      expect(r1.text).toBe('200');
      expect(r1.color).toBe('Red');

      const r2 = formatCellRich(50, fmt);
      expect(r2.text).toBe('50');
      expect(r2.color).toBe('Blue');
    });
  });

  describe('bracket color codes', () => {
    it('extracts named colors', () => {
      expect(formatCellRich(42, '[Red]0').color).toBe('Red');
      expect(formatCellRich(42, '[Blue]0.00').color).toBe('Blue');
      expect(formatCellRich(42, '[Green]#,##0').color).toBe('Green');
    });

    it('extracts indexed colors', () => {
      expect(formatCellRich(42, '[Color3]0').color).toBe('Color3');
      expect(formatCellRich(42, '[Color56]0').color).toBe('Color56');
    });

    it('returns no color when none specified', () => {
      expect(formatCellRich(42, '0.00').color).toBeUndefined();
    });
  });

  describe('loadFormat / loadFormatTable', () => {
    it('registers and uses a custom format', () => {
      loadFormat('#,##0.000', 200);
      expect(formatCellRich(1234.5, 200).text).toBe('1,234.500');
    });

    it('bulk-registers formats', () => {
      loadFormatTable({ 201: '0.0%', 202: '#,##0.00' });
      expect(formatCellRich(0.456, 201).text).toBe('45.6%');
      expect(formatCellRich(1234, 202).text).toBe('1,234.00');
    });
  });
});

/** 2023-03-15 12:00, a Wednesday; SheetJS SSF 0.20.3 gives the same text. */
const MID_MARCH = 45000.5;

describe('year tokens', () => {
  it('shows a one-letter year as two digits', () => {
    expect(formatCell(MID_MARCH, 'd/m/y')).toBe('15/3/23');
  });

  it('shows three or more y as the four-digit year', () => {
    expect(formatCell(MID_MARCH, 'yyy')).toBe('2023');
    expect(formatCell(MID_MARCH, 'yyyyy')).toBe('2023');
  });

  it('keeps yy and yyyy', () => {
    expect(formatCell(MID_MARCH, 'yy')).toBe('23');
    expect(formatCell(MID_MARCH, 'yyyy')).toBe('2023');
  });

  it('shows the era year e as the four-digit year', () => {
    expect(formatCell(MID_MARCH, 'd-m-e')).toBe('15-3-2023');
  });

  it('shows a run of e as the year once', () => {
    expect(formatCell(45000.5, 'd ee')).toBe('15 2023');
    expect(formatCell(45000.5, 'd eee')).toBe('15 2023');
  });

  it('keeps E in a scientific code out of the date path', () => {
    expect(formatCell(1234.5, '0.00E+00')).toBe('1.23E+03');
  });
});

describe('A/P meridiem token', () => {
  it('shows the 12-hour clock and P after noon in the code letter case', () => {
    expect(formatCell(45000.5541, 'h:mm A/P')).toBe('1:17 P');
  });

  it('shows a before noon in lower case', () => {
    expect(formatCell(0.25, 'h:mm a/p')).toBe('6:00 a');
  });

  it('keeps the case of each letter', () => {
    expect(formatCell(45000.5541, 'h:mm A/p')).toBe('1:17 p');
  });

  it('leaves AM/PM unchanged', () => {
    expect(formatCell(45000.5541, 'h:mm AM/PM')).toBe('1:17 PM');
    expect(formatCell(0.25, 18)).toBe('6:00 AM');
  });

  it('counts A/P as a date token', () => {
    expect(isDateFormatCode('h A/P')).toBe(true);
    expect(isDateFormatCode('A/P')).toBe(true);
  });
});

describe('built-in date ids without a code', () => {
  it.each([27, 30, 50, 58])('id %i renders like id 14', (id) => {
    expect(formatCell(MID_MARCH, id)).toBe(formatCell(MID_MARCH, 14));
  });

  it('keeps General for an unknown non-date id', () => {
    expect(formatCell(MID_MARCH, 5)).toBe('45000.5');
  });

  it('lets a registration for a date id win', () => {
    loadFormat('0.000', 30);
    try {
      expect(formatCell(1, 30)).toBe('1.000');
    } finally {
      loadFormat(getBuiltinFormat(14) ?? 'General', 30);
    }
  });
});
