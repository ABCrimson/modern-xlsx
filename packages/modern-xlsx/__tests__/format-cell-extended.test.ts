import { describe, expect, it } from 'vitest';
import { formatCellRich, loadFormat, loadFormatTable } from '../src/index.js';

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
