import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('extended cell operations', () => {
  describe('cell.dateValue', () => {
    it('returns Date for date-formatted number', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');

      // Apply a date style
      const styleIdx = wb.createStyle().numberFormat('yyyy-mm-dd').build(wb.styles);
      const cell = ws.cell('A1');
      cell.value = 46082; // 2026-03-01
      cell.styleIndex = styleIdx;

      const d = cell.dateValue;
      expect(d).toBeInstanceOf(Date);
      expect(d?.getUTCFullYear()).toBe(2026);
      expect(d?.getUTCMonth()).toBe(2); // March = 2
      expect(d?.getUTCDate()).toBe(1);
    });

    it('returns null for non-date cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 42;
      expect(ws.cell('A1').dateValue).toBeNull();
    });

    it('returns null for string cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'hello';
      expect(ws.cell('A1').dateValue).toBeNull();
    });
  });

  describe('cell.numberFormat', () => {
    it('reads the number format from style', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      const styleIdx = wb.createStyle().numberFormat('#,##0.00').build(wb.styles);
      const cell = ws.cell('A1');
      cell.value = 42;
      cell.styleIndex = styleIdx;

      expect(cell.numberFormat).toBe('#,##0.00');
    });

    it('returns null when no custom format', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 42;
      expect(ws.cell('A1').numberFormat).toBeNull();
    });

    it('reads builtin format by id', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      // numFmtId=14 is builtin 'm/d/yy'
      wb.styles.cellXfs.push({
        numFmtId: 14,
        fontId: 0,
        fillId: 0,
        borderId: 0,
      });
      const cell = ws.cell('A1');
      cell.value = 46082;
      cell.styleIndex = wb.styles.cellXfs.length - 1;

      expect(cell.numberFormat).toBe('mm-dd-yy');
    });
  });
});
