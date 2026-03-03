import { describe, expect, it } from 'vitest';
import { sheetToFormulae, sheetToTxt, Workbook } from '../src/index.js';

describe('extended sheet conversion utilities', () => {
  describe('sheetToTxt', () => {
    it('produces tab-separated output', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'Name';
      ws.cell('B1').value = 'Age';
      ws.cell('A2').value = 'Alice';
      ws.cell('B2').value = 30;

      const txt = sheetToTxt(ws);
      const lines = txt.split('\n');
      expect(lines[0]).toBe('Name\tAge');
      expect(lines[1]).toBe('Alice\t30');
    });

    it('respects sheetRows limit', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'r1';
      ws.cell('A2').value = 'r2';
      ws.cell('A3').value = 'r3';

      const txt = sheetToTxt(ws, { sheetRows: 2 });
      expect(txt.split('\n')).toHaveLength(2);
    });
  });

  describe('sheetToFormulae', () => {
    it('returns cell references with values and formulas', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 100;
      ws.cell('A2').value = 200;
      ws.cell('A3').formula = 'SUM(A1:A2)';

      const formulae = sheetToFormulae(ws);
      expect(formulae).toContain('A1=100');
      expect(formulae).toContain('A2=200');
      expect(formulae).toContain("A3='SUM(A1:A2)");
    });

    it('handles string values', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'hello';

      const formulae = sheetToFormulae(ws);
      expect(formulae).toContain("A1='hello");
    });

    it('handles empty sheet', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(sheetToFormulae(ws)).toEqual([]);
    });
  });
});
