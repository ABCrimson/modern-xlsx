import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('extended worksheet operations', () => {
  describe('usedRange', () => {
    it('computes used range from cells', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('B2').value = 'hello';
      ws.cell('D5').value = 42;

      expect(ws.usedRange).toBe('B2:D5');
    });

    it('returns null for empty sheet', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(ws.usedRange).toBeNull();
    });

    it('handles single cell', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('C3').value = 'only';
      expect(ws.usedRange).toBe('C3:C3');
    });

    it('handles wide range', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.cell('A1').value = 'start';
      ws.cell('Z100').value = 'end';
      expect(ws.usedRange).toBe('A1:Z100');
    });
  });

  describe('tabColor', () => {
    it('roundtrips tab color', async () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Colored');
      ws.cell('A1').value = 1;
      ws.tabColor = 'FF0000';

      expect(ws.tabColor).toBe('FF0000');

      const buf = await wb.toBuffer();
      const wb2 = await readBuffer(buf);
      expect(wb2.getSheet('Colored')?.tabColor).toBe('FF0000');
    });

    it('defaults to null', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      expect(ws.tabColor).toBeNull();
    });

    it('can be cleared', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.tabColor = 'FF0000';
      ws.tabColor = null;
      expect(ws.tabColor).toBeNull();
    });
  });
});
