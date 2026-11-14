import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('removeSheet rejects non-integer indexes', () => {
  for (const bad of [Number.NaN, 0.5, Number.POSITIVE_INFINITY]) {
    it(`removeSheet(${bad}) returns false and changes nothing`, () => {
      const wb = new Workbook();
      wb.addSheet('X');
      wb.addSheet('Y');
      wb.addNamedRange('OnX', 'X!$A$1', 0);

      expect(wb.removeSheet(bad)).toBe(false);
      expect(wb.sheetNames).toEqual(['X', 'Y']);
      expect(wb.namedRanges).toEqual([{ name: 'OnX', value: 'X!$A$1', sheetId: 0 }]);
    });
  }
});

describe('Worksheet remove* rejects non-integer indexes', () => {
  for (const bad of [Number.NaN, 0.5]) {
    it(`removeChart(${bad}) returns false and keeps the chart`, () => {
      const ws = new Workbook().addSheet('Sheet1');
      ws.addChart('bar', (b) => {
        b.title('Only').addSeries({ valRef: 'Sheet1!$A$1:$A$5' });
      });

      expect(ws.removeChart(bad)).toBe(false);
      expect(ws.charts).toHaveLength(1);
    });
  }
});
