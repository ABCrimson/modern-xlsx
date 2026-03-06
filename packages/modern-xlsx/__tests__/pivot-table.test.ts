import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('Pivot Tables', () => {
  it('empty sheet has no pivot tables', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.pivotTables).toHaveLength(0);
  });

  it('pivot tables getter returns readonly array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const pivots = ws.pivotTables;
    expect(Array.isArray(pivots)).toBe(true);
    expect(pivots).toHaveLength(0);
  });
});
