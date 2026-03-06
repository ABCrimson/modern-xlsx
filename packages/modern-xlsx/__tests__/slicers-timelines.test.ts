import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/index.js';

describe('Slicers', () => {
  it('empty sheet has no slicers', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.slicers).toHaveLength(0);
  });

  it('slicers getter returns readonly array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const slicers = ws.slicers;
    expect(Array.isArray(slicers)).toBe(true);
    expect(slicers).toEqual([]);
  });
});

describe('Timelines', () => {
  it('empty sheet has no timelines', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.timelines).toHaveLength(0);
  });

  it('timelines getter returns readonly array', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const timelines = ws.timelines;
    expect(Array.isArray(timelines)).toBe(true);
    expect(timelines).toEqual([]);
  });
});
