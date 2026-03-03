import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('Sheet Management', () => {
  // --- Sheet State ---

  it('default sheet state is visible', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.state).toBe('visible');
  });

  it('set sheet state to hidden', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws2.state = 'hidden';
    expect(ws2.state).toBe('hidden');
  });

  it('hidden sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    const ws2 = wb.addSheet('Sheet2');
    ws2.cell('A1').value = 'hidden';
    ws2.state = 'hidden';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Sheet2')?.state).toBe('hidden');
    expect(wb2.getSheet('Sheet1')?.state).toBe('visible');
  });

  it('cannot hide last visible sheet', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => wb.hideSheet('Sheet1')).toThrow('last visible');
  });

  it('unhide restores visible state', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws2.state = 'hidden';
    wb.unhideSheet('Sheet2');
    expect(wb.getSheet('Sheet2')?.state).toBe('visible');
  });

  // --- Move Sheet ---

  it('move sheet changes order', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');
    wb.moveSheet(2, 0);
    expect(wb.sheetNames).toEqual(['C', 'A', 'B']);
  });

  // --- Clone Sheet ---

  it('clone sheet duplicates content', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Original');
    ws.cell('A1').value = 'hello';
    const clone = wb.cloneSheet(0, 'Copy');
    expect(clone.cell('A1').value).toBe('hello');
    expect(wb.sheetCount).toBe(2);
  });

  it('clone sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'original';
    wb.cloneSheet(0, 'Clone');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(2);
    expect(wb2.getSheet('Clone')?.cell('A1').value).toBe('original');
  });

  // --- Rename Sheet ---

  it('rename sheet updates name', () => {
    const wb = new Workbook();
    wb.addSheet('OldName');
    wb.renameSheet('OldName', 'NewName');
    expect(wb.sheetNames).toEqual(['NewName']);
    expect(wb.getSheet('NewName')).toBeDefined();
  });

  it('rename rejects duplicate name', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(() => wb.renameSheet('A', 'B')).toThrow('already exists');
  });
});
