import { describe, expect, it } from 'vitest';
import { sheetToCsv, sheetToJson } from '../src/utils.js';
import { Workbook } from '../src/workbook.js';

/**
 * CLI unit tests.
 *
 * Rather than spawning a subprocess (fragile), we test the underlying
 * functions that the CLI calls: `sheetToJson`, `sheetToCsv`, and the
 * `Worksheet.dimension` / `Worksheet.rowCount` getters.
 */
describe('CLI underlying functions', () => {
  it('Worksheet.dimension returns the source dimension', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    // New sheet has no dimension
    expect(ws.dimension).toBeNull();
  });

  it('Worksheet.rowCount returns the number of rows', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    expect(ws.rowCount).toBe(0);
    ws.cell('A1').value = 'hello';
    expect(ws.rowCount).toBe(1);
    ws.cell('A5').value = 'world';
    expect(ws.rowCount).toBe(2);
  });

  it('sheetToJson converts worksheet to JSON objects', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;
    ws.cell('A3').value = 'Bob';
    ws.cell('B3').value = 25;

    const json = sheetToJson(ws);
    expect(json).toHaveLength(2);
    expect(json[0]).toEqual({ Name: 'Alice', Age: 30 });
    expect(json[1]).toEqual({ Name: 'Bob', Age: 25 });
  });

  it('sheetToCsv converts worksheet to CSV string', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;

    const csv = sheetToCsv(ws);
    expect(csv).toContain('Name');
    expect(csv).toContain('Alice');
    expect(csv).toContain('30');
  });

  it('info command properties are accessible', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addSheet('Sheet2');
    expect(wb.sheetCount).toBe(2);
    expect(wb.sheetNames).toEqual(['Sheet1', 'Sheet2']);
    for (let i = 0; i < wb.sheetCount; i++) {
      const ws = wb.getSheetByIndex(i);
      expect(ws).toBeDefined();
      expect(ws?.rowCount).toBe(0);
      expect(ws?.dimension).toBeNull();
    }
  });
});
