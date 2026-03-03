import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('dynamic array formulas', () => {
  it('roundtrips dynamicArray flag', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 1;
    ws.cell('A2').value = 2;

    // Set up the raw cell data to have dynamicArray flag
    const cellData = ws.rows[0]?.cells[0];
    if (cellData) {
      cellData.formula = 'SORT(A1:A2)';
      cellData.formulaType = 'array';
      cellData.dynamicArray = true;
    }

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const cell2 = wb2.getSheet('S')?.rows[0]?.cells[0];
    expect(cell2?.dynamicArray).toBe(true);
  });

  it('does not set dynamicArray for non-dynamic formulas', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    const cell = ws.cell('A1');
    cell.formula = 'SUM(B1:B10)';
    cell.value = 55;

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const cell2 = wb2.getSheet('S')?.rows[0]?.cells[0];
    expect(cell2?.dynamicArray).toBeUndefined();
  });
});
