import { describe, expect, it } from 'vitest';
import type { TableColumnData, TableDefinitionData } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

/** A `ref`-sized table named `Table1` with a header row and a totals row. */
function table(ref: string, columns: Omit<TableColumnData, 'id'>[]): TableDefinitionData {
  return {
    id: 1,
    name: 'Table1',
    displayName: 'Table1',
    ref,
    headerRowCount: 1,
    totalsRowCount: 1,
    totalsRowShown: true,
    columns: columns.map((c, i) => ({ id: i + 1, ...c })),
  };
}

/** Locations of the totals-cell errors `wb.validate()` reports. */
function totalsErrors(wb: Workbook): string[] {
  return wb
    .validate()
    .issues.filter((i) => i.severity === 'error' && i.message.includes('totals cell'))
    .map((i) => i.location);
}

describe('addTable writes the totals row cells', () => {
  // Function numbers and structured references as Excel 365 saves them (COM ground truth).
  it.each([
    ['sum', 109],
    ['average', 101],
    ['count', 103],
    ['countNums', 102],
    ['max', 104],
    ['min', 105],
    ['stdDev', 107],
    ['var', 110],
  ])('writes SUBTOTAL(%s -> %i)', (totalsRowFunction, code) => {
    const ws = new Workbook().addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.addTable(table('A1:A3', [{ name: 'Amount', totalsRowFunction }]));
    expect(ws.cell('A3').formula).toBe(`SUBTOTAL(${code},Table1[Amount])`);
  });

  it('places each column in its own cell of the last row, from the table start column', () => {
    const ws = new Workbook().addSheet('Data');
    ws.addTable(
      table('C2:E6', [
        { name: 'Item', totalsRowLabel: 'Total' },
        { name: 'Qty', totalsRowFunction: 'sum' },
        { name: 'Price', totalsRowFunction: 'average' },
      ]),
    );
    expect(ws.cell('C6').value).toBe('Total');
    expect(ws.cell('D6').formula).toBe('SUBTOTAL(109,Table1[Qty])');
    expect(ws.cell('E6').formula).toBe('SUBTOTAL(101,Table1[Price])');
    expect(ws.cell('C2').value).toBeNull();
  });

  it("escapes [ ] # and ' in a column name with a leading apostrophe", () => {
    const ws = new Workbook().addSheet('Data');
    ws.addTable(
      table('A1:E3', [
        { name: 'Space Name', totalsRowFunction: 'sum' },
        { name: 'Br[ack]et', totalsRowFunction: 'sum' },
        { name: 'Hash#', totalsRowFunction: 'sum' },
        { name: "Apos'trophe", totalsRowFunction: 'sum' },
        { name: 'a.b,c:d(e)', totalsRowFunction: 'sum' },
      ]),
    );
    expect(ws.cell('A3').formula).toBe('SUBTOTAL(109,Table1[Space Name])');
    expect(ws.cell('B3').formula).toBe("SUBTOTAL(109,Table1[Br'[ack']et])");
    expect(ws.cell('C3').formula).toBe("SUBTOTAL(109,Table1[Hash'#])");
    expect(ws.cell('D3').formula).toBe("SUBTOTAL(109,Table1[Apos''trophe])");
    expect(ws.cell('E3').formula).toBe('SUBTOTAL(109,Table1[a.b,c:d(e)])');
  });

  it('writes a custom formula without a leading =', () => {
    const ws = new Workbook().addSheet('Data');
    ws.addTable(
      table('A1:B3', [
        { name: 'Custom', totalsRowFunction: 'custom', totalsRowFormula: '=SUM(Table1[Custom])*2' },
        { name: 'Other', totalsRowFunction: 'custom', totalsRowFormula: 'MAX(Table1[Other])' },
      ]),
    );
    expect(ws.cell('A3').formula).toBe('SUM(Table1[Custom])*2');
    expect(ws.cell('B3').formula).toBe('MAX(Table1[Other])');
  });

  it('keeps a cell the caller filled, and validate() reports the kept cells Excel refuses', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A3').value = 'mine';
    ws.cell('B3').formula = 'A1+1';
    ws.cell('C3').value = 7;
    ws.addTable(
      table('A1:D3', [
        { name: 'A', totalsRowLabel: 'Total' },
        { name: 'B', totalsRowFunction: 'sum' },
        { name: 'C', totalsRowFunction: 'sum' },
        { name: 'D', totalsRowFunction: 'sum' },
      ]),
    );
    expect(ws.cell('A3').value).toBe('mine');
    expect(ws.cell('B3').formula).toBe('A1+1');
    expect(ws.cell('C3').value).toBe(7);
    expect(ws.cell('C3').formula).toBeNull();
    expect(ws.cell('D3').formula).toBe('SUBTOTAL(109,Table1[D])');
    // Excel refuses a label other than totalsRowLabel and a formula other than
    // the column's SUBTOTAL; it opens a plain value in a SUBTOTAL column.
    expect(totalsErrors(wb)).toEqual(['Data!A3', 'Data!B3']);
  });

  it('writes nothing for a table without a totals row or for columns without a function', async () => {
    const ws = new Workbook().addSheet('Data');
    ws.addTable({
      ...table('A1:B3', [
        { name: 'A', totalsRowFunction: 'sum' },
        { name: 'B', totalsRowFunction: 'none' },
      ]),
      totalsRowCount: 0,
    });
    ws.addTable({
      ...table('D1:E3', [{ name: 'D' }, { name: 'E', totalsRowFunction: 'none' }]),
      id: 2,
      name: 'Table2',
      displayName: 'Table2',
    });
    expect(ws.rows).toHaveLength(0);
  });

  it('writes the formulas into the sheet part', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.cell('A2').value = 5;
    ws.addTable(table('A1:A3', [{ name: 'Amount', totalsRowFunction: 'sum' }]));
    const xml = readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'];
    expect(xml).toContain('<f>SUBTOTAL(109,Table1[Amount])</f>');
  });

  it('does not touch a table read from a file', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.addTable(table('A1:A3', [{ name: 'Amount', totalsRowFunction: 'sum' }]));
    const back = await readBuffer(await wb.toBuffer());
    const read = back.getSheet('Data');
    expect(read?.cell('A3').formula).toBe('SUBTOTAL(109,Table1[Amount])');
    // The caller edits the totals cell of the read table to a plain value (a
    // SUBTOTAL column holds its formula or a value in a file Excel opens); a
    // second write keeps the edit and does not bring the formula back.
    if (read) {
      read.cell('A3').formula = null;
      read.cell('A3').value = 42;
    }
    expect(back.validate().errorCount).toBe(0);
    const again = await readBuffer(await back.toBuffer());
    expect(again.getSheet('Data')?.cell('A3').formula).toBeNull();
    expect(again.getSheet('Data')?.cell('A3').value).toBe(42);
    // A formula other than the column's SUBTOTAL is kept too, and validate() reports it.
    const edited = again.getSheet('Data');
    if (edited) edited.cell('A3').formula = 'SUM(A2:A2)';
    expect(totalsErrors(again)).toEqual(['Data!A3']);
    const kept = await readBuffer(await again.toBuffer());
    expect(kept.getSheet('Data')?.cell('A3').formula).toBe('SUM(A2:A2)');
    // A read table with an empty totals cell stays empty.
    const bare = new Workbook();
    bare.addSheet('Data').cell('A1').value = 'x';
    bare.addSheet('Data2');
    const raw = bare.getSheet('Data');
    if (raw) {
      (
        raw as unknown as { data: { worksheet: { tables: TableDefinitionData[] } } }
      ).data.worksheet.tables = [table('A1:A3', [{ name: 'x', totalsRowFunction: 'sum' }])];
    }
    const reread = await readBuffer(await bare.toBuffer());
    expect(reread.getSheet('Data')?.tables).toHaveLength(1);
    expect(reread.getSheet('Data')?.cell('A3').formula).toBeNull();
  });
});

describe('validate() reports a totals cell Excel refuses', () => {
  /** The workbooks these cases share: Table1 over A1:B4, Item (label Total) and Amount (sum). */
  function amounts(before: (ws: ReturnType<Workbook['addSheet']>) => void = () => {}) {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    for (const [r, row] of [
      ['Item', 'Amount'],
      ['a', 1],
      ['b', 2],
    ].entries()) {
      ws.cell(`A${r + 1}`).value = row[0] ?? null;
      ws.cell(`B${r + 1}`).value = row[1] ?? null;
    }
    before(ws);
    ws.addTable(
      table('A1:B4', [
        { name: 'Item', totalsRowLabel: 'Total' },
        { name: 'Amount', totalsRowFunction: 'sum' },
      ]),
    );
    return wb;
  }

  // Excel 365: v5 and v8 open; v6, v7 and the clone v3 are refused.
  it.each([
    ['v5: a number in a SUBTOTAL column', (ws) => (ws.cell('B4').value = 7), []],
    ['v8: text in a SUBTOTAL column', (ws) => (ws.cell('B4').value = 'n/a'), []],
    ['v6: a label other than totalsRowLabel', (ws) => (ws.cell('A4').value = 'mine'), ['Data!A4']],
    [
      'v7: a formula other than the SUBTOTAL',
      (ws) => (ws.cell('B4').formula = 'B2+B3'),
      ['Data!B4'],
    ],
  ] satisfies [string, (ws: ReturnType<Workbook['addSheet']>) => unknown, string[]][])(
    '%s',
    (_, before, expected) => {
      expect(totalsErrors(amounts(before))).toEqual(expected);
    },
  );

  it('v3: a plain formula in the totals cell of a cloned table', () => {
    const wb = amounts();
    wb.cloneSheet(0, 'Copy').cell('B4').formula = 'SUM(B2:B3)';
    expect(totalsErrors(wb)).toEqual(['Copy!B4']);
  });

  it('reports nothing for the totals rows addTable writes, nor after a round trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.addTable(
      table('A1:L3', [
        { name: 'Item', totalsRowLabel: 'Total' },
        ...['sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var'].map(
          (totalsRowFunction) => ({ name: `F ${totalsRowFunction}`, totalsRowFunction }),
        ),
        { name: 'Custom', totalsRowFunction: 'custom', totalsRowFormula: '=SUM(Table1[Custom])*2' },
        { name: "Br[ack]et#'", totalsRowFunction: 'sum' },
        { name: 'None', totalsRowFunction: 'none' },
      ]),
    );
    expect(wb.validate().issues.filter((i) => i.severity === 'error')).toEqual([]);
    const back = await readBuffer(await wb.toBuffer());
    expect(back.validate().issues.filter((i) => i.severity === 'error')).toEqual([]);
  });
});
