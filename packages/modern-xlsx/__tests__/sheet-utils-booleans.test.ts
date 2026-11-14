import { describe, expect, it } from 'vitest';
import {
  aoaToSheet,
  jsonToSheet,
  readBuffer,
  sheetAddAoa,
  sheetAddJson,
  sheetToCsv,
  sheetToHtml,
  sheetToJson,
  sheetToTxt,
  Workbook,
  type Worksheet,
} from '../src/index.js';
import type { CellType } from '../src/model/types.js';
import { buildZip, readZipText } from './helpers/zip.js';

/** A one-row sheet whose cells carry exactly these declared types and stored values. */
function typedRow(cells: readonly (readonly [CellType, string])[]): Worksheet {
  const ws = aoaToSheet([cells.map(() => 'placeholder')]);
  const row = ws.rows[0];
  if (!row) throw new Error('no row');
  cells.forEach(([cellType, value], i) => {
    const cell = row.cells[i];
    if (!cell) throw new Error(`no cell ${i}`);
    cell.cellType = cellType;
    cell.value = value;
  });
  return ws;
}

const html = (cells: readonly string[]) =>
  `<table><tr>${cells.map((c) => `<td>${c}</td>`).join('')}</tr></table>`;

describe('sheetToCsv, sheetToTxt and sheetToHtml write booleans as Excel does', () => {
  // Excel 365 saving TRUE, FALSE, the text "true" and the text "TRUE" as CSV writes
  // `TRUE,FALSE,true,TRUE` and as Unicode or tab-delimited text the same fields joined by TAB.
  const excelRow = () => aoaToSheet([[true, false, 'true', 'TRUE']]);

  it('sheetToCsv writes TRUE and FALSE for boolean cells and leaves text alone', () => {
    expect(sheetToCsv(excelRow())).toBe('TRUE,FALSE,true,TRUE');
  });

  it('sheetToTxt writes TRUE and FALSE for boolean cells and leaves text alone', () => {
    expect(sheetToTxt(excelRow())).toBe('TRUE\tFALSE\ttrue\tTRUE');
  });

  it('sheetToHtml writes TRUE and FALSE for boolean cells and leaves text alone', () => {
    expect(sheetToHtml(excelRow())).toBe(html(['TRUE', 'FALSE', 'true', 'TRUE']));
  });

  it.each([
    ['1', 'TRUE'],
    ['0', 'FALSE'],
    ['true', 'TRUE'],
    ['True', 'TRUE'],
    ['false', 'FALSE'],
    ['fAlSe', 'FALSE'],
  ])('a boolean cell storing %s renders as %s in all three', (stored, shown) => {
    const ws = typedRow([['boolean', stored]]);
    expect([sheetToCsv(ws), sheetToTxt(ws), sheetToHtml(ws)]).toEqual([
      shown,
      shown,
      html([shown]),
    ]);
  });

  it('a boolean cell with any other stored value passes through unchanged', () => {
    const ws = typedRow([
      ['boolean', '1'],
      ['boolean', '2'],
      ['boolean', 'yes'],
      ['boolean', ''],
      ['boolean', 'a<b'],
    ]);
    expect(sheetToCsv(ws)).toBe('TRUE,2,yes,,a<b');
    expect(sheetToTxt(ws)).toBe('TRUE\t2\tyes\t\ta<b');
    expect(sheetToHtml(ws)).toBe(html(['TRUE', '2', 'yes', '', 'a&lt;b']));
  });

  it('only boolean cells change: a number 1 and a string 1 stay 1', () => {
    const ws = typedRow([
      ['boolean', '1'],
      ['number', '1'],
      ['sharedString', '1'],
      ['sharedString', 'FALSE'],
    ]);
    expect(sheetToCsv(ws)).toBe('TRUE,1,1,FALSE');
    expect(sheetToHtml(ws, { includeStyle: true })).toBe(
      '<table><tr><td style="text-align:center">TRUE</td><td style="text-align:right">1</td><td>1</td><td>FALSE</td></tr></table>',
    );
  });

  it('escapeFormulas leaves TRUE and FALSE unquoted and still defuses a boolean-typed trigger', () => {
    const ws = typedRow([
      ['boolean', '1'],
      ['boolean', '0'],
      ['boolean', '=x'],
      ['sharedString', 'TRUE'],
    ]);
    expect(sheetToCsv(ws, { escapeFormulas: true })).toBe(`TRUE,FALSE,"'=x",TRUE`);
    expect(sheetToTxt(ws, { escapeFormulas: true })).toBe(`TRUE\tFALSE\t"'=x"\tTRUE`);
  });

  it('a value built by hand that is not a string (true, 1, 5) renders instead of throwing', () => {
    const ws = typedRow([
      ['boolean', '1'],
      ['boolean', '1'],
      ['number', '1'],
    ]);
    const cells = ws.rows[0]?.cells ?? [];
    // Outside the types: CellData.value is a string, but hand-built data may hold true or 1.
    Object.assign(cells[0] ?? {}, { value: true });
    Object.assign(cells[1] ?? {}, { value: 1 });
    Object.assign(cells[2] ?? {}, { value: 5 });
    expect(sheetToHtml(ws)).toBe(html(['TRUE', 'TRUE', '5']));
    expect(sheetToCsv(ws)).toBe('TRUE,TRUE,5');
    expect(sheetToTxt(ws)).toBe('TRUE\tTRUE\t5');
  });

  it('forceQuote quotes TRUE and FALSE like any other field', () => {
    expect(sheetToCsv(aoaToSheet([[true, false]]), { forceQuote: true })).toBe('"TRUE","FALSE"');
  });
});

describe('booleans round-trip through the import helpers', () => {
  it('aoaToSheet, jsonToSheet, sheetAddAoa and sheetAddJson store true and false as booleans', () => {
    const fromAoa = aoaToSheet([[true, false]]);
    const fromJson = jsonToSheet([{ a: true, b: false }], { skipHeader: true });
    const added = aoaToSheet([]);
    sheetAddAoa(added, [[true, false]]);
    const addedJson = aoaToSheet([]);
    sheetAddJson(addedJson, [{ a: true, b: false }], { skipHeader: true });
    for (const ws of [fromAoa, fromJson, added, addedJson]) {
      expect(sheetToCsv(ws)).toBe('TRUE,FALSE');
      expect(sheetToJson(ws, { header: 'A' })).toEqual([{ A: true, B: false }]);
    }
  });

  it('the text "TRUE" stays text: aoaToSheet does not parse it into a boolean', () => {
    const ws = aoaToSheet([['TRUE', 'false', true]]);
    expect(ws.rows[0]?.cells.map((c) => c.cellType)).toEqual([
      'sharedString',
      'sharedString',
      'boolean',
    ]);
    expect(sheetToCsv(ws)).toBe('TRUE,false,TRUE');
    expect(sheetToJson(ws, { header: 'A' })).toEqual([{ A: 'TRUE', B: 'false', C: true }]);
  });

  it('a written and re-read workbook still renders TRUE and FALSE', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = true;
    ws.cell('B1').value = false;
    ws.cell('C1').value = 'TRUE';
    const back = (await readBuffer(await wb.toBuffer())).getSheet('S');
    if (!back) throw new Error('no sheet S');
    expect(sheetToCsv(back)).toBe('TRUE,FALSE,TRUE');
    expect(sheetToTxt(back)).toBe('TRUE\tFALSE\tTRUE');
    expect(sheetToHtml(back)).toBe(html(['TRUE', 'FALSE', 'TRUE']));
  });

  it('a sheet as Excel saves it (t="b" values and cached =1=1 results) matches Excel’s own CSV', async () => {
    // The sheetData Excel 365 wrote for TRUE, FALSE, text "TRUE", =1=1, =1=2, text "true" and
    // text "x,y" (shared strings 0, 1 and 2); Excel saved that sheet as CSV as
    // `TRUE,FALSE,TRUE,TRUE,FALSE,true,"x,y"`.
    const sheetData =
      '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c><c r="C1" t="s"><v>0</v></c>' +
      '<c r="D1" t="b"><f>1=1</f><v>1</v></c><c r="E1" t="b"><f>1=2</f><v>0</v></c>' +
      '<c r="F1" t="s"><v>1</v></c><c r="G1" t="s"><v>2</v></c></row>';
    const wb = new Workbook();
    const seed = wb.addSheet('S');
    seed.cell('A1').value = 'TRUE';
    seed.cell('B1').value = 'true';
    seed.cell('C1').value = 'x,y';
    const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${main}><sheetData>${sheetData}</sheetData></worksheet>`;
    const enc = new TextEncoder();
    const parts = Object.entries(readZipText(await wb.toBuffer())).map(
      ([name, body]) =>
        [name, enc.encode(name === 'xl/worksheets/sheet1.xml' ? sheet : body)] as const,
    );
    const ws = (await readBuffer(buildZip(parts))).getSheet('S');
    if (!ws) throw new Error('no sheet S');
    expect(sheetToCsv(ws)).toBe('TRUE,FALSE,TRUE,TRUE,FALSE,true,"x,y"');
    expect(sheetToTxt(ws)).toBe('TRUE\tFALSE\tTRUE\tTRUE\tFALSE\ttrue\tx,y');
  });
});
