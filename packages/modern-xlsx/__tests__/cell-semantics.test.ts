import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { INVALID_ARGUMENT, INVALID_CELL_VALUE, ModernXlsxError } from '../src/model/errors.js';
import type { CellData, ColumnInfo, SheetData, WorkbookData } from '../src/model/types.js';
import { readZipText } from './helpers/zip.js';

function seeded(
  rows: { index: number; cells: CellData[] }[],
  extra: Partial<WorkbookData> = {},
  columns: ColumnInfo[] = [],
): Workbook {
  const sheet: SheetData = {
    name: 'S',
    worksheet: {
      dimension: null,
      rows: rows.map((r) => ({ ...r, height: null, hidden: false })),
      mergeCells: [],
      autoFilter: null,
      frozenPane: null,
      columns,
    },
  };
  return new Workbook({ sheets: [sheet], ...extra });
}

function num(reference: string, value: string): CellData {
  return { reference, cellType: 'number', value, formula: null, styleIndex: null };
}

describe('cell value semantics', () => {
  it('clears stale richText when the value is set', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    const cell = ws.cell('A1');
    cell.richText = [{ text: 'a', bold: true }, { text: 'b' }];
    expect(cell.richText).toHaveLength(2);
    cell.value = 'x';
    expect(cell.richText).toBeUndefined();
    expect(cell.value).toBe('x');
    cell.richText = [{ text: 'q' }];
    cell.value = null;
    expect(cell.richText).toBeUndefined();
  });

  it('writes a workbook after a string cell is set to null', async () => {
    const wb = new Workbook();
    const cell = wb.addSheet('S').cell('A1');
    cell.value = 'text';
    cell.value = null;
    expect(cell.type).toBe('number');
    const wb2 = await readBuffer(await wb.toBuffer());
    expect(wb2.getSheet('S')?.cell('A1').value).toBeNull();
  });

  it('rejects non-finite numbers with INVALID_CELL_VALUE', () => {
    const cell = new Workbook().addSheet('S').cell('A1');
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expect(() => {
        cell.value = bad;
      }).toThrow(ModernXlsxError);
      try {
        cell.value = bad;
      } catch (e) {
        expect((e as ModernXlsxError).code).toBe(INVALID_CELL_VALUE);
        expect((e as ModernXlsxError).message).toBe(
          `Cell value must be a finite number, got ${bad}`,
        );
      }
    }
  });

  it('reads negative zero back as 0', () => {
    const cell = new Workbook().addSheet('S').cell('A1');
    cell.value = -0;
    expect(Object.is(cell.value, 0)).toBe(true);
  });
});

describe('formula cell types', () => {
  it('is a number for a numeric cached value, before and after a round trip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    const cell = ws.cell('A1');
    cell.formula = 'SUM(1,1)';
    cell.value = 2;
    expect(cell.type).toBe('number');
    const wb2 = await readBuffer(await wb.toBuffer());
    const back = wb2.getSheet('S')?.cell('A1');
    expect(back?.type).toBe('number');
    expect(back?.value).toBe(2);
    expect(back?.formula).toBe('SUM(1,1)');
  });

  it('follows the cached value type', () => {
    const cell = new Workbook().addSheet('S').cell('A1');
    cell.formula = 'A2';
    expect(cell.type).toBe('number');
    cell.value = 'text';
    expect(cell.type).toBe('formulaStr');
    cell.value = true;
    expect(cell.type).toBe('boolean');
    cell.value = 3;
    expect(cell.type).toBe('number');
    cell.value = null;
    expect(cell.type).toBe('number');
  });

  it('keeps the error type while a formula is set', () => {
    const wb = seeded([
      {
        index: 1,
        cells: [
          { reference: 'A1', cellType: 'error', value: '#N/A', formula: 'NA()', styleIndex: null },
        ],
      },
    ]);
    const cell = wb.getSheet('S')?.cell('A1');
    expect(cell?.type).toBe('error');
    if (cell) cell.formula = 'NA()';
    expect(cell?.type).toBe('error');
  });

  it('turns a plain error cell into a sharedString when a string is set', () => {
    const wb = seeded([
      {
        index: 1,
        cells: [
          { reference: 'A1', cellType: 'error', value: '#N/A', formula: null, styleIndex: null },
        ],
      },
    ]);
    const cell = wb.getSheet('S')?.cell('A1');
    if (cell) cell.value = 'hello';
    expect(cell?.type).toBe('sharedString');
  });

  it('turns formulaStr into sharedString (or number) when the formula is cleared', () => {
    const cell = new Workbook().addSheet('S').cell('A1');
    cell.formula = 'A2';
    cell.value = 'text';
    cell.formula = null;
    expect(cell.type).toBe('sharedString');

    const empty = new Workbook().addSheet('S').cell('A1');
    empty.formula = 'A2';
    empty.formula = null;
    expect(empty.type).toBe('number');
  });
});

describe('formula keys and the cell setters', () => {
  const master: CellData = {
    ...num('B1', '2'),
    formula: 'A1*2',
    formulaType: 'shared',
    formulaRef: 'B1:B3',
    sharedIndex: 0,
  };
  const child = (reference: string): CellData => ({
    ...num(reference, '4'),
    formulaType: 'shared',
    sharedIndex: 0,
  });
  const keysOf = (c: CellData | undefined) => [c?.formulaType, c?.formulaRef, c?.sharedIndex];
  const cells = (wb: Workbook) => wb.toJSON().sheets[0]?.worksheet.rows.flatMap((r) => r.cells);

  it('formula = null or an empty string drops formulaType, formulaRef and sharedIndex', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }, { ...master, reference: 'C1' }] }]);
    const ws = wb.getSheet('S');
    if (ws) {
      ws.cell('B1').formula = null;
      ws.cell('C1').formula = '';
    }
    const [b1, c1] = cells(wb) ?? [];
    expect(keysOf(b1)).toEqual([undefined, undefined, undefined]);
    expect(keysOf(c1)).toEqual([undefined, undefined, undefined]);
    expect(b1?.formula).toBeNull();
    expect(b1?.value).toBe('2');
  });

  it('a value set on a shared-formula child replaces its formula', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }, child('C1'), child('D1')] }]);
    const ws = wb.getSheet('S');
    if (ws) {
      ws.cell('C1').value = 42;
      ws.cell('D1').value = null;
    }
    const [, c1, d1] = cells(wb) ?? [];
    expect(keysOf(c1)).toEqual([undefined, undefined, undefined]);
    expect(c1?.value).toBe('42');
    expect(keysOf(d1)).toEqual([undefined, undefined, undefined]);
    expect(d1?.value).toBeNull();
  });

  it('a value set on a cell with formula text keeps its formula keys', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }] }]);
    const b1 = wb.getSheet('S')?.cell('B1');
    if (b1) b1.value = 7;
    expect(keysOf(cells(wb)?.[0])).toEqual(['shared', 'B1:B3', 0]);
    expect(b1?.formula).toBe('A1*2');
  });

  it('a value set on a data-table cell without text keeps its formula keys', () => {
    const table: CellData = { ...num('B2', '5'), formulaType: 'dataTable', formulaRef: 'B2:C3' };
    const wb = seeded([{ index: 2, cells: [table] }]);
    const b2 = wb.getSheet('S')?.cell('B2');
    if (b2) b2.value = 6;
    expect(keysOf(cells(wb)?.[0])).toEqual(['dataTable', 'B2:C3', undefined]);
  });

  it('formula text set on a shared-formula child drops its keys; on a master it keeps them', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }, child('C1')] }]);
    const ws = wb.getSheet('S');
    if (ws) {
      ws.cell('B1').formula = 'A1*3';
      ws.cell('C1').formula = 'A1*5';
    }
    const [b1, c1] = cells(wb) ?? [];
    expect(keysOf(b1)).toEqual(['shared', 'B1:B3', 0]);
    expect(b1?.formula).toBe('A1*3');
    expect(keysOf(c1)).toEqual([undefined, undefined, undefined]);
    expect(c1?.formula).toBe('A1*5');
  });

  it('formula = "" leaves no formula: the getter gives null and the write no <f>', async () => {
    const wb = seeded([{ index: 1, cells: [{ ...num('A1', '3'), formula: 'B1+1' }] }]);
    const a1 = wb.getSheet('S')?.cell('A1');
    if (a1) a1.formula = '';
    expect(a1?.formula).toBeNull();
    const xml = readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'] ?? '';
    expect(xml).toContain('<c r="A1"><v>3</v></c>');
    expect(xml).not.toContain('<f');
  });

  it('a shared cell without formulaRef is a child for the formula setter, text or not', () => {
    const textChild: CellData = { ...child('C1'), formula: 'A1*2' };
    const wb = seeded([
      { index: 1, cells: [{ ...master }, textChild, { ...textChild, reference: 'D1' }] },
    ]);
    const ws = wb.getSheet('S');
    if (ws) {
      ws.cell('C1').formula = 'A1*100';
      ws.cell('D1').formula = null;
    }
    const [, c1, d1] = cells(wb) ?? [];
    expect(keysOf(c1)).toEqual([undefined, undefined, undefined]);
    expect(c1?.formula).toBe('A1*100');
    expect(keysOf(d1)).toEqual([undefined, undefined, undefined]);
    expect(d1?.formula).toBeNull();
  });

  it('a value set on a shared child that carries text keeps its formula and keys', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }, { ...child('C1'), formula: 'A1*2' }] }]);
    const c1 = wb.getSheet('S')?.cell('C1');
    if (c1) c1.value = 9;
    expect(keysOf(cells(wb)?.[1])).toEqual(['shared', undefined, 0]);
    expect(c1?.formula).toBe('A1*2');
  });

  it('rich text set on a shared-formula child drops its keys', () => {
    const wb = seeded([{ index: 1, cells: [{ ...master }, child('C1')] }]);
    const c1 = wb.getSheet('S')?.cell('C1');
    if (c1) c1.richText = [{ text: 'x' }];
    expect(keysOf(cells(wb)?.[1])).toEqual([undefined, undefined, undefined]);
    expect(c1?.value).toBe('x');
  });
});

describe('date system and number formats', () => {
  it('honours the 1904 date system', () => {
    const wb = seeded([{ index: 1, cells: [{ ...num('A1', '0'), styleIndex: 0 }] }], {
      dateSystem: 'date1904',
      styles: {
        numFmts: [],
        fonts: [],
        fills: [],
        borders: [],
        cellXfs: [{ numFmtId: 14, fontId: 0, fillId: 0, borderId: 0 }],
      },
    });
    const date = wb.getSheet('S')?.cell('A1').dateValue;
    expect(date?.toISOString()).toBe('1904-01-01T00:00:00.000Z');
  });

  it('prefers the file numFmts over the built-in table', () => {
    const wb = seeded(
      [
        {
          index: 1,
          cells: [
            { ...num('A1', '45000'), styleIndex: 0 },
            { ...num('B1', '45000'), styleIndex: 1 },
          ],
        },
      ],
      {
        styles: {
          numFmts: [{ id: 14, formatCode: 'dd/mm/yyyy' }],
          fonts: [],
          fills: [],
          borders: [],
          cellXfs: [
            { numFmtId: 14, fontId: 0, fillId: 0, borderId: 0 },
            { numFmtId: 31, fontId: 0, fillId: 0, borderId: 0 },
          ],
        },
      },
    );
    const ws = wb.getSheet('S');
    expect(ws?.cell('A1').numberFormat).toBe('dd/mm/yyyy');
    expect(ws?.cell('B1').dateValue).toBeInstanceOf(Date);
  });
});

describe('sorted cells', () => {
  it('keeps cells in column order when written out of order', () => {
    const ws = new Workbook().addSheet('S');
    ws.cell('D2').value = 1;
    ws.cell('C2').value = 2;
    expect(ws.rows[0]?.cells.map((c) => c.reference)).toEqual(['C2', 'D2']);
  });

  it('does not duplicate a cell when a read row was unsorted', () => {
    const wb = seeded([{ index: 2, cells: [num('D2', '1'), num('C2', '2')] }]);
    const ws = wb.getSheet('S');
    expect(ws?.cell('C2').value).toBe(2);
    expect(ws?.rows[0]?.cells).toHaveLength(2);
    expect(ws?.rows[0]?.cells.map((c) => c.reference)).toEqual(['C2', 'D2']);
  });
});

describe('column spans', () => {
  const span = (min: number, max: number): ColumnInfo => ({
    min,
    max,
    width: 10,
    hidden: true,
    customWidth: true,
    outlineLevel: 2,
  });

  it('splits a covering span into up to three copies', () => {
    const ws = seeded([], {}, [span(1, 5)]).getSheet('S');
    ws?.setColumnWidth(3, 20);
    expect(ws?.columns.map((c) => [c.min, c.max])).toEqual([
      [1, 2],
      [3, 3],
      [4, 5],
    ]);
    expect(ws?.columns[1]?.width).toBe(20);
    expect(ws?.columns[0]?.width).toBe(10);
    expect(ws?.columns[0]?.hidden).toBe(true);
    expect(ws?.columns[1]?.outlineLevel).toBe(2);
  });

  it('ungroups only the requested column of a span', () => {
    const ws = seeded([], {}, [span(1, 5)]).getSheet('S');
    ws?.ungroupColumns(3, 3);
    const levels = ws?.columns.map((c) => [c.min, c.max, c.outlineLevel ?? null]);
    expect(levels).toEqual([
      [1, 2, 2],
      [3, 3, null],
      [4, 5, 2],
    ]);
  });

  it('inserts a missing column in sorted position', () => {
    const ws = seeded([], {}, [span(1, 1), span(5, 6)]).getSheet('S');
    ws?.setColumnWidth(3, 12);
    expect(ws?.columns.map((c) => [c.min, c.max])).toEqual([
      [1, 1],
      [3, 3],
      [5, 6],
    ]);
  });
});

describe('image ids', () => {
  it('skips media already present in the file', () => {
    const wb = seeded([], { preservedEntries: { 'xl/media/image1.png': [1, 2, 3] } });
    wb.addImage('S', { fromCol: 0, fromRow: 0, toCol: 1, toRow: 1 }, new Uint8Array([4, 5]));
    const keys = Object.keys(wb.toJSON().preservedEntries ?? {});
    expect(keys).toContain('xl/media/image1.png');
    expect(keys).toContain('xl/media/image2.png');
  });
});

describe('error classes', () => {
  it('throws ModernXlsxError for a detached worksheet', () => {
    const ws = new Workbook().addSheet('S');
    const detached = new (ws.constructor as new (...a: unknown[]) => typeof ws)({
      name: 'X',
      worksheet: {
        dimension: null,
        rows: [],
        mergeCells: [],
        autoFilter: null,
        frozenPane: null,
        columns: [],
        threadedComments: [
          { id: 'c1', refCell: 'A1', personId: 'p1', text: 't', timestamp: '2026-01-01T00:00:00Z' },
        ],
      },
    });
    for (const fn of [
      () => detached.addThreadedComment('A1', 't', 'a'),
      () => detached.replyToComment('c1', 't', 'a'),
    ]) {
      expect(fn).toThrow(expect.objectContaining({ code: INVALID_ARGUMENT }));
      expect(fn).toThrow(ModernXlsxError);
    }
  });
});
