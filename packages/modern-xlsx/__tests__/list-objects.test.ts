import { describe, expect, it } from 'vitest';
import type { TableDefinitionData } from '../src/index.js';
import { readBuffer, TABLE_STYLES, VALID_TABLE_STYLES, Workbook } from '../src/index.js';

// ---------------------------------------------------------------------------
// Tables (Excel ListObjects)
// ---------------------------------------------------------------------------

describe('Excel Tables (ListObjects)', () => {
  it('adds and retrieves a table on a worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Name';
    ws.cell('B1').value = 'Age';
    ws.cell('A2').value = 'Alice';
    ws.cell('B2').value = 30;

    const table: TableDefinitionData = {
      id: 1,
      displayName: 'People',
      ref: 'A1:B2',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Age' },
      ],
      styleInfo: {
        name: 'TableStyleMedium2',
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    };

    ws.addTable(table);
    expect(ws.tables).toHaveLength(1);
    expect(ws.getTable('People')).toBeDefined();
    expect(ws.getTable('People')?.ref).toBe('A1:B2');
  });

  it('removes a table by display name', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.addTable({
      id: 1,
      displayName: 'Removable',
      ref: 'A1:B2',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Col1' },
        { id: 2, name: 'Col2' },
      ],
    });

    expect(ws.removeTable('Removable')).toBe(true);
    expect(ws.tables).toHaveLength(0);
    expect(ws.removeTable('NonExistent')).toBe(false);
  });

  it('returns undefined for non-existent table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    expect(ws.getTable('DoesNotExist')).toBeUndefined();
  });

  it('supports table with totals row and calculated columns', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sales');
    const table: TableDefinitionData = {
      id: 2,
      name: 'Sales',
      displayName: 'Sales',
      ref: 'A1:D6',
      headerRowCount: 1,
      totalsRowCount: 1,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Product' },
        { id: 2, name: 'Qty', totalsRowFunction: 'sum' },
        { id: 3, name: 'Price', totalsRowFunction: 'average' },
        {
          id: 4,
          name: 'Total',
          totalsRowFunction: 'sum',
          calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
        },
      ],
      autoFilterRef: 'A1:D5',
    };

    ws.addTable(table);
    expect(ws.getTable('Sales')?.columns[3]?.calculatedColumnFormula).toBe(
      'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
    );
    expect(ws.getTable('Sales')?.totalsRowCount).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Table Styles
// ---------------------------------------------------------------------------

describe('Table Styles', () => {
  it('has 60 built-in styles', () => {
    expect(TABLE_STYLES.light).toHaveLength(21);
    expect(TABLE_STYLES.medium).toHaveLength(28);
    expect(TABLE_STYLES.dark).toHaveLength(11);
    expect(VALID_TABLE_STYLES.size).toBe(60);
  });

  it('includes expected style names', () => {
    expect(VALID_TABLE_STYLES.has('TableStyleLight1')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleMedium14')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleDark11')).toBe(true);
    expect(VALID_TABLE_STYLES.has('TableStyleFake1')).toBe(false);
  });
});

describe('Roundtrip: Excel tables', () => {
  it('keeps the table definition and style info', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sales');
    ws.cell('A1').value = 'Product';
    const table: TableDefinitionData = {
      id: 2,
      name: 'Sales',
      displayName: 'Sales',
      ref: 'A1:D6',
      headerRowCount: 1,
      totalsRowCount: 1,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Product' },
        { id: 2, name: 'Qty', totalsRowFunction: 'sum' },
        { id: 3, name: 'Price', totalsRowFunction: 'average' },
        {
          id: 4,
          name: 'Total',
          totalsRowFunction: 'sum',
          calculatedColumnFormula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]',
        },
      ],
      autoFilterRef: 'A1:D5',
      styleInfo: {
        name: 'TableStyleMedium9',
        showFirstColumn: true,
        showLastColumn: false,
        showRowStripes: true,
        showColumnStripes: false,
      },
    };
    ws.addTable(table);

    const wb2 = await readBuffer(await wb.toBuffer());
    const t = wb2.getSheet('Sales')?.getTable('Sales');

    expect(t?.columns).toEqual(table.columns);
    expect(t?.ref).toBe('A1:D6');
    expect(t?.totalsRowCount).toBe(1);
    expect(t?.autoFilterRef).toBe('A1:D5');
    expect(t?.styleInfo).toMatchObject({
      name: 'TableStyleMedium9',
      showFirstColumn: true,
      showRowStripes: true,
    });
    // An absent flag means false.
    expect(t?.styleInfo?.showLastColumn ?? false).toBe(false);
  });
});
