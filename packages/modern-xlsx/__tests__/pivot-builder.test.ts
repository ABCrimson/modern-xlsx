import { describe, expect, it } from 'vitest';
import { PivotTableBuilder, readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

describe('PivotTableBuilder contract', () => {
  it('rejects a duplicate row field', () => {
    const b = new PivotTableBuilder().addRowField({ fieldIndex: 0 });
    expect(() => b.addRowField({ fieldIndex: 0 })).toThrow(RangeError);
  });

  it('rejects a field used for both row and col', () => {
    const b = new PivotTableBuilder().addRowField({ fieldIndex: 1 });
    expect(() => b.addColField({ fieldIndex: 1 })).toThrow(RangeError);
    expect(() => b.addPageField({ fieldIndex: 1 })).toThrow(RangeError);
  });

  it('rejects a non-u32 field index', () => {
    for (const fieldIndex of [-1, 1.5, Number.NaN, 2 ** 32]) {
      expect(() => new PivotTableBuilder().addRowField({ fieldIndex })).toThrow(RangeError);
    }
  });

  it('allows a data field and a row field on the same index', () => {
    const data = new PivotTableBuilder()
      .addDataField({ fieldIndex: 2, subtotal: 'sum', name: 'Total' })
      .addRowField({ fieldIndex: 2 })
      .build();
    expect(data.pivotFields[2]).toMatchObject({ dataField: true, axis: 'axisRow' });
    expect(data.pivotFields[2]?.name).toBeUndefined();
  });

  it('marks a data-only field with dataField and no axis or name', () => {
    const data = new PivotTableBuilder()
      .addRowField({ fieldIndex: 0 })
      .addDataField({ fieldIndex: 2, name: 'Total Revenue' })
      .build();
    const pf = data.pivotFields[2];
    expect(pf?.dataField).toBe(true);
    expect(pf?.axis).toBeUndefined();
    expect(pf?.name).toBeUndefined();
    expect(data.dataFields[0]?.name).toBe('Total Revenue');
  });

  it('allows Sum and Count on one field', () => {
    const data = new PivotTableBuilder()
      .addDataField({ fieldIndex: 2, subtotal: 'sum' })
      .addDataField({ fieldIndex: 2, subtotal: 'count' })
      .build();
    expect(data.dataFields.map((d) => d.subtotal)).toEqual(['sum', 'count']);
  });

  it('builds the exact pivotFields for a row field plus a data field', () => {
    const data = new PivotTableBuilder()
      .addRowField({ fieldIndex: 0, name: 'Region' })
      .addDataField({ fieldIndex: 2, name: 'Total Revenue' })
      .build();
    expect(data.pivotFields[2]).toEqual({
      items: [],
      subtotals: [],
      compact: true,
      outline: true,
      dataField: true,
    });
  });

  it('a data placeholder gets its axis from a later row field, without a data caption', () => {
    const data = new PivotTableBuilder()
      .addDataField({ fieldIndex: 0, name: 'Sum of X' })
      .addRowField({ fieldIndex: 0, name: 'X' })
      .build();
    expect(data.pivotFields[0]).toMatchObject({ axis: 'axisRow', dataField: true, name: 'X' });
    expect(data.dataFields[0]?.name).toBe('Sum of X');
  });

  it('round-trips dataField through write and read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Region';
    ws.cell('B1').value = 'Product';
    ws.cell('C1').value = 'Revenue';
    ws.cell('A2').value = 'East';
    ws.cell('B2').value = 'Pen';
    ws.cell('C2').value = 10;
    wb.addPivotCache({
      source: { ref: 'A1:C2', sheet: 'Sheet1' },
      fields: [
        { name: 'Region', sharedItems: [{ type: 'string', v: 'East' }] },
        { name: 'Product', sharedItems: [{ type: 'string', v: 'Pen' }] },
        { name: 'Revenue', sharedItems: [] },
      ],
      recordCount: 1,
    });
    wb.addPivotCacheRecords({
      records: [
        [
          { type: 'index', v: 0 },
          { type: 'index', v: 0 },
          { type: 'number', v: 10 },
        ],
      ],
    });
    ws.addPivotTableFromBuilder((b) => {
      b.name('P')
        .cacheId(0)
        .location('E3:F5')
        .addRowField({ fieldIndex: 0, name: 'Region' })
        .addDataField({ fieldIndex: 2, name: 'Total Revenue' });
    });

    const bytes = await wb.toBuffer();
    const xml = readZipText(bytes)['xl/pivotTables/pivotTable1.xml'];
    expect(xml).toContain('<pivotField dataField="1"');
    const wb2 = await readBuffer(bytes);
    expect(wb2.getSheet('Sheet1')?.pivotTables[0]?.pivotFields[2]?.dataField).toBe(true);
  });

  it('build() returns an independent deep copy', () => {
    const b = new PivotTableBuilder().addRowField({ fieldIndex: 0 });
    const p1 = b.build();
    const snapshot = structuredClone(p1);
    b.addRowField({ fieldIndex: 3 })
      .addDataField({ fieldIndex: 4 })
      .addPageField({ fieldIndex: 5 });
    expect(p1).toEqual(snapshot);
    const p2 = b.build();
    expect(p2.rowFields).toHaveLength(2);
    expect(p2.dataFields).toHaveLength(1);
    expect(p2.pageFields).toHaveLength(1);
  });

  it('caps fieldIndex at the 16384 sheet columns', () => {
    expect(() => new PivotTableBuilder().addRowField({ fieldIndex: 16384 })).toThrow(RangeError);
    const p = new PivotTableBuilder().addDataField({ fieldIndex: 16383 }).build();
    expect(p.pivotFields).toHaveLength(16384);
  });
});
