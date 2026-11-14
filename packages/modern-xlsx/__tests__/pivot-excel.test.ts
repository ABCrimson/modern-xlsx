import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const TABLE = 'xl/pivotTables/pivotTable1.xml';
const CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml';

/**
 * A pivot table built with addPivotCache, addPivotCacheRecords and
 * PivotTableBuilder. Written with a row field without items, Excel 16 shows
 * it with empty cells and crashes on PivotCache().Refresh(), so the writer
 * completes the field from its cache.
 */
function repro(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Sheet1');
  const rows: (string | number)[][] = [
    ['Region', 'Product', 'Revenue'],
    ['East', 'Pen', 10],
    ['East', 'Ink', 20],
    ['West', 'Pen', 30],
    ['West', 'Ink', 40],
  ];
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      ws.cell(`${'ABC'[c]}${r + 1}`).value = value;
    });
  });
  wb.addPivotCache({
    source: { ref: 'A1:C5', sheet: 'Sheet1' },
    fields: [
      {
        name: 'Region',
        sharedItems: [
          { type: 'string', v: 'East' },
          { type: 'string', v: 'West' },
        ],
      },
      {
        name: 'Product',
        sharedItems: [
          { type: 'string', v: 'Pen' },
          { type: 'string', v: 'Ink' },
        ],
      },
      { name: 'Revenue', sharedItems: [] },
    ],
    recordCount: 4,
  });
  wb.addPivotCacheRecords({
    records: [
      [0, 0, 10],
      [0, 1, 20],
      [1, 0, 30],
      [1, 1, 40],
    ].map(([region, product, revenue]) => [
      { type: 'index' as const, v: region as number },
      { type: 'index' as const, v: product as number },
      { type: 'number' as const, v: revenue as number },
    ]),
  });
  ws.addPivotTableFromBuilder((b) => {
    b.name('PT1')
      .cacheId(0)
      .location('E3:F6')
      .addRowField({ fieldIndex: 0, name: 'Region' })
      .addDataField({ fieldIndex: 2, name: 'Total Revenue' });
  });
  return wb;
}

describe('a pivot table built through the API', () => {
  it('writes the row field items Excel needs to refresh, and refreshes its cache on open', async () => {
    const wb = repro();
    const before = structuredClone(wb.toJSON());
    const parts = readZipText(await wb.toBuffer());

    expect(parts[TABLE]).toContain(
      '<pivotFields count="3"><pivotField axis="axisRow" name="Region"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField/><pivotField dataField="1"/></pivotFields>',
    );
    expect(parts[CACHE]).toContain(' refreshOnLoad="1"');
    // The workbook model is not changed by the write.
    expect(wb.toJSON()).toEqual(before);
    expect(wb.getSheet('Sheet1')?.pivotTables[0]?.pivotFields[0]?.items).toEqual([]);
  });

  it('reads back with the items and writes the same pivot table again', async () => {
    const first = readZipText(await repro().toBuffer());
    const read = await readBuffer(await repro().toBuffer());
    expect(read.getSheet('Sheet1')?.pivotTables[0]?.pivotFields[0]?.items).toEqual([
      { x: 0 },
      { x: 1 },
      { t: 'default' },
    ]);

    const second = readZipText(await read.toBuffer());
    expect(second[TABLE]).toBe(first[TABLE]);
    // A cache read from a file is not one the API built: its pivot table has
    // items now, so the cache no longer refreshes on open.
    expect(second[CACHE]).toBe(first[CACHE]?.replace(' refreshOnLoad="1"', ''));

    const third = readZipText(await (await readBuffer(await read.toBuffer())).toBuffer());
    expect(third[TABLE]).toBe(second[TABLE]);
    expect(third[CACHE]).toBe(second[CACHE]);
  });

  it('puts the values field on the columns of a table with two data fields', async () => {
    const wb = repro();
    wb.getSheet('Sheet1')?.removePivotTable(0);
    wb.getSheet('Sheet1')?.addPivotTableFromBuilder((b) => {
      b.name('PT2')
        .cacheId(0)
        .location('E3:H6')
        .addRowField({ fieldIndex: 0 })
        .addColField({ fieldIndex: 1 })
        .addDataField({ fieldIndex: 2, name: 'Total Revenue' })
        .addDataField({ fieldIndex: 2, subtotal: 'count', name: 'Sales' });
    });
    const xml = readZipText(await wb.toBuffer())[TABLE];
    expect(xml).toContain('<colFields count="2"><field x="1"/><field x="-2"/></colFields>');
    expect(xml).toContain(
      '<pivotField axis="axisCol"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>',
    );
    expect(wb.getSheet('Sheet1')?.pivotTables[0]?.colFields).toEqual([{ x: 1 }]);
  });

  it('writes one pivot field per cache field when the builder uses fewer', async () => {
    // A 5-field cache with fields 0 and 3 used: the builder makes 4 pivot
    // fields, and Excel 16 refuses `<pivotFields count="4">` on that cache.
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Region';
    wb.addPivotCache({
      source: { ref: 'A1:E3', sheet: 'Sheet1' },
      fields: [
        {
          name: 'Region',
          sharedItems: [
            { type: 'string', v: 'East' },
            { type: 'string', v: 'West' },
          ],
        },
        { name: 'Product', sharedItems: [] },
        { name: 'Year', sharedItems: [] },
        { name: 'Revenue', sharedItems: [] },
        { name: 'Units', sharedItems: [] },
      ],
      recordCount: 0,
    });
    ws.addPivotTableFromBuilder((b) => {
      b.name('PT1')
        .cacheId(0)
        .location('G3:H6')
        .addRowField({ fieldIndex: 0 })
        .addDataField({ fieldIndex: 3, name: 'Total Revenue' });
    });
    expect(ws.pivotTables[0]?.pivotFields).toHaveLength(4);

    const xml = readZipText(await wb.toBuffer())[TABLE];
    expect(xml).toContain(
      '<pivotFields count="5"><pivotField axis="axisRow"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField/><pivotField/><pivotField dataField="1"/><pivotField/></pivotFields>',
    );
    // The model is not changed; read back, the table has all five fields.
    expect(ws.pivotTables[0]?.pivotFields).toHaveLength(4);
    const read = await readBuffer(await wb.toBuffer());
    expect(read.getSheet('Sheet1')?.pivotTables[0]?.pivotFields).toHaveLength(5);
  });

  it('validate() warns about a pivot table whose cacheId names no cache', () => {
    const wb = repro();
    expect(wb.validate().issues.filter((i) => i.location.includes('pivotTables'))).toEqual([]);
    wb.getSheet('Sheet1')?.addPivotTableFromBuilder((b) => {
      b.name('Orphan').cacheId(4).location('J3').addRowField({ fieldIndex: 0 });
    });
    const issues = wb.validate().issues.filter((i) => i.location.includes('pivotTables'));
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({
      severity: 'warning',
      category: 'structure',
      location: 'Sheet1!pivotTables[1]',
      autoFixable: false,
    });
    expect(issues[0]?.message).toContain('cacheId 4');
  });
});
