import { describe, expect, it } from 'vitest';
import { type CacheSource, readBuffer, Workbook } from '../src/index.js';
import { buildZip, readZipText } from './helpers/zip.js';

const CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml';

/**
 * The cache definition Excel 16 saved for a pivot table inserted from the table `Sales`
 * (`PivotCaches().Create(1, "Sales")` through COM), with `refreshedBy` anonymised. Excel
 * refuses the file if a read and write turns its source into `<worksheetSource ref="" sheet=""/>`.
 */
const EXCEL_TABLE_SOURCE_CACHE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" r:id="rId1" refreshedBy="Author" refreshedDate="46338.560222453707" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="4" xr:uid="{61517C26-2271-41C6-BDD1-C7A40A20D77D}"><cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource><cacheFields count="2"><cacheField name="Region" numFmtId="0"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField><cacheField name="Revenue" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="40"/></cacheField></cacheFields><extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition/></ext></extLst></pivotCacheDefinition>';

/** The `<cacheSource>` Excel 16 saved for a pivot table over two consolidation ranges. */
const EXCEL_CONSOLIDATION_SOURCE =
  '<cacheSource type="consolidation"><consolidation><pages count="1"><page count="2"><pageItem name="Item1"/><pageItem name="Item2"/></page></pages><rangeSets count="2"><rangeSet i1="0" ref="A1:B5" sheet="Q1"/><rangeSet i1="1" ref="A1:B5" sheet="Q2"/></rangeSets></consolidation></cacheSource>';

/**
 * The `<cacheSource>` Excel 16 saved for a pivot table over a consolidation range in another workbook: its
 * `r:id` names the `externalLinkPath` relationship of {@link EXCEL_EXTERNAL_CACHE_RELS}.
 */
const EXTERNAL_CONSOLIDATION_SOURCE =
  '<cacheSource type="consolidation"><consolidation autoPage="0"><rangeSets count="1"><rangeSet ref="A1:C3" sheet="Data" r:id="rId2"/></rangeSets></consolidation></cacheSource>';

const CACHE_RELS = 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels';

/** The `.rels` Excel 16 saved beside a cache whose source is in another workbook. */
const EXCEL_EXTERNAL_CACHE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="xcons-source.xlsx" TargetMode="External"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>';

/**
 * The workbook Excel built the cache above from: the table `Sales` (A1:B5 of `Data Sheet`) and a pivot table on
 * `Pivot` that sums Revenue by Region, over a cache whose source is given.
 */
function salesWorkbook(source: CacheSource): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data Sheet');
  const rows: (string | number)[][] = [
    ['Region', 'Revenue'],
    ['East', 10],
    ['East', 20],
    ['West', 30],
    ['West', 40],
  ];
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      ws.cell(`${'AB'[c]}${r + 1}`).value = value;
    });
  });
  ws.addTable({
    id: 1,
    displayName: 'Sales',
    ref: 'A1:B5',
    headerRowCount: 1,
    totalsRowCount: 0,
    totalsRowShown: false,
    columns: [
      { id: 1, name: 'Region' },
      { id: 2, name: 'Revenue' },
    ],
  });
  wb.addPivotCache({
    source,
    fields: [
      {
        name: 'Region',
        sharedItems: [
          { type: 'string', v: 'East' },
          { type: 'string', v: 'West' },
        ],
      },
      { name: 'Revenue', sharedItems: [] },
    ],
    recordCount: 4,
  });
  wb.addPivotCacheRecords({
    records: [
      [0, 10],
      [0, 20],
      [1, 30],
      [1, 40],
    ].map(([region, revenue]) => [
      { type: 'index' as const, v: region as number },
      { type: 'number' as const, v: revenue as number },
    ]),
  });
  wb.addSheet('Pivot').addPivotTableFromBuilder((b) => {
    b.name('PT1')
      .cacheId(0)
      .location('A3:B6')
      .addRowField({ fieldIndex: 0, name: 'Region' })
      .addDataField({ fieldIndex: 1, name: 'Sum of Revenue' });
  });
  return wb;
}

/** `bytes` with each part named in `replaced`, which it must hold, replaced by its text. */
function withParts(bytes: Uint8Array, replaced: Record<string, string>): Uint8Array {
  const enc = new TextEncoder();
  const parts = readZipText(bytes);
  for (const [name, xml] of Object.entries(replaced)) {
    expect(parts[name]).toBeDefined();
    parts[name] = xml;
  }
  return buildZip(Object.entries(parts).map(([name, text]) => [name, enc.encode(text)] as const));
}

/** `bytes` with the cache definition part replaced by `xml`. */
function withCacheDefinition(bytes: Uint8Array, xml: string): Uint8Array {
  return withParts(bytes, { [CACHE]: xml });
}

/** The `<cacheSource>` element of a written package. */
function cacheSource(bytes: Uint8Array): string | undefined {
  return readZipText(bytes)[CACHE]?.match(
    /<cacheSource[\s\S]*?(?:<\/cacheSource>|\/>)(?=<cacheFields)/,
  )?.[0];
}

describe('pivot cache sources', () => {
  it('keeps the Excel table source of a cache through a read and write', async () => {
    const pkg = withCacheDefinition(
      await salesWorkbook({ ref: 'A1:B5', sheet: 'Data Sheet' }).toBuffer(),
      EXCEL_TABLE_SOURCE_CACHE,
    );
    const read = await readBuffer(pkg);
    expect(read.pivotCaches[0]?.source).toStrictEqual({ ref: '', sheet: '', name: 'Sales' });

    const once = await read.toBuffer();
    const xml = readZipText(once)[CACHE] ?? '';
    expect(cacheSource(once)).toBe(
      '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>',
    );
    // No attribute is written empty.
    expect(xml).not.toMatch(/\s[\w:]+=""/);

    // A second generation writes the same part.
    const twice = await (await readBuffer(once)).toBuffer();
    expect(readZipText(twice)[CACHE]).toBe(xml);
  });

  it('writes a table source built through the API, and a range source as before', async () => {
    const byName = await salesWorkbook({ ref: '', sheet: '', name: 'Sales' }).toBuffer();
    expect(cacheSource(byName)).toBe(
      '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>',
    );
    expect(readZipText(byName)[CACHE]).not.toMatch(/\s[\w:]+=""/);

    const byRange = await salesWorkbook({ ref: 'A1:B5', sheet: 'Data Sheet' }).toBuffer();
    expect(cacheSource(byRange)).toBe(
      '<cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="Data Sheet"/></cacheSource>',
    );
    // The JSON bridge gains no key for a range source.
    expect((await readBuffer(byRange)).pivotCaches[0]?.source).toStrictEqual({
      ref: 'A1:B5',
      sheet: 'Data Sheet',
    });
  });

  it('adds no relationship keys for a pivot without other relationships', async () => {
    const read = await readBuffer(
      await salesWorkbook({ ref: 'A1:B5', sheet: 'Data Sheet' }).toBuffer(),
    );
    expect(read.pivotCaches[0]).not.toHaveProperty('relationships');
    expect(read.pivotCaches[0]?.source).not.toHaveProperty('rId');
    expect(readZipText(await read.toBuffer())[CACHE_RELS]).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>',
    );
  });

  it.each([
    [
      'a worksheet source',
      '<cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="Data" r:id="rId2"/></cacheSource>',
    ],
    ['a consolidation source', EXTERNAL_CONSOLIDATION_SOURCE],
  ])('keeps the relationships of %s in another workbook through two writes', async (_, source) => {
    const pkg = withParts(await salesWorkbook({ ref: 'A1:B5', sheet: 'Data Sheet' }).toBuffer(), {
      [CACHE]: EXCEL_TABLE_SOURCE_CACHE.replace(
        '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>',
        source,
      ),
      [CACHE_RELS]: EXCEL_EXTERNAL_CACHE_RELS,
    });
    const read = await readBuffer(pkg);
    const cache = read.pivotCaches[0];
    expect(cache?.relationships).toStrictEqual([
      {
        id: 'rId2',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',
        target: 'xcons-source.xlsx',
        targetMode: 'External',
      },
    ]);
    if (source === EXTERNAL_CONSOLIDATION_SOURCE) {
      expect(cache?.source.rawXml).toBe(source);
    } else {
      expect(cache?.source).toStrictEqual({ ref: 'A1:B5', sheet: 'Data', rId: 'rId2' });
    }

    const once = await read.toBuffer();
    const parts = readZipText(once);
    expect(cacheSource(once)).toBe(source);
    // The records keep rId1; the external link keeps the id the source names.
    expect(parts[CACHE]).toContain(' r:id="rId1"');
    expect(parts[CACHE_RELS]).toContain(
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="xcons-source.xlsx" TargetMode="External"/>',
    );
    expect(parts[CACHE_RELS]).toContain(
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>',
    );

    // A second generation writes the same parts.
    const twice = readZipText(await (await readBuffer(once)).toBuffer());
    expect(twice[CACHE]).toBe(parts[CACHE]);
    expect(twice[CACHE_RELS]).toBe(parts[CACHE_RELS]);
  });

  it('drops a source r:id that names no kept relationship', async () => {
    const wb = salesWorkbook({ ref: '', sheet: '', rawXml: EXTERNAL_CONSOLIDATION_SOURCE });
    const written = await wb.toBuffer();
    expect(cacheSource(written)).toBe(EXTERNAL_CONSOLIDATION_SOURCE.replace(' r:id="rId2"', ''));
  });

  it('keeps a consolidation source verbatim', async () => {
    const pkg = withCacheDefinition(
      await salesWorkbook({ ref: 'A1:B5', sheet: 'Data Sheet' }).toBuffer(),
      EXCEL_TABLE_SOURCE_CACHE.replace(
        '<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>',
        EXCEL_CONSOLIDATION_SOURCE,
      ),
    );
    const read = await readBuffer(pkg);
    expect(read.pivotCaches[0]?.source.rawXml).toBe(EXCEL_CONSOLIDATION_SOURCE);
    expect(cacheSource(await read.toBuffer())).toBe(EXCEL_CONSOLIDATION_SOURCE);
  });
});
