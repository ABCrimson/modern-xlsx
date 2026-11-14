import { describe, expect, it } from 'vitest';
import type { CellData, ChartSeriesData, WorkbookData } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

/**
 * The additive wire fields are typed in src/model/types/*.ts (re-exported through src/model/types.ts) as optional properties. Each case
 * builds a workbook JSON that uses them, writes it and reads it back.
 */
function sample(): WorkbookData {
  const seed = new Workbook();
  const ws = seed.addSheet('S');
  ws.cell('A1').value = 'Region';
  ws.cell('A2').value = 3;
  ws.cell('B2').value = 4;
  const json = seed.toJSON();
  const series: ChartSeriesData = {
    idx: 0,
    order: 0,
    nameLiteral: 'Revenue',
    catRef: 'S!$A$2:$A$2',
    valRef: 'S!$B$2:$B$2',
  };
  const sheet = json.sheets[0];
  if (!sheet) throw new Error('sample sheet missing');
  const worksheet = sheet.worksheet;
  worksheet.tabColorSpec = { theme: 4 };
  worksheet.conditionalFormatting = [
    {
      sqref: 'B2',
      rules: [
        {
          ruleType: 'dataBar',
          priority: 1,
          operator: null,
          formula: null,
          dxfId: null,
          dataBar: {
            cfvos: [
              { cfvoType: 'min', val: null },
              { cfvoType: 'max', val: null },
            ],
            color: '',
            colorSpec: { theme: 5 },
          },
        },
      ],
    },
  ];
  worksheet.tables = [
    {
      id: 1,
      displayName: 'T1',
      ref: 'A1:B2',
      headerRowCount: 1,
      totalsRowCount: 1,
      totalsRowShown: true,
      columns: [
        { id: 1, name: 'Region' },
        { id: 2, name: 'N', totalsRowFormula: 'SUM(T1[N])' },
      ],
    },
  ];
  worksheet.charts = [
    {
      chart: { chartType: 'column', series: [series] },
      anchor: { fromCol: 3, fromRow: 0, toCol: 9, toRow: 12 },
    },
  ];
  json.styles.fonts[0] = {
    ...json.styles.fonts[0],
    name: 'Calibri',
    size: 11,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color: 'FF000000',
    colorSpec: { rgb: 'FF000000', theme: 1, tint: -0.25 },
  };
  json.definedNames = [{ name: 'Hid', value: 'S!$A$1', sheetId: null, hidden: true }];
  json.calcPr = { fullCalcOnLoad: true };
  return json;
}

describe('additive wire fields round trip', () => {
  it('keeps colour specs, hidden names, calcPr, totalsRowFormula and nameLiteral', async () => {
    const bytes = await new Workbook(sample()).toBuffer();
    const wb = await readBuffer(bytes);
    const json = wb.toJSON();
    const worksheet = json.sheets[0]?.worksheet;

    expect(json.styles.fonts[0]?.colorSpec).toEqual({ rgb: 'FF000000', theme: 1, tint: -0.25 });
    expect(worksheet?.tabColorSpec).toEqual({ theme: 4 });
    expect(worksheet?.conditionalFormatting?.[0]?.rules[0]?.dataBar?.colorSpec).toEqual({
      theme: 5,
    });
    expect(json.definedNames?.find((n) => n.name === 'Hid')?.hidden).toBe(true);
    expect(json.calcPr).toMatchObject({ fullCalcOnLoad: true });
    expect(worksheet?.tables?.[0]?.columns[1]?.totalsRowFormula).toBe('SUM(T1[N])');
    expect(worksheet?.charts?.[0]?.chart.series[0]?.nameLiteral).toBe('Revenue');
  });

  it('writes a series nameLiteral as c:v', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('B2').value = 4;
    ws.addChart('column', (b) =>
      b.addSeries({ nameLiteral: 'Revenue', catRef: 'S!$A$2:$A$2', valRef: 'S!$B$2:$B$2' }),
    );
    const parts = readZipText(await wb.toBuffer());
    const xml = parts['xl/charts/chart1.xml'];
    expect(xml).toContain('<c:v>Revenue</c:v>');
  });

  it('still emits the deprecated dynamicArray for a dynamic-array cell', async () => {
    const json = sample();
    const cell: CellData = {
      reference: 'C2',
      cellType: 'number',
      styleIndex: null,
      value: '7',
      formula: 'SEQUENCE(1)',
      cellMetadata: 1,
    };
    json.sheets[0]?.worksheet.rows[1]?.cells.push(cell);
    const wb = await readBuffer(await new Workbook(json).toBuffer());
    const read = wb.toJSON().sheets[0]?.worksheet.rows[1]?.cells.find((c) => c.reference === 'C2');
    expect(read?.cellMetadata).toBe(1);
    expect(read?.dynamicArray).toBe(true);
  });
});
