import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

async function roundtrip(wb: Workbook): Promise<Workbook> {
  return readBuffer(await wb.toBuffer());
}

describe('SpreadsheetML colours are written as ARGB', () => {
  it("'#FF0000' reads back as 'FFFF0000'", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'x';
    ws.cell('A1').styleIndex = wb.createStyle().font({ color: '#FF0000' }).build(wb.styles);
    const wb2 = await roundtrip(wb);
    const xf = wb2.styles.cellXfs[wb2.getSheet('S')?.cell('A1').styleIndex ?? 0];
    expect(wb2.styles.fonts[xf?.fontId ?? 0]?.color).toBe('FFFF0000');
  });

  it("fgColor 'ff0000' reads back as 'FFFF0000'", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'x';
    ws.cell('A1').styleIndex = wb
      .createStyle()
      .fill({ pattern: 'solid', fgColor: 'ff0000' })
      .build(wb.styles);
    const wb2 = await roundtrip(wb);
    const xf = wb2.styles.cellXfs[wb2.getSheet('S')?.cell('A1').styleIndex ?? 0];
    expect(wb2.styles.fills[xf?.fillId ?? 0]?.fgColor).toBe('FFFF0000');
  });

  it("tabColor '0000FF' reads back as 'FF0000FF'", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 1;
    ws.tabColor = '0000FF';
    const wb2 = await roundtrip(wb);
    expect(wb2.getSheet('S')?.tabColor).toBe('FF0000FF');
  });

  it("run colour 'ABCDEF' reads back as 'FFABCDEF'", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').richText = [{ text: 'a', color: 'ABCDEF' }, { text: 'b' }];
    const wb2 = await roundtrip(wb);
    expect(wb2.getSheet('S')?.cell('A1').richText?.[0]?.color).toBe('FFABCDEF');
  });

  it("'80FF0000' stays unchanged", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 1;
    ws.tabColor = '80FF0000';
    const wb2 = await roundtrip(wb);
    expect(wb2.getSheet('S')?.tabColor).toBe('80FF0000');
  });

  it('chart fillColor stays RRGGBB in DrawingML a:srgbClr', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Q1';
    ws.cell('B1').value = 1;
    ws.addChart('bar', (b) => {
      b.addSeries({
        name: 'Revenue',
        catRef: 'Sheet1!$A$1:$A$1',
        valRef: 'Sheet1!$B$1:$B$1',
        fillColor: '4472C4',
      }).anchor({ col: 4, row: 0 }, { col: 12, row: 15 });
    });
    const parts = readZipText(await wb.toBuffer());
    const chartXml = Object.entries(parts).find(([name]) =>
      /^xl\/charts\/chart\d+\.xml$/.test(name),
    )?.[1];
    expect(chartXml).toContain('<a:srgbClr val="4472C4"/>');
    expect(chartXml).not.toContain('FF4472C4');
  });

  it("sparkline colours 'ff0000' and '#00FF00' read back as ARGB", async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 1;
    ws.addSparklineGroup({
      sparklines: [{ formula: 'S!A1:A5', sqref: 'C1' }],
      colorSeries: 'ff0000',
      colorNegative: '#00FF00',
    });
    const g = (await roundtrip(wb)).getSheet('S')?.sparklineGroups[0];
    expect(g?.colorSeries).toBe('FFFF0000');
    expect(g?.colorNegative).toBe('FF00FF00');
  });
});
