import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

describe('column attributes round trip', () => {
  it('keeps a column styleIndex through toBuffer -> readBuffer', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const styleIndex = wb.createStyle().numberFormat('#,##0.000').build(wb.styles);
    expect(styleIndex).toBeGreaterThan(0);
    ws.columns = [
      { min: 1, max: 1, width: 10, hidden: false, customWidth: true, styleIndex },
      { min: 2, max: 2, width: 12, hidden: false, customWidth: true },
    ];

    const bytes = await wb.toBuffer();
    expect(readZipText(bytes)['xl/worksheets/sheet1.xml']).toContain(
      `<col min="1" max="1" width="10" style="${styleIndex}" customWidth="1"/>`,
    );

    const back = await readBuffer(bytes);
    const [styled, plain] = back.getSheet('Sheet1')?.columns ?? [];
    expect(styled?.styleIndex).toBe(styleIndex);
    const numFmtId = back.styles.cellXfs[styleIndex]?.numFmtId;
    expect(back.styles.numFmts.find((f) => f.id === numFmtId)?.formatCode).toBe('#,##0.000');
    expect(plain).toBeDefined();
    expect(plain).not.toHaveProperty('styleIndex');
    expect(plain).not.toHaveProperty('bestFit');
    expect(plain).not.toHaveProperty('phonetic');
  });

  it('keeps bestFit and phonetic flags', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').columns = [
      {
        min: 3,
        max: 4,
        width: 9,
        hidden: false,
        customWidth: false,
        bestFit: true,
        phonetic: true,
      },
    ];
    const back = await readBuffer(await wb.toBuffer());
    expect(back.getSheet('Sheet1')?.columns[0]).toMatchObject({ bestFit: true, phonetic: true });
  });
});
