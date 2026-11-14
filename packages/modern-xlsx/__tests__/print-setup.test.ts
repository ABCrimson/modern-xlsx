import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

// ---------------------------------------------------------------------------
// Print Titles & Print Areas
// ---------------------------------------------------------------------------

describe('Print Titles', () => {
  it('sets and gets print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$2');
  });

  it('clears print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    wb.setPrintTitles(0, null);
    expect(wb.getPrintTitles(0)).toBeNull();
  });

  it('updates existing print titles', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintTitles(0, 'Sheet1!$1:$2');
    wb.setPrintTitles(0, 'Sheet1!$1:$3');
    expect(wb.getPrintTitles(0)).toBe('Sheet1!$1:$3');
    // Should not have duplicates
    expect(wb.namedRanges.filter((d) => d.name === '_xlnm.Print_Titles')).toHaveLength(1);
  });
});

describe('Print Area', () => {
  it('sets and gets print area', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$D$50');
  });

  it('clears print area', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    wb.setPrintArea(0, null);
    expect(wb.getPrintArea(0)).toBeNull();
  });

  it('handles multiple sheets independently', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    wb.addSheet('Sheet2');

    wb.setPrintArea(0, 'Sheet1!$A$1:$D$50');
    wb.setPrintArea(1, 'Sheet2!$A$1:$F$100');

    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$D$50');
    expect(wb.getPrintArea(1)).toBe('Sheet2!$A$1:$F$100');
  });
});

// ---------------------------------------------------------------------------
// Page Breaks
// ---------------------------------------------------------------------------

describe('Page Breaks', () => {
  it('sets and gets page breaks', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.pageBreaks = {
      rowBreaks: [{ id: 10, max: 16383, man: true }],
      colBreaks: [{ id: 5, max: 1048575, man: true }],
    };

    expect(ws.pageBreaks).toEqual({
      rowBreaks: [{ id: 10, max: 16383, man: true }],
      colBreaks: [{ id: 5, max: 1048575, man: true }],
    });
  });

  it('adds row breaks via helper', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addRowBreak(10);
    ws.addRowBreak(25);

    expect(ws.pageBreaks?.rowBreaks).toHaveLength(2);
    expect(ws.pageBreaks?.rowBreaks?.[0]?.id).toBe(10);
    expect(ws.pageBreaks?.rowBreaks?.[0]?.man).toBe(true);
    expect(ws.pageBreaks?.rowBreaks?.[1]?.id).toBe(25);
  });

  it('adds col breaks via helper', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addColBreak(3);

    expect(ws.pageBreaks?.colBreaks).toHaveLength(1);
    expect(ws.pageBreaks?.colBreaks?.[0]?.id).toBe(3);
    expect(ws.pageBreaks?.colBreaks?.[0]?.max).toBe(1048575);
  });

  it('clears page breaks with null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.addRowBreak(10);
    ws.pageBreaks = null;

    expect(ws.pageBreaks).toBeNull();
  });

  it('returns null when no page breaks set', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.pageBreaks).toBeNull();
  });
});

describe('Worksheet Print Area (convenience)', () => {
  it('sets and gets print area on worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.setPrintArea('Sheet1!$A$1:$D$50');
    expect(ws.getPrintArea()).toBe('Sheet1!$A$1:$D$50');
    // Also accessible via workbook-level API
    expect(wb.getPrintArea(0)).toBe('Sheet1!$A$1:$D$50');
  });

  it('clears print area on worksheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.setPrintArea('Sheet1!$A$1:$D$50');
    ws.setPrintArea(null);
    expect(ws.getPrintArea()).toBeNull();
  });
});

describe('Roundtrip: Page Breaks', () => {
  it('roundtrips row and column page breaks through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';

    ws.addRowBreak(10);
    ws.addRowBreak(20);
    ws.addColBreak(5);

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.pageBreaks).not.toBeNull();
    expect(ws2?.pageBreaks?.rowBreaks).toHaveLength(2);
    expect(ws2?.pageBreaks?.rowBreaks?.[0]?.id).toBe(10);
    expect(ws2?.pageBreaks?.rowBreaks?.[0]?.man).toBe(true);
    expect(ws2?.pageBreaks?.rowBreaks?.[1]?.id).toBe(20);
    expect(ws2?.pageBreaks?.colBreaks).toHaveLength(1);
    expect(ws2?.pageBreaks?.colBreaks?.[0]?.id).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// Page Margins
// ---------------------------------------------------------------------------

describe('worksheet page margins', () => {
  it('round-trip all six margins and are written before <pageSetup>', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'x';
    const margins = { top: 1.5, bottom: 1.5, left: 1, right: 1, header: 0.5, footer: 0.5 };
    ws.pageSetup = { orientation: 'landscape' };
    ws.pageMargins = margins;

    const bytes = await wb.toBuffer();
    expect((await readBuffer(bytes)).getSheet('S')?.pageMargins).toEqual(margins);

    const xml = readZipText(bytes)['xl/worksheets/sheet1.xml'] ?? '';
    expect(xml).toContain('<pageMargins ');
    expect(xml.indexOf('<pageMargins ')).toBeLessThan(xml.indexOf('<pageSetup '));
  });

  it('margins alone write no empty <pageSetup/>', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.cell('A1').value = 'x';
    ws.pageMargins = { top: 1.5, bottom: 1.5, left: 1, right: 1, header: 0.5, footer: 0.5 };

    const xml = readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'] ?? '';
    expect(xml).toContain('<pageMargins ');
    expect(xml).not.toContain('<pageSetup');
  });
});

describe('Roundtrip: fit to page', () => {
  it('writes pageSetUpPr fitToPage and reads it back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    const buf = await wb.toBuffer();
    const sheetXml = readZipText(buf)['xl/worksheets/sheet1.xml'];
    expect(sheetXml).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(sheetXml).toContain('fitToWidth="1"');

    const ps = (await readBuffer(buf)).getSheet('Sheet1')?.pageSetup;
    expect(ps?.fitToPage).toBe(true);
    expect(ps?.fitToWidth).toBe(1);
    expect(ps?.fitToHeight).toBe(0);
  });

  it('writes no sheetPr when fitToPage is unset', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.pageSetup = { fitToWidth: 1 };

    const sheetXml = readZipText(await wb.toBuffer())['xl/worksheets/sheet1.xml'];
    expect(sheetXml).not.toContain('pageSetUpPr');
    expect(sheetXml).not.toContain('<sheetPr');
  });
});
