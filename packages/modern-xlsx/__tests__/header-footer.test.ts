import { describe, expect, it } from 'vitest';
import type { HeaderFooterData } from '../src/index.js';
import { HeaderFooterBuilder, readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

// ---------------------------------------------------------------------------
// Headers / Footers
// ---------------------------------------------------------------------------

describe('Header/Footer', () => {
  it('sets and gets header/footer data', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    const hf: HeaderFooterData = {
      oddHeader: '&CConfidential',
      oddFooter: '&LPage &P&R&D',
    };

    ws.headerFooter = hf;
    expect(ws.headerFooter).toEqual(hf);
  });

  it('clears header/footer with null', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.headerFooter = { oddHeader: '&CTest' };
    ws.headerFooter = null;
    expect(ws.headerFooter).toBeNull();
  });
});

describe('HeaderFooterBuilder', () => {
  it('builds a basic header string', () => {
    const result = new HeaderFooterBuilder()
      .left('Left Text')
      .center('Center Text')
      .right('Right Text')
      .build();

    expect(result).toBe('&LLeft Text&CCenter Text&RRight Text');
  });

  it('uses static code helpers', () => {
    expect(HeaderFooterBuilder.pageNumber()).toBe('&P');
    expect(HeaderFooterBuilder.totalPages()).toBe('&N');
    expect(HeaderFooterBuilder.date()).toBe('&D');
    expect(HeaderFooterBuilder.time()).toBe('&T');
    expect(HeaderFooterBuilder.fileName()).toBe('&F');
    expect(HeaderFooterBuilder.sheetName()).toBe('&A');
    expect(HeaderFooterBuilder.filePath()).toBe('&Z');
  });

  it('uses font formatting helpers', () => {
    expect(HeaderFooterBuilder.bold('test')).toBe('&Btest&B');
    expect(HeaderFooterBuilder.italic('test')).toBe('&Itest&I');
    expect(HeaderFooterBuilder.underline('test')).toBe('&Utest&U');
    expect(HeaderFooterBuilder.strikethrough('test')).toBe('&Stest&S');
    expect(HeaderFooterBuilder.fontSize(12)).toBe('&12');
    expect(HeaderFooterBuilder.fontName('Arial')).toBe('&"Arial"');
    expect(HeaderFooterBuilder.color('FF0000')).toBe('&KFF0000');
  });

  it('composes a complex header/footer', () => {
    const hf = new HeaderFooterBuilder()
      .left(`Printed: ${HeaderFooterBuilder.date()}`)
      .center(HeaderFooterBuilder.bold('Confidential'))
      .right(`Page ${HeaderFooterBuilder.pageNumber()} of ${HeaderFooterBuilder.totalPages()}`)
      .build();

    expect(hf).toBe('&LPrinted: &D&C&BConfidential&B&RPage &P of &N');
  });
});

// ---------------------------------------------------------------------------
// Roundtrip tests (require WASM)
// ---------------------------------------------------------------------------

describe('Roundtrip: Header/Footer', () => {
  it('roundtrips header/footer through WASM', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'data';
    ws.headerFooter = {
      oddHeader: '&CConfidential',
      oddFooter: '&LPage &P&RDate: &D',
      differentOddEven: false,
      differentFirst: false,
    };

    const buf = await wb.toBuffer();
    const wb2 = await readBuffer(buf);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.headerFooter?.oddHeader).toBe('&CConfidential');
    expect(ws2?.headerFooter?.oddFooter).toBe('&LPage &P&RDate: &D');
  });

  it('roundtrips all six parts and flags, and escapes & in the XML', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'x';
    const headerFooter: HeaderFooterData = {
      oddHeader: '&CR&&D',
      oddFooter: '&Lodd',
      evenHeader: '&Ceven',
      evenFooter: '&Leven',
      firstHeader: '&Cfirst',
      firstFooter: '&Rfirst',
      differentOddEven: true,
      differentFirst: true,
      scaleWithDoc: false,
      alignWithMargins: false,
    };
    ws.headerFooter = headerFooter;

    const buf = await wb.toBuffer();
    expect((await readBuffer(buf)).getSheet('Sheet1')?.headerFooter).toEqual(headerFooter);
    expect(readZipText(buf)['xl/worksheets/sheet1.xml']).toContain(
      '<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0"><oddHeader>&amp;CR&amp;&amp;D</oddHeader><oddFooter>&amp;Lodd</oddFooter><evenHeader>&amp;Ceven</evenHeader><evenFooter>&amp;Leven</evenFooter><firstHeader>&amp;Cfirst</firstHeader><firstFooter>&amp;Rfirst</firstFooter></headerFooter>',
    );
  });
});
