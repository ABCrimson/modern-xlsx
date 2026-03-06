import { describe, expect, it } from 'vitest';
import { RichTextBuilder, readBuffer, Workbook } from '../src/index.js';
import type { RichTextRun } from '../src/types.js';

describe('RichTextBuilder', () => {
  it('builds plain and bold runs', () => {
    const runs = new RichTextBuilder().text('Hello ').bold('World').build();
    expect(runs).toHaveLength(2);
    expect(runs[0]).toEqual({ text: 'Hello ' });
    expect(runs[1]).toEqual({ text: 'World', bold: true });
  });

  it('builds italic and bold+italic runs', () => {
    const runs = new RichTextBuilder().italic('foo').boldItalic('bar').build();
    expect(runs).toHaveLength(2);
    expect(runs[0]).toEqual({ text: 'foo', italic: true });
    expect(runs[1]).toEqual({ text: 'bar', bold: true, italic: true });
  });

  it('builds underline runs', () => {
    const runs = new RichTextBuilder().underline('underlined').build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({ text: 'underlined', underline: true });
  });

  it('builds strike runs', () => {
    const runs = new RichTextBuilder().strike('deleted').build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({ text: 'deleted', strike: true });
  });

  it('builds colored runs', () => {
    const runs = new RichTextBuilder().colored('red text', 'FF0000').build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({ text: 'red text', color: 'FF0000' });
  });

  it('builds styled runs with all options', () => {
    const runs = new RichTextBuilder()
      .styled('fancy', {
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        fontName: 'Arial',
        fontSize: 14,
        color: 'FF0000',
      })
      .build();
    expect(runs).toHaveLength(1);
    expect(runs[0]).toEqual({
      text: 'fancy',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      fontName: 'Arial',
      fontSize: 14,
      color: 'FF0000',
    });
  });

  it('plainText() returns concatenated text', () => {
    const builder = new RichTextBuilder().bold('Hello').text(' ').italic('World');
    expect(builder.plainText()).toBe('Hello World');
  });
});

describe('Cell.richText', () => {
  it('is undefined for a new cell', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');
    cell.value = 'plain text';
    expect(cell.richText).toBeUndefined();
  });

  it('can be set and read back', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');

    const runs: RichTextRun[] = [
      { text: 'Hello ', bold: true },
      { text: 'World', italic: true },
    ];
    cell.richText = runs;

    expect(cell.richText).toHaveLength(2);
    expect(cell.richText?.[0]).toEqual({ text: 'Hello ', bold: true });
    expect(cell.richText?.[1]).toEqual({ text: 'World', italic: true });
    expect(cell.value).toBe('Hello World');
    expect(cell.type).toBe('sharedString');
  });

  it('auto-updates value from runs text', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');

    cell.richText = [{ text: 'Part1' }, { text: ' Part2', color: 'FF0000' }];

    expect(cell.value).toBe('Part1 Part2');
  });

  it('can be cleared by setting undefined', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');

    cell.richText = [{ text: 'bold', bold: true }];
    expect(cell.richText).toBeDefined();

    cell.richText = undefined;
    expect(cell.richText).toBeUndefined();
  });

  it('integrates with RichTextBuilder', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    const cell = ws.cell('A1');

    const runs = new RichTextBuilder().text('Normal ').bold('Bold ').italic('Italic').build();
    cell.richText = runs;

    expect(cell.richText).toHaveLength(3);
    expect(cell.value).toBe('Normal Bold Italic');
  });
});

describe('Rich text roundtrip', () => {
  it('preserves rich text through write/read cycle', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    // Set up a cell with rich text
    ws.cell('A1').richText = [
      { text: 'Bold', bold: true },
      { text: ' and ' },
      { text: 'Italic', italic: true },
    ];

    // Plain text cell for comparison
    ws.cell('A2').value = 'plain text';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');
    expect(ws2).toBeDefined();

    // Rich text cell should preserve runs
    const cell1 = ws2?.cell('A1');
    expect(cell1?.value).toBe('Bold and Italic');
    expect(cell1?.richText).toBeDefined();
    expect(cell1?.richText).toHaveLength(3);
    expect(cell1?.richText?.[0]).toMatchObject({ text: 'Bold', bold: true });
    expect(cell1?.richText?.[1]).toMatchObject({ text: ' and ' });
    expect(cell1?.richText?.[2]).toMatchObject({ text: 'Italic', italic: true });

    // Plain text cell should not have rich text
    const cell2 = ws2?.cell('A2');
    expect(cell2.value).toBe('plain text');
    expect(cell2.richText).toBeUndefined();
  });

  it('preserves underline and strike through roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.cell('A1').richText = [
      { text: 'underlined', underline: true },
      { text: ' and ' },
      { text: 'struck', strike: true },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');

    const cell = ws2?.cell('A1');
    expect(cell?.richText).toHaveLength(3);
    expect(cell?.richText?.[0]).toMatchObject({ text: 'underlined', underline: true });
    expect(cell?.richText?.[2]).toMatchObject({ text: 'struck', strike: true });
  });

  it('preserves font properties through roundtrip', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.cell('A1').richText = [
      {
        text: 'styled',
        bold: true,
        italic: true,
        underline: true,
        strike: true,
        fontName: 'Arial',
        fontSize: 16,
        color: 'FF0000',
      },
    ];

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');

    const runs = ws2?.cell('A1').richText;
    expect(runs).toHaveLength(1);
    expect(runs?.[0]).toMatchObject({
      text: 'styled',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      fontName: 'Arial',
      fontSize: 16,
      color: 'FF0000',
    });
  });

  it('handles multiple cells with different rich text', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');

    ws.cell('A1').richText = [{ text: 'bold', bold: true }];
    ws.cell('A2').richText = [{ text: 'italic', italic: true }];
    ws.cell('A3').value = 'plain';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Sheet1');

    expect(ws2?.cell('A1').richText?.[0]).toMatchObject({ text: 'bold', bold: true });
    expect(ws2?.cell('A2').richText?.[0]).toMatchObject({ text: 'italic', italic: true });
    expect(ws2?.cell('A3').richText).toBeUndefined();
  });
});
