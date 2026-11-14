import { assert, describe, expect, it } from 'vitest';
import {
  drawTable,
  HeaderFooterBuilder,
  RichTextBuilder,
  StyleBuilder,
  Workbook,
} from '../src/index.js';

describe('HeaderFooterBuilder contract', () => {
  it('text() doubles ampersands', () => {
    expect(HeaderFooterBuilder.text('R&D')).toBe('R&&D');
    expect(HeaderFooterBuilder.text('plain')).toBe('plain');
  });

  it('fontSize() separates a digit-leading text and validates the size', () => {
    expect(HeaderFooterBuilder.fontSize(12, '2026')).toBe('&12 2026');
    expect(HeaderFooterBuilder.fontSize(12, 'Hi')).toBe('&12Hi');
    expect(HeaderFooterBuilder.fontSize(12)).toBe('&12');
    for (const size of [0, 410, 1.5, Number.NaN, -3]) {
      expect(() => HeaderFooterBuilder.fontSize(size)).toThrow(RangeError);
    }
    expect(HeaderFooterBuilder.fontSize(409)).toBe('&409');
  });

  it('build() enforces the 255 character limit', () => {
    expect(new HeaderFooterBuilder().left('x'.repeat(253)).build()).toHaveLength(255);
    expect(() => new HeaderFooterBuilder().left('x'.repeat(254)).build()).toThrow(
      'header/footer string is 256 characters; Excel allows 255',
    );
  });

  it('color() accepts RGB and ARGB and rejects anything else', () => {
    expect(HeaderFooterBuilder.color('ff0000')).toBe('&KFF0000');
    expect(HeaderFooterBuilder.color('80FF0000')).toBe('&KFF0000');
    expect(() => HeaderFooterBuilder.color('red')).toThrow(RangeError);
    expect(() => HeaderFooterBuilder.color('FF00')).toThrow(RangeError);
  });
});

describe('RichTextBuilder contract', () => {
  it('build() returns a frozen copy of frozen runs', () => {
    const b = new RichTextBuilder().text('a').bold('b');
    const runs = b.build();
    expect(Object.isFrozen(runs)).toBe(true);
    expect(runs.every((r) => Object.isFrozen(r))).toBe(true);
    b.italic('c');
    expect(runs).toHaveLength(2);
    expect(b.build()).toHaveLength(3);
  });

  it('assigning the runs to a cell stores runs that are not frozen', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    const cell = ws.cell('A1');
    cell.richText = new RichTextBuilder().text('a').build();
    const stored = cell.richText;
    expect(stored).toBeDefined();
    expect(Object.isFrozen(stored?.[0])).toBe(false);
  });
});

describe('builders keep their state private', () => {
  it('exposes no own enumerable keys', () => {
    expect(Object.keys(new StyleBuilder())).toEqual([]);
    expect(Object.keys(new RichTextBuilder())).toEqual([]);
    expect(Object.keys(new HeaderFooterBuilder())).toEqual([]);
  });

  describe('StyleBuilder interning', () => {
    const sizes = (wb: Workbook) => [
      wb.styles.fonts.length,
      wb.styles.fills.length,
      wb.styles.borders.length,
      wb.styles.cellXfs.length,
    ];

    it('identical builds return the same index and add nothing the second time', () => {
      const wb = new Workbook();
      const make = () =>
        new StyleBuilder()
          .font({ bold: true })
          .fill({ pattern: 'solid', fgColor: 'FFFF00' })
          .build(wb.styles);
      const before = sizes(wb);
      const a = make();
      const after = sizes(wb);
      for (let i = 0; i < 4; i++) expect(after[i]).toBeLessThanOrEqual((before[i] ?? 0) + 1);
      expect(make()).toBe(a);
      expect(sizes(wb)).toEqual(after);
    });

    it('does not reuse an entry edited in place', () => {
      const wb = new Workbook();
      const make = () => new StyleBuilder().font({ bold: true }).build(wb.styles);
      const a = make();
      const xf = wb.styles.cellXfs[a];
      assert.isDefined(xf);
      xf.numFmtId = 9;
      const b = make();
      expect(b).not.toBe(a);
      expect(wb.styles.cellXfs[b]?.numFmtId).toBe(0);
    });

    it('drawTable with one override on 100 cells adds at most one xf', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      const cellStyles: Record<string, { font: { bold: boolean } }> = {};
      for (let r = 0; r < 100; r++) cellStyles[`${r},0`] = { font: { bold: true } };
      const rows = Array.from({ length: 100 }, (_, i) => [i]);
      const noOverride = new Workbook();
      drawTable(noOverride, noOverride.addSheet('S'), { headers: ['N'], rows });
      drawTable(wb, ws, { headers: ['N'], rows, cellStyles });
      expect(wb.styles.cellXfs.length - noOverride.styles.cellXfs.length).toBeLessThanOrEqual(1);
    });
  });
});
