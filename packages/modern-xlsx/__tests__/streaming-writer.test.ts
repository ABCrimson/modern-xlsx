import { describe, expect, it } from 'vitest';
import { ModernXlsxError, readBuffer, StreamingXlsxWriter, Workbook } from '../src/index.js';
import type { StreamingCellInput } from '../src/io/streaming-writer.js';

describe('StreamingXlsxWriter', () => {
  // -----------------------------------------------------------------------
  // 1. Basic: write 1K rows and verify output is valid XLSX
  // -----------------------------------------------------------------------
  it('writes 1K rows and produces a valid XLSX', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Data');
    const ROWS = 1000;
    const COLS = 5;

    for (let r = 0; r < ROWS; r++) {
      const cells: StreamingCellInput[] = [];
      for (let c = 0; c < COLS; c++) {
        cells.push({
          value: String(r * COLS + c),
          cellType: 'number',
        });
      }
      writer.writeRow(cells);
    }

    const xlsx = writer.finish();
    expect(xlsx).toBeInstanceOf(Uint8Array);
    expect(xlsx.length).toBeGreaterThan(0);

    // Read back and verify
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Data');
    expect(ws).toBeDefined();
    expect(ws?.rows).toHaveLength(ROWS);
    // First row, first cell
    expect(ws?.cell('A1').value).toBe(0);
    // Last row, last cell
    expect(ws?.cell('E1000').value).toBe(999 * 5 + 4);
  });

  // -----------------------------------------------------------------------
  // 2. Large: write 100K rows without OOM
  // -----------------------------------------------------------------------
  it('writes 100K rows without OOM', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('BigData');
    const ROWS = 100_000;

    for (let r = 0; r < ROWS; r++) {
      writer.writeRow([
        { value: String(r), cellType: 'number' },
        { value: `row_${r}`, cellType: 'sharedString' },
      ]);
    }

    const xlsx = writer.finish();
    expect(xlsx).toBeInstanceOf(Uint8Array);
    // The output should be non-trivial in size
    expect(xlsx.length).toBeGreaterThan(1_000_000);
  });

  // -----------------------------------------------------------------------
  // 3. Multi-sheet streaming
  // -----------------------------------------------------------------------
  it('writes multiple sheets', async () => {
    const writer = StreamingXlsxWriter.create();

    writer.startSheet('First');
    writer.writeRow([{ value: 'hello', cellType: 'sharedString' }]);
    writer.writeRow([{ value: '42', cellType: 'number' }]);

    writer.startSheet('Second');
    writer.writeRow([{ value: 'world', cellType: 'sharedString' }]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);

    expect(wb.sheetNames).toEqual(['First', 'Second']);

    const s1 = wb.getSheet('First');
    expect(s1?.rows).toHaveLength(2);
    expect(s1?.cell('A1').value).toBe('hello');
    expect(s1?.cell('A2').value).toBe(42);

    const s2 = wb.getSheet('Second');
    expect(s2?.rows).toHaveLength(1);
    expect(s2?.cell('A1').value).toBe('world');
  });

  // -----------------------------------------------------------------------
  // 4. SST deduplication
  // -----------------------------------------------------------------------
  it('builds SST correctly with deduplication', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('SST');

    // Write duplicate strings
    writer.writeRow([
      { value: 'Alpha', cellType: 'sharedString' },
      { value: 'Beta', cellType: 'sharedString' },
    ]);
    writer.writeRow([
      { value: 'Alpha', cellType: 'sharedString' },
      { value: 'Gamma', cellType: 'sharedString' },
    ]);
    writer.writeRow([
      { value: 'Beta', cellType: 'sharedString' },
      { value: 'Beta', cellType: 'sharedString' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);

    // Only 3 unique strings in the SST
    expect(wb.toJSON().sharedStrings?.strings).toHaveLength(3);

    // Verify cell values
    const ws = wb.getSheet('SST');
    expect(ws).toBeDefined();
    expect(ws?.cell('A1').value).toBe('Alpha');
    expect(ws?.cell('B1').value).toBe('Beta');
    expect(ws?.cell('A2').value).toBe('Alpha');
    expect(ws?.cell('B2').value).toBe('Gamma');
    expect(ws?.cell('A3').value).toBe('Beta');
    expect(ws?.cell('B3').value).toBe('Beta');
  });

  // -----------------------------------------------------------------------
  // 5. Mixed cell types
  // -----------------------------------------------------------------------
  it('supports mixed cell types in a row', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Mixed');

    writer.writeRow([
      { value: '3.14', cellType: 'number' },
      { value: 'text', cellType: 'sharedString' },
      { value: 'true', cellType: 'boolean' },
      { value: 'inline text', cellType: 'inlineStr' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Mixed');
    expect(ws).toBeDefined();

    expect(ws?.cell('A1').value).toBe(3.14);
    expect(ws?.cell('B1').value).toBe('text');
    expect(ws?.cell('C1').value).toBe(true);
    expect(ws?.cell('D1').value).toBe('inline text');
  });

  // -----------------------------------------------------------------------
  // 6. Empty cells are skipped (sparse rows)
  // -----------------------------------------------------------------------
  it('skips empty cells', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Sparse');

    writer.writeRow([
      { value: 'A', cellType: 'sharedString' },
      // @ts-expect-error -- deliberate: exactOptionalPropertyTypes forbids an explicit `undefined`, but callers from plain JS can still pass one and the writer must skip it.
      { value: undefined },
      { value: 'C', cellType: 'sharedString' },
    ]);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Sparse');
    expect(ws).toBeDefined();

    expect(ws?.cell('A1').value).toBe('A');
    expect(ws?.cell('B1').value).toBeNull();
    expect(ws?.cell('C1').value).toBe('C');
  });

  // -----------------------------------------------------------------------
  // 6b. Error: writeRow before startSheet
  // -----------------------------------------------------------------------
  it('throws XML_WRITE when writeRow is called before startSheet', () => {
    const writer = StreamingXlsxWriter.create();
    expect(() => writer.writeRow([{ value: '1' }])).toThrow(
      expect.objectContaining({ code: 'XML_WRITE' }),
    );
  });

  // -----------------------------------------------------------------------
  // 7. Error: finish without sheets
  // -----------------------------------------------------------------------
  it('throws when finishing without any sheets', () => {
    const writer = StreamingXlsxWriter.create();
    expect(() => writer.finish()).toThrow();
  });

  // -----------------------------------------------------------------------
  // 8. Error: use after finish
  // -----------------------------------------------------------------------
  it('throws when using writer after finish', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S1');
    writer.writeRow([{ value: '1', cellType: 'number' }]);
    writer.finish();

    const calls = [
      () => writer.startSheet('S2'),
      () => writer.writeRow([{ value: '2' }]),
      () => writer.finish(),
      () => writer.setStylesXml('<styleSheet/>'),
    ];
    for (const call of calls) {
      expect(call).toThrow('StreamingWriter already finished');
      expect(call).toThrow(expect.objectContaining({ code: 'INVALID_ARGUMENT' }));
    }
  });

  // -----------------------------------------------------------------------
  // 8b. Custom styles via setStylesXml
  // -----------------------------------------------------------------------
  it('applies custom styles set with setStylesXml', async () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs>' +
      '</styleSheet>';

    const writer = StreamingXlsxWriter.create();
    writer.setStylesXml(stylesXml);
    writer.startSheet('Styled');
    writer.writeRow([{ value: '1.5', cellType: 'number', style: 1 }]);
    const wb = await readBuffer(writer.finish());

    const cell = wb.getSheet('Styled')?.cell('A1');
    expect(cell?.styleIndex).toBe(1);
    expect(cell?.value).toBe(1.5);
    expect(wb.styles.cellXfs[1]).toMatchObject({ numFmtId: 164, fontId: 1 });
    expect(wb.styles.numFmts).toEqual([{ id: 164, formatCode: '0.000' }]);
    expect(wb.styles.fonts[1]?.bold).toBe(true);
  });

  // -----------------------------------------------------------------------
  // 9. Wide columns (beyond Z -> AA, AB, ...)
  // -----------------------------------------------------------------------
  it('handles wide rows (30+ columns)', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Wide');

    const cells: StreamingCellInput[] = [];
    for (let c = 0; c < 30; c++) {
      cells.push({ value: String(c), cellType: 'number' });
    }
    writer.writeRow(cells);

    const xlsx = writer.finish();
    const wb = await readBuffer(xlsx);
    const ws = wb.getSheet('Wide');
    expect(ws).toBeDefined();

    // Column 0 = A, column 25 = Z, column 26 = AA, column 27 = AB, etc.
    expect(ws?.cell('A1').value).toBe(0);
    expect(ws?.cell('Z1').value).toBe(25);
    expect(ws?.cell('AA1').value).toBe(26);
    expect(ws?.cell('AD1').value).toBe(29);
  });

  // -----------------------------------------------------------------------
  // 10. Performance: 100K rows with 5 columns
  // -----------------------------------------------------------------------
  it('writes 100K x 5 cols in reasonable time', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Perf');
    const ROWS = 100_000;

    for (let r = 0; r < ROWS; r++) {
      writer.writeRow([
        { value: String(r), cellType: 'number' },
        { value: String(r * 2), cellType: 'number' },
        { value: String(r * 3), cellType: 'number' },
        { value: `name_${r % 100}`, cellType: 'sharedString' },
        { value: r % 2 === 0 ? 'true' : 'false', cellType: 'boolean' },
      ]);
    }
    const xlsx = writer.finish();

    expect(xlsx.length).toBeGreaterThan(0);
  });

  it('rejects an injected number value and keeps the stream usable', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S');
    let caught: unknown;
    try {
      writer.writeRow([{ value: '1</v></c><c r="Z9"><v>2', cellType: 'number' }]);
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(ModernXlsxError);
    expect((caught as ModernXlsxError).code).toBe('INVALID_CELL_VALUE');

    writer.writeRow([{ value: '7', cellType: 'number' }]);
    const wb = await readBuffer(writer.finish());
    const rows = wb.getSheet('S')?.rows ?? [];
    expect(rows).toHaveLength(1);
    expect(wb.getSheet('S')?.cell('A1').value).toBe(7);
  });
});

/**
 * `startSheet` refuses a sheet name that repeats an earlier one as
 * `Workbook.addSheet` compares them: ignoring case and canonical
 * equivalence, final sigma folded. Excel 365 (build 20527, 2026-11-13)
 * refused to open the files the writer wrote before (sheets café with U+00E9
 * and cafe + U+0301; AΣ and aσ; bσx and bςx), and addSheet refuses the
 * second of each pair.
 */
describe('StreamingXlsxWriter sheet names compared as addSheet compares', () => {
  const pairs: ReadonlyArray<readonly [string, string, string]> = [
    ['canonical', 'caf\u00E9', 'cafe\u0301'],
    ['sigma', 'A\u03A3', 'a\u03C3'],
    ['final sigma', 'b\u03C3x', 'b\u03C2x'],
    ['case', 'Data', 'DATA'],
  ];

  it.each(pairs)(
    'refuses the second of a %s pair and keeps the first sheet open',
    async (_, first, second) => {
      const writer = StreamingXlsxWriter.create();
      writer.startSheet(first);
      writer.writeRow([{ value: '1', cellType: 'number' }]);
      let caught: unknown;
      try {
        writer.startSheet(second);
      } catch (e) {
        caught = e;
      }
      expect(caught).toBeInstanceOf(ModernXlsxError);
      expect((caught as ModernXlsxError).code).toBe('INVALID_CELL_VALUE');
      expect((caught as ModernXlsxError).message).toBe(
        `invalid cell value: Invalid sheet name '${second}': duplicate of '${first}'`,
      );
      writer.writeRow([{ value: '2', cellType: 'number' }]);
      const wb = await readBuffer(writer.finish());
      expect(wb.sheetNames).toEqual([first]);
      expect(wb.getSheet(first)?.cell('A2').value).toBe(2);
      const both = new Workbook();
      both.addSheet(first);
      expect(() => both.addSheet(second)).toThrow('already exists');
    },
  );

  it('does not count a name refused for another reason as started', async () => {
    // 27 + 5 = 32 UTF-16 units: too long; the composed spelling (31 units) has the same key.
    const tooLong = `${'x'.repeat(27)}cafe\u0301`;
    const composed = `${'x'.repeat(27)}caf\u00E9`;
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S1');
    expect(() => writer.startSheet(tooLong)).toThrow(
      expect.objectContaining({ code: 'INVALID_CELL_VALUE' }),
    );
    expect(() => writer.startSheet('Bad:Name')).toThrow(
      expect.objectContaining({ code: 'INVALID_CELL_VALUE' }),
    );
    writer.startSheet(composed);
    writer.startSheet('Bad Name');
    expect(() => writer.startSheet(tooLong)).toThrow(`duplicate of '${composed}'`);
    const wb = await readBuffer(writer.finish());
    expect(wb.sheetNames).toEqual(['S1', composed, 'Bad Name']);
  });

  it.each([
    ['a number', 2024],
    ['undefined', undefined],
    ['null', null],
    ['a String object', Object('Data')],
  ])('refuses a name that is %s with INVALID_CELL_VALUE and keeps writing', async (_, name) => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('Data');
    writer.writeRow([{ value: '1', cellType: 'number' }]);
    let caught: unknown;
    try {
      writer.startSheet(name as string);
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(ModernXlsxError);
    expect((caught as ModernXlsxError).code).toBe('INVALID_CELL_VALUE');
    expect((caught as ModernXlsxError).message).toBe(
      'invalid cell value: Invalid sheet name: not a string',
    );
    writer.startSheet('Next');
    const wb = await readBuffer(writer.finish());
    expect(wb.sheetNames).toEqual(['Data', 'Next']);
    expect(wb.getSheet('Data')?.cell('A1').value).toBe(1);
  });

  it('reports a disposed writer with INVALID_ARGUMENT', () => {
    const writer = StreamingXlsxWriter.create();
    writer[Symbol.dispose]();
    for (const name of ['S1', 2024]) {
      expect(() => writer.startSheet(name as string)).toThrow(
        expect.objectContaining({
          code: 'INVALID_ARGUMENT',
          message: 'StreamingXlsxWriter has been disposed',
        }),
      );
    }
  });

  it('still reports a finished writer before a repeated name', () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S1');
    writer.finish();
    expect(() => writer.startSheet('s1')).toThrow('StreamingWriter already finished');
  });
});
