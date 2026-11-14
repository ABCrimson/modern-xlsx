import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook, Worksheet } from '../src/index.js';
import { nameKey } from '../src/model/name-rules.js';
import type { WorksheetData } from '../src/model/types.js';
import { readZipText } from './helpers/zip.js';

describe('Sheet Management', () => {
  // --- Sheet State ---

  it('default sheet state is visible', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    expect(ws.state).toBe('visible');
  });

  it('set sheet state to hidden', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws2.state = 'hidden';
    expect(ws2.state).toBe('hidden');
  });

  it('hidden sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'data';
    const ws2 = wb.addSheet('Sheet2');
    ws2.cell('A1').value = 'hidden';
    ws2.state = 'hidden';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.getSheet('Sheet2')?.state).toBe('hidden');
    expect(wb2.getSheet('Sheet1')?.state).toBe('visible');
  });

  it('cannot hide last visible sheet', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    expect(() => wb.hideSheet('Sheet1')).toThrow('last visible');
  });

  it('unhide restores visible state', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const ws2 = wb.addSheet('Sheet2');
    ws2.state = 'hidden';
    wb.unhideSheet('Sheet2');
    expect(wb.getSheet('Sheet2')?.state).toBe('visible');
  });

  it('veryHidden sheet survives roundtrip and is written as state="veryHidden"', async () => {
    const wb = new Workbook();
    wb.addSheet('A').cell('A1').value = 'x';
    wb.addSheet('B').state = 'veryHidden';

    const bytes = await wb.toBuffer();
    expect(readZipText(bytes)['xl/workbook.xml']).toContain('state="veryHidden"');
    expect((await readBuffer(bytes)).getSheet('B')?.state).toBe('veryHidden');
  });

  it('veryHidden does not count as visible for the last-visible guard', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B').state = 'veryHidden';
    expect(() => wb.hideSheet('A')).toThrow('last visible');
  });

  it('hideSheet keeps veryHidden and unhideSheet makes it visible', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B').state = 'veryHidden';
    wb.hideSheet('B');
    expect(wb.getSheet('B')?.state).toBe('veryHidden');
    wb.unhideSheet('B');
    expect(wb.getSheet('B')?.state).toBe('visible');
  });

  // --- Move Sheet ---

  it('move sheet changes order', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');
    wb.moveSheet(2, 0);
    expect(wb.sheetNames).toEqual(['C', 'A', 'B']);
  });

  // --- Clone Sheet ---

  it('clone sheet duplicates content', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Original');
    ws.cell('A1').value = 'hello';
    const clone = wb.cloneSheet(0, 'Copy');
    expect(clone.cell('A1').value).toBe('hello');
    expect(wb.sheetCount).toBe(2);
  });

  it('clone sheet survives roundtrip', async () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1').cell('A1').value = 'original';
    wb.cloneSheet(0, 'Clone');

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    expect(wb2.sheetCount).toBe(2);
    expect(wb2.getSheet('Clone')?.cell('A1').value).toBe('original');
  });

  // --- Rename Sheet ---

  it('rename sheet updates name', () => {
    const wb = new Workbook();
    wb.addSheet('OldName');
    wb.renameSheet('OldName', 'NewName');
    expect(wb.sheetNames).toEqual(['NewName']);
    expect(wb.getSheet('NewName')).toBeDefined();
  });

  it('rename rejects duplicate name', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(() => wb.renameSheet('A', 'B')).toThrow('already exists');
  });

  // --- Name integrity ---

  it('rejects a sheet name that differs only in case', () => {
    const wb = new Workbook();
    wb.addSheet('data');
    expect(() => wb.addSheet('DATA')).toThrow('already exists');
    expect(() => wb.cloneSheet(0, 'Data')).toThrow('already exists');
    wb.addSheet('Other');
    expect(() => wb.renameSheet('Other', 'dAtA')).toThrow('already exists');
    expect(wb.sheetNames).toEqual(['data', 'Other']);
  });

  it("rejects the reserved name 'History' in any case", () => {
    const wb = new Workbook();
    expect(() => wb.addSheet('History')).toThrow("'History' is a reserved sheet name");
    expect(() => wb.addSheet('history')).toThrow("'History' is a reserved sheet name");
    const ws = wb.addSheet('Log');
    expect(() => wb.renameSheet('Log', 'HISTORY')).toThrow('reserved');
    expect(() => {
      ws.name = 'History';
    }).toThrow('reserved');
    expect(wb.sheetNames).toEqual(['Log']);
  });

  it('allows a case-only rename of the same sheet', () => {
    const wb = new Workbook();
    wb.addSheet('data');
    wb.renameSheet('data', 'Data');
    expect(wb.sheetNames).toEqual(['Data']);
    expect(wb.getSheet('Data')).toBeDefined();
  });

  it('name setter renames through the workbook', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Old');
    wb.addNamedRange('Ref', 'Old!$A$1');
    ws.name = 'New';
    expect(ws.name).toBe('New');
    expect(wb.getSheet('New')?.name).toBe('New');
    expect(wb.getSheet('Old')).toBeUndefined();
    expect(wb.getNamedRange('Ref')?.value).toBe('New!$A$1');
  });

  it('name setter rejects a duplicate of another sheet', () => {
    const wb = new Workbook();
    wb.addSheet('New');
    const other = wb.addSheet('Other');
    expect(() => wb.renameSheet('Other', 'NEW')).toThrow('already exists');
    expect(() => {
      other.name = 'new';
    }).toThrow('already exists');
    expect(wb.sheetNames).toEqual(['New', 'Other']);
  });

  it('name setter still works on a standalone worksheet', () => {
    const ws = new Worksheet({
      name: 'Loose',
      worksheet: {
        dimension: null,
        rows: [],
        mergeCells: [],
        autoFilter: null,
        frozenPane: null,
        columns: [],
      },
    });
    ws.name = 'T';
    expect(ws.name).toBe('T');
    expect(() => {
      ws.name = 'history';
    }).toThrow('reserved');
  });

  // --- Sheet index remapping ---

  it('removeSheet remaps scoped names and drops the removed sheet names', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.setPrintArea(0, 'A!$A$1:$B$2');
    wb.setPrintArea(1, 'B!$A$1:$C$3');
    wb.addNamedRange('LocalA', 'A!$A$1', 0);
    wb.addNamedRange('Global', 'B!$A$1');
    wb.workbookViews = [
      {
        activeTab: 1,
        firstSheet: 1,
        showHorizontalScroll: true,
        showVerticalScroll: true,
        showSheetTabs: true,
      },
    ];
    wb.removeSheet(0);
    expect(wb.namedRanges.map((d) => [d.name, d.sheetId])).toEqual([
      ['_xlnm.Print_Area', 0],
      ['Global', null],
    ]);
    expect(wb.getPrintArea(0)).toBe('B!$A$1:$C$3');
    expect(wb.workbookViews[0]?.activeTab).toBe(0);
    expect(wb.workbookViews[0]?.firstSheet).toBe(0);
  });

  it('removing the active sheet clamps activeTab into range', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.workbookViews = [
      {
        activeTab: 1,
        firstSheet: 0,
        showHorizontalScroll: true,
        showVerticalScroll: true,
        showSheetTabs: true,
      },
    ];
    wb.removeSheet(1);
    expect(wb.workbookViews[0]?.activeTab).toBe(0);
  });

  it('moveSheet(2, 0) remaps activeTab and scoped names', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addSheet('C');
    wb.addNamedRange('OnA', 'A!$A$1', 0);
    wb.addNamedRange('OnB', 'B!$A$1', 1);
    wb.addNamedRange('OnC', 'C!$A$1', 2);
    wb.workbookViews = [
      {
        activeTab: 2,
        firstSheet: 0,
        showHorizontalScroll: true,
        showVerticalScroll: true,
        showSheetTabs: true,
      },
    ];
    wb.moveSheet(2, 0);
    expect(wb.sheetNames).toEqual(['C', 'A', 'B']);
    expect(wb.namedRanges.map((d) => [d.name, d.sheetId])).toEqual([
      ['OnA', 1],
      ['OnB', 2],
      ['OnC', 0],
    ]);
    expect(wb.workbookViews[0]?.activeTab).toBe(0);
    expect(wb.workbookViews[0]?.firstSheet).toBe(1);
  });

  it('cloneSheet with an insert index shifts later scoped names', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.addNamedRange('OnB', 'B!$A$1', 1);
    wb.cloneSheet(0, 'A2', 1);
    expect(wb.sheetNames).toEqual(['A', 'A2', 'B']);
    expect(wb.getNamedRange('OnB')?.sheetId).toBe(2);
  });

  it('drops calcChain when sheet positions change', () => {
    const chain = [{ cellRef: 'A1', sheetId: 2 }];
    const seeded = () => {
      const wb = new Workbook({ calcChain: chain });
      wb.addSheet('A');
      wb.addSheet('B');
      return wb;
    };

    const removed = seeded();
    removed.removeSheet(0);
    expect(removed.calcChain).toEqual([]);

    const moved = seeded();
    moved.moveSheet(1, 0);
    expect(moved.calcChain).toEqual([]);

    const sameSpot = seeded();
    sameSpot.moveSheet(1, 1);
    expect(sameSpot.calcChain).toEqual(chain);

    const cloned = seeded();
    cloned.cloneSheet(1, 'B2', 0);
    expect(cloned.calcChain).toEqual([]);

    const appended = seeded();
    appended.cloneSheet(1, 'B2');
    expect(appended.calcChain).toEqual(chain);
  });

  // --- Clone and move checks ---

  it('moveSheet rejects bad indexes', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(() => wb.moveSheet(9, 0)).toThrow('Invalid source index: 9');
    expect(() => wb.moveSheet(0, 9)).toThrow('Invalid target index: 9');
    expect(() => wb.moveSheet(0.5, 1)).toThrow('Invalid source index: 0.5');
    expect(() => wb.moveSheet(0, -1)).toThrow('Invalid target index: -1');
  });

  it('cloneSheet rejects bad indexes and a duplicate name', () => {
    const wb = new Workbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(() => wb.cloneSheet(9, 'X')).toThrow('Invalid source index: 9');
    expect(() => wb.cloneSheet(0, 'X', -1)).toThrow('Invalid insert index: -1');
    expect(() => wb.cloneSheet(0, 'X', 1.5)).toThrow('Invalid insert index: 1.5');
    expect(() => wb.cloneSheet(0, 'X', 3)).toThrow('Invalid insert index: 3');
    expect(() => wb.cloneSheet(0, 'b')).toThrow('already exists');
    expect(wb.sheetNames).toEqual(['A', 'B']);
    wb.cloneSheet(0, 'X', 2);
    expect(wb.sheetNames).toEqual(['A', 'B', 'X']);
  });

  it('cloneSheet gives cloned tables a new id and a unique name', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Src');
    ws.addTable({
      id: 3,
      name: 'Sales',
      displayName: 'Sales',
      ref: 'A1:B3',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: false,
      columns: [
        { id: 1, name: 'Item' },
        { id: 2, name: 'Qty' },
      ],
    });
    wb.addNamedRange('sales_2', 'Src!$A$1');
    const clone = wb.cloneSheet(0, 'Copy');
    const [copied] = clone.tables;
    expect(copied?.id).toBe(4);
    expect(copied?.displayName).toBe('Sales_3');
    expect(copied?.name).toBe('Sales_3');
    expect(ws.tables[0]?.displayName).toBe('Sales');
    expect(ws.tables[0]?.id).toBe(3);
  });

  it('cloneSheet re-ids threaded comments and remaps reply parentId', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Src');
    const rootId = ws.addThreadedComment('A1', 'Check this', 'Alice');
    const replyId = ws.replyToComment(rootId, 'Done', 'Bob');
    const clone = wb.cloneSheet(0, 'Copy');
    const [root, reply] = clone.threadedComments;
    expect(root?.id).not.toBe(rootId);
    expect(reply?.id).not.toBe(replyId);
    expect(reply?.parentId).toBe(root?.id);
    expect(root?.id).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/,
    );
    expect(ws.threadedComments.map((c) => c.id)).toEqual([rootId, replyId]);
  });

  it('cloneSheet keeps the braced upper-case id format Excel writes', () => {
    const rootId = '{5A6E3C1B-0D2F-4E8A-9B7C-1D2E3F4A5B6C}';
    const wb = new Workbook({
      sheets: [
        {
          name: 'Src',
          worksheet: {
            dimension: null,
            rows: [],
            mergeCells: [],
            autoFilter: null,
            frozenPane: null,
            columns: [],
            threadedComments: [
              { id: rootId, refCell: 'A1', personId: 'p', text: 'Q', timestamp: 't' },
              {
                id: '{0F1E2D3C-4B5A-4978-8695-A4B3C2D1E0F9}',
                refCell: 'A1',
                personId: 'p',
                text: 'A',
                timestamp: 't',
                parentId: rootId,
              },
            ],
          },
        },
      ],
    });
    const [root, reply] = wb.cloneSheet(0, 'Copy').threadedComments;
    expect(root?.id).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/,
    );
    expect(root?.id).not.toBe(rootId);
    expect(reply?.parentId).toBe(root?.id);
  });

  it("cloneSheet points legacy 'tc=' comment authors at the new thread ids", () => {
    const rootId = '{5A6E3C1B-0D2F-4E8A-9B7C-1D2E3F4A5B6C}';
    const wb = new Workbook({
      sheets: [
        {
          name: 'Src',
          worksheet: {
            dimension: null,
            rows: [],
            mergeCells: [],
            autoFilter: null,
            frozenPane: null,
            columns: [],
            comments: [
              { cellRef: 'A1', author: `tc=${rootId}`, text: 'Q' },
              { cellRef: 'B1', author: 'Alice', text: 'Plain' },
            ],
            threadedComments: [
              { id: rootId, refCell: 'A1', personId: 'p', text: 'Q', timestamp: 't' },
            ],
          },
        },
      ],
    });
    const clone = wb.cloneSheet(0, 'Copy');
    const [root] = clone.threadedComments;
    expect(root?.id).not.toBe(rootId);
    expect(clone.comments.map((c) => c.author)).toEqual([`tc=${root?.id}`, 'Alice']);
    expect(wb.getSheet('Src')?.comments[0]?.author).toBe(`tc=${rootId}`);
  });
});

describe('sheet names compared ignoring case and canonical equivalence', () => {
  const NFC = 'caf\u{E9}';
  const NFD = 'cafe\u{301}';

  it('addSheet, cloneSheet, renameSheet and the name setter refuse a canonically equal name', () => {
    const wb = new Workbook();
    wb.addSheet(NFC);
    const other = wb.addSheet('Other');
    expect(() => wb.addSheet(NFD)).toThrow('already exists');
    expect(() => wb.addSheet('CAFE\u{301}')).toThrow('already exists');
    expect(() => wb.cloneSheet(1, NFD)).toThrow('already exists');
    expect(() => wb.renameSheet('Other', NFD)).toThrow('already exists');
    expect(() => {
      other.name = NFD;
    }).toThrow('already exists');
    expect(wb.sheetNames).toEqual([NFC, 'Other']);
  });

  it('allows a rename of a sheet to another canonical form of its own name', () => {
    const wb = new Workbook();
    const ws = wb.addSheet(NFC);
    wb.addSheet('Other');
    wb.renameSheet(NFC, NFD);
    expect(wb.sheetNames).toEqual([NFD, 'Other']);
    ws.name = NFC;
    expect(wb.sheetNames).toEqual([NFC, 'Other']);
  });

  /**
   * Names that differ only in case, where composing the lower-case form
   * creates a character with no precomposed capital: `t` + U+0308 composes to
   * U+1E97 but `T` + U+0308 stays as it is. Excel 365 (build 20527,
   * 2026-11-13) refused to open a hand-built file holding each pair as two
   * sheets.
   */
  it.each([
    ['T\u{308}a', 't\u{308}a'],
    ['\u{391}\u{342}a', '\u{3B1}\u{342}a'],
    ['J\u{30C}x', 'j\u{30C}x'],
  ])(
    'addSheet, cloneSheet, renameSheet and the name setter refuse %j beside %j',
    (first, second) => {
      const wb = new Workbook();
      wb.addSheet(first);
      const other = wb.addSheet('Other');
      expect(() => wb.addSheet(second)).toThrow('already exists');
      expect(() => wb.cloneSheet(1, second)).toThrow('already exists');
      expect(() => wb.renameSheet('Other', second)).toThrow('already exists');
      expect(() => {
        other.name = second;
      }).toThrow('already exists');
      expect(wb.sheetNames).toEqual([first, 'Other']);
    },
  );

  // Each pair is equal ignoring case and canonical equivalence (directly or
  // through a third spelling), or differs only in a final sigma (U+03C2)
  // against U+03C3, so the second sheet is refused. Excel 365 (build 20527,
  // 2026-11-13) refuses to open a file holding any of the three sigma pairs
  // as two sheets.
  it.each([
    ['T\u{308}a', 't\u{308}a'],
    ['T\u{308}a', '\u{1E97}a'],
    ['\u{391}\u{342}a', '\u{3B1}\u{342}a'],
    ['\u{391}\u{342}', '\u{1FB6}'],
    ['J\u{30C}x', '\u{1F0}x'],
    ['caf\u{E9}', 'CAFE\u{301}'],
    ['\u{212B}', '\u{E5}'],
    ['W\u{30A}', '\u{1E98}'],
    ['\u{212A}', 'k'],
    ['\u{2126}', '\u{3C9}'],
    ['\u{3C3}', '\u{3C2}'],
    ['\u{391}\u{3A3}', '\u{3B1}\u{3C3}'],
    ['a\u{3A3}b', 'a\u{3C2}b'],
  ])('takes %j and %j as one sheet name', (first, second) => {
    const wb = new Workbook();
    wb.addSheet(first);
    expect(() => wb.addSheet(second)).toThrow('already exists');
  });

  // Each pair differs ignoring case and canonical equivalence: another mark,
  // a lower case that is not the other's (U+0131, U+0130), or a letter with no
  // capital beside the letter it is a variant of. Excel 365 (build 20527,
  // 2026-11-13) opens a file holding U+017F beside s, U+03D0 beside U+03B2 or
  // U+03D1 beside U+03B8 as two sheets.
  it.each([
    ['a', 'b'],
    ['e\u{301}', 'e\u{300}'],
    ['\u{E9}', 'e'],
    ['\u{17F}', 's'],
    ['\u{3D0}', '\u{3B2}'],
    ['\u{3D1}', '\u{3B8}'],
    ['\u{131}', 'i'],
    ['\u{130}', 'I'],
    ['A\u{308}', 'A\u{304}'],
  ])('keeps %j and %j apart', (first, second) => {
    const wb = new Workbook();
    wb.addSheet(first);
    expect(() => wb.addSheet(second)).not.toThrow();
  });
});

/**
 * `nameKey` returns the lower case of a name of ASCII characters at once:
 * the full key (NFC, lower case, final sigma folded, NFC) gives the same
 * string for every such name.
 */
describe('nameKey of an ASCII name', () => {
  const fullKey = (name: string): string =>
    name.normalize('NFC').toLowerCase().replaceAll('\u03C2', '\u03C3').normalize('NFC');

  it('equals the full key for each of the 128 code units, alone and in a name', () => {
    for (let code = 0; code < 128; code++) {
      const ch = String.fromCharCode(code);
      expect(nameKey(ch)).toBe(fullKey(ch));
      expect(nameKey(`A${ch}z`)).toBe(fullKey(`A${ch}z`));
    }
  });

  it('equals the full key for 20,000 generated ASCII names', () => {
    let seed = 12345;
    for (let i = 0; i < 20_000; i++) {
      let name = '';
      const length = 1 + (i % 40);
      for (let j = 0; j < length; j++) {
        seed = (seed * 1103515245 + 12345) % 2147483648;
        name += String.fromCharCode(seed % 128);
      }
      expect(nameKey(name)).toBe(fullKey(name));
    }
  });

  it('still takes the full key once a name holds a character past ASCII', () => {
    expect(nameKey('CAFE\u0301')).toBe('caf\u00E9');
    expect(nameKey('\u039F\u0394\u039F\u03A3')).toBe('\u03BF\u03B4\u03BF\u03C3');
  });
});

/**
 * The name index of the constructor leaves out a name that is not a string
 * (data built by hand, outside the types) instead of calling `toLowerCase`
 * on it.
 */
describe('a sheet name that is not a string', () => {
  it('does not make the constructor throw', () => {
    const worksheet = { rows: [] } as unknown as WorksheetData;
    let wb: Workbook | undefined;
    expect(() => {
      wb = new Workbook({ sheets: [{ name: 2024 as unknown as string, worksheet }] });
    }).not.toThrow();
    expect(wb?.sheetNames).toEqual([2024]);
  });
});

/** A workbook with one empty sheet per name. */
function bookOf(...names: string[]): Workbook {
  const wb = new Workbook();
  for (const name of names) wb.addSheet(name);
  return wb;
}

const namesOf = (wb: Workbook): string[][] => wb.namedRanges.map((d) => [d.name, d.value]);

/**
 * Excel reads a path into another workbook only in a QUOTED colon prefix.
 * Unquoted, `First:Sheet1!$C$10` is the name `First`, the range operator
 * and a sheet-qualified cell. Excel 365 (build 20527, 2026-11-13) on a
 * workbook it saved: its rename of Sheet1 to Data turned
 * `First:Sheet1!$C$10` and `SUM(First:Sheet1!$C$10)` into
 * `First:Data!$C$10` and `SUM(First:Data!$C$10)` as `Name.RefersTo` shows
 * them; the file it saves holds `First:'Data'!$C$10` and
 * `SUM(First:'Data'!$C$10)` (2026-11-14). Its rename of Data to My Data
 * gave `First:'My Data'!$C$10`, and the cell `=SUM(Span_a)` kept its value
 * (165). Over values written by hand, its rename of Sheet1 to Data turned
 * `SUM(Nosuch:Sheet1!$A$1)` into `SUM(Nosuch:Data!$A$1)` (saved as
 * `SUM(Nosuch:'Data'!$A$1)`) and kept `SUM(Sheet1:Nosuch!$A$1)` (shown as
 * `SUM(Sheet1:'Sheet1:[Nosuch]Nosuch'!$A$1)`) and `'Nosuch:Sheet1'!$A$1`
 * (shown as `'Nosuch:[Sheet1]Sheet1'!$A$1`), both read as references into
 * another workbook.
 */
describe('renameSheet over an unquoted colon prefix', () => {
  const values: ReadonlyArray<readonly [string, string]> = [
    ['First', 'Sheet1!$A$1'],
    ['Span_a', 'First:Sheet1!$C$10'],
    ['Span_c', 'SUM(First:Sheet1!$C$10)'],
    ['nosuchLast', 'SUM(Nosuch:Sheet1!$A$1)'],
    ['nosuchFirst', 'SUM(Sheet1:Nosuch!$A$1)'],
    ['quoted', "'Nosuch:Sheet1'!$A$1"],
  ];
  const kept = values.slice(4);

  it('follows the sheet when the last end is the renamed sheet, twice and after a write and read', async () => {
    const wb = bookOf('Sheet1', 'Sheet3');
    for (const [name, value] of values) wb.addNamedRange(name, value);
    wb.renameSheet('Sheet1', 'Data');
    expect(namesOf(wb)).toEqual([
      ['First', 'Data!$A$1'],
      ['Span_a', "First:'Data'!$C$10"],
      ['Span_c', "SUM(First:'Data'!$C$10)"],
      ['nosuchLast', "SUM(Nosuch:'Data'!$A$1)"],
      ...kept,
    ]);
    const back = await readBuffer(await wb.toBuffer());
    back.renameSheet('Data', 'My Final');
    expect(namesOf(back)).toEqual([
      ['First', "'My Final'!$A$1"],
      ['Span_a', "First:'My Final'!$C$10"],
      ['Span_c', "SUM(First:'My Final'!$C$10)"],
      ['nosuchLast', "SUM(Nosuch:'My Final'!$A$1)"],
      ...kept,
    ]);
    back.renameSheet('My Final', 'Data');
    expect(back.getNamedRange('Span_a')?.value).toBe("First:'Data'!$C$10");
  });

  it("rewrites the form Excel stores, First:'Sheet1'!$C$10", () => {
    const wb = bookOf('Sheet1');
    wb.addNamedRange('First', 'Sheet1!$A$1');
    wb.addNamedRange('Span_a', "First:'Sheet1'!$C$10");
    wb.renameSheet('Sheet1', 'My Data');
    expect(wb.getNamedRange('Span_a')?.value).toBe("First:'My Data'!$C$10");
  });
});

/**
 * Excel 365 (build 20527, 2026-11-14) stores a sheet-qualified reference
 * right after the range operator quoted: names added by `Names.Add` over
 * sheets Sheet1, Sheet2, Sheet3 and Other are saved as `SRC` below, and the
 * file it saves after each of its own renames holds exactly the texts of
 * `EXCEL_STORED`. Unquoted, `First:First!$C$10` would be the one-sheet span
 * First:First, so the defined name First would no longer take part.
 */
describe('renameSheet stores the text Excel 365 stores after the range operator', () => {
  const SRC: ReadonlyArray<readonly [string, string]> = [
    ['First', 'Sheet1!$A$1'],
    ['pp_12', "Other!$A$1:'Sheet1'!$B$2"],
    ['pp_13', "Sheet1!$A$1:'Sheet1'!$B$2"],
    ['pp_18', "SUM(Sheet1!$A$1,First:'Sheet1'!$B$2,Sheet3!$A$1)"],
    ['pp_23', "First:'Sheet3'!$C$10"],
    ['pp_6', "First:'Sheet1'!$C$10"],
    ['pp_7', "SUM(First:'Sheet1'!$C$10)"],
    ['pp_9', "SUM(Nosuch:'Sheet1'!$A$1)"],
    ['pp_span', 'SUM(Sheet1:Sheet3!$A$1)'],
  ];
  const EXCEL_STORED: Readonly<Record<string, ReadonlyArray<readonly [string, string]>>> = {
    'Sheet1 to First': [
      ['First', 'First!$A$1'],
      ['pp_12', "Other!$A$1:'First'!$B$2"],
      ['pp_13', "First!$A$1:'First'!$B$2"],
      ['pp_18', "SUM(First!$A$1,First:'First'!$B$2,Sheet3!$A$1)"],
      ['pp_23', "First:'Sheet3'!$C$10"],
      ['pp_6', "First:'First'!$C$10"],
      ['pp_7', "SUM(First:'First'!$C$10)"],
      ['pp_9', "SUM(Nosuch:'First'!$A$1)"],
      ['pp_span', 'SUM(First:Sheet3!$A$1)'],
    ],
    'Sheet3 to First': [
      ['First', 'Sheet1!$A$1'],
      ['pp_12', "Other!$A$1:'Sheet1'!$B$2"],
      ['pp_13', "Sheet1!$A$1:'Sheet1'!$B$2"],
      ['pp_18', "SUM(Sheet1!$A$1,First:'Sheet1'!$B$2,First!$A$1)"],
      ['pp_23', "First:'First'!$C$10"],
      ['pp_6', "First:'Sheet1'!$C$10"],
      ['pp_7', "SUM(First:'Sheet1'!$C$10)"],
      ['pp_9', "SUM(Nosuch:'Sheet1'!$A$1)"],
      ['pp_span', 'SUM(Sheet1:First!$A$1)'],
    ],
    'Sheet1 to Data': [
      ['First', 'Data!$A$1'],
      ['pp_12', "Other!$A$1:'Data'!$B$2"],
      ['pp_13', "Data!$A$1:'Data'!$B$2"],
      ['pp_18', "SUM(Data!$A$1,First:'Data'!$B$2,Sheet3!$A$1)"],
      ['pp_23', "First:'Sheet3'!$C$10"],
      ['pp_6', "First:'Data'!$C$10"],
      ['pp_7', "SUM(First:'Data'!$C$10)"],
      ['pp_9', "SUM(Nosuch:'Data'!$A$1)"],
      ['pp_span', 'SUM(Data:Sheet3!$A$1)'],
    ],
    'Sheet1 to B, a new sheet First, First to Q': [
      ['First', 'B!$A$1'],
      ['pp_12', "Other!$A$1:'B'!$B$2"],
      ['pp_13', "B!$A$1:'B'!$B$2"],
      ['pp_18', "SUM(B!$A$1,First:'B'!$B$2,Sheet3!$A$1)"],
      ['pp_23', "First:'Sheet3'!$C$10"],
      ['pp_6', "First:'B'!$C$10"],
      ['pp_7', "SUM(First:'B'!$C$10)"],
      ['pp_9', "SUM(Nosuch:'B'!$A$1)"],
      ['pp_span', 'SUM(B:Sheet3!$A$1)'],
    ],
    'Sheet1 to First, First to Z': [
      ['First', 'Z!$A$1'],
      ['pp_12', "Other!$A$1:'Z'!$B$2"],
      ['pp_13', "Z!$A$1:'Z'!$B$2"],
      ['pp_18', "SUM(Z!$A$1,First:'Z'!$B$2,Sheet3!$A$1)"],
      ['pp_23', "First:'Sheet3'!$C$10"],
      ['pp_6', "First:'Z'!$C$10"],
      ['pp_7', "SUM(First:'Z'!$C$10)"],
      ['pp_9', "SUM(Nosuch:'Z'!$A$1)"],
      ['pp_span', 'SUM(Z:Sheet3!$A$1)'],
    ],
  };
  const STEPS: Readonly<Record<string, (wb: Workbook) => void>> = {
    'Sheet1 to First': (wb) => wb.renameSheet('Sheet1', 'First'),
    'Sheet3 to First': (wb) => wb.renameSheet('Sheet3', 'First'),
    'Sheet1 to Data': (wb) => wb.renameSheet('Sheet1', 'Data'),
    'Sheet1 to B, a new sheet First, First to Q': (wb) => {
      wb.renameSheet('Sheet1', 'B');
      wb.addSheet('First');
      wb.renameSheet('First', 'Q');
    },
    'Sheet1 to First, First to Z': (wb) => {
      wb.renameSheet('Sheet1', 'First');
      wb.renameSheet('First', 'Z');
    },
  };
  const source = (values: ReadonlyArray<readonly [string, string]>): Workbook => {
    const wb = bookOf('Sheet1', 'Sheet2', 'Sheet3', 'Other');
    for (const [name, value] of values) wb.addNamedRange(name, value);
    return wb;
  };

  it.each(Object.keys(STEPS))('%s gives the texts Excel saves', (step) => {
    const wb = source(SRC);
    STEPS[step]?.(wb);
    expect(namesOf(wb)).toEqual(EXCEL_STORED[step]);
  });

  it('quotes the last end of the unquoted form written by hand, and keeps a span from forming', () => {
    const byHand: ReadonlyArray<readonly [string, string]> = [
      ['First', 'Sheet1!$A$1'],
      ['dn', 'First:Sheet1!$C$10'],
      ['sum', 'SUM(Sheet1!$A$1,First:Sheet1!$B$2,Sheet3!$A$1)'],
      ['other', 'First:Sheet3!$C$10'],
    ];
    const toFirst = source(byHand);
    toFirst.renameSheet('Sheet1', 'First');
    expect(namesOf(toFirst)).toEqual([
      ['First', 'First!$A$1'],
      ['dn', "First:'First'!$C$10"],
      ['sum', "SUM(First!$A$1,First:'First'!$B$2,Sheet3!$A$1)"],
      ['other', "First:'Sheet3'!$C$10"],
    ]);
    const threeToFirst = source(byHand);
    threeToFirst.renameSheet('Sheet3', 'First');
    expect(namesOf(threeToFirst)).toEqual([
      ['First', 'Sheet1!$A$1'],
      ['dn', "First:'Sheet1'!$C$10"],
      ['sum', "SUM(Sheet1!$A$1,First:'Sheet1'!$B$2,First!$A$1)"],
      ['other', "First:'First'!$C$10"],
    ]);
  });

  it('keeps the quoted text through a write, a read and a later sheet named First', async () => {
    const wb = source([
      ['First', 'Sheet1!$A$1'],
      ['dn', 'First:Sheet1!$C$10'],
    ]);
    wb.renameSheet('Sheet1', 'B');
    expect(wb.getNamedRange('dn')?.value).toBe("First:'B'!$C$10");
    const back = await readBuffer(await wb.toBuffer());
    back.addSheet('First');
    back.renameSheet('First', 'Q');
    expect(back.getNamedRange('dn')?.value).toBe("First:'B'!$C$10");
    back.renameSheet('B', "It's");
    expect(back.getNamedRange('dn')?.value).toBe("First:'It''s'!$C$10");
  });
});

/**
 * U+0130 (capital I with dot above) lower-cases to `i` + U+0307, and Excel
 * 365 (build 20527, 2026-11-13) does not take that spelling for the sheet:
 * with sheets U+0130 `a` and Other, its rename of the first sheet to X kept
 * `'i` U+0307 `a'!$A$1` (stored as the external `[1]i` U+0307 `a!$A$1`),
 * and kept `ia` beside U+0130 `a`, `IA` beside U+0131 `a`, `STRASSE`
 * beside `Straße`, `CAFE` + U+0301 beside `Café` and `FIX` beside U+FB01
 * `x`, while it followed `οδος` (final sigma) for ΟΔΟΣ, `ka` for U+212A
 * `a`, U+212A `A` for `ka` and `ßa` for U+1E9E `a`.
 */
describe('renameSheet compares a prefix with the sheet name as Excel 365 does', () => {
  const probes: ReadonlyArray<readonly [string, string, string, boolean]> = [
    ['dotted I spelled i + U+0307', 'İa', 'i̇a', false],
    ['dotted I spelled U+0130', 'İa', 'İa', true],
    ['dotted I spelled i', 'İa', 'ia', false],
    ['dotless i', 'ıa', 'IA', false],
    ['final sigma', 'ΟΔΟΣ', 'οδος', true],
    ['Kelvin sign sheet', 'Ka', 'ka', true],
    ['Kelvin sign prefix', 'ka', 'KA', true],
    ['sharp s', 'Straße', 'STRASSE', false],
    ['capital sharp s', 'ẞa', 'ßa', true],
    ['normal form', 'Café', 'CAFÉ', false],
    ['ligature', 'ﬁx', 'FIX', false],
  ];

  it.each(probes)('%s', (_, sheet, prefix, follows) => {
    const wb = bookOf(sheet, 'Other');
    wb.addNamedRange('dn', `'${prefix}'!$A$1`);
    wb.renameSheet(0, 'X');
    expect(wb.getNamedRange('dn')?.value).toBe(follows ? 'X!$A$1' : `'${prefix}'!$A$1`);
  });
});

/**
 * Data built by hand, outside the types: a sheet whose name is a number
 * neither stops a rename nor is taken for a sheet, and a rename that throws
 * leaves the workbook as it was.
 */
describe('renameSheet beside a sheet name that is not a string', () => {
  const worksheet = { rows: [] } as unknown as WorksheetData;
  const handBuilt = (): Workbook =>
    new Workbook({
      sheets: [
        { name: 'Sheet1', worksheet },
        { name: 2024 as unknown as string, worksheet },
        { name: 'Sheet3', worksheet },
      ],
      definedNames: [
        { name: 'a', value: 'Sheet1!$A$1', sheetId: null },
        { name: 'b', value: "SUM('Sheet1:Sheet3'!$A$1)", sheetId: null },
        { name: 'c', value: "SUM('Sheet1:2024'!$A$1)", sheetId: null },
      ],
    });

  /** Runs the rename and returns whether it completed; when it throws, nothing may have changed. */
  const renameOrNothing = (wb: Workbook, from: number, to: string): boolean => {
    const before = JSON.stringify(wb.toJSON());
    try {
      wb.renameSheet(from, to);
      return true;
    } catch {
      expect(JSON.stringify(wb.toJSON())).toBe(before);
      return false;
    }
  };

  it('renames another sheet and its references', () => {
    const wb = handBuilt();
    expect(renameOrNothing(wb, 0, 'Data')).toBe(true);
    expect(wb.sheetNames).toEqual(['Data', 2024, 'Sheet3']);
    expect(namesOf(wb)).toEqual([
      ['a', 'Data!$A$1'],
      ['b', 'SUM(Data:Sheet3!$A$1)'],
      ['c', "SUM('Sheet1:2024'!$A$1)"],
    ]);
  });

  it('either renames the sheet named by a number or changes nothing', () => {
    const wb = handBuilt();
    if (renameOrNothing(wb, 1, 'Mid')) expect(wb.sheetNames).toEqual(['Sheet1', 'Mid', 'Sheet3']);
    expect(wb.getNamedRange('a')?.value).toBe('Sheet1!$A$1');
  });
});

/**
 * A sheet whose own stored name is not a string (data built by hand,
 * outside the types) is renamed by index, as 1.2.1 renamed it; no sheet
 * prefix can name it, so the defined names are kept as they are.
 */
describe('renameSheet of a sheet whose own name is not a string', () => {
  it('renames it in a workbook without defined names', () => {
    const wb = bookOf('Sheet1', 'Sheet2');
    (wb.toJSON().sheets[0] as { name: unknown }).name = 5;
    wb.renameSheet(0, 'Fixed');
    expect(wb.sheetNames).toEqual(['Fixed', 'Sheet2']);
    expect(wb.getSheet('Fixed')).toBeDefined();
  });

  it('renames it and keeps the defined names', () => {
    const wb = bookOf('Sheet1', 'Sheet2');
    wb.addNamedRange('a', 'Sheet2!$A$1');
    wb.addNamedRange('b', "SUM('5:Sheet2'!$A$1)");
    (wb.toJSON().sheets[0] as { name: unknown }).name = 5;
    wb.renameSheet(0, 'Fixed');
    expect(wb.sheetNames).toEqual(['Fixed', 'Sheet2']);
    expect(namesOf(wb)).toEqual([
      ['a', 'Sheet2!$A$1'],
      ['b', "SUM('5:Sheet2'!$A$1)"],
    ]);
  });
});

/**
 * A name held as a String object (data built by hand, outside the types)
 * is compared by its string value, as a name held as a string is.
 */
describe('a name held as a String object', () => {
  it('is seen by the sheet-name check', () => {
    const worksheet = { rows: [] } as unknown as WorksheetData;
    const name = Object('Data') as string;
    const wb = new Workbook({ sheets: [{ name, worksheet }] });
    expect(() => wb.addSheet('DATA')).toThrow('already exists');
    expect(wb.sheetNames).toHaveLength(1);
  });

  it('is seen by the table-name check', () => {
    const wb = bookOf('Data');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('missing sheet');
    const table = (id: number, displayName: string, ref: string) => ({
      id,
      displayName,
      ref,
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: false,
      columns: [{ id: 1, name: 'h' }],
    });
    ws.addTable(table(1, 'Tx', 'A1:A3'));
    const stored = ws.tables[0];
    if (!stored) throw new Error('missing table');
    stored.displayName = Object('Tx') as string;
    expect(() => ws.addTable(table(2, 'TX', 'C1:C3'))).toThrow('already used by table');
    expect(ws.tables).toHaveLength(1);
  });
});
