import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

function workbookWith(...names: string[]): Workbook {
  const wb = new Workbook();
  for (const name of names) wb.addSheet(name);
  return wb;
}

describe('addNamedRange validation', () => {
  it.each([
    'A1',
    'xfd1048576',
    'R1C1',
    'R',
    'c',
    'R12',
    'C3',
    '1abc',
    'a b',
    'a!b',
    '',
    '.dot',
    'x'.repeat(256),
  ])('rejects %j', (name) => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange(name, 'Data!$A$1')).toThrow(/defined name/i);
    expect(wb.namedRanges).toEqual([]);
  });

  it.each(['Revenue', '_hidden', '\\back', 'Tax.Rate_2', 'ABCD1', 'Umsätze', 'x'.repeat(255)])(
    'accepts %j',
    (name) => {
      const wb = workbookWith('Data');
      wb.addNamedRange(name, 'Data!$A$1');
      expect(wb.getNamedRange(name)?.value).toBe('Data!$A$1');
    },
  );

  it('rejects a case-insensitive duplicate in the same scope', () => {
    const wb = workbookWith('Data', 'Other');
    wb.addNamedRange('X', 'Data!$A$1');
    expect(() => wb.addNamedRange('x', 'Data!$B$1')).toThrow('already exists');
    wb.addNamedRange('x', 'Data!$B$1', 0);
    wb.addNamedRange('X', 'Other!$B$1', 1);
    expect(() => wb.addNamedRange('X', 'Data!$C$1', 0)).toThrow('already exists');
    expect(wb.namedRanges.map((d) => [d.name, d.sheetId])).toEqual([
      ['X', null],
      ['x', 0],
      ['X', 1],
    ]);
  });

  it.each([5, -1, 1.5, 2])('rejects sheetId %d outside the sheet range', (sheetId) => {
    const wb = workbookWith('Data', 'Other');
    expect(() => wb.addNamedRange('Local', 'Data!$A$1', sheetId)).toThrow('sheet index');
  });

  it("allows only the supported '_xlnm.' built-in names", () => {
    const wb = workbookWith('Data');
    wb.addNamedRange('_xlnm.Print_Area', 'Data!$A$1:$B$2', 0);
    wb.addNamedRange('_xlnm.Print_Titles', 'Data!$1:$1', 0);
    wb.addNamedRange('_xlnm._FilterDatabase', 'Data!$A$1:$B$2', 0);
    expect(() => wb.addNamedRange('_xlnm.Foo', 'Data!$A$1')).toThrow('_xlnm.');
    expect(wb.getPrintArea(0)).toBe('Data!$A$1:$B$2');
  });

  it('leaves setPrintArea and setPrintTitles unchanged', () => {
    const wb = workbookWith('Data');
    wb.setPrintArea(0, 'Data!$A$1:$B$2');
    wb.setPrintArea(0, 'Data!$A$1:$C$3');
    wb.setPrintTitles(0, 'Data!$1:$1');
    expect(wb.namedRanges).toHaveLength(2);
    expect(wb.getPrintArea(0)).toBe('Data!$A$1:$C$3');
  });
});

describe('renameSheet rewrites defined names', () => {
  it("renaming Data to Q1 gives 'Q1'!A1", () => {
    const wb = workbookWith('Data');
    wb.addNamedRange('Quoted', "'Data'!A1");
    wb.addNamedRange('Bare', 'Data!A1');
    wb.renameSheet('Data', 'Q1');
    expect(wb.getNamedRange('Quoted')?.value).toBe("'Q1'!A1");
    expect(wb.getNamedRange('Bare')?.value).toBe("'Q1'!A1");
  });

  it("renaming Data to 'My Data' quotes the prefix", () => {
    const wb = workbookWith('Data');
    wb.addNamedRange('Rev', 'Data!$A$1');
    wb.addNamedRange('Span', 'SUM(Data!$A$1:$A$5,data!B:B)');
    wb.renameSheet('Data', 'My Data');
    expect(wb.getNamedRange('Rev')?.value).toBe("'My Data'!$A$1");
    expect(wb.getNamedRange('Span')?.value).toBe("SUM('My Data'!$A$1:$A$5,'My Data'!B:B)");
  });

  it('escapes an apostrophe in the new name', () => {
    const wb = workbookWith('Data');
    wb.addNamedRange('Rev', 'Data!$A$1');
    wb.renameSheet(0, "O'Brien");
    expect(wb.getNamedRange('Rev')?.value).toBe("'O''Brien'!$A$1");
  });

  it('matches an escaped quoted old name', () => {
    const wb = workbookWith("O'Brien");
    wb.addNamedRange('Rev', "'O''Brien'!$A$1");
    wb.renameSheet(0, 'Plain');
    expect(wb.getNamedRange('Rev')?.value).toBe('Plain!$A$1');
  });

  it("leaves MyData!A1 and 'Data2'!A1 unchanged", () => {
    const wb = workbookWith('Data', 'MyData', 'Data2');
    wb.addNamedRange('Mine', 'MyData!A1');
    wb.addNamedRange('Two', "'Data2'!A1");
    wb.addNamedRange('Bare', 'A1+Data');
    wb.renameSheet('Data', 'Q1');
    expect(wb.getNamedRange('Mine')?.value).toBe('MyData!A1');
    expect(wb.getNamedRange('Two')?.value).toBe("'Data2'!A1");
    expect(wb.getNamedRange('Bare')?.value).toBe('A1+Data');
  });

  it('rewrites an end of a 3-D span and leaves values with tokenizer errors unchanged', () => {
    const wb = workbookWith('Data', 'Data2', 'Data3');
    wb.addNamedRange('ThreeD', 'SUM(Data:Data3!A1)');
    wb.addNamedRange('Broken', 'Data!A1+"open');
    wb.renameSheet('Data', 'First');
    expect(wb.getNamedRange('ThreeD')?.value).toBe('SUM(First:Data3!A1)');
    expect(wb.getNamedRange('Broken')?.value).toBe('Data!A1+"open');
    wb.renameSheet('Data2', 'Middle');
    expect(wb.getNamedRange('ThreeD')?.value).toBe('SUM(First:Data3!A1)');
  });

  it('quotes a 3-D span as one unit when either end needs it', () => {
    const wb = workbookWith('My Sheet', 'Sheet2', 'Sheet3');
    wb.addNamedRange('Span', "'My Sheet:Sheet3'!$A$1");
    wb.addNamedRange('Span2', 'Sheet2:Sheet3!$A$1');
    wb.renameSheet('My Sheet', 'Plain');
    expect(wb.getNamedRange('Span')?.value).toBe('Plain:Sheet3!$A$1');
    wb.renameSheet('Sheet3', 'Q 3');
    expect(wb.getNamedRange('Span')?.value).toBe("'Plain:Q 3'!$A$1");
    expect(wb.getNamedRange('Span2')?.value).toBe("'Sheet2:Q 3'!$A$1");
    wb.renameSheet('Plain', "O'Brien");
    expect(wb.getNamedRange('Span')?.value).toBe("'O''Brien:Q 3'!$A$1");
  });

  it('rewrites Print_Titles and Print_Area', () => {
    const wb = workbookWith('Data');
    wb.setPrintTitles(0, 'Data!$1:$1,Data!$A:$B');
    wb.setPrintArea(0, 'Data!$A$1:$C$10');
    wb.renameSheet('Data', 'Report 2026');
    expect(wb.getPrintTitles(0)).toBe("'Report 2026'!$1:$1,'Report 2026'!$A:$B");
    expect(wb.getPrintArea(0)).toBe("'Report 2026'!$A$1:$C$10");
  });

  it("rewrites a sheet-qualified error literal: Data!#REF! becomes 'Q1'!#REF!", () => {
    const wb = workbookWith('Data');
    wb.addNamedRange('Err', 'Data!$A$1+Data!#REF!');
    wb.addNamedRange('Plain', '#REF!');
    wb.renameSheet('Data', 'Q1');
    expect(wb.getNamedRange('Err')?.value).toBe("'Q1'!$A$1+'Q1'!#REF!");
    expect(wb.getNamedRange('Plain')?.value).toBe('#REF!');
  });

  it('renames a sheet with stored _xlnm.Criteria and case-duplicate names', () => {
    const wb = new Workbook({
      sheets: [
        {
          name: 'Data',
          worksheet: {
            dimension: null,
            rows: [],
            mergeCells: [],
            autoFilter: null,
            frozenPane: null,
            columns: [],
          },
        },
      ],
      definedNames: [
        { name: '_xlnm.Criteria', value: 'Data!$E$1:$E$2', sheetId: 0 },
        { name: '_xlnm.Extract', value: 'Data!$G$1', sheetId: 0 },
        { name: 'Rate', value: 'Data!$A$1', sheetId: null },
        { name: 'RATE', value: 'Data!$B$1', sheetId: null },
      ],
    });
    wb.renameSheet('Data', 'Sales');
    expect(wb.namedRanges.map((d) => d.value)).toEqual([
      'Sales!$E$1:$E$2',
      'Sales!$G$1',
      'Sales!$A$1',
      'Sales!$B$1',
    ]);
    const ws = wb.getSheet('Sales');
    if (!ws) throw new Error('missing sheet');
    ws.name = 'Data';
    expect(wb.namedRanges[0]?.value).toBe('Data!$E$1:$E$2');
    expect(() => wb.addNamedRange('_xlnm.Criteria', 'Data!$A$1', 0)).toThrow('_xlnm.');
  });

  it('does not rewrite cell formulas', () => {
    const wb = workbookWith('Data', 'Other');
    const other = wb.getSheet('Other');
    if (!other) throw new Error('missing sheet');
    other.cell('A1').formula = 'Data!A1+1';
    wb.renameSheet('Data', 'Q1');
    expect(other.cell('A1').formula).toBe('Data!A1+1');
  });
});

describe('defined-name rule order', () => {
  it('reports the shape rule before the scope rule', () => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange('A1', 'Data!$A$1', 99)).toThrow(/cell reference/);
  });
  it('reports the reserved-prefix rule before the scope rule', () => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange('_xlnm.Bogus', 'Data!$A$1', 99)).toThrow(/_xlnm\./);
  });
  it('reports the first failing shape rule (start character before later characters)', () => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange('1 a', 'Data!$A$1')).toThrow(/must start with/);
  });
  it('reports the shape rule before the reserved-prefix rule', () => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange('_xlnm.Foo bar', 'Data!$A$1')).toThrow(/may contain only/);
  });
  it('reports the length rule before the start-character rule', () => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange(`1${'x'.repeat(255)}`, 'Data!$A$1')).toThrow(/1-255 characters/);
  });
});

/*
 * Defined names as Excel 365 (Windows, build 20527) proved them on 2026-11-12,
 * in one Excel instance per run: `Names.Add` through COM for the API (it
 * refuses what Excel's own name box refuses), and one workbook per name,
 * written by modern-xlsx with the name stored as given and opened with alerts
 * off, for validate() (it reports what Excel refuses to open). Excel opens a
 * file whose defined name is an A1 cell reference (`A1`, `R1`, `XFD1`,
 * `C16385`) and keeps the name, though `Names.Add` refuses it; it refuses to
 * open one that is an R1C1 reference or starts with one (`R1x`, `C1_total`,
 * `R2D2`). The API keeps refusing the cell references.
 *
 * Where the API and `Names.Add` still differ: `Names.Add` refuses `\a` and
 * `\1` (the files open, and the API takes both), and takes `a€b`, `a·b` and
 * `a°b` (the files open too; the API takes letters, then also combining
 * marks, digits, `_`, `.`, `\` and `?`).
 */

/** Rule fragments of the messages for a name shaped like a reference. */
const DN_BOOL = 'TRUE and FALSE';
const DN_A1 = 'cell reference';
const DN_R1C1 = 'R1C1';

/** [name, rule fragment] for each name `Names.Add` refused; a `/defined name/` fragment for the character rules. */
const DN_API_REFUSED: [string, string | RegExp][] = [
  ...['R', 'C', 'r', 'c', 'RC', 'R1C1', 'R1x', 'C1_', 'C1_total', 'RC2x', 'R2D2', 'C3PO'].map(
    (name): [string, string] => [name, DN_R1C1],
  ),
  ...['R12abc', 'RC1data', 'R01x', 'C001_', 'c16384x', 'r1048576x', 'R1048576x', 'R1C'].map(
    (name): [string, string] => [name, DN_R1C1],
  ),
  // R1, C1, C16385 and RC16385 are cell references (column R, C or RC).
  ...['R1', 'C1', 'C16385', 'RC16385', 'A1', 'A01', 'XFD1', 'T2'].map((name): [string, string] => [
    name,
    DN_A1,
  ]),
  ['TRUE', DN_BOOL],
  ['FALSE', DN_BOOL],
  ['true', DN_BOOL],
  ['1a', /must start with/],
  ['?a', /must start with/],
  ['.a', /must start with/],
  ['a b', /may contain only/],
  ['a-b', /may contain only/],
  ['x'.repeat(256), /1-255 characters/],
];

/** Names `Names.Add` accepted; the API takes each. */
const DN_API_ACCEPTED: string[] = [
  'R0x',
  'R1048577',
  'C16385x',
  'A0',
  'XFE1',
  'AAAA1',
  'TRUEx',
  'R1.a',
  'a?b',
  'a\\b',
  '_a',
  'a.b',
  'x'.repeat(255),
  'Rx',
  'Cx',
  'R_1',
  'C5.x',
  'rc5.x',
  'Año2024',
  '表1',
  'Таблица1',
  'RC0x',
  'C0x',
  'r1048576.x',
  'Rate1',
  '\\abc',
  '\\',
];

/** [name, rule fragment] for each defined name Excel refused to open in a file. */
const DN_FILE_REFUSED: [string, string][] = [
  ...[
    'R',
    'C',
    'r',
    'c',
    'RC',
    'R1C1',
    'R1x',
    'C1_',
    'C1_total',
    'RC2x',
    'R2D2',
    'C3PO',
    'R12abc',
    'RC1data',
    'R01x',
    'C001_',
    'c16384x',
    'r1048576x',
    'R1048576x',
    'R1C',
  ].map((name): [string, string] => [name, DN_R1C1]),
  ['TRUE', DN_BOOL],
  ['FALSE', DN_BOOL],
  ['true', DN_BOOL],
  ['False', DN_BOOL],
  ['1a', 'starts with'],
  ['?a', 'starts with'],
  ['.a', 'starts with'],
  // Every ASCII punctuation character but `_ . \ ? :` (a `:` opened in `a:b`), and the whitespace controls.
  ...[...' !"#$%&\'()*+,-/;<=>@[]^`{|}~\t\n\r\u007f'].map((ch): [string, string] => [
    `a${ch}b`,
    'character Excel refuses',
  ]),
  ['x'.repeat(256), 'longer than 255'],
];

/** Defined names Excel opened in a file: validate() reports nothing for them. */
const DN_FILE_OPENED: string[] = [
  'R1',
  'C1',
  'A1',
  'A01',
  'XFD1',
  'C16385',
  'RC16385',
  'T2',
  'R0x',
  'R1048577',
  'C16385x',
  'A0',
  'XFE1',
  'AAAA1',
  'TRUEx',
  'TRUE.x',
  'R1.a',
  'C1.x',
  'a?b',
  'a\\b',
  '_a',
  '\\a',
  'a.b',
  'x'.repeat(255),
  'Rx',
  'Cx',
  'R_1',
  'C5.x',
  'rc5.x',
  'a€b',
  'Año2024',
  '表1',
  'Таблица1',
  'a·b',
  'a°b',
  'RC0x',
  'C0x',
  'r1048576.x',
  'Rate1',
  'a:b',
  '\\1',
  '\\R1',
  '\\',
  '_',
  'a\u0001b',
];

/** A workbook holding one sheet `S` and the defined name `name` as a read file would. */
function rawWorkbookWithName(name: string): Workbook {
  return new Workbook({
    sheets: [
      {
        name: 'S',
        worksheet: {
          dimension: null,
          rows: [],
          mergeCells: [],
          autoFilter: null,
          frozenPane: null,
          columns: [],
        },
      },
    ],
    definedNames: [{ name, value: 'S!$A$1', sheetId: null }],
  });
}

describe('defined names Excel refuses', () => {
  it.each(DN_API_REFUSED)('addNamedRange refuses %j (%s)', (name, rule) => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange(name, 'Data!$A$1')).toThrow(rule);
    expect(() => wb.addNamedRange(name, 'Data!$A$1')).toThrow(/defined name/i);
    expect(wb.namedRanges).toEqual([]);
  });

  it.each(DN_FILE_REFUSED)('validate() reports %j (%s) as a defined-name error', (name, rule) => {
    const issues = rawWorkbookWithName(name)
      .validate()
      .issues.filter((issue) => issue.severity === 'error');
    expect(issues).toHaveLength(1);
    expect(issues[0]?.category).toBe('definedName');
    expect(issues[0]?.location).toBe('definedNames[0]');
    expect(issues[0]?.message).toContain(rule);
  });
});

describe('defined names Excel accepts', () => {
  it.each(DN_API_ACCEPTED)('addNamedRange takes %j and validate() reports nothing', (name) => {
    const wb = workbookWith('Data');
    wb.addNamedRange(name, 'Data!$A$1');
    expect(wb.getNamedRange(name)?.value).toBe('Data!$A$1');
    expect(wb.validate().errorCount).toBe(0);
  });

  it.each(DN_FILE_OPENED)('validate() takes %j in a raw model', (name) => {
    expect(rawWorkbookWithName(name).validate().errorCount).toBe(0);
  });

  it('validate() skips the _xlnm. built-in names', () => {
    const wb = rawWorkbookWithName('_xlnm._FilterDatabase');
    expect(wb.validate().errorCount).toBe(0);
  });
});

/*
 * Combining marks, as Excel 365 (build 20527) proved them on 2026-11-12:
 * every accepted name below opens with the name kept in a file modern-xlsx
 * wrote. `Names.Add` takes each one except `\` + U+0301 (as with `\a`), and
 * Excel's own save drops a lone U+FE0F. Each refused name was refused by
 * `Names.Add` (a leading mark, or a space), and the files holding a leading U+0301 or U+093F did not open. Excel also
 * takes a leading Thai mark (U+0E34), symbols such as `€`, `·`, `°`, `±`,
 * `²` and `Ⅻ`, emoji, and U+200C/U+200D inside Persian and Devanagari words,
 * which the API still refuses.
 */

/** [script, a name Excel accepts, a name Excel refuses, rule fragment of the refusal]. */
const DN_MARKS: [string, string, string, RegExp][] = [
  [
    'Thai',
    '\u{E23}\u{E32}\u{E04}\u{E32}\u{E2A}\u{E34}\u{E19}\u{E04}\u{E49}\u{E32}',
    '\u{E23}\u{E32}\u{E04}\u{E32} \u{E2A}\u{E34}\u{E19}\u{E04}\u{E49}\u{E32}',
    /may contain only/,
  ],
  ['Hindi', '\u{92E}\u{942}\u{932}\u{94D}\u{92F}', '\u{93F}\u{915}', /must start with/],
  ['Hindi (spacing mark)', '\u{915}\u{940}\u{92E}\u{924}', '\u{940}\u{915}', /must start with/],
  ['Bengali', '\u{9AE}\u{9C2}\u{9B2}\u{9CD}\u{9AF}', '\u{9BE}\u{9A6}', /must start with/],
  ['Tamil', '\u{BB5}\u{BBF}\u{BB2}\u{BC8}', '\u{BBF}\u{BB5}', /must start with/],
  [
    'Arabic',
    '\u{645}\u{64E}\u{628}\u{652}\u{644}\u{64E}\u{63A}',
    '\u{64E}\u{645}',
    /must start with/,
  ],
  ['Hebrew', '\u{5DE}\u{5B0}\u{5D7}\u{5B4}\u{5D9}\u{5E8}', '\u{5B0}\u{5DE}', /must start with/],
  ['Latin', 'cafe\u{301}', '\u{301}cafe', /must start with/],
  ['Vietnamese', 'Tie\u{302}\u{300}n', '\u{300}Tien', /must start with/],
  ['Korean', '\u{1100}\u{1161}\u{1100}\u{1167}\u{11A8}', '\u{AC00} \u{AC69}', /may contain only/],
  ['enclosing mark', 'a\u{20DD}b', '\u{20DD}ab', /must start with/],
];

describe('defined names with combining marks', () => {
  it.each(DN_MARKS)('%s: addNamedRange takes %j', (_script, accepted) => {
    const wb = workbookWith('Data');
    wb.addNamedRange(accepted, 'Data!$A$1');
    expect(wb.getNamedRange(accepted)?.value).toBe('Data!$A$1');
    expect(wb.validate().errorCount).toBe(0);
  });

  it.each(DN_MARKS)('%s: addNamedRange refuses %3$j', (_script, _accepted, refused, rule) => {
    const wb = workbookWith('Data');
    expect(() => wb.addNamedRange(refused, 'Data!$A$1')).toThrow(rule);
    expect(wb.namedRanges).toEqual([]);
  });

  it.each([
    'a1\u{301}',
    '_\u{301}',
    'a.\u{301}',
    '\\\u{301}',
    'A1\u{301}',
    'R\u{301}',
    'R1.\u{301}',
    'TRUE\u{301}',
    'a\u{FE0F}b',
  ])(
    'takes a mark after a digit, `_`, `.`, `\\`, a reference-like start or a letter: %j',
    (name) => {
      const wb = workbookWith('Data');
      wb.addNamedRange(name, 'Data!$A$1');
      expect(wb.getNamedRange(name)?.value).toBe('Data!$A$1');
    },
  );

  it.each(['R1\u{301}', 'C1\u{301}', 'RC1\u{301}'])('refuses %j, an R1C1 reference', (name) => {
    expect(() => workbookWith('Data').addNamedRange(name, 'Data!$A$1')).toThrow('R1C1');
  });

  it.each([
    '\u{E34}abc',
    'a\u{20AC}b',
    'a\u{B7}b',
    'a\u{B0}b',
    'a\u{1F600}b',
    '\u{645}\u{6CC}\u{200C}\u{62E}\u{648}\u{627}\u{647}\u{645}',
  ])('still refuses %j, which Excel takes (planned for 1.6.0)', (name) => {
    expect(() => workbookWith('Data').addNamedRange(name, 'Data!$A$1')).toThrow(/defined name/i);
  });

  // Excel 365 (review probe, 2026-11-12) kept only one of two defined names
  // `caf\u{E9}` and `cafe\u{301}` in one scope, and did not open a file
  // holding a table and a defined name spelled so.
  it.each([
    ['caf\u{E9}', 'cafe\u{301}'],
    ['cafe\u{301}', 'CAF\u{C9}'],
  ])('takes %j, then refuses the canonically equivalent %j in one scope', (first, second) => {
    const wb = workbookWith('Data');
    wb.addNamedRange(first, 'Data!$A$1');
    expect(() => wb.addNamedRange(second, 'Data!$A$2')).toThrow(/already exists/);
    expect(wb.namedRanges.map((d) => d.name)).toEqual([first]);
  });

  it('refuses a name canonically equivalent to a table name', () => {
    const wb = workbookWith('Data');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('missing sheet');
    ws.addTable({
      id: 1,
      displayName: 'caf\u{E9}',
      ref: 'A1:A3',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: false,
      columns: [{ id: 1, name: 'h' }],
    });
    expect(() => wb.addNamedRange('cafe\u{301}', 'Data!$A$2')).toThrow(/already used by table/);
    expect(wb.namedRanges).toEqual([]);
  });
});

/**
 * Made by Excel 365 (`Names.Add`, saved by Excel, 2026-11-12): sheets `Data`
 * and `Other`, and eight workbook-scope names that all refer to
 * `Data!$A$1`: `a€b`, `a°b`, `a·b`, `e` + U+0301 + `te`, `Plain_Name`, an Arabic,
 * a Hindi and a Thai word with combining marks.
 */
const XL_NAMES = readFileSync(
  new URL('../../../crates/modern-xlsx-core/tests/fixtures/excel/xl-names.xlsx', import.meta.url),
);

describe('renaming a sheet of a workbook Excel saved', () => {
  it('renameSheet rewrites every value and keeps every name, even ones addNamedRange refuses', async () => {
    const wb = await readBuffer(XL_NAMES);
    const names = wb.namedRanges.map((d) => d.name);
    expect(names).toHaveLength(8);
    expect(names).toContain('a€b');
    expect(wb.namedRanges.every((d) => d.value === 'Data!$A$1')).toBe(true);
    wb.renameSheet('Data', 'New');
    expect(wb.sheetNames).toEqual(['New', 'Other']);
    expect(wb.namedRanges.map((d) => d.name)).toEqual(names);
    expect(wb.namedRanges.map((d) => d.value)).toEqual(names.map(() => 'New!$A$1'));
    const back = await readBuffer(await wb.toBuffer());
    expect(back.namedRanges.map((d) => [d.name, d.value])).toEqual(
      names.map((name) => [name, 'New!$A$1']),
    );
    expect(back.validate().errorCount).toBe(0);
  });

  it('the name setter rewrites them likewise', async () => {
    const wb = await readBuffer(XL_NAMES);
    const names = wb.namedRanges.map((d) => d.name);
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('missing sheet');
    ws.name = 'Renamed';
    expect(wb.sheetNames).toEqual(['Renamed', 'Other']);
    expect(wb.namedRanges.map((d) => d.name)).toEqual(names);
    expect(wb.namedRanges.map((d) => d.value)).toEqual(names.map(() => 'Renamed!$A$1'));
  });
});

/**
 * The defined names of a workbook Excel 365 made (build 20527, `Names.Add`,
 * saved by Excel, 2026-11-13) with sheets `Sheet1`, `Sheet2`, `Sheet3`,
 * `My Sheet`, `O'Neil` and `Sheet10` and two other workbooks: `src.xlsx`
 * (open while the names were made, external link 1) and `dir sp/src2.xlsx`
 * (closed, external link 2). Each value is as `xl/workbook.xml` stores it.
 */
const EXCEL_BOOK_NAMES: ReadonlyArray<readonly [string, string, number | null]> = [
  ['extBookNm', '[1]!srcName', null],
  ['extClosedQ', '[2]Sheet1!$A$1', null],
  ['extOpen', '[1]Sheet1!$A$1', null],
  ['extOpenAp', "'[1]O''Neil'!$A$1", null],
  ['extOpenQ', "'[1]My Sheet'!$A$1", null],
  ['extRange', '[1]Sheet1!$A$1:$B$2', null],
  ['loc', 'Sheet1!$B$2', null],
  ['locAp', "'O''Neil'!$B$2", null],
  ['locQ', "'My Sheet'!$B$2", null],
  ['longer', 'Sheet10!$A$1', null],
  ['mixed', 'Sheet1!$A$1+[1]Sheet1!$A$1', null],
  ['myName', 'Sheet1!$C$3', 0],
  ['str', '"Sheet1!A1"', null],
  ['strMix', 'IF(Sheet1!$A$1=1,"Sheet1!A1",Sheet10!$A$1)', null],
  ['threeD', 'Sheet1:Sheet3!$A$1', null],
  ['threeDSum', 'SUM(Sheet1:Sheet3!$A$1)', null],
];

function excelBook(): Workbook {
  const wb = workbookWith('Sheet1', 'Sheet2', 'Sheet3', 'My Sheet', "O'Neil", 'Sheet10');
  for (const [name, value, sheetId] of EXCEL_BOOK_NAMES) {
    wb.addNamedRange(name, value, sheetId ?? undefined);
  }
  return wb;
}

/**
 * The bound for renaming one value of 100,000 references: about 70 times the
 * whole test's time on the development host (57-59 ms over three runs on
 * 2026-11-13), room for a slower CI runner.
 */
const RENAME_100K_BOUND_MS = 4000;

/** What Excel stores after its own rename of the sheet in that workbook: only the changed values. */
const EXCEL_RENAMES: ReadonlyArray<readonly [string, string, Readonly<Record<string, string>>]> = [
  [
    'Sheet1',
    'Data',
    {
      loc: 'Data!$B$2',
      mixed: 'Data!$A$1+[1]Sheet1!$A$1',
      myName: 'Data!$C$3',
      strMix: 'IF(Data!$A$1=1,"Sheet1!A1",Sheet10!$A$1)',
      threeD: 'Data:Sheet3!$A$1',
      threeDSum: 'SUM(Data:Sheet3!$A$1)',
    },
  ],
  ['Sheet3', 'Data', { threeD: 'Sheet1:Data!$A$1', threeDSum: 'SUM(Sheet1:Data!$A$1)' }],
  ['Sheet2', 'Data', {}],
  ['My Sheet', 'New Name', { locQ: "'New Name'!$B$2" }],
  ["O'Neil", "O'Brien's", { locAp: "'O''Brien''s'!$B$2" }],
];

describe('renameSheet follows what Excel stores after its own rename', () => {
  it.each(EXCEL_RENAMES)('renaming %j to %j', (from, to, changed) => {
    const wb = excelBook();
    wb.renameSheet(from, to);
    const expected = EXCEL_BOOK_NAMES.map(([name, value]) => [name, changed[name] ?? value]);
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual(expected);
  });

  it('keeps [1]Sheet1!$A$1, a sheet of another workbook, while Sheet1!$B$2 follows the rename', () => {
    const wb = workbookWith('Sheet1');
    wb.addNamedRange('ext', '[1]Sheet1!$A$1');
    wb.addNamedRange('loc', 'Sheet1!$B$2');
    wb.addNamedRange('both', '[1]sheet1!$A$1+sheet1!$A$1+[Book2.xlsx]Sheet1!A1');
    wb.renameSheet('Sheet1', 'Data');
    expect(wb.getNamedRange('ext')?.value).toBe('[1]Sheet1!$A$1');
    expect(wb.getNamedRange('loc')?.value).toBe('Data!$B$2');
    expect(wb.getNamedRange('both')?.value).toBe('[1]sheet1!$A$1+Data!$A$1+[Book2.xlsx]Sheet1!A1');
  });

  it.each([
    ['Sheet3', 'Data'],
    ['C', 'Data'],
  ])('keeps a quoted prefix that holds a bracket when %j is renamed to %j', (from, to) => {
    const wb = workbookWith('C', 'Sheet3', 'My Sheet');
    const values: ReadonlyArray<readonly [string, string]> = [
      ['extSpan', "'[1]My Sheet:Sheet3'!$A$1"],
      ['extPath', "'C:/dir/[Book.xlsx]Sheet1'!$A$1"],
      ['extSum', "SUM('[1]My Sheet:Sheet3'!$A$1)"],
    ];
    for (const [name, value] of values) wb.addNamedRange(name, value);
    wb.renameSheet(from, to);
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual(values);
  });

  // A quoted prefix with a colon is a 3-D span only when both ends name
  // sheets of the workbook; an end no sheet has (an empty one, one longer
  // than 31 UTF-16 units, one with / \ ? * [ ], or a 31-unit name that is
  // no sheet) leaves the prefix as it is.
  it('keeps a quoted prefix whose colon does not part two sheet names when C is renamed', () => {
    const wb = workbookWith('C', 'Sheet3');
    const values: ReadonlyArray<readonly [string, string]> = [
      ['path', "'C:/dir/Book.xlsx'!A1"],
      ['back', "'C:\\dir\\Book.xlsx'!A1"],
      ['ask', "'C:a?b'!A1"],
      ['star', "'C:a*b'!A1"],
      ['long', `'C:${'x'.repeat(32)}'!A1`],
      ['lastPath', "'dir/x:C'!A1"],
    ];
    for (const [name, value] of values) wb.addNamedRange(name, value);
    wb.addNamedRange('span', "'C:Sheet3'!A1");
    wb.addNamedRange('longEnd', `'C:${'x'.repeat(31)}'!A1`);
    wb.renameSheet('C', 'Data');
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual([
      ...values,
      ['span', 'Data:Sheet3!A1'],
      ['longEnd', `'C:${'x'.repeat(31)}'!A1`],
    ]);
  });

  /**
   * Excel 365 (build 20527, 2026-11-13; sheets Sheet1, Sheet3, My Sheet,
   * O'Brien, C, a 31-letter name and a 15-emoji name) reads a quoted prefix
   * with a colon as a 3-D span only when both ends are sheets of the
   * workbook; otherwise it reads a path into another workbook and its own
   * rename keeps it: `'Sheet1:Nosuch'!$A$1` showed as
   * `='Sheet1:[Nosuch]Nosuch'!$A$1` before and after renaming Sheet1 to
   * Data, `'C:Nosuch'!$A$1` as `='C:\Users\...\[Nosuch]Nosuch'!$A$1` before
   * and after renaming C to Drive, `'Sheet1:''x'!$A$1` as
   * `='Sheet1:[''x]''x'!$A$1` and `'Sheet1:x'''!$A$1` as
   * `='Sheet1:[x'']x'''!$A$1`, all kept; while `'Sheet1:Sheet3'!$A$1` became
   * `Data:Sheet3!$A$1`, `SUM(Sheet1:Sheet3!$A$1)` became
   * `SUM(Data:Sheet3!$A$1)`, `'Sheet1:QQQ…'` (31 letters, a sheet) became
   * `Data:QQQ…`, and `'C:Sheet3'!$A$1` named Drive and Sheet3.
   */
  it('rewrites a quoted colon prefix only when both ends are sheets of the workbook, as Excel does', () => {
    const q31 = 'Q'.repeat(31);
    const wb = workbookWith('Sheet1', 'Sheet3', 'C', q31);
    const kept: ReadonlyArray<readonly [string, string]> = [
      ['noSuch', "'Sheet1:Nosuch'!$A$1"],
      ['drive', "'C:Nosuch'!$A$1"],
      ['quoteFirst', "'Sheet1:''x'!$A$1"],
      ['quoteLast', "'Sheet1:x'''!$A$1"],
      ['long32', `'Sheet1:${q31}Q'!$A$1`],
    ];
    for (const [name, value] of kept) wb.addNamedRange(name, value);
    wb.addNamedRange('quoted', "'Sheet1:Sheet3'!$A$1");
    wb.addNamedRange('sum', 'SUM(Sheet1:Sheet3!$A$1)');
    wb.addNamedRange('reversed', "Sheet1!$A$1+'Sheet3:Sheet1'!$B$2");
    wb.addNamedRange('long31', `'Sheet1:${q31}'!$A$1`);
    wb.addNamedRange('cSpan', "'C:Sheet3'!$A$1");
    wb.renameSheet('Sheet1', 'Data');
    wb.renameSheet('C', 'Drive');
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual([
      ...kept,
      ['quoted', 'Data:Sheet3!$A$1'],
      ['sum', 'SUM(Data:Sheet3!$A$1)'],
      ['reversed', 'Data!$A$1+Sheet3:Data!$B$2'],
      ['long31', `Data:${q31}!$A$1`],
      ['cSpan', 'Drive:Sheet3!$A$1'],
    ]);
  });

  /**
   * `toLowerCase` ends ΟΔΟΣ (U+039F U+0394 U+039F U+03A3) in U+03C2, and
   * Excel takes U+03C2 and U+03C3 as one letter: Excel 365 (build 20527,
   * 2026-11-13) renamed that sheet to X and turned `ΟΔΟΣ!$A$2`, `οδος!$A$2`
   * (U+03C2), `οδοσ!$A$2` (U+03C3), `SUM(οδος:Other!$A$1)` and
   * `SUM(οδοσ:Other!$A$1)` into `X!$A$2` and `SUM(X:Other!$A$1)`.
   */
  it('folds the final sigma when it matches a prefix, as Excel does', () => {
    const wb = workbookWith('ΟΔΟΣ', 'Other');
    wb.addNamedRange('nU', 'ΟΔΟΣ!$A$2');
    wb.addNamedRange('nF', 'οδος!$A$2');
    wb.addNamedRange('nS', 'οδοσ!$A$2');
    wb.addNamedRange('spanF', 'SUM(οδος:Other!$A$1)');
    wb.addNamedRange('spanS', 'SUM(οδοσ:Other!$A$1)');
    wb.addNamedRange('spanLast', "SUM('Other:οδοσ'!$A$1)");
    wb.renameSheet(0, 'X');
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual([
      ['nU', 'X!$A$2'],
      ['nF', 'X!$A$2'],
      ['nS', 'X!$A$2'],
      ['spanF', 'SUM(X:Other!$A$1)'],
      ['spanS', 'SUM(X:Other!$A$1)'],
      ['spanLast', 'SUM(Other:X!$A$1)'],
    ]);
  });

  /**
   * Excel 365 (build 20527, 2026-11-13) does not take a prefix spelled in
   * another normal form for the sheet: its own rename of sheet `café`
   * (U+00E9) to `X` turned `café!$A$1` and `'CAFÉ'!$A$1` into `X!$A$1` and
   * left `cafe` + U+0301 `!$A$1` (stored as the external `[1]cafe` + U+0301
   * `!$A$1`), and its rename of sheet U+1E97 `a` left `T` + U+0308 `a!$A$1`,
   * typed in Excel or read from the file alike. The rename matches the
   * prefix ignoring case only, as Excel does.
   */
  it('leaves a prefix spelled in another normal form, as Excel does', () => {
    const wb = workbookWith('caf\u{E9}', '\u{1E97}b');
    wb.addNamedRange('nfc', 'caf\u{E9}!$A$1');
    wb.addNamedRange('upper', "'CAF\u{C9}'!$A$1");
    wb.addNamedRange('nfd', "'cafe\u{301}'!$A$1");
    wb.addNamedRange('same', "'\u{1E97}b'!$A$1");
    wb.addNamedRange('cap', "'T\u{308}b'!$A$1");
    wb.renameSheet('caf\u{E9}', 'X');
    wb.renameSheet('\u{1E97}b', 'Y');
    expect(wb.namedRanges.map((d) => [d.name, d.value])).toEqual([
      ['nfc', 'X!$A$1'],
      ['upper', 'X!$A$1'],
      ['nfd', "'cafe\u{301}'!$A$1"],
      ['same', 'Y!$A$1'],
      ['cap', "'T\u{308}b'!$A$1"],
    ]);
  });

  it('renames one value with 100,000 references in linear time', () => {
    const wb = workbookWith('Data');
    const count = 100_000;
    wb.addNamedRange('Many', Array.from({ length: count }, () => 'Data!$A$1').join('+'));
    const start = performance.now();
    wb.renameSheet('Data', 'New');
    const ms = performance.now() - start;
    expect(wb.getNamedRange('Many')?.value).toBe(
      Array.from({ length: count }, () => 'New!$A$1').join('+'),
    );
    expect(ms).toBeLessThan(RENAME_100K_BOUND_MS);
  });
});

describe('table names in table formulas are compared in canonical form', () => {
  it('a formula naming its table in NFD names the NFC table', () => {
    const wb = workbookWith('Data');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('missing sheet');
    const table = (displayName: string, id: number, ref: string, formula: string) => ({
      id,
      displayName,
      ref,
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: false,
      columns: [
        { id: 1, name: '@Home' },
        { id: 2, name: 'Total', calculatedColumnFormula: formula },
      ],
    });
    expect(() =>
      ws.addTable(table('Caf\u{E9}', 1, 'A1:B3', 'SUM(Cafe\u{301}[@Home])')),
    ).not.toThrow();
    expect(() =>
      ws.addTable(table('Caf\u{E9}2', 2, 'D1:E3', 'SUM(CAFE\u{301}2[@Home])')),
    ).not.toThrow();
    expect(() => ws.addTable(table('Other', 3, 'G1:H3', 'SUM(Other[@Away])'))).toThrow(
      /@ shorthand/,
    );
  });

  // The table being added is not in the sheet's list yet, so only the key of
  // its own display name finds it: `Cafe` + U+0301 (NFD) against the formula's
  // U+00E9, and `T` + U+0308 (NFC: no precomposed capital) against the
  // formula's `t` + U+0308, which composes to U+1E97.
  it.each([
    ['Cafe\u{301}', 'SUM(Caf\u{E9}[@Home])'],
    ['T\u{308}b', 'SUM(t\u{308}b[@Home])'],
  ])('a formula of table %j naming it as in %j names the table itself', (displayName, formula) => {
    const ws = workbookWith('Data').getSheet('Data');
    if (!ws) throw new Error('missing sheet');
    expect(() =>
      ws.addTable({
        id: 1,
        displayName,
        ref: 'A1:B3',
        headerRowCount: 1,
        totalsRowCount: 0,
        totalsRowShown: false,
        columns: [
          { id: 1, name: '@Home' },
          { id: 2, name: 'Total', calculatedColumnFormula: formula },
        ],
      }),
    ).not.toThrow();
  });
});
