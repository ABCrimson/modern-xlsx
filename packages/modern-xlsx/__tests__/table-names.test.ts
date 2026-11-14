import { describe, expect, it } from 'vitest';
import type { SheetData, TableDefinitionData } from '../src/index.js';
import { ModernXlsxError, Workbook } from '../src/index.js';
import { nameKey } from '../src/model/name-rules.js';

/*
 * The table-name rules, as Excel 365 (Windows, build of 2026-11-12) proved them:
 * one workbook per name, written by modern-xlsx and opened through COM with
 * alerts off (OK = opens, ERROR = "Unable to get the Open property"). Every
 * refused row below opened with ERROR and every accepted row with OK. Excel
 * also refuses some non-ASCII names (a no-break space or U+0085 anywhere; a
 * leading Arabic-Indic, Devanagari, Thai or Extended Arabic-Indic digit or
 * euro sign) without a rule the probes could pin down, so they are not
 * checked and not listed here. A leading combining mark (U+0301 + `abc`,
 * U+093F + `x`) was refused too; addTable refuses one, validate() does not
 * check it yet (see LEADING_MARK below).
 */

/** Rule fragments, as the messages of addTable and validate() name them. */
const EMPTY = 'empty';
const LONG = 'longer than 255';
const CHAR = 'character Excel refuses';
const FIRST = 'starts with';
const BOOL = 'TRUE and FALSE';
const A1 = 'cell reference';
const R1C1 = 'R1C1';

/** Every printable ASCII character a table name may not hold anywhere. */
const FORBIDDEN_ASCII = [...' !"#$%&\'()*+,-/:;<=>@[]^`{|}~'];

/** [name, rule] for each name Excel refused. */
const REFUSED: [string, string][] = [
  ['', EMPTY],
  [`T${'a'.repeat(255)}`, LONG],
  ['é'.repeat(256), LONG],
  // Length counts UTF-16 code units: 1 + 2 x 128 = 257.
  [`T${'𝐀'.repeat(128)}`, LONG],
  ...FORBIDDEN_ASCII.map((ch): [string, string] => [`Ta${ch}b`, CHAR]),
  ...FORBIDDEN_ASCII.map((ch): [string, string] => [`${ch}Tab`, CHAR]),
  ['My Table', CHAR],
  ['My-Table', CHAR],
  ['Ta\tb', CHAR],
  ['Ta\nb', CHAR],
  ['Ta\rb', CHAR],
  ['Ta\u007fb', CHAR],
  ['1Table', FIRST],
  ['123', FIRST],
  ['.Table', FIRST],
  ['?Table', FIRST],
  ['TRUE', BOOL],
  ['FALSE', BOOL],
  ['true', BOOL],
  ...['T2', 'Q1', 'A1', 'a1', 'AB12', 'xfd1', 'XFD1048576', 'A01', 'xA1', 'Rx1', 'Tab123'].map(
    (name): [string, string] => [name, A1],
  ),
  // R1, C1, RC1 and C16385 are also cell references (column R, C or RC; row 1 or 16385).
  ...['R1', 'C1', 'R01', 'RC1', 'C16385'].map((name): [string, string] => [name, A1]),
  ...[
    'R',
    'C',
    'r',
    'c',
    'RC',
    'rc',
    'R1C',
    'R1C1',
    'r2c3',
    'R1048576C16384',
    'R1C16385',
    'R1C0',
    'R1x',
    'R1_',
    'R1é',
    'R1\\',
    'R1?',
    'R10x',
    'R0001x',
    'R1048576x',
    'R1Cx',
    'R1C.',
    'R1C0x',
    'R1C1x',
    'R1C1_',
    'R1C1.a',
    'R1C1R',
    'R1C16384x',
    'R1C16385x',
    'R1C1048576x',
    'R1R',
    'RC1x',
    'RC16384x',
    'rC1x',
    'Rc1x',
    'C1x',
    'C1_',
    'C01x',
    'C1C',
    'C1R',
    'C1R1',
    'C16384x',
    'r1x',
    'c1x',
  ].map((name): [string, string] => [name, R1C1]),
];

/**
 * Names starting with a combining mark, which Excel refused: addTable
 * refuses them with the FIRST rule, but validate() does not check them yet,
 * so they are kept out of REFUSED.
 */
const LEADING_MARK: string[] = ['́abc', 'िx'];

/** Names Excel opened. */
const ACCEPTED: string[] = [
  'Table1',
  'Sales_2',
  `T${'a'.repeat(254)}`,
  'é'.repeat(255),
  `T${'𝐀'.repeat(127)}`,
  'Ta.b',
  'Ta?b',
  'Ta\\b',
  'Ta_b',
  'My.Table',
  '_Table',
  '\\Table',
  'Tab\\le',
  '\\',
  'Table?',
  'TRUE1',
  'TRUEx',
  'True_',
  // Cell-shaped but outside the grid, or more than a cell reference.
  'XFE1',
  'ABCD1',
  'A1048577',
  'A0',
  'A1x',
  'A1_',
  'T2x',
  'A1.b',
  'XFD1x',
  'A1é',
  'A1\\',
  'A1?',
  // R and C names that are no R1C1 reference: a row or column out of range,
  // a '.' after the number, or no number after R, C or RC.
  'R0',
  'C0',
  'R0x',
  'C0x',
  'R1048577',
  'R1048577x',
  'R1048577C1',
  'R99999999999x',
  'C16385x',
  'C99999x',
  'RC0',
  'RC16385x',
  'R1.',
  'R1.a',
  'R1048576.x',
  'C1.',
  'C1.x',
  'RC1.',
  'RC1.x',
  'R.',
  'R\\',
  'C?',
  'RC_',
  'RC.',
  'RC\\',
  'RC?',
  'RC.x',
  'RCx',
  'RCC',
  'Rx',
  'Rr',
  'RR',
  'RRx',
  'R_1',
  'CC',
  'CR',
  'CRx',
  'CR1x',
  'Cat',
  // Non-ASCII names Excel opened.
  'Données',
  'Таблица',
  '売上',
  'Étude',
  'بTab',
  '中Tab',
  '１Tab',
  'Ⅷb',
  'T𝐀b',
  'Taéb',
  'Ta€b',
  'Ta·b',
  '·Tab',
  'Ta–b',
  '–Tab',
  'Ta²b',
  '²Tab',
  'Ta٣b',
  'Ta😀b',
  '😀Tab',
  'Ta　b',
  '　Tab',
  'Ta b',
  'Ta b',
  'Ta​b',
  'Ta b',
  'Ta b',
  'Ta﻿b',
  'Ta b',
  'Ta b',
  'Ta­b',
  'Ta×b',
  'Ta¿b',
  // A combining mark after the first character: `a` + U+0301 + `bc`.
  'ábc',
];

/** A one-column table over `ref`; `totals` adds a summed totals row. */
function table(
  displayName: string,
  extra: Partial<TableDefinitionData> = {},
  ref = 'A1:A3',
): TableDefinitionData {
  return {
    id: 1,
    displayName,
    ref,
    headerRowCount: 1,
    totalsRowCount: 1,
    totalsRowShown: true,
    columns: [{ id: 1, name: 'h', totalsRowFunction: 'sum' }],
    ...extra,
  };
}

/** A sheet holding `tables` as a read file would, without going through addTable. */
function rawSheet(name: string, tables: TableDefinitionData[]): SheetData {
  return {
    name,
    worksheet: {
      dimension: null,
      rows: [],
      mergeCells: [],
      autoFilter: null,
      frozenPane: null,
      columns: [],
      tables,
    },
  };
}

/** The error issues of `wb.validate()` whose message names `rule`. */
function errorsNaming(wb: Workbook, rule: string) {
  return wb
    .validate()
    .issues.filter((issue) => issue.severity === 'error' && issue.message.includes(rule));
}

/** Calls `fn` and returns the ModernXlsxError it throws. */
function thrown(fn: () => unknown): ModernXlsxError {
  try {
    fn();
  } catch (e) {
    if (e instanceof ModernXlsxError) return e;
    throw e;
  }
  throw new Error('expected a ModernXlsxError');
}

describe('table names Excel refuses', () => {
  it.each(REFUSED)('addTable refuses %j (%s) and leaves the sheet unchanged', (name, rule) => {
    const ws = new Workbook().addSheet('S');
    ws.cell('A1').value = 'h';
    const error = thrown(() => ws.addTable(table(name)));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain(rule);
    expect(error.message).toContain(JSON.stringify(name));
    expect(ws.tables).toHaveLength(0);
    // The totals row was not written either.
    expect(ws.rows.map((r) => r.index)).toEqual([1]);
  });

  it.each(REFUSED)('validate() reports %j (%s) as a structure error', (name, rule) => {
    const wb = new Workbook({ sheets: [rawSheet('S', [table(name, { totalsRowCount: 0 })])] });
    const issues = errorsNaming(wb, rule);
    expect(issues).toHaveLength(1);
    expect(issues[0]?.category).toBe('structure');
    expect(issues[0]?.location).toBe(`S!table(${name})`);
  });

  it.each(LEADING_MARK)('addTable refuses %j, which starts with a combining mark', (name) => {
    const ws = new Workbook().addSheet('S');
    const error = thrown(() => ws.addTable(table(name)));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain(FIRST);
    expect(error.message).toContain('combining mark');
    expect(ws.tables).toHaveLength(0);
  });
});

describe('table names Excel accepts', () => {
  it.each(ACCEPTED)('addTable takes %j and validate() reports nothing', (name) => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addTable(table(name));
    expect(ws.tables[0]?.displayName).toBe(name);
    expect(wb.validate().errorCount).toBe(0);
  });

  it.each(ACCEPTED)('validate() takes %j in a raw model', (name) => {
    const wb = new Workbook({ sheets: [rawSheet('S', [table(name, { totalsRowCount: 0 })])] });
    expect(wb.validate().errorCount).toBe(0);
  });
});

describe('the name attribute', () => {
  // Excel opened each of these as the `name` of a table whose displayName is valid.
  it.each(['T2', 'My Table', 'R', '', 'A&B', '1x', 'TRUE', `T${'a'.repeat(254)}`])(
    'takes %j',
    (name) => {
      const wb = new Workbook();
      wb.addSheet('S').addTable(table('GoodName', { name }));
      expect(wb.validate().errorCount).toBe(0);
    },
  );

  it('refuses one longer than 255 characters', () => {
    const name = `T${'a'.repeat(255)}`;
    const ws = new Workbook().addSheet('S');
    const error = thrown(() => ws.addTable(table('GoodName', { name })));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain(LONG);
    expect(ws.tables).toHaveLength(0);
    const wb = new Workbook({ sheets: [rawSheet('S', [table('GoodName', { name })])] });
    expect(errorsNaming(wb, LONG)).toHaveLength(1);
  });
});

describe('table names must be unique', () => {
  /** [label, first table, second table, second sheet?, rule] Excel refused. */
  const clashes: [string, TableDefinitionData, TableDefinitionData, boolean, string][] = [
    ['the same name', table('Sales'), table('Sales', { id: 2 }, 'C1:C3'), false, 'table'],
    [
      'a name differing in case',
      table('Sales'),
      table('SALES', { id: 2 }, 'C1:C3'),
      false,
      'table',
    ],
    ['a name on another sheet', table('Sales'), table('sales', { id: 2 }), true, 'table'],
    ['a non-ASCII case pair', table('Täble'), table('TÄBLE', { id: 2 }, 'C1:C3'), false, 'table'],
    [
      'an equal name attribute',
      table('First', { name: 'Same' }),
      table('Second', { id: 2, name: 'Same' }, 'C1:C3'),
      false,
      'name attribute',
    ],
    [
      'a name attribute differing in case',
      table('First', { name: 'same' }),
      table('Second', { id: 2, name: 'SAME' }, 'C1:C3'),
      false,
      'name attribute',
    ],
  ];

  it.each(clashes)('addTable refuses %s', (_label, first, second, otherSheet, rule) => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    ws.addTable(first);
    const target = otherSheet ? wb.addSheet('T') : ws;
    const error = thrown(() => target.addTable(second));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain(`already used by ${rule}`);
    expect(target.tables).toHaveLength(otherSheet ? 0 : 1);
  });

  it.each(clashes)('validate() reports %s', (_label, first, second, otherSheet, rule) => {
    const sheets = otherSheet
      ? [rawSheet('S', [first]), rawSheet('T', [second])]
      : [rawSheet('S', [first, second])];
    const wb = new Workbook({ sheets });
    const issues = errorsNaming(wb, `already used by ${rule}`);
    expect(issues).toHaveLength(1);
    expect(issues[0]?.category).toBe('structure');
  });

  it.each([
    ['a workbook name', 'Rates', undefined],
    ['a name differing in case', 'RATES', undefined],
    ['a sheet-scoped name', 'Rates', 0],
  ] as const)('refuses a table named like %s', (_label, definedName, sheetId) => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    wb.addNamedRange(definedName, 'S!$B$1', sheetId);
    const error = thrown(() => ws.addTable(table('Rates')));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain('already used by defined name');
    expect(ws.tables).toHaveLength(0);
    const raw = new Workbook({
      sheets: [rawSheet('S', [table('Rates')])],
      definedNames: wb.namedRanges.map((d) => ({ ...d })),
    });
    expect(errorsNaming(raw, 'already used by defined name')).toHaveLength(1);
  });

  it('takes names Excel opened: a name attribute equal to a defined name or to a later display name, and a table named like its sheet', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    wb.addNamedRange('Rates', 'S!$B$1');
    ws.addTable(table('GoodName', { name: 'Rates' }));
    ws.addTable(table('First', { id: 2, name: 'Other' }, 'C1:C3'));
    ws.addTable(table('Other', { id: 3 }, 'E1:E3'));
    ws.addTable(table('S', { id: 4, name: 'S' }, 'G1:G3'));
    ws.addTable(table('Own', { id: 5, name: 'OWN' }, 'I1:I3'));
    expect(ws.tables).toHaveLength(5);
    expect(wb.validate().errorCount).toBe(0);
  });
});

describe('a defined name may not repeat a table name', () => {
  // Excel 365 refused each pair at every scope: `Names.Add` beside a table of
  // that name, and a written file holding both (validate() reports it).
  it.each([
    ['a workbook name', 'Sales', undefined],
    ['a name differing in case', 'SALES', undefined],
    ['a name scoped to the table sheet', 'sales', 0],
    ['a name scoped to another sheet', 'Sales', 1],
  ] as const)('addNamedRange refuses %s', (_label, definedName, sheetId) => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    wb.addSheet('T');
    ws.addTable(table('Sales'));
    const error = thrown(() => wb.addNamedRange(definedName, 'S!$B$1', sheetId));
    expect(error.code).toBe('INVALID_ARGUMENT');
    expect(error.message).toContain('already used by table "Sales"');
    expect(error.message).toContain(JSON.stringify(definedName));
    expect(wb.namedRanges).toEqual([]);
  });

  it('refuses both orders: addTable after addNamedRange, addNamedRange after addTable', () => {
    const first = new Workbook();
    const s1 = first.addSheet('S');
    first.addNamedRange('Rates', 'S!$B$1');
    expect(() => s1.addTable(table('rates'))).toThrow('already used by defined name');
    const second = new Workbook();
    second.addSheet('S').addTable(table('rates'));
    expect(() => second.addNamedRange('Rates', 'S!$B$1')).toThrow('already used by table');
  });

  it('a stored pair does not stop renameSheet, and validate() reports it', () => {
    const wb = new Workbook({
      sheets: [rawSheet('S', [table('Sales')])],
      definedNames: [{ name: 'sales', value: 'S!$B$1', sheetId: null }],
    });
    wb.renameSheet('S', 'Data');
    expect(wb.namedRanges[0]?.value).toBe('Data!$B$1');
    expect(errorsNaming(wb, 'already used by defined name')).toHaveLength(1);
  });

  describe('canonically equivalent names are the same name', () => {
    // `café` with U+00E9 and `cafe` + U+0301 are canonically equivalent; Excel
    // refuses a file holding a table and a defined name, or two tables, named so.
    const precomposed = 'café';
    const decomposed = 'café';

    it('addTable refuses a table equivalent to a defined name, and addNamedRange the reverse', () => {
      const first = new Workbook();
      const s1 = first.addSheet('S');
      first.addNamedRange(precomposed, 'S!$B$1');
      const error = thrown(() => s1.addTable(table(decomposed)));
      expect(error.code).toBe('INVALID_ARGUMENT');
      expect(error.message).toContain('already used by defined name');
      expect(s1.tables).toHaveLength(0);
      const second = new Workbook();
      second.addSheet('S').addTable(table(decomposed));
      expect(() => second.addNamedRange(precomposed, 'S!$B$1')).toThrow('already used by table');
    });

    it('addTable refuses a table equivalent to another table, by display name or name attribute', () => {
      const wb = new Workbook();
      const ws = wb.addSheet('S');
      ws.addTable(table(precomposed, { name: `N${precomposed}` }));
      const byDisplay = thrown(() =>
        ws.addTable(table(decomposed.toUpperCase(), { id: 2 }, 'C1:C3')),
      );
      expect(byDisplay.message).toContain('already used by table');
      const byName = thrown(() =>
        ws.addTable(table('Other', { id: 3, name: `N${decomposed}` }, 'E1:E3')),
      );
      expect(byName.message).toContain('already used by name attribute');
      expect(ws.tables).toHaveLength(1);
    });
  });

  it('takes a defined name no table has', () => {
    const wb = new Workbook();
    wb.addSheet('S').addTable(table('Sales'));
    wb.addNamedRange('Sales2', 'S!$B$1');
    wb.addNamedRange('Sale', 'S!$B$1');
    expect(wb.namedRanges.map((d) => d.name)).toEqual(['Sales2', 'Sale']);
    expect(wb.validate().errorCount).toBe(0);
  });
});

/**
 * The sheet-name, table-name and defined-name checks look the stored names'
 * keys up in indexes instead of computing them on every call. The indexes
 * must answer as a scan of the lists would after anything that changes a
 * name: the API's add, rename, remove, move, clone and load paths, and
 * changes made through the objects the API hands out (`toJSON()`,
 * `namedRanges`, `getNamedRange`, `tables`, `getTable`), which the
 * workbook does not see.
 */
describe('the indexed name checks answer as a scan of the lists does', () => {
  const POOL = [
    'Data',
    'DATA',
    'caf\u00E9',
    'cafe\u0301',
    'CAF\u00C9',
    'A\u03A3x',
    'a\u03C3x',
    'b\u03C2x',
    'b\u03C3x',
    'Sales',
    'SALES',
    'Sales_Q',
    'sales_q',
    'Other',
    'other',
  ];
  const keyOf = (name: string | null | undefined): string | undefined =>
    name ? nameKey(name) : undefined;
  const allTables = (wb: Workbook): TableDefinitionData[] =>
    wb.toJSON().sheets.flatMap((sheet) => sheet.worksheet.tables ?? []);
  const sheetTaken = (wb: Workbook, name: string, except?: number): boolean =>
    wb.toJSON().sheets.some((sheet, i) => i !== except && keyOf(sheet.name) === nameKey(name));
  const definedTaken = (wb: Workbook, name: string, scope: number | undefined): boolean =>
    (wb.toJSON().definedNames ?? []).some(
      (d) => keyOf(d.name) === nameKey(name) && (d.sheetId ?? null) === (scope ?? null),
    ) || allTables(wb).some((t) => keyOf(t.displayName) === nameKey(name));
  const tableTaken = (wb: Workbook, name: string, attribute: string | undefined): boolean =>
    allTables(wb).some(
      (t) =>
        keyOf(t.displayName) === nameKey(name) ||
        (attribute !== undefined && keyOf(t.name) === nameKey(attribute)),
    ) || (wb.toJSON().definedNames ?? []).some((d) => keyOf(d.name) === nameKey(name));

  let seed = 20261113;
  const rand = (n: number): number => {
    seed = (seed * 1103515245 + 12345) % 2147483648;
    return Math.floor((seed / 2147483648) * n);
  };
  const pick = (): string => {
    const base = POOL[rand(POOL.length)] ?? 'Data';
    return rand(3) === 0 ? base : `${base}_${rand(3)}`;
  };
  /** `true` when `run` throws a duplicate error; any other error fails the test. */
  const refused = (run: () => unknown): boolean => {
    try {
      run();
      return false;
    } catch (e) {
      expect(e).toBeInstanceOf(ModernXlsxError);
      expect((e as Error).message).toMatch(/already/);
      return true;
    }
  };

  /** One step's state: the workbook, its data, a sheet position `at` and a name. */
  interface Step {
    wb: Workbook;
    data: ReturnType<Workbook['toJSON']>;
    count: number;
    at: number;
    name: string;
  }
  /** Each operation returns the workbook to go on with and whether it compared a check with the scan. */
  type Operation = (step: Step) => { wb: Workbook; checked: boolean };
  const same = (step: Step, checked = false) => ({ wb: step.wb, checked });
  const check = (step: Step, label: string, run: () => unknown, expected: boolean) => {
    expect(refused(run), `${label} ${step.name}`).toBe(expected);
    return same(step, true);
  };

  const operations: Operation[] = [
    (st) => check(st, 'addSheet', () => st.wb.addSheet(st.name), sheetTaken(st.wb, st.name)),
    (st) =>
      st.count === 0
        ? same(st)
        : check(
            st,
            'renameSheet',
            () => st.wb.renameSheet(st.at, st.name),
            sheetTaken(st.wb, st.name, st.at),
          ),
    (st) => {
      if (st.count > 1) st.wb.removeSheet(st.at);
      return same(st);
    },
    (st) =>
      st.count === 0
        ? same(st)
        : check(
            st,
            'cloneSheet',
            () => st.wb.cloneSheet(st.at, st.name, rand(st.count + 1)),
            sheetTaken(st.wb, st.name),
          ),
    (st) => {
      if (st.count > 1) st.wb.moveSheet(st.at, rand(st.count));
      return same(st);
    },
    (st) => {
      const scope = rand(2) === 0 || st.count === 0 ? undefined : st.at;
      return check(
        st,
        'addNamedRange',
        () => st.wb.addNamedRange(st.name, '$A$1', scope),
        definedTaken(st.wb, st.name, scope),
      );
    },
    (st) => {
      const ws = st.wb.getSheetByIndex(st.at);
      if (!ws) return same(st);
      const attribute = rand(2) === 0 ? undefined : pick();
      const added = table(st.name, {
        totalsRowCount: 0,
        ...(attribute ? { name: attribute } : {}),
      });
      return check(st, 'addTable', () => ws.addTable(added), tableTaken(st.wb, st.name, attribute));
    },
    (st) => {
      const names = st.wb.namedRanges;
      const victim = names[rand(Math.max(names.length, 1))];
      if (victim) st.wb.removeNamedRange(victim.name);
      return same(st);
    },
    (st) => {
      const ws = st.wb.getSheetByIndex(st.at);
      const victim = ws?.tables[rand(Math.max(ws.tables.length, 1))];
      if (victim) ws?.removeTable(victim.displayName);
      return same(st);
    },
    // Behind the workbook's back: a sheet renamed in place or its list entry replaced.
    (st) => {
      const sheet = st.data.sheets[st.at];
      if (sheet && rand(2) === 0) sheet.name = st.name;
      else if (sheet) st.data.sheets[st.at] = { ...sheet, name: st.name };
      return same(st);
    },
    // A defined name renamed, removed or rescoped in place, or the list replaced.
    (st) => {
      const names = st.wb.namedRanges as DefinedNameLike[];
      const victim = names[rand(Math.max(names.length, 1))];
      const roll = rand(3);
      if (!victim) st.data.definedNames = [{ name: st.name, value: '$A$1', sheetId: null }];
      else if (roll === 0) victim.name = st.name;
      else if (roll === 1) names.splice(names.indexOf(victim), 1);
      else victim.sheetId = st.count > 0 ? rand(st.count) : null;
      return same(st);
    },
    // A table renamed in place (either name), one pushed, or a sheet's list dropped.
    (st) => {
      const tables = st.wb.getSheetByIndex(st.at)?.tables as TableDefinitionData[] | undefined;
      const victim = tables?.[rand(Math.max(tables.length, 1))];
      const roll = rand(4);
      if (victim && roll === 0) victim.displayName = st.name;
      else if (victim && roll === 1) victim.name = st.name;
      else if (tables && roll === 2) tables.push(table(st.name, { totalsRowCount: 0 }));
      else if (st.data.sheets[st.at]) delete st.data.sheets[st.at]?.worksheet.tables;
      return same(st);
    },
    // A load: a new workbook over a copy of the data, or over the same data.
    (st) => ({
      wb: new Workbook(rand(2) === 0 ? structuredClone(st.data) : st.data),
      checked: false,
    }),
  ];
  /** The checks run twice as often as the other operations. */
  const weights = [0, 0, 1, 2, 3, 4, 5, 5, 6, 6, 7, 8, 9, 10, 11, 12];

  it('over 4,000 seeded operations of the API and of the objects it hands out', () => {
    let wb = new Workbook();
    wb.addSheet('Start');
    let checks = 0;
    for (let i = 0; i < 4000; i++) {
      const data = wb.toJSON();
      const count = data.sheets.length;
      const step: Step = { wb, data, count, at: rand(Math.max(count, 1)), name: pick() };
      const operation = operations[weights[rand(weights.length)] ?? 0];
      const result = operation?.(step) ?? same(step);
      wb = result.wb;
      if (result.checked) checks++;
    }
    expect(checks).toBeGreaterThan(1500);
  });
});

/** A defined name as the API hands it out, written to as a caller can. */
interface DefinedNameLike {
  name: string;
  value: string;
  sheetId?: number | null;
}
