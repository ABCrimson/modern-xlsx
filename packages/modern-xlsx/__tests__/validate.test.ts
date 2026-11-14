import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import type { TableDefinitionData } from '../src/index.js';
import { ModernXlsxError, Workbook } from '../src/index.js';

describe('OOXML Validation & Compliance', () => {
  it('validates a minimal workbook as clean', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();
    expect(report.isValid).toBe(true);
    expect(report.errorCount).toBe(0);
  });

  it('uses Calibri as the default font of a new workbook', () => {
    expect(new Workbook().styles.fonts[0]?.name).toBe('Calibri');
  });

  it('detects dangling style index on a cell', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'test';
    ws.cell('A1').styleIndex = 999;
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.errorCount).toBeGreaterThan(0);
    const issue = report.issues.find(
      (i) => i.category === 'styleIndex' && i.message.includes('999'),
    );
    expect(issue).toBeDefined();
    expect(issue?.autoFixable).toBe(true);
  });

  it('detects overlapping merge regions', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.addMergeCell('A1:B2');
    ws.addMergeCell('B2:C3');
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.issues.some((i) => i.category === 'mergeCell')).toBe(true);
  });

  it('warns about missing theme colors (info level)', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();
    // New workbooks set no theme colors; validation reports that as an info issue, never an error.
    const themeIssue = report.issues.find((i) => i.category === 'theme');
    expect(themeIssue).toMatchObject({ severity: 'info', category: 'theme' });
  });

  it('detects duplicate sheet names (case-insensitive)', () => {
    const wb = new Workbook({
      sheets: [
        { name: 'Data', worksheet: emptyWorksheet() },
        { name: 'data', worksheet: emptyWorksheet() },
      ],
      dateSystem: 'date1900',
      styles: defaultStyles(),
    });
    const report = wb.validate();
    expect(report.isValid).toBe(false);
    expect(report.issues.some((i) => i.message.includes('Duplicate'))).toBe(true);
  });

  it('repairs dangling style indices', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 42;
    ws.cell('A1').styleIndex = 999;

    const { workbook, report, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    expect(report.isValid).toBe(true);
    expect(workbook.getSheet('Sheet1')?.cell('A1').styleIndex).toBe(0);
  });

  it('repairs missing default styles', () => {
    const wb = new Workbook({
      sheets: [{ name: 'Sheet1', worksheet: emptyWorksheet() }],
      dateSystem: 'date1900',
      styles: {
        numFmts: [],
        fonts: [],
        fills: [],
        borders: [],
        cellXfs: [],
      },
    });

    const reportBefore = wb.validate();
    expect(reportBefore.isValid).toBe(false);

    const { workbook, report, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    expect(report.isValid).toBe(true);
    expect(workbook.styles.fonts.length).toBeGreaterThan(0);
    expect(workbook.styles.fills.length).toBeGreaterThanOrEqual(2);
    expect(workbook.styles.cellXfs.length).toBeGreaterThan(0);
  });

  it('repairs bad metadata dates', () => {
    const wb = new Workbook({
      sheets: [{ name: 'Sheet1', worksheet: emptyWorksheet() }],
      dateSystem: 'date1900',
      styles: defaultStyles(),
      docProperties: {
        created: 'not-a-date',
        title: '   ', // whitespace-only
      },
    });

    const { workbook, repairCount } = wb.repair();
    expect(repairCount).toBeGreaterThan(0);
    // serde skips None fields (undefined in JS), not null
    expect(workbook.docProperties?.created).toBeUndefined();
    expect(workbook.docProperties?.title).toBeUndefined();
  });

  it('repaired workbook can be written and read back', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Sheet1');
    ws.cell('A1').value = 'Hello';
    ws.cell('A1').styleIndex = 999;

    const { workbook } = wb.repair();
    const buffer = await workbook.toBuffer();
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('validation report has correct structure', () => {
    const wb = new Workbook();
    wb.addSheet('Sheet1');
    const report = wb.validate();

    expect(typeof report.isValid).toBe('boolean');
    expect(typeof report.errorCount).toBe('number');
    expect(typeof report.warningCount).toBe('number');
    expect(typeof report.infoCount).toBe('number');
    expect(Array.isArray(report.issues)).toBe(true);

    for (const issue of report.issues) {
      expect(typeof issue.severity).toBe('string');
      expect(typeof issue.category).toBe('string');
      expect(typeof issue.message).toBe('string');
      expect(typeof issue.location).toBe('string');
      expect(typeof issue.suggestion).toBe('string');
      expect(typeof issue.autoFixable).toBe('boolean');
    }
  });
});

// Helpers

function emptyWorksheet() {
  return {
    dimension: null,
    rows: [],
    mergeCells: [],
    autoFilter: null,
    frozenPane: null,
    columns: [],
  };
}

function defaultStyles() {
  return {
    numFmts: [],
    fonts: [
      {
        name: 'Calibri',
        size: 11,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        color: null,
      },
    ],
    fills: [
      { patternType: 'none', fgColor: null, bgColor: null },
      { patternType: 'gray125', fgColor: null, bgColor: null },
    ],
    borders: [{ left: null, right: null, top: null, bottom: null }],
    cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
  };
}

/**
 * Most milliseconds one check of a 16,384-column table may take: each took 5 to
 * 12 ms on the release host (Node 26, 2026-11-12), and the bound leaves 40x
 * that for a slow runner. Before the check was linear, validate() took 5.7 s.
 */
const BOUND_MS = 500;

/** A case of the shared table `this-row-shorthand-cases.json` (see its `about`). */
interface ShorthandCase {
  readonly formula: string;
  readonly columns: readonly string[];
  readonly tables: Readonly<Record<string, readonly string[]>>;
  readonly reported: boolean;
}

/** The case table the Rust `tests/hostile_validate.rs` runs too, so the two checks cannot drift. */
const SHORTHAND_CASES: readonly ShorthandCase[] = JSON.parse(
  readFileSync(
    new URL(
      '../../../crates/modern-xlsx-core/tests/fixtures/this-row-shorthand-cases.json',
      import.meta.url,
    ),
    'utf8',
  ),
).cases;

/** Table `displayName` over a header and one data row from row `row + 1`, with the columns `names` (at most 26). */
function formulaTable(
  id: number,
  displayName: string,
  row: number,
  names: readonly string[],
): TableDefinitionData {
  const last = String.fromCharCode(64 + names.length);
  return {
    id,
    displayName,
    name: displayName,
    ref: `A${row + 1}:${last}${row + 2}`,
    headerRowCount: 1,
    totalsRowCount: 0,
    totalsRowShown: false,
    columns: names.map((name, c) => ({ id: c + 1, name })),
  };
}

/** Case `c`'s table `T` (its columns and `Out`, holding the formula) and the other tables. */
function caseTables(c: ShorthandCase): [TableDefinitionData, TableDefinitionData[]] {
  const own = formulaTable(1, 'T', 0, [...c.columns, 'Out']);
  const out = own.columns.at(-1);
  if (out) out.calculatedColumnFormula = c.formula;
  const others = Object.entries(c.tables).map(([name, columns], i) =>
    formulaTable(i + 2, name, 10 * (i + 1), columns),
  );
  return [own, others];
}

/** The `@ shorthand` errors validate() reports for a raw model holding `tables` on sheet `S`, as read from a file. */
function rawShorthandErrors(tables: TableDefinitionData[]) {
  const wb = new Workbook();
  wb.addSheet('S');
  const sheet = wb.toJSON().sheets[0];
  if (!sheet) throw new Error('missing sheet');
  sheet.worksheet.tables = tables;
  return wb.validate().issues.filter((issue) => issue.message.includes('@ shorthand'));
}

/** Whether addTable refuses `table` as the `@` shorthand on a sheet that already holds `others`. */
function addTableRefuses(table: TableDefinitionData, others: TableDefinitionData[]): boolean {
  const ws = new Workbook().addSheet('S');
  for (const other of others) ws.addTable(other);
  try {
    ws.addTable(table);
    return false;
  } catch (e) {
    if (e instanceof ModernXlsxError && e.message.includes('@ shorthand')) return true;
    throw e;
  }
}

describe('the @ this-row shorthand in table formulas', () => {
  it('has a shared case table of at least 30 cases', () => {
    expect(SHORTHAND_CASES.length).toBeGreaterThanOrEqual(30);
  });

  it.each(SHORTHAND_CASES)('addTable and validate() agree on %j', (c) => {
    const [own, others] = caseTables(c);
    expect(addTableRefuses(own, others)).toBe(c.reported);
    expect(rawShorthandErrors([own, ...others])).toHaveLength(c.reported ? 1 : 0);
  });

  /** The table part Excel 16.0 (build 20527) saved for table `LocT` beside a link to another workbook, which Excel opens. */
  function excelSavedTable(): TableDefinitionData {
    const table = formulaTable(1, 'LocT', 0, [
      'Item',
      'Qty',
      '@a_x005f_x000a_b',
      '@t_x0009_b',
      'E1',
      'E2',
      'E3',
      'E4',
      'E5',
    ]);
    const formulas = [
      'SUM([1]!Ext[@Home])',
      'LocT[[#This Row],[Qty]]+SUM([1]!Ext[Qty])',
      'SUM(LocT[@a_x005F_x000a_b])',
      'SUM([1]!Ext[[#All],[@Home]])',
      'SUM(LocT[@t\tb])',
    ];
    formulas.forEach((formula, i) => {
      const column = table.columns[4 + i];
      if (column) column.calculatedColumnFormula = formula;
    });
    const qty = table.columns[1];
    if (qty) {
      qty.totalsRowFunction = 'custom';
      qty.totalsRowFormula = 'SUM([1]!Ext[@Home])';
    }
    return { ...table, ref: 'A1:I5', totalsRowCount: 1, totalsRowShown: true };
  }

  it("takes the formulas Excel saved for another workbook's table and an escaped column name", () => {
    const ws = new Workbook().addSheet('Loc');
    ws.addTable(excelSavedTable());
    expect(ws.cell('B5').formula).toBe('SUM([1]!Ext[@Home])');
    expect(rawShorthandErrors([excelSavedTable()])).toEqual([]);
  });

  /** A table of 16,384 columns, the last named `last`, each holding `formula(c)`. */
  function wideTable(last: string, formula: (c: number) => string): TableDefinitionData {
    const columns = Array.from({ length: 16_384 }, (_, c) => ({
      id: c + 1,
      name: c === 16_383 ? last : `C${c}`,
      calculatedColumnFormula: formula(c),
    }));
    return {
      id: 1,
      displayName: 'Wide',
      name: 'Wide',
      ref: 'A1:XFD2',
      headerRowCount: 1,
      totalsRowCount: 0,
      totalsRowShown: false,
      columns,
    };
  }

  it('checks a 16,384-column table with a [@...] in every column in linear time', () => {
    const accepted = wideTable('@x', () => '[@x]');
    // Every column passes but the last, whose `[@y]` is the shorthand.
    const refused = wideTable('@x', (c) => (c === 16_383 ? '[@x]+[@y]' : '[@x]'));
    const time = (run: () => void) => {
      const start = performance.now();
      run();
      return performance.now() - start;
    };
    const add = time(() => new Workbook().addSheet('S').addTable(accepted));
    const refuse = time(() =>
      expect(() => new Workbook().addSheet('S').addTable(refused)).toThrow('@ shorthand [@y]'),
    );
    let errors: unknown[] = [];
    const check = time(() => {
      errors = rawShorthandErrors([refused]);
    });
    expect(errors).toHaveLength(1);
    expect(add, 'addTable, accepted').toBeLessThan(BOUND_MS);
    expect(refuse, 'addTable, refused').toBeLessThan(BOUND_MS);
    expect(check, 'validate()').toBeLessThan(BOUND_MS);
  });
});
