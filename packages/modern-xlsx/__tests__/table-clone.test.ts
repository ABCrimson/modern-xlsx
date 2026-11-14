import { describe, expect, it } from 'vitest';
import type { TableColumnData, TableDefinitionData } from '../src/index.js';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

/** A table over `ref` with a header row and a totals row. */
function table(
  displayName: string,
  ref: string,
  columns: Omit<TableColumnData, 'id'>[],
  id = 1,
): TableDefinitionData {
  return {
    id,
    name: displayName,
    displayName,
    ref,
    headerRowCount: 1,
    totalsRowCount: 1,
    totalsRowShown: true,
    columns: columns.map((c, i) => ({ id: i + 1, ...c })),
  };
}

/** Sheet `Data` with Item/Amount over A1:B3 and table `Table1` over A1:B4 (label + sum totals). */
function authored(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  for (const [r, row] of [
    ['Item', 'Amount'],
    ['a', 1],
    ['b', 2],
  ].entries()) {
    ws.cell(`A${r + 1}`).value = row[0] ?? null;
    ws.cell(`B${r + 1}`).value = row[1] ?? null;
  }
  ws.addTable(
    table('Table1', 'A1:B4', [
      { name: 'Item', totalsRowLabel: 'Total' },
      { name: 'Amount', totalsRowFunction: 'sum' },
    ]),
  );
  return wb;
}

describe('cloneSheet renames the clone table in its formulas', () => {
  it('points the totals row of the clone at the clone table and leaves the source alone', () => {
    const wb = authored();
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.tables[0]?.displayName).toBe('Table1_2');
    expect(copy.cell('B4').formula).toBe('SUBTOTAL(109,Table1_2[Amount])');
    expect(copy.cell('A4').value).toBe('Total');
    const source = wb.getSheet('Data');
    expect(source?.tables[0]?.displayName).toBe('Table1');
    expect(source?.cell('B4').formula).toBe('SUBTOTAL(109,Table1[Amount])');
  });

  it('writes the renamed totals formula into the clone sheet part', async () => {
    const wb = authored();
    wb.cloneSheet(0, 'Copy');
    const parts = readZipText(await wb.toBuffer());
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<f>SUBTOTAL(109,Table1[Amount])</f>');
    expect(parts['xl/worksheets/sheet2.xml']).toContain('<f>SUBTOTAL(109,Table1_2[Amount])</f>');
  });

  it('renames the table in calculated-column and custom totals formulas of the clone table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.cell('B1').value = 'Double';
    ws.cell('A2').value = 3;
    ws.cell('B2').formula = 'Table1[[#This Row],[Amount]]*2';
    ws.addTable(
      table('Table1', 'A1:B3', [
        { name: 'Amount', totalsRowFunction: 'sum' },
        {
          name: 'Double',
          calculatedColumnFormula: 'Table1[[#This Row],[Amount]]*2',
          totalsRowFunction: 'custom',
          totalsRowFormula: 'SUM(Table1[Double])/COUNT(Table1[Amount])',
        },
      ]),
    );
    const copy = wb.cloneSheet(0, 'Copy');
    const column = copy.tables[0]?.columns[1];
    expect(column?.calculatedColumnFormula).toBe('Table1_2[[#This Row],[Amount]]*2');
    expect(column?.totalsRowFormula).toBe('SUM(Table1_2[Double])/COUNT(Table1_2[Amount])');
    expect(copy.cell('B2').formula).toBe('Table1_2[[#This Row],[Amount]]*2');
    expect(copy.cell('B3').formula).toBe('SUM(Table1_2[Double])/COUNT(Table1_2[Amount])');
    const sourceColumn = wb.getSheet('Data')?.tables[0]?.columns[1];
    expect(sourceColumn?.calculatedColumnFormula).toBe('Table1[[#This Row],[Amount]]*2');
    expect(sourceColumn?.totalsRowFormula).toBe('SUM(Table1[Double])/COUNT(Table1[Amount])');
  });

  it('renames the table in every cell formula of the cloned sheet, ignoring case', () => {
    const wb = authored();
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('no sheet');
    ws.cell('D1').formula = 'SUM(Table1[Amount])+COUNTA(Table1[[#All],[Item]])';
    ws.cell('D2').formula = '=MAX(table1[Amount])';
    ws.cell('D3').formula = 'ROWS(TABLE1[#Data])';
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('D1').formula).toBe('SUM(Table1_2[Amount])+COUNTA(Table1_2[[#All],[Item]])');
    expect(copy.cell('D2').formula).toBe('=MAX(Table1_2[Amount])');
    expect(copy.cell('D3').formula).toBe('ROWS(Table1_2[#Data])');
    expect(ws.cell('D1').formula).toBe('SUM(Table1[Amount])+COUNTA(Table1[[#All],[Item]])');
    expect(ws.cell('D2').formula).toBe('=MAX(table1[Amount])');
  });

  it('leaves string literals, longer names and tables on other sheets alone', () => {
    const wb = authored();
    const other = wb.addSheet('Other');
    other.cell('A1').value = 'V';
    other.cell('A2').value = 5;
    other.addTable({
      ...table('Table10', 'A1:A3', [{ name: 'V', totalsRowFunction: 'sum' }]),
      id: 2,
    });
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('no sheet');
    ws.cell('D1').formula = '"Table1[Amount]"&Table1[[#Headers],[Amount]]';
    ws.cell('D2').formula = 'SUM(Table10[V])+Table1x';
    ws.cell('D3').formula = 'Table1[Amount]+Data!';
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('D1').formula).toBe('"Table1[Amount]"&Table1_2[[#Headers],[Amount]]');
    expect(copy.cell('D2').formula).toBe('SUM(Table10[V])+Table1x');
    // A formula the tokenizer rejects is kept as it is.
    expect(copy.cell('D3').formula).toBe('Table1[Amount]+Data!');
    expect(other.cell('A3').formula).toBe('SUBTOTAL(109,Table10[V])');
  });

  it('renames each of several tables on the cloned sheet to its own new name', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'X';
    ws.cell('C1').value = 'Y';
    ws.addTable(table('East', 'A1:A3', [{ name: 'X', totalsRowFunction: 'sum' }]));
    ws.addTable(table('West', 'C1:C3', [{ name: 'Y', totalsRowFunction: 'count' }], 2));
    ws.cell('E1').formula = 'East[#Totals]+West[#Totals]';
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.tables.map((t) => t.displayName)).toEqual(['East_2', 'West_2']);
    expect(copy.cell('A3').formula).toBe('SUBTOTAL(109,East_2[X])');
    expect(copy.cell('C3').formula).toBe('SUBTOTAL(103,West_2[Y])');
    expect(copy.cell('E1').formula).toBe('East_2[#Totals]+West_2[#Totals]');
  });

  it('renames the totals formula of a table read from a file', async () => {
    const wb = await readBuffer(await authored().toBuffer());
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('B4').formula).toBe('SUBTOTAL(109,Table1_2[Amount])');
    expect(wb.getSheet('Data')?.cell('B4').formula).toBe('SUBTOTAL(109,Table1[Amount])');
  });
});

/** Error-severity issues `validate()` reports, as `location: message` lines. */
function validationErrors(wb: Workbook): string[] {
  return wb
    .validate()
    .issues.filter((issue) => issue.severity === 'error')
    .map((issue) => `${issue.location}: ${issue.message}`);
}

/**
 * Sheet `Data` with table `name` over A1:C4 (Item, Amount, and a calculated
 * column Double with a custom totals formula; label + sum totals), a formula
 * outside the table in E1:E4, and sheet `Other` with table `Elsewhere\x`.
 */
function oddNamed(name: string): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  for (const [r, row] of [
    ['Item', 'Amount', 'Double'],
    ['a', 1, null],
    ['b', 2, null],
  ].entries()) {
    ws.cell(`A${r + 1}`).value = row[0] ?? null;
    ws.cell(`B${r + 1}`).value = row[1] ?? null;
    ws.cell(`C${r + 1}`).value = row[2] ?? null;
  }
  for (const r of [2, 3]) ws.cell(`C${r}`).formula = `${name}[[#This Row],[Amount]]*2`;
  ws.addTable(
    table(name, 'A1:C4', [
      { name: 'Item', totalsRowLabel: 'Total' },
      { name: 'Amount', totalsRowFunction: 'sum' },
      {
        name: 'Double',
        calculatedColumnFormula: `${name}[[#This Row],[Amount]]*2`,
        totalsRowFunction: 'custom',
        totalsRowFormula: `SUM(${name}[Double])`,
      },
    ]),
  );
  const other = wb.addSheet('Other');
  other.cell('A1').value = 'V';
  other.cell('A2').value = 5;
  other.addTable({
    ...table('Elsewhere\\x', 'A1:A3', [{ name: 'V', totalsRowFunction: 'sum' }]),
    id: 2,
  });
  ws.cell('E1').formula = `SUM(${name}[Amount])+Elsewhere\\x[V]`;
  ws.cell('E2').formula = `"${name}[Amount]"&${name}[[#Headers],[Amount]]`;
  ws.cell('E3').formula = `Data!${name}[Amount]`;
  ws.cell('E4').formula = `${name}[Amount]+Data!`;
  return wb;
}

/** Every formula of a sheet and its tables, keyed by cell reference or `table/column/kind`. */
function formulas(wb: Workbook, sheet: string): Record<string, string> {
  const data = wb.toJSON().sheets.find((s) => s.name === sheet);
  const out: Record<string, string> = {};
  for (const cell of data?.worksheet.rows.flatMap((row) => row.cells) ?? []) {
    if (cell.formula) out[cell.reference] = cell.formula;
  }
  for (const t of data?.worksheet.tables ?? []) {
    for (const c of t.columns) {
      if (c.calculatedColumnFormula) out[`${c.name}/calculated`] = c.calculatedColumnFormula;
      if (c.totalsRowFormula) out[`${c.name}/totals`] = c.totalsRowFormula;
    }
  }
  return out;
}

describe('cloneSheet renames tables whose names the formula tokenizer cannot scan', () => {
  for (const name of ['a\\b', '\\abc', 'a?b', 'Ventes€', 'a·b']) {
    it(`points every formula of the clone at the clone of ${name}`, () => {
      const wb = oddNamed(name);
      const before = formulas(wb, 'Data');
      const copy = wb.cloneSheet(0, 'Copy');
      const renamed = `${name}_2`;
      expect(copy.tables[0]?.displayName).toBe(renamed);
      expect(formulas(wb, 'Copy')).toEqual({
        B4: `SUBTOTAL(109,${renamed}[Amount])`,
        C2: `${renamed}[[#This Row],[Amount]]*2`,
        C3: `${renamed}[[#This Row],[Amount]]*2`,
        C4: `SUM(${renamed}[Double])`,
        'Double/calculated': `${renamed}[[#This Row],[Amount]]*2`,
        'Double/totals': `SUM(${renamed}[Double])`,
        // A table on another sheet keeps its name.
        E1: `SUM(${renamed}[Amount])+Elsewhere\\x[V]`,
        // Text in a string literal is not a reference.
        E2: `"${name}[Amount]"&${renamed}[[#Headers],[Amount]]`,
        // A sheet-qualified name and a dangling `Data!` make the formula one
        // the tokenizer rejects with any table name: it is kept as it is.
        E3: `Data!${name}[Amount]`,
        E4: `${name}[Amount]+Data!`,
      });
      expect(formulas(wb, 'Data')).toEqual(before);
      expect(formulas(wb, 'Other')).toEqual({ A3: 'SUBTOTAL(109,Elsewhere\\x[V])' });
      expect(validationErrors(wb)).toEqual([]);
    });
  }

  it('writes the renamed totals formula of a backslash name into the clone sheet part', async () => {
    const wb = oddNamed('a\\b');
    wb.cloneSheet(0, 'Copy');
    const parts = readZipText(await wb.toBuffer());
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<f>SUBTOTAL(109,a\\b[Amount])</f>');
    expect(parts['xl/worksheets/sheet3.xml']).toContain('<f>SUBTOTAL(109,a\\b_2[Amount])</f>');
  });

  it('keeps a formula with an unclosed bracket or string as it is', () => {
    const wb = oddNamed('a?b');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('no sheet');
    ws.cell('F1').formula = 'a?b[Amount';
    ws.cell('F2').formula = 'a?b[Amount]&"open';
    ws.cell('F3').formula = "a?b[Amount]&'open";
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('F1').formula).toBe('a?b[Amount');
    expect(copy.cell('F2').formula).toBe('a?b[Amount]&"open');
    expect(copy.cell('F3').formula).toBe("a?b[Amount]&'open");
  });

  it('leaves sheet- and workbook-qualified references as the tokenized path does', () => {
    const wb = oddNamed('a\\b');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('no sheet');
    ws.cell('F1').formula = 'Data!Table1[a]';
    ws.cell('F2').formula = 'Data!a\\b[Amount]';
    ws.cell('F3').formula = "'Data'!a\\b[Amount]";
    ws.cell('F4').formula = '[1]Sheet1!a\\b[Amount]';
    ws.cell('F5').formula = '[1]!a\\b[Amount]';
    const copy = wb.cloneSheet(0, 'Copy');
    for (const [ref, formula] of [
      ['F1', 'Data!Table1[a]'],
      ['F2', 'Data!a\\b[Amount]'],
      ['F3', "'Data'!a\\b[Amount]"],
      ['F4', '[1]Sheet1!a\\b[Amount]'],
      ['F5', '[1]!a\\b[Amount]'],
    ] as const) {
      expect(copy.cell(ref).formula).toBe(formula);
    }
  });

  it('skips escaped brackets and quotes inside the brackets of a reference', () => {
    const wb = oddNamed('a€b');
    const ws = wb.getSheet('Data');
    if (!ws) throw new Error('no sheet');
    ws.cell('F1').formula = "a€b[[#This Row],['[x']]]&a€b['']+a€b[Amount]";
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('F1').formula).toBe("a€b_2[[#This Row],['[x']]]&a€b_2['']+a€b_2[Amount]");
  });
});

describe('cloneSheet keeps clone table names within 255 UTF-16 units', () => {
  /** Lone surrogates: a high one not followed by a low one, or a low one not preceded by a high one. */
  const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

  for (const [label, name] of [
    ['255 units', `L${'x'.repeat(254)}`],
    ['254 units', `L${'x'.repeat(253)}`],
    ['an astral character at the cut', `L${'x'.repeat(251)}😀y`],
  ] as const) {
    it(`cuts a source name of ${label} so every clone name fits`, () => {
      expect(name.length).toBeGreaterThanOrEqual(254);
      const wb = new Workbook();
      const ws = wb.addSheet('Data');
      ws.cell('A1').value = 'Amount';
      ws.cell('A2').value = 1;
      ws.addTable(table(name, 'A1:A3', [{ name: 'Amount', totalsRowFunction: 'sum' }]));
      const names = new Set<string>();
      // Ten clones reach the two-digit suffix `_11`.
      for (let i = 1; i <= 10; i++) {
        const copy = wb.cloneSheet(0, `Copy${i}`);
        const clone = copy.tables[0];
        const displayName = clone?.displayName ?? '';
        expect(displayName.length).toBeLessThanOrEqual(255);
        expect(displayName).toMatch(new RegExp(`_${i + 1}$`));
        expect(LONE_SURROGATE.test(displayName)).toBe(false);
        expect(clone?.name).toBe(displayName);
        expect(copy.cell('A3').formula).toBe(`SUBTOTAL(109,${displayName}[Amount])`);
        names.add(displayName.toLowerCase());
      }
      expect(names.size).toBe(10);
      expect(validationErrors(wb)).toEqual([]);
    });
  }

  it('puts "_" before a cut name that would read as an R1C1 reference', () => {
    // A name only because a `.` follows its R<n> number; the cut drops the `.`.
    const name = `R${'0'.repeat(251)}1.a`;
    expect(name.length).toBe(255);
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.cell('A2').value = 1;
    ws.addTable(table(name, 'A1:A3', [{ name: 'Amount', totalsRowFunction: 'sum' }]));
    const expected = [`_R${'0'.repeat(251)}_2`, `_R${'0'.repeat(251)}_3`];
    for (const [i, want] of expected.entries()) {
      const copy = wb.cloneSheet(0, `Copy${i + 1}`);
      const clone = copy.tables[0];
      expect(clone?.displayName).toBe(want);
      expect(want.length).toBe(255);
      expect(clone?.name).toBe(want);
      expect(copy.cell('A3').formula).toBe(`SUBTOTAL(109,${want}[Amount])`);
      // addTable takes the clone name in a workbook of its own.
      const fresh = new Workbook().addSheet('S');
      expect(() => fresh.addTable(table(want, 'A1:A3', [{ name: 'Amount' }]))).not.toThrow();
    }
    expect(validationErrors(wb)).toEqual([]);
  });
});

describe('cloneSheet picks a clone table name free under canonical equivalence', () => {
  // `café` with U+00E9 and `cafe` + U+0301 are one name to Excel.
  for (const [label, source, existing, want] of [
    ['a precomposed source beside a decomposed _2', 'café', 'café_2', 'café_3'],
    ['a decomposed source beside a precomposed _2', 'café', 'CAFÉ_2', 'café_3'],
  ] as const) {
    it(`skips the taken name for ${label}`, () => {
      const wb = new Workbook();
      const ws = wb.addSheet('Data');
      ws.cell('A1').value = 'Amount';
      ws.cell('A2').value = 1;
      ws.addTable(table(source, 'A1:A3', [{ name: 'Amount', totalsRowFunction: 'sum' }]));
      wb.addSheet('Other').addTable(table(existing, 'A1:A3', [{ name: 'Amount' }], 2));
      const copy = wb.cloneSheet(0, 'Copy');
      expect(copy.tables[0]?.displayName).toBe(want);
      expect(copy.tables[0]?.name).toBe(want);
      expect(copy.cell('A3').formula).toBe(`SUBTOTAL(109,${want}[Amount])`);
    });
  }

  it('skips a defined name equivalent to the first candidate', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Amount';
    ws.addTable(table('café', 'A1:A3', [{ name: 'Amount' }]));
    wb.addNamedRange('café_2', 'Data!$B$1');
    expect(wb.cloneSheet(0, 'Copy').tables[0]?.displayName).toBe('café_3');
  });

  /**
   * Excel 365 (build 20527, 2026-11-13) reads `cafe` + U+0301 `[Qty]` and
   * `CAFE` + U+0301 `[Qty]` as table `café` (U+00E9), and its own copy of the
   * sheet points all four spellings at the copy's table, as the clone does.
   */
  it('points a table reference spelled in another normal form at the clone table', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Data');
    ws.cell('A1').value = 'Qty';
    ws.cell('A2').value = 1;
    const spellings = ['cafe\u{301}', 'caf\u{E9}', 'CAF\u{C9}', 'CAFE\u{301}'];
    spellings.forEach((name, i) => {
      ws.cell(`C${i + 2}`).formula = `SUM(${name}[Qty])`;
    });
    ws.addTable(table('caf\u{E9}', 'A1:A3', [{ name: 'Qty' }]));
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.tables[0]?.displayName).toBe('caf\u{E9}_2');
    expect(spellings.map((_, i) => copy.cell(`C${i + 2}`).formula)).toEqual(
      spellings.map(() => 'SUM(caf\u{E9}_2[Qty])'),
    );
    expect(spellings.map((_, i) => ws.cell(`C${i + 2}`).formula)).toEqual(
      spellings.map((name) => `SUM(${name}[Qty])`),
    );
  });
});

describe('cloneSheet rewrites long formulas in linear time', () => {
  it('rewrites a formula with 20,000 table references', () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Src');
    ws.cell('A1').value = 'a';
    ws.addTable(table('T_1', 'A1:A3', [{ name: 'a', totalsRowFunction: 'sum' }]));
    const refs = 20_000;
    ws.cell('B1').formula = Array.from({ length: refs }, () => 'T_1[a]').join('+');
    const copy = wb.cloneSheet(0, 'Copy');
    expect(copy.cell('B1').formula).toBe(Array.from({ length: refs }, () => 'T_1_2[a]').join('+'));
    expect(ws.cell('B1').formula).toBe(Array.from({ length: refs }, () => 'T_1[a]').join('+'));
  });
});
