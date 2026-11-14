import { describe, expect, it } from 'vitest';
import { expandSharedFormula } from '../src/formula/shared.js';
import { readBuffer, Workbook } from '../src/index.js';
import { ModernXlsxError } from '../src/model/errors.js';
import type { CellData } from '../src/model/types.js';
import { readZipText } from './helpers/zip.js';

describe('expandSharedFormula', () => {
  // -----------------------------------------------------------------------
  // No offset
  // -----------------------------------------------------------------------
  it('returns same formula when master and child are the same cell', () => {
    expect(expandSharedFormula('A1*2', 'B1', 'B1')).toBe('A1*2');
  });

  // -----------------------------------------------------------------------
  // Row offset only
  // -----------------------------------------------------------------------
  it('shifts relative row refs by row delta', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('A1*2', 'B1', 'B3')).toBe('A3*2');
  });

  // -----------------------------------------------------------------------
  // Column offset only
  // -----------------------------------------------------------------------
  it('shifts relative col refs by col delta', () => {
    // master B1, child D1 → colDelta=2
    expect(expandSharedFormula('A1*2', 'B1', 'D1')).toBe('C1*2');
  });

  // -----------------------------------------------------------------------
  // Both offsets
  // -----------------------------------------------------------------------
  it('shifts both row and col for multiple refs', () => {
    // master B1, child D3 → rowDelta=2, colDelta=2
    expect(expandSharedFormula('A1+B1', 'B1', 'D3')).toBe('C3+D3');
  });

  // -----------------------------------------------------------------------
  // Absolute references preserved
  // -----------------------------------------------------------------------
  it('preserves fully absolute refs', () => {
    // $A$1 should not shift, A1 shifts by rowDelta=2
    expect(expandSharedFormula('$A$1*A1', 'B1', 'B3')).toBe('$A$1*A3');
  });

  // -----------------------------------------------------------------------
  // Mixed absolute references
  // -----------------------------------------------------------------------
  it('handles mixed absolute/relative correctly', () => {
    // $A1: absCol=true (no col shift), absRow=false (row 1→3)
    // A$1: absCol=false (col A stays A, delta=0), absRow=true (no row shift)
    expect(expandSharedFormula('$A1+A$1', 'B1', 'B3')).toBe('$A3+A$1');
  });

  // -----------------------------------------------------------------------
  // Range adjustment
  // -----------------------------------------------------------------------
  it('shifts range references', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('SUM(A1:A10)', 'B1', 'B3')).toBe('SUM(A3:A12)');
  });

  // -----------------------------------------------------------------------
  // Function with mixed refs
  // -----------------------------------------------------------------------
  it('preserves absolute refs in functions', () => {
    // $A$1 stays, B1 shifts to B3
    expect(expandSharedFormula('IF($A$1>0,B1,0)', 'B1', 'B3')).toBe('IF($A$1>0,B3,0)');
  });

  // -----------------------------------------------------------------------
  // Edge case: #REF! for negative row
  // -----------------------------------------------------------------------
  it('produces #REF! when adjusted row goes below 1', () => {
    // master B2, child B1 → rowDelta=-1
    // A1: row 1 + (-1) = 0 < 1 → #REF!
    expect(expandSharedFormula('A1*2', 'B2', 'B1')).toBe('#REF!*2');
  });

  // -----------------------------------------------------------------------
  // Edge case: #REF! for negative column
  // -----------------------------------------------------------------------
  it('produces #REF! when adjusted col goes below A', () => {
    // master B1, child A1 → colDelta=-1
    // A1: col A (0) + (-1) = -1 < 0 → #REF!
    expect(expandSharedFormula('A1*2', 'B1', 'A1')).toBe('#REF!*2');
  });

  // -----------------------------------------------------------------------
  // Complex formula
  // -----------------------------------------------------------------------
  it('handles complex formula with nested functions', () => {
    // master B1, child C2 → rowDelta=1, colDelta=1
    expect(expandSharedFormula('IF(A1>0,SUM(A1:A10),0)', 'B1', 'C2')).toBe(
      'IF(B2>0,SUM(B2:B11),0)',
    );
  });

  // -----------------------------------------------------------------------
  // Sheet-qualified references
  // -----------------------------------------------------------------------
  it('shifts sheet-qualified relative refs', () => {
    // master B1, child B3 → rowDelta=2
    expect(expandSharedFormula('Sheet1!A1+A1', 'B1', 'B3')).toBe('Sheet1!A3+A3');
  });

  // -----------------------------------------------------------------------
  // Parse error returns formula unchanged
  // -----------------------------------------------------------------------
  it('returns original formula on parse error', () => {
    expect(expandSharedFormula('=??invalid', 'B1', 'B3')).toBe('=??invalid');
  });
});

describe('expandSharedFormula sheet edges', () => {
  it('shifting past XFD gives #REF!', () => {
    expect(expandSharedFormula('XFD1', 'A1', 'B1')).toBe('#REF!');
  });

  it('shifts through modern error literals and empty arguments', () => {
    expect(expandSharedFormula('IF(A1,#SPILL!,1)', 'B1', 'B2')).toBe('IF(A2,#SPILL!,1)');
    expect(expandSharedFormula('IF(A1,,2)', 'B1', 'B2')).toBe('IF(A2,,2)');
  });

  it('keeps a structured reference and shifts the cells around it', () => {
    expect(expandSharedFormula('SUM(Table1[Col])+A1', 'B1', 'B2')).toBe('SUM(Table1[Col])+A2');
  });

  it('shifting past row 1048576 gives #REF!', () => {
    expect(expandSharedFormula('A1048576', 'A1', 'A2')).toBe('#REF!');
  });

  it('shifting a cell past row 1048576 from another master gives #REF!', () => {
    expect(expandSharedFormula('A1048576', 'B1', 'B2')).toBe('#REF!');
  });

  it('a range shifted above row 1 becomes one #REF!', () => {
    expect(expandSharedFormula('SUM(A1:A10)', 'B2', 'B1')).toBe('SUM(#REF!)');
    expect(expandSharedFormula('SUM(A1:A10)', 'B12', 'B1')).toBe('SUM(#REF!)');
  });

  it('a range whose end leaves the sheet becomes one #REF!', () => {
    expect(expandSharedFormula('SUM(A2:A1048576)', 'B1', 'B2')).toBe('SUM(#REF!)');
  });

  it('whole-column and whole-row ranges shift their relative parts', () => {
    expect(expandSharedFormula('SUM(A:A)+A1', 'B1', 'C2')).toBe('SUM(B:B)+B2');
    expect(expandSharedFormula('SUM(1:1)', 'A1', 'A3')).toBe('SUM(3:3)');
    expect(expandSharedFormula('SUM($A:B)+SUM($1:2)', 'A1', 'B2')).toBe('SUM($A:C)+SUM($1:3)');
  });

  it('whole-column and whole-row ranges shifted off the sheet become #REF!', () => {
    expect(expandSharedFormula('SUM(XFD:XFD)', 'A1', 'B1')).toBe('SUM(#REF!)');
    expect(expandSharedFormula('SUM(1:1)', 'A2', 'A1')).toBe('SUM(#REF!)');
  });

  it('an in-bounds shift is unchanged', () => {
    expect(expandSharedFormula('A1', 'A1', 'B2')).toBe('B2');
  });

  it('shifts a range with a function endpoint', () => {
    expect(expandSharedFormula('SUM(A2:INDEX(A:A,B1))', 'C1', 'C2')).toBe('SUM(A3:INDEX(A:A,B2))');
  });

  it('shifts an unquoted non-ASCII sheet prefix and writes it quoted', () => {
    expect(expandSharedFormula('Лист2!A1*2', 'B1', 'B2')).toBe("'Лист2'!A2*2");
    expect(expandSharedFormula('Données!A1+売上', 'B1', 'C1')).toBe("'Données'!B1+売上");
  });

  it('an out-of-bounds child address throws INVALID_CELL_REF', () => {
    let caught: unknown;
    try {
      expandSharedFormula('A1', 'A1', 'A1048577');
    } catch (e) {
      caught = e;
    }
    expect(caught).toBeInstanceOf(ModernXlsxError);
    expect((caught as ModernXlsxError).code).toBe('INVALID_CELL_REF');
  });
});

describe('shared-formula children without a cached value', () => {
  it('keep formulaType and sharedIndex through write and read', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('S');
    for (let i = 1; i <= 4; i++) ws.cell(`A${i}`).value = i;
    const rows = wb.toJSON().sheets[0]?.worksheet.rows ?? [];
    rows[0]?.cells.push({
      reference: 'D1',
      cellType: 'number',
      styleIndex: null,
      value: '11',
      formula: 'A1+10',
      formulaType: 'shared',
      formulaRef: 'D1:D4',
      sharedIndex: 0,
    });
    for (let i = 2; i <= 4; i++) {
      rows[i - 1]?.cells.push({
        reference: `D${i}`,
        cellType: 'number',
        styleIndex: null,
        value: i === 3 ? '13' : null,
        formula: null,
        formulaType: 'shared',
        sharedIndex: 0,
      });
    }
    const buf = await wb.toBuffer();
    const xml = readZipText(buf)['xl/worksheets/sheet1.xml'] ?? '';
    expect(xml).toContain('<c r="D2"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/><v>13</v></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/></c>');

    const back = await readBuffer(buf);
    const cells = back.toJSON().sheets[0]?.worksheet.rows.flatMap((r) => r.cells) ?? [];
    for (const ref of ['D2', 'D3', 'D4']) {
      const c = cells.find((x) => x.reference === ref);
      expect(c?.formulaType, ref).toBe('shared');
      expect(c?.sharedIndex, ref).toBe(0);
    }
  });
});

/**
 * Column D of sheet "S" holds shared group 0 as Excel saves a filled-down
 * formula: the master D1 (`A1+10`, ref D1:D4) and the children D2..D4 with
 * cached values, except D3 (a producer that omits it). With `d2Text`, the
 * child D2 carries that formula text as well (`<f t="shared" si="0">text</f>`,
 * which ECMA-376 allows and modern-xlsx 1.2.x wrote after an edit of a child).
 * Written and read back, so the workbook is one read from a file.
 */
async function readSharedGroup(d2Text: string | null = null) {
  const wb = new Workbook();
  const ws = wb.addSheet('S');
  for (let i = 1; i <= 4; i++) ws.cell(`A${i}`).value = i;
  const rows = wb.toJSON().sheets[0]?.worksheet.rows ?? [];
  rows[0]?.cells.push({
    reference: 'D1',
    cellType: 'number',
    styleIndex: null,
    value: '11',
    formula: 'A1+10',
    formulaType: 'shared',
    formulaRef: 'D1:D4',
    sharedIndex: 0,
  });
  for (const [i, value] of [
    [2, '12'],
    [3, null],
    [4, '14'],
  ] as const) {
    rows[i - 1]?.cells.push({
      reference: `D${i}`,
      cellType: 'number',
      styleIndex: null,
      value,
      formula: i === 2 ? d2Text : null,
      formulaType: 'shared',
      sharedIndex: 0,
    });
  }
  const read = await readBuffer(await wb.toBuffer());
  const sheet = read.getSheet('S');
  if (!sheet) throw new Error('sheet S missing');
  return { wb: read, ws: sheet };
}

/** Writes `wb`; returns the sheet XML and the read-back cells of column D. */
async function writeColumnD(wb: Workbook) {
  const buf = await wb.toBuffer();
  const xml = readZipText(buf)['xl/worksheets/sheet1.xml'] ?? '';
  const back = await readBuffer(buf);
  const cells = back.toJSON().sheets[0]?.worksheet.rows.flatMap((r) => r.cells) ?? [];
  const d = (ref: string) => cells.find((c) => c.reference === ref);
  return { xml, d };
}

/** The cell `ref` read back with no formula and no shared keys, holding `value`. */
function expectPlain(cell: CellData | undefined, ref: string, value: string | null) {
  expect(cell?.formula ?? null, ref).toBeNull();
  expect(cell?.formulaType ?? null, ref).toBeNull();
  expect(cell?.formulaRef ?? null, ref).toBeNull();
  expect(cell?.sharedIndex ?? null, ref).toBeNull();
  expect(cell?.value ?? null, ref).toBe(value);
}

describe('a cleared shared formula stays cleared', () => {
  it('the source group is written with its master and children', async () => {
    const { wb } = await readSharedGroup();
    const { xml } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>');
    expect(xml).toContain('<c r="D2"><f t="shared" si="0"/><v>12</v></c>');
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
  });

  it('value = null on one child clears it and leaves the group', async () => {
    const { wb, ws } = await readSharedGroup();
    ws.cell('D2').value = null;
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D2"/>');
    expect(xml).toContain('<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>');
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
    expectPlain(d('D2'), 'D2', null);
    expect(d('D4')?.sharedIndex).toBe(0);
  });

  it('formula = null and value = null on every cell of the group clear it', async () => {
    const { wb, ws } = await readSharedGroup();
    for (let i = 1; i <= 4; i++) {
      const c = ws.cell(`D${i}`);
      c.formula = null;
      c.value = null;
    }
    const { xml, d } = await writeColumnD(wb);
    for (let i = 1; i <= 4; i++) {
      expect(xml).toContain(`<c r="D${i}"/>`);
      expectPlain(d(`D${i}`), `D${i}`, null);
    }
    expect(xml).not.toContain('<f');
  });

  it('a child whose master was removed from the model is written without <f>', async () => {
    const { wb } = await readSharedGroup();
    const row1 = wb.toJSON().sheets[0]?.worksheet.rows.find((r) => r.index === 1);
    const master = row1?.cells.findIndex((c) => c.reference === 'D1') ?? -1;
    expect(master).toBeGreaterThanOrEqual(0);
    row1?.cells.splice(master, 1);
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D2"><v>12</v></c>');
    expect(xml).toContain('<c r="D3"/>');
    expect(xml).toContain('<c r="D4"><v>14</v></c>');
    expect(xml).not.toContain('<f');
    expectPlain(d('D2'), 'D2', '12');
    expectPlain(d('D3'), 'D3', null);
    expectPlain(d('D4'), 'D4', '14');
  });

  it('a cleared master leaves its valued children as plain values', async () => {
    const { wb, ws } = await readSharedGroup();
    ws.cell('D1').formula = null;
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D1"><v>11</v></c>');
    expect(xml).toContain('<c r="D2"><v>12</v></c>');
    expect(xml).toContain('<c r="D3"/>');
    expect(xml).toContain('<c r="D4"><v>14</v></c>');
    expect(xml).not.toContain('<f');
    expectPlain(d('D1'), 'D1', '11');
    expectPlain(d('D2'), 'D2', '12');
    expectPlain(d('D4'), 'D4', '14');
  });
});

describe('formula text or rich text on a shared-formula child', () => {
  it('formula text gives a child a plain formula; the master keeps its group', async () => {
    const { wb, ws } = await readSharedGroup();
    ws.cell('D2').formula = 'A2*5';
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D2"><f>A2*5</f><v>12</v></c>');
    expect(xml).toContain('<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>');
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
    expect(d('D2')?.formula).toBe('A2*5');
    expect(d('D2')?.formulaType ?? null).toBeNull();
    expect(d('D2')?.sharedIndex ?? null).toBeNull();
    expect(d('D4')?.sharedIndex).toBe(0);
  });

  it('formula text on a child keeps it a plain formula when the master is cleared', async () => {
    const { wb, ws } = await readSharedGroup();
    ws.cell('D1').formula = null;
    ws.cell('D2').formula = 'A2*5';
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D2"><f>A2*5</f><v>12</v></c>');
    expect(xml).toContain('<c r="D1"><v>11</v></c>');
    expect(xml).toContain('<c r="D3"/>');
    expect(xml).toContain('<c r="D4"><v>14</v></c>');
    expect(xml).not.toContain('t="shared"');
    expect(d('D2')?.formula).toBe('A2*5');
    expect(d('D2')?.formulaType ?? null).toBeNull();
    expect(d('D2')?.sharedIndex ?? null).toBeNull();
    expectPlain(d('D4'), 'D4', '14');
  });

  it('rich text on a child replaces its formula', async () => {
    const { wb, ws } = await readSharedGroup();
    ws.cell('D2').richText = [{ text: 'rich', bold: true }];
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toMatch(/<c r="D2" t="s"><v>\d+<\/v><\/c>/);
    expect(xml).toContain('<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
    expect(d('D2')?.formula ?? null).toBeNull();
    expect(d('D2')?.formulaType ?? null).toBeNull();
    expect(d('D2')?.sharedIndex ?? null).toBeNull();
    expect(d('D2')?.value).toBe('rich');
    expect(d('D2')?.richText?.[0]?.text).toBe('rich');
  });
});

describe('a shared-formula child that carries its own text', () => {
  const master = '<c r="D1"><f t="shared" ref="D1:D4" si="0">A1+10</f><v>11</v></c>';

  it('is written as read while its master is in the sheet', async () => {
    const { wb } = await readSharedGroup('A2+10');
    const { xml } = await writeColumnD(wb);
    expect(xml).toContain(master);
    expect(xml).toContain('<c r="D2"><f t="shared" si="0">A2+10</f><v>12</v></c>');
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
  });

  it('becomes a plain formula when the master is cleared; the other children keep values', async () => {
    const { wb, ws } = await readSharedGroup('A2+10');
    ws.cell('D1').formula = null;
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D1"><v>11</v></c>');
    expect(xml).toContain('<c r="D2"><f>A2+10</f><v>12</v></c>');
    expect(xml).toContain('<c r="D3"/>');
    expect(xml).toContain('<c r="D4"><v>14</v></c>');
    expect(xml).not.toContain('t="shared"');
    expect(d('D2')?.formula).toBe('A2+10');
    expect(d('D2')?.sharedIndex ?? null).toBeNull();
    expectPlain(d('D4'), 'D4', '14');
  });

  it('new formula text gives it a plain formula of its own', async () => {
    const { wb, ws } = await readSharedGroup('A2+10');
    ws.cell('D2').formula = 'A2*100';
    expect(ws.cell('D2').formula).toBe('A2*100');
    const { xml, d } = await writeColumnD(wb);
    expect(xml).toContain('<c r="D2"><f>A2*100</f><v>12</v></c>');
    expect(xml).toContain(master);
    expect(xml).toContain('<c r="D4"><f t="shared" si="0"/><v>14</v></c>');
    expect(d('D2')?.formula).toBe('A2*100');
    expect(d('D2')?.formulaType ?? null).toBeNull();
    expect(d('D2')?.sharedIndex ?? null).toBeNull();
  });

  it('formula = null or an empty string clears its formula and its shared keys', async () => {
    for (const cleared of [null, ''] as const) {
      const { wb, ws } = await readSharedGroup('A2+10');
      ws.cell('D2').formula = cleared;
      expect(ws.cell('D2').formula, String(cleared)).toBeNull();
      const { xml, d } = await writeColumnD(wb);
      expect(xml).toContain('<c r="D2"><v>12</v></c>');
      expect(xml).toContain(master);
      expectPlain(d('D2'), 'D2', '12');
    }
  });
});
