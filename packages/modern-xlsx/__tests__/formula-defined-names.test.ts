import { describe, expect, expectTypeOf, it } from 'vitest';
import {
  createDefaultFunctions,
  type EvalContext,
  encodeCellRef,
  evaluateFormula,
  Workbook,
} from '../src/index.js';

/**
 * The evaluator has no defined-name table and `EvalContext` has no member that
 * supplies one: a name evaluates to `#NAME?` whatever the workbook defines.
 */
function contextFor(wb: Workbook): EvalContext {
  return {
    getCell: (sheet, col, row) =>
      wb.getSheet(sheet)?.cell(encodeCellRef(row - 1, col)).value ?? null,
    currentSheet: 'Data',
    functions: createDefaultFunctions(),
  };
}

function workbookWithMyRange(): Workbook {
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  ws.cell('A1').value = 1;
  ws.cell('A2').value = 2;
  ws.cell('A3').value = 3;
  wb.addNamedRange('MyRange', 'Data!$A$1:$A$3');
  return wb;
}

describe('defined names in formulas', () => {
  it('evaluates =MyRange to #NAME? although the workbook defines MyRange', () => {
    const wb = workbookWithMyRange();
    expect(wb.namedRanges.map((n) => n.name)).toContain('MyRange');
    expect(evaluateFormula('=MyRange', contextFor(wb))).toBe('#NAME?');
  });

  it('propagates #NAME? through =SUM(MyRange)', () => {
    expect(evaluateFormula('=SUM(MyRange)', contextFor(workbookWithMyRange()))).toBe('#NAME?');
  });

  it('evaluates the range the name stands for when it is written out', () => {
    expect(evaluateFormula('=SUM(Data!$A$1:$A$3)', contextFor(workbookWithMyRange()))).toBe(6);
  });

  it('has no EvalContext member through which names are supplied', () => {
    expectTypeOf<EvalContext>().not.toHaveProperty('names');
    const ctx = {
      ...contextFor(workbookWithMyRange()),
      names: new Map([['MyRange', 'Data!$A$1:$A$3']]),
    };
    expect(evaluateFormula('=MyRange', ctx)).toBe('#NAME?');
  });
});
