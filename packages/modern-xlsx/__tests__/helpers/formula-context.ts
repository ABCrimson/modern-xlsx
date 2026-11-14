import { columnToLetter } from '../../src/format/cell-ref.js';
import { createDefaultFunctions } from '../../src/formula/functions/index.js';
import type { CellValue } from '../../src/index.js';
import { evaluateFormula } from '../../src/index.js';

const functions = createDefaultFunctions();

/**
 * Evaluates `formula` on 'Sheet1' against a sparse cell map keyed 'Sheet!A1'
 * (sheet name, `columnToLetter` of the 0-based column, then the row exactly as
 * the evaluator passes it). Unlisted cells read as `null`.
 */
export function ev(formula: string, cells?: Record<string, CellValue>): CellValue {
  return evaluateFormula(formula, {
    currentSheet: 'Sheet1',
    functions,
    getCell: (sheet: string, col: number, row: number): CellValue =>
      cells?.[`${sheet}!${columnToLetter(col)}${row}`] ?? null,
  });
}
