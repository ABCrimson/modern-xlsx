import { columnToLetter } from '../../src/format/cell-ref.js';
import type { SheetProtectionData, Worksheet } from '../../src/index.js';

/**
 * Fills the A1-anchored block `rows` x `cols` with numbers: the cell at 1-based
 * row `r` and 0-based column `c` gets `r * (c + 1)`.
 */
export function fillGrid(ws: Worksheet, rows: number, cols: number): void {
  for (let r = 1; r <= rows; r++) {
    for (let c = 0; c < cols; c++) {
      ws.cell(`${columnToLetter(c)}${r}`).value = r * (c + 1);
    }
  }
}

/** A protected sheet with every permission flag switched off. Treat it as read-only. */
export const SHEET_PROTECTION_ALL_OFF: Readonly<SheetProtectionData> = {
  sheet: true,
  objects: false,
  scenarios: false,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
};
