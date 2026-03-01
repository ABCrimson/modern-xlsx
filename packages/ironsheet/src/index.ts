export type {
  BorderData,
  BorderSideData,
  BorderStyle,
  CellData,
  CellType,
  CellXfData,
  ColumnInfo,
  DateSystem,
  FillData,
  FontData,
  FrozenPane,
  NumFmt,
  PatternType,
  RowData,
  SheetData,
  StylesData,
  WorkbookData,
  WorksheetData,
} from './types.js';
export { initWasm } from './wasm-loader.js';
export { Cell, Workbook, Worksheet } from './workbook.js';

import { ensureInitialized, wasmRead } from './wasm-loader.js';
import { Workbook } from './workbook.js';

export async function readBuffer(data: Uint8Array): Promise<Workbook> {
  ensureInitialized();
  const raw = wasmRead(data);
  return new Workbook(raw);
}

export const VERSION = '0.1.0' as const;
