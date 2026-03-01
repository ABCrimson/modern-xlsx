export type {
  BorderData,
  BorderSideData,
  CellData,
  CellType,
  CellXfData,
  ColumnInfo,
  DateSystem,
  FillData,
  FontData,
  FrozenPane,
  NumFmt,
  RowData,
  SheetData,
  StylesData,
  WorkbookData,
  WorksheetData,
} from './types.js';
export { initWasm } from './wasm-loader.js';
export { Cell, Workbook, Worksheet } from './workbook.js';

export async function readBuffer(data: Uint8Array): Promise<import('./workbook.js').Workbook> {
  const { ensureInitialized, wasmRead } = await import('./wasm-loader.js');
  const { Workbook } = await import('./workbook.js');
  ensureInitialized();
  const raw = wasmRead(data);
  return new Workbook(raw);
}

export const VERSION = '0.1.0';
