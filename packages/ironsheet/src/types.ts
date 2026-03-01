export type DateSystem = 'date1900' | 'date1904';

export type CellType = 'sharedString' | 'number' | 'boolean' | 'error' | 'formulaStr' | 'inlineStr';

export interface CellData {
  reference: string;
  cellType: CellType;
  styleIndex?: number | null;
  value?: string | null;
  formula?: string | null;
}

export interface RowData {
  index: number;
  cells: CellData[];
  height?: number | null;
  hidden: boolean;
}

export interface FrozenPane {
  xSplit: number;
  ySplit: number;
  topLeftCell: string;
}

export interface ColumnInfo {
  min: number;
  max: number;
  width: number;
  hidden: boolean;
}

export interface WorksheetData {
  dimension?: string | null;
  rows: RowData[];
  mergeCells: string[];
  autoFilter?: string | null;
  frozenPane?: FrozenPane | null;
  columns: ColumnInfo[];
}

export interface NumFmt {
  id: number;
  formatCode: string;
}

export interface FontData {
  name?: string | null;
  size?: number | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color?: string | null;
}

export interface FillData {
  patternType: string;
  fgColor?: string | null;
  bgColor?: string | null;
}

export interface BorderSideData {
  style: string;
  color?: string | null;
}

export interface BorderData {
  left?: BorderSideData | null;
  right?: BorderSideData | null;
  top?: BorderSideData | null;
  bottom?: BorderSideData | null;
}

export interface CellXfData {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
}

export interface StylesData {
  numFmts: NumFmt[];
  fonts: FontData[];
  fills: FillData[];
  borders: BorderData[];
  cellXfs: CellXfData[];
}

export interface SheetData {
  name: string;
  worksheet: WorksheetData;
}

export interface WorkbookData {
  sheets: SheetData[];
  dateSystem: DateSystem;
  styles: StylesData;
  sharedStrings?: { strings: string[] };
}
