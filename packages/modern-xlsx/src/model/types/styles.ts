import type { BorderStyle, ColorData, PatternType } from './cells.js';

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

/** A number format the workbook defines (`<numFmt>`), referenced by its ID. */
export interface NumFmt {
  /**
   * The number-format ID that {@link CellXfData.numFmtId} refers to. A workbook's own
   * definition of an ID wins over the built-in format with that ID; `StyleBuilder` gives a
   * new format the ID after the highest defined one, at least 164.
   */
  id: number;
  /** The Excel format code string (e.g., `'#,##0.00'`, `'yyyy-mm-dd'`). */
  formatCode: string;
}

/** Cell text alignment properties. */
export interface AlignmentData {
  /** Horizontal alignment. */
  horizontal?:
    | 'general'
    | 'left'
    | 'center'
    | 'right'
    | 'fill'
    | 'justify'
    | 'centerContinuous'
    | 'distributed'
    | null;
  /** Vertical alignment. */
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed' | null;
  /** Whether text wraps within the cell. */
  wrapText?: boolean;
  /** Text rotation in degrees (0-180, or 255 for vertical text). */
  textRotation?: number | null;
  /** Text indent level. */
  indent?: number | null;
  /** Whether text shrinks to fit the cell width. */
  shrinkToFit?: boolean;
  /** Reading order: 0 context-dependent, 1 left-to-right, 2 right-to-left. */
  readingOrder?: number | null;
  /** Justify the last line of distributed text. */
  justifyLastLine?: boolean;
  /** Indent change applied by a differential format. */
  relativeIndent?: number | null;
}

/** Cell protection properties (effective only when sheet protection is enabled). */
export interface ProtectionData {
  /** Whether the cell is locked (prevents editing when sheet is protected). */
  locked: boolean;
  /** Whether the cell formula is hidden from the user. */
  hidden: boolean;
}

/** Font properties for cell text rendering. */
export interface FontData {
  /** Font family name (e.g., `'Calibri'`, `'Arial'`). */
  name: string | null;
  /** Font size in points. */
  size: number | null;
  /** Whether the text is bold. */
  bold: boolean;
  /** Whether the text is italic. */
  italic: boolean;
  /** Whether the text is underlined. */
  underline: boolean;
  /**
   * The underline type when it is not single. Written as `<u val="..."/>`
   * whenever set; with it unset, `underline` writes a single `<u/>`.
   */
  underlineStyle?: 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | null;
  /** Whether the text has strikethrough. */
  strike: boolean;
  /** Font color as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form (e.g., `'FF0000'` for red, read back as `'FFFF0000'`). */
  color: string | null;
  /** The full font colour (theme, indexed, tint or auto); see {@link ColorData} for the precedence rule against `color`. */
  colorSpec?: ColorData | null;
  /** Vertical alignment: baseline, superscript, or subscript. */
  vertAlign?: 'baseline' | 'superscript' | 'subscript' | null;
  /** Font family number (1=Roman, 2=Swiss, 3=Modern, etc.). */
  family?: number | null;
  /** Character set identifier. */
  charset?: number | null;
  /** Font scheme (e.g., `'major'`, `'minor'`). */
  scheme?: string | null;
  /** Condense font spacing. */
  condense?: boolean;
  /** Extend font spacing. */
  extend?: boolean;
  /**
   * The boolean properties written explicitly off (`<b val="0"/>`, `<u val="none"/>`, ...),
   * so a differential format keeps its override instead of inheriting the cell's value.
   * Filled when a dxf font is read; on write each listed property whose flag is off is
   * written explicitly.
   */
  explicitOff?: Array<'b' | 'i' | 'strike' | 'condense' | 'extend' | 'u'>;
}

/** A `<stop>` of a gradient fill: its position and colour. */
export interface GradientStopData {
  position: number;
  /** The `rgb` of the stop colour, or `''` when it has none. */
  color: string;
  /** The full stop colour; see {@link ColorData} for the precedence rule against `color`. */
  colorSpec?: ColorData | null;
}

/** A `<gradientFill>`: the gradient's `type`, angle and colour stops. */
export interface GradientFillData {
  /** The `type` attribute. */
  gradientType?: string | null;
  degree?: number | null;
  stops?: readonly GradientStopData[];
}

/** Cell background fill properties. */
export interface FillData {
  /**
   * Fill pattern type (e.g., `'none'`, `'solid'`, `'gray125'`); `''` when the
   * `patternType` is absent (a gradient fill, or a bare `<patternFill/>`).
   */
  patternType: PatternType | (string & {});
  /** Foreground color as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form. */
  fgColor: string | null;
  /** Background color (used with pattern fills) as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form. */
  bgColor: string | null;
  /** The full foreground colour; see {@link ColorData} for the precedence rule against `fgColor`. */
  fgColorSpec?: ColorData | null;
  /** The full background colour; see {@link ColorData} for the precedence rule against `bgColor`. */
  bgColorSpec?: ColorData | null;
  /** Gradient fill configuration (alternative to pattern fill). */
  gradientFill?: GradientFillData | null;
}

/** A single border edge (left, right, top, or bottom). */
export interface BorderSideData {
  /** The border line style (e.g., `'thin'`, `'medium'`, `'double'`). */
  style: BorderStyle;
  /** Border color as hex RGB (RRGGBB) or ARGB (AARRGGBB), with or without a leading #; written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form, or `null` for automatic. */
  color: string | null;
  /** The full side colour; see {@link ColorData} for the precedence rule against `color`. */
  colorSpec?: ColorData | null;
}

/** Cell border configuration for all four sides plus optional diagonal. */
export interface BorderData {
  /** Left border, or `null` for none. */
  left: BorderSideData | null;
  /** Right border, or `null` for none. */
  right: BorderSideData | null;
  /** Top border, or `null` for none. */
  top: BorderSideData | null;
  /** Bottom border, or `null` for none. */
  bottom: BorderSideData | null;
  /** Diagonal border, or `null` for none. */
  diagonal?: BorderSideData | null;
  /** Whether the diagonal goes from bottom-left to top-right. */
  diagonalUp?: boolean;
  /** Whether the diagonal goes from top-left to bottom-right. */
  diagonalDown?: boolean;
}

/** A cell format (XF) record referencing its font, fill and border by index and its number format by ID. */
export interface CellXfData {
  /**
   * Number-format ID, not an index into `StylesData.numFmts`; `0` is General. An ID that
   * matches a {@link NumFmt.id} names a format the workbook defines, which wins over a
   * built-in format with the same ID; any other ID names a built-in format, whose code
   * `getBuiltinFormat` returns. IDs from 164 up are custom: `Workbook.validate()` reports one
   * without a matching `NumFmt` as an error.
   */
  numFmtId: number;
  /** Index into `StylesData.fonts`. */
  fontId: number;
  /** Index into `StylesData.fills`. */
  fillId: number;
  /** Index into `StylesData.borders`. */
  borderId: number;
  /** Text alignment configuration. */
  alignment?: AlignmentData | null;
  /** Cell protection configuration. */
  protection?: ProtectionData | null;
  /** Index into `StylesData.cellStyleXfs`, the style record this format derives from. */
  xfId?: number | null;
  /** Whether the font from this XF should be applied. */
  applyFont?: boolean;
  /** Whether the fill from this XF should be applied. */
  applyFill?: boolean;
  /** Whether the border from this XF should be applied. */
  applyBorder?: boolean;
  /** Whether the number format from this XF should be applied. */
  applyNumberFormat?: boolean;
  /** Whether the alignment from this XF should be applied. */
  applyAlignment?: boolean;
  /** Whether the protection from this XF should be applied. */
  applyProtection?: boolean;
  /** Whether the cell text is shown as entered, with a leading apostrophe (`quotePrefix`). */
  quotePrefix?: boolean;
  /** Whether the cell shows a pivot table field button (`pivotButton`). */
  pivotButton?: boolean;
}

/**
 * A differential format (`<dxf>`), which conditional-formatting rules and table columns
 * reference by its index in `StylesData.dxfs`.
 */
export interface DxfStyleData {
  font?: FontData | null;
  fill?: FillData | null;
  border?: BorderData | null;
  numFmt?: NumFmt | null;
  alignment?: AlignmentData | null;
  protection?: ProtectionData | null;
}

/** A named cell style (`<cellStyle>`), such as `'Normal'`. */
export interface CellStyleData {
  name: string;
  /** Zero-based index into `StylesData.cellStyleXfs` of the style's format record. */
  xfId: number;
  /** The `builtinId` of a built-in style. */
  builtinId?: number | null;
}

/** The shared style tables for the entire workbook. */
export interface StylesData {
  /** The number formats the workbook defines (`<numFmts>`), looked up by {@link NumFmt.id}. */
  numFmts: NumFmt[];
  /** Font definitions referenced by CellXfData.fontId. */
  fonts: FontData[];
  /** Fill definitions referenced by CellXfData.fillId. */
  fills: FillData[];
  /** Border definitions referenced by CellXfData.borderId. */
  borders: BorderData[];
  /** Cell format records -- each cell's `styleIndex` indexes into this array. */
  cellXfs: CellXfData[];
  /** Differential formatting styles (used by conditional formatting and tables). */
  dxfs?: DxfStyleData[];
  /** Named cell styles (e.g., `'Normal'`, `'Heading 1'`). */
  cellStyles?: CellStyleData[];
  /**
   * The `cellStyleXfs` records that `cellStyles` and `CellXfData.xfId` point at.
   * When empty, the writer emits the single default record Excel expects.
   */
  cellStyleXfs?: CellXfData[];
  /** The source `<tableStyles>` element, verbatim. */
  tableStylesXml?: string | null;
  /** The source `<colors>` element (indexed palette, MRU colours), verbatim. */
  colorsXml?: string | null;
  /** The `<styleSheet>`-level `<extLst>` element, verbatim, with the namespace declarations it inherited. */
  extLstXml?: string | null;
}
