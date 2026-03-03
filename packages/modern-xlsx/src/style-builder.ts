import type {
  AlignmentData,
  BorderData,
  BorderStyle,
  CellXfData,
  FillData,
  FontData,
  PatternType,
  ProtectionData,
  StylesData,
} from './types.js';

/**
 * Fluent builder for constructing cell styles (font, fill, border, alignment, etc.).
 *
 * Chain setter methods then call {@link build} to register the style
 * in a `StylesData` object and obtain a reusable style index.
 */
export class StyleBuilder {
  private fontData: Partial<FontData> = {};
  private fillData: Partial<FillData> = {};
  private borderData: Partial<BorderData> = {};
  private numFmtCode: string | null = null;
  private alignmentData: Partial<AlignmentData> | null = null;
  private protectionData: Partial<ProtectionData> | null = null;

  /** Set font properties (name, size, bold, italic, color, etc.). */
  font(opts: Partial<FontData>): this {
    Object.assign(this.fontData, opts);
    return this;
  }

  /** Set fill properties (pattern type, foreground color, background color). */
  fill(opts: { pattern?: PatternType; fgColor?: string | null; bgColor?: string | null }): this {
    if (opts.pattern !== undefined) this.fillData.patternType = opts.pattern;
    if (opts.fgColor !== undefined) this.fillData.fgColor = opts.fgColor;
    if (opts.bgColor !== undefined) this.fillData.bgColor = opts.bgColor;
    return this;
  }

  /** Set border styles and colors for each side (left, right, top, bottom). */
  border(opts: {
    left?: { style: BorderStyle; color?: string | null };
    right?: { style: BorderStyle; color?: string | null };
    top?: { style: BorderStyle; color?: string | null };
    bottom?: { style: BorderStyle; color?: string | null };
  }): this {
    if (opts.left)
      this.borderData.left = { style: opts.left.style, color: opts.left.color ?? null };
    if (opts.right)
      this.borderData.right = { style: opts.right.style, color: opts.right.color ?? null };
    if (opts.top) this.borderData.top = { style: opts.top.style, color: opts.top.color ?? null };
    if (opts.bottom)
      this.borderData.bottom = { style: opts.bottom.style, color: opts.bottom.color ?? null };
    return this;
  }

  /** Set alignment properties (horizontal, vertical, wrap text, rotation, indent). */
  alignment(opts: Partial<AlignmentData>): this {
    if (!this.alignmentData) this.alignmentData = {};
    Object.assign(this.alignmentData, opts);
    return this;
  }

  /** Set cell protection properties (locked, hidden). */
  protection(opts: Partial<ProtectionData>): this {
    if (!this.protectionData) this.protectionData = {};
    Object.assign(this.protectionData, opts);
    return this;
  }

  /** Set a custom number format code (e.g. `"#,##0.00"`, `"yyyy-mm-dd"`). */
  numberFormat(code: string): this {
    this.numFmtCode = code;
    return this;
  }

  /**
   * Build the style and add it to the styles data. Returns the style index
   * that can be assigned to cell.styleIndex.
   */
  build(styles: StylesData): number {
    const fontId = pushFont(styles, this.fontData);
    const fillId = pushFill(styles, this.fillData);
    const borderId = pushBorder(styles, this.borderData);
    const numFmtId = resolveNumFmtId(styles, this.numFmtCode);

    const xf: CellXfData = { numFmtId, fontId, fillId, borderId };

    if (this.alignmentData) {
      xf.alignment = {
        horizontal: this.alignmentData.horizontal ?? null,
        vertical: this.alignmentData.vertical ?? null,
        wrapText: this.alignmentData.wrapText ?? false,
        textRotation: this.alignmentData.textRotation ?? null,
        indent: this.alignmentData.indent ?? null,
        shrinkToFit: this.alignmentData.shrinkToFit ?? false,
      };
      xf.applyAlignment = true;
    }

    if (this.protectionData) {
      xf.protection = {
        locked: this.protectionData.locked ?? true,
        hidden: this.protectionData.hidden ?? false,
      };
      xf.applyProtection = true;
    }

    const styleIndex = styles.cellXfs.length;
    styles.cellXfs.push(xf);
    return styleIndex;
  }
}

function pushFont(styles: StylesData, partial: Partial<FontData>): number {
  const font: FontData = {
    name: partial.name ?? 'Aptos',
    size: partial.size ?? 11,
    bold: partial.bold ?? false,
    italic: partial.italic ?? false,
    underline: partial.underline ?? false,
    strike: partial.strike ?? false,
    color: partial.color ?? null,
  };
  const id = styles.fonts.length;
  styles.fonts.push(font);
  return id;
}

function pushFill(styles: StylesData, partial: Partial<FillData>): number {
  const fill: FillData = {
    patternType: partial.patternType ?? 'none',
    fgColor: partial.fgColor ?? null,
    bgColor: partial.bgColor ?? null,
  };
  const id = styles.fills.length;
  styles.fills.push(fill);
  return id;
}

function pushBorder(styles: StylesData, partial: Partial<BorderData>): number {
  const border: BorderData = {
    left: partial.left ?? null,
    right: partial.right ?? null,
    top: partial.top ?? null,
    bottom: partial.bottom ?? null,
  };
  const id = styles.borders.length;
  styles.borders.push(border);
  return id;
}

function resolveNumFmtId(styles: StylesData, code: string | null): number {
  if (!code) return 0; // General

  // Custom number formats start at 164
  const existing = styles.numFmts.find((f) => f.formatCode === code);
  if (existing) return existing.id;

  const maxId =
    styles.numFmts.length > 0 ? styles.numFmts.reduce((m, f) => (f.id > m ? f.id : m), 0) : 163;
  const id = maxId + 1;
  styles.numFmts.push({ id, formatCode: code });
  return id;
}
