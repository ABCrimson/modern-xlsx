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
 *
 * @example
 * ```ts
 * const idx = new StyleBuilder()
 *   .font({ bold: true, color: 'FF0000' })
 *   .fill({ pattern: 'solid', fgColor: 'FFFF00' })
 *   .border({ bottom: { style: 'thin', color: '000000' } })
 *   .build(wb.styles);
 * ws.cell('A1').styleIndex = idx;
 * ```
 */
export class StyleBuilder {
  private fontData: Partial<FontData> = {};
  private fillData: Partial<FillData> = {};
  private borderData: Partial<BorderData> = {};
  private numFmtCode: string | null = null;
  private alignmentData: Partial<AlignmentData> | null = null;
  private protectionData: Partial<ProtectionData> | null = null;

  /**
   * Set font properties (name, size, bold, italic, color, etc.).
   *
   * @param opts - Partial font data to merge into the current font.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.font({ name: 'Arial', size: 14, bold: true });
   * ```
   */
  font(opts: Partial<FontData>): this {
    Object.assign(this.fontData, opts);
    return this;
  }

  /**
   * Set fill properties (pattern type, foreground color, background color).
   *
   * @param opts - Fill options with pattern type and colors as hex RGB strings.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.fill({ pattern: 'solid', fgColor: '4472C4' });
   * ```
   */
  fill(opts: { pattern?: PatternType; fgColor?: string | null; bgColor?: string | null }): this {
    if (opts.pattern !== undefined) this.fillData.patternType = opts.pattern;
    if (opts.fgColor !== undefined) this.fillData.fgColor = opts.fgColor;
    if (opts.bgColor !== undefined) this.fillData.bgColor = opts.bgColor;
    return this;
  }

  /**
   * Set border styles and colors for each side (left, right, top, bottom).
   *
   * @param opts - Border definitions for each side.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.border({
   *   bottom: { style: 'double', color: '000000' },
   *   top: { style: 'thin' },
   * });
   * ```
   */
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

  /**
   * Set alignment properties (horizontal, vertical, wrap text, rotation, indent).
   *
   * @param opts - Alignment options to merge into the current alignment.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.alignment({ horizontal: 'center', wrapText: true });
   * ```
   */
  alignment(opts: Partial<AlignmentData>): this {
    this.alignmentData ??= {};
    Object.assign(this.alignmentData, opts);
    return this;
  }

  /**
   * Set cell protection properties (locked, hidden).
   *
   * @param opts - Protection options.
   * @returns `this` for chaining.
   */
  protection(opts: Partial<ProtectionData>): this {
    this.protectionData ??= {};
    Object.assign(this.protectionData, opts);
    return this;
  }

  /**
   * Set a custom number format code (e.g., `'#,##0.00'`, `'yyyy-mm-dd'`).
   *
   * @param code - The Excel number format code string.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.numberFormat('#,##0.00');
   * ```
   */
  numberFormat(code: string): this {
    this.numFmtCode = code;
    return this;
  }

  /**
   * Build the style and register it in the given styles data object.
   *
   * @param styles - The workbook's shared StylesData (typically `wb.styles`).
   * @returns The zero-based style index to assign to `cell.styleIndex`.
   *
   * @example
   * ```ts
   * const idx = new StyleBuilder()
   *   .font({ bold: true })
   *   .build(wb.styles);
   * ws.cell('A1').styleIndex = idx;
   * ```
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
    styles.numFmts.length > 0
      ? Math.max(
          163,
          styles.numFmts.reduce((m, f) => (f.id > m ? f.id : m), 0),
        )
      : 163;
  const id = maxId + 1;
  styles.numFmts.push({ id, formatCode: code });
  return id;
}
