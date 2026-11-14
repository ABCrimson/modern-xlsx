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
} from '../model/types.js';

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
  #fontData: Partial<FontData> = {};
  #fillData: Partial<FillData> = {};
  #borderData: Partial<BorderData> = {};
  #numFmtCode: string | null = null;
  #alignmentData: Partial<AlignmentData> | null = null;
  #protectionData: Partial<ProtectionData> | null = null;

  /**
   * Set font properties (name, size, bold, italic, color, etc.).
   *
   * {@link build} writes `name`, `size`, `bold`, `italic`, `underline`, `strike` and `color`,
   * with `'Calibri'`, 11, `false` and no colour as the defaults; the other `FontData` members
   * (`underlineStyle`, `colorSpec`, `vertAlign`, ...) are not carried into the built font.
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
    Object.assign(this.#fontData, opts);
    return this;
  }

  /**
   * Set fill properties (pattern type, foreground color, background color).
   *
   * @param opts - Fill options with pattern type and colors as hex RGB (RRGGBB) or ARGB (AARRGGBB) strings, with or without a leading #, written as uppercase ARGB (FF alpha added for RRGGBB) and read back in that form.
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.fill({ pattern: 'solid', fgColor: '4472C4' });
   * ```
   */
  fill(opts: { pattern?: PatternType; fgColor?: string | null; bgColor?: string | null }): this {
    if (opts.pattern !== undefined) this.#fillData.patternType = opts.pattern;
    if (opts.fgColor !== undefined) this.#fillData.fgColor = opts.fgColor;
    if (opts.bgColor !== undefined) this.#fillData.bgColor = opts.bgColor;
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
      this.#borderData.left = { style: opts.left.style, color: opts.left.color ?? null };
    if (opts.right)
      this.#borderData.right = { style: opts.right.style, color: opts.right.color ?? null };
    if (opts.top) this.#borderData.top = { style: opts.top.style, color: opts.top.color ?? null };
    if (opts.bottom)
      this.#borderData.bottom = { style: opts.bottom.style, color: opts.bottom.color ?? null };
    return this;
  }

  /**
   * Set alignment properties (horizontal, vertical, wrap text, rotation, indent).
   *
   * {@link build} writes `horizontal`, `vertical`, `wrapText`, `textRotation`, `indent` and
   * `shrinkToFit`; `readingOrder`, `justifyLastLine` and `relativeIndent` are not carried over.
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
    this.#alignmentData ??= {};
    Object.assign(this.#alignmentData, opts);
    return this;
  }

  /**
   * Set cell protection properties (locked, hidden).
   *
   * An unset `locked` is written as `true` and an unset `hidden` as `false`.
   *
   * @param opts - Protection options.
   * @returns `this` for chaining.
   */
  protection(opts: Partial<ProtectionData>): this {
    this.#protectionData ??= {};
    Object.assign(this.#protectionData, opts);
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
    this.#numFmtCode = code;
    return this;
  }

  /**
   * Build the style and register it in the given styles data object.
   *
   * Equal fonts, fills, borders and cell formats are reused, so identical builds return the
   * same index; an entry that differs from every existing one is appended.
   * A number format code already in `styles.numFmts` is reused; a new code gets the id after
   * the highest existing one (at least 164). Without `numberFormat` the format is General (id 0).
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
    const fontId = pushFont(styles, this.#fontData);
    const fillId = pushFill(styles, this.#fillData);
    const borderId = pushBorder(styles, this.#borderData);
    const numFmtId = resolveNumFmtId(styles, this.#numFmtCode);

    const xf: CellXfData = { numFmtId, fontId, fillId, borderId };

    if (this.#alignmentData) {
      xf.alignment = {
        horizontal: this.#alignmentData.horizontal ?? null,
        vertical: this.#alignmentData.vertical ?? null,
        wrapText: this.#alignmentData.wrapText ?? false,
        textRotation: this.#alignmentData.textRotation ?? null,
        indent: this.#alignmentData.indent ?? null,
        shrinkToFit: this.#alignmentData.shrinkToFit ?? false,
      };
      xf.applyAlignment = true;
    }

    if (this.#protectionData) {
      xf.protection = {
        locked: this.#protectionData.locked ?? true,
        hidden: this.#protectionData.hidden ?? false,
      };
      xf.applyProtection = true;
    }

    return intern(styles.cellXfs, internMaps(styles).xfs, xf);
  }
}

/** JSON with object keys sorted recursively, so equal values give equal keys whatever their key order. */
function canonical(v: unknown): string {
  return JSON.stringify(v, (_key, value: unknown) => {
    if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
      const sorted: Record<string, unknown> = {};
      for (const k of Object.keys(value).sort()) sorted[k] = (value as Record<string, unknown>)[k];
      return sorted;
    }
    return value;
  });
}

interface InternMaps {
  fonts: Map<string, number>;
  fills: Map<string, number>;
  borders: Map<string, number>;
  xfs: Map<string, number>;
}

const INTERNED = new WeakMap<StylesData, InternMaps>();

function seed(list: readonly unknown[]): Map<string, number> {
  const map = new Map<string, number>();
  list.forEach((item, idx) => {
    const key = canonical(item);
    if (!map.has(key)) map.set(key, idx);
  });
  return map;
}

/** Key maps for one `StylesData`, seeded lazily from its existing entries (first occurrence wins). */
function internMaps(styles: StylesData): InternMaps {
  let maps = INTERNED.get(styles);
  if (!maps) {
    maps = {
      fonts: seed(styles.fonts),
      fills: seed(styles.fills),
      borders: seed(styles.borders),
      xfs: seed(styles.cellXfs),
    };
    INTERNED.set(styles, maps);
  }
  return maps;
}

/** Index of an equal entry in `list`, or the index of `item` after appending it. A cached index is reused only while the entry still there equals the key (entries can be edited in place). */
function intern<T>(list: T[], map: Map<string, number>, item: T): number {
  const key = canonical(item);
  const idx = map.get(key);
  if (idx !== undefined && idx < list.length && canonical(list[idx]) === key) return idx;
  list.push(item);
  const id = list.length - 1;
  map.set(key, id);
  return id;
}

function pushFont(styles: StylesData, partial: Partial<FontData>): number {
  const font: FontData = {
    name: partial.name ?? 'Calibri',
    size: partial.size ?? 11,
    bold: partial.bold ?? false,
    italic: partial.italic ?? false,
    underline: partial.underline ?? false,
    strike: partial.strike ?? false,
    color: partial.color ?? null,
  };
  return intern(styles.fonts, internMaps(styles).fonts, font);
}

function pushFill(styles: StylesData, partial: Partial<FillData>): number {
  const fill: FillData = {
    patternType: partial.patternType ?? 'none',
    fgColor: partial.fgColor ?? null,
    bgColor: partial.bgColor ?? null,
  };
  return intern(styles.fills, internMaps(styles).fills, fill);
}

function pushBorder(styles: StylesData, partial: Partial<BorderData>): number {
  const border: BorderData = {
    left: partial.left ?? null,
    right: partial.right ?? null,
    top: partial.top ?? null,
    bottom: partial.bottom ?? null,
  };
  return intern(styles.borders, internMaps(styles).borders, border);
}

function resolveNumFmtId(styles: StylesData, code: string | null): number {
  if (!code) return 0; // General

  // Custom number formats start at 164
  const existing = styles.numFmts.find((f) => f.formatCode === code);
  if (existing) return existing.id;

  let maxId = 163;
  for (const f of styles.numFmts) if (f.id > maxId) maxId = f.id;
  const id = maxId + 1;
  styles.numFmts.push({ id, formatCode: code });
  return id;
}
