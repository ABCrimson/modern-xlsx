/**
 * Fluent builder for Excel header/footer format strings.
 *
 * Excel headers/footers use special `&`-prefixed codes for positioning,
 * dynamic content, and font formatting. This builder provides a type-safe
 * API for constructing these format strings.
 *
 * @example
 * ```ts
 * const hf = new HeaderFooterBuilder()
 *   .left(`Printed: ${HeaderFooterBuilder.date()}`)
 *   .center(HeaderFooterBuilder.bold('Confidential'))
 *   .right(`Page ${HeaderFooterBuilder.pageNumber()} of ${HeaderFooterBuilder.totalPages()}`)
 *   .build();
 * ```
 */
export class HeaderFooterBuilder {
  readonly #parts: string[] = [];

  /**
   * Add left-aligned content.
   *
   * Literal text must go through {@link HeaderFooterBuilder.text} so a literal `&` is escaped.
   *
   * @param text - The format codes and (escaped) text for the left section.
   * @returns `this` for chaining.
   */
  left(text: string): this {
    this.#parts.push(`&L${text}`);
    return this;
  }

  /**
   * Add center-aligned content.
   *
   * Literal text must go through {@link HeaderFooterBuilder.text} so a literal `&` is escaped.
   *
   * @param text - The format codes and (escaped) text for the center section.
   * @returns `this` for chaining.
   */
  center(text: string): this {
    this.#parts.push(`&C${text}`);
    return this;
  }

  /**
   * Add right-aligned content.
   *
   * Literal text must go through {@link HeaderFooterBuilder.text} so a literal `&` is escaped.
   *
   * @param text - The format codes and (escaped) text for the right section.
   * @returns `this` for chaining.
   */
  right(text: string): this {
    this.#parts.push(`&R${text}`);
    return this;
  }

  /**
   * Build the final format string.
   *
   * @returns The complete header/footer format string for Excel.
   * @throws {RangeError} When the string is longer than the 255 characters Excel allows.
   */
  build(): string {
    const str = this.#parts.join('');
    if (str.length > 255) {
      throw new RangeError(`header/footer string is ${str.length} characters; Excel allows 255`);
    }
    return str;
  }

  /**
   * Escape literal text for embedding in a section (`&` becomes `&&`).
   *
   * @param str - The literal text.
   * @returns The text with every `&` doubled.
   */
  static text(str: string): string {
    return str.replaceAll('&', '&&');
  }

  // ---------------------------------------------------------------------------
  // Static code helpers — return format code strings for embedding
  // ---------------------------------------------------------------------------

  /**
   * Current page number (`&P`).
   *
   * @returns The `&P` format code string.
   */
  static pageNumber(): string {
    return '&P';
  }

  /**
   * Total number of pages (`&N`).
   *
   * @returns The `&N` format code string.
   */
  static totalPages(): string {
    return '&N';
  }

  /**
   * Current date (`&D`).
   *
   * @returns The `&D` format code string.
   */
  static date(): string {
    return '&D';
  }

  /**
   * Current time (`&T`).
   *
   * @returns The `&T` format code string.
   */
  static time(): string {
    return '&T';
  }

  /**
   * File name without path (`&F`).
   *
   * @returns The `&F` format code string.
   */
  static fileName(): string {
    return '&F';
  }

  /**
   * Sheet tab name (`&A`).
   *
   * @returns The `&A` format code string.
   */
  static sheetName(): string {
    return '&A';
  }

  /**
   * File path (`&Z`).
   *
   * @returns The `&Z` format code string.
   */
  static filePath(): string {
    return '&Z';
  }

  // ---------------------------------------------------------------------------
  // Font formatting helpers — wrap text with toggle codes
  // ---------------------------------------------------------------------------

  /**
   * Wrap text in bold toggle (`&B...&B`).
   *
   * @param text - The text to make bold.
   * @returns The text wrapped in bold format codes.
   */
  static bold(text: string): string {
    return `&B${text}&B`;
  }

  /**
   * Wrap text in italic toggle (`&I...&I`).
   *
   * @param text - The text to italicize.
   * @returns The text wrapped in italic format codes.
   */
  static italic(text: string): string {
    return `&I${text}&I`;
  }

  /**
   * Wrap text in underline toggle (`&U...&U`).
   *
   * @param text - The text to underline.
   * @returns The text wrapped in underline format codes.
   */
  static underline(text: string): string {
    return `&U${text}&U`;
  }

  /**
   * Wrap text in strikethrough toggle (`&S...&S`).
   *
   * @param text - The text to strike through.
   * @returns The text wrapped in strikethrough format codes.
   */
  static strikethrough(text: string): string {
    return `&S${text}&S`;
  }

  /**
   * Set font size in points (e.g., `&12` for 12pt).
   *
   * @param size - The font size in points, an integer from 1 to 409.
   * @param text - Optional text that follows; a space is inserted when it starts
   *   with a digit so the digits are not read as part of the size.
   * @returns The font size format code string, followed by `text` when given.
   * @throws {RangeError} When `size` is not an integer from 1 to 409.
   */
  static fontSize(size: number, text?: string): string {
    if (!Number.isInteger(size) || size < 1 || size > 409) {
      throw new RangeError('fontSize: size must be an integer from 1 to 409');
    }
    if (text === undefined) return `&${size}`;
    return `&${size}${/^[0-9]/.test(text) ? ' ' : ''}${text}`;
  }

  /**
   * Set font name (e.g., `&"Arial"`).
   *
   * @param name - The font family name.
   * @returns The font name format code string.
   */
  static fontName(name: string): string {
    return `&"${name}"`;
  }

  /**
   * Set text color using hex RGB (e.g., `&KFF0000` for red).
   *
   * @param rgb - Six hex digits (`RRGGBB`) or eight (`AARRGGBB`, alpha dropped), without `#`.
   * @returns The color format code string with the digits in uppercase.
   * @throws {RangeError} When `rgb` is not six or eight hex digits.
   */
  static color(rgb: string): string {
    if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return `&K${rgb.toUpperCase()}`;
    if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return `&K${rgb.slice(2).toUpperCase()}`;
    throw new RangeError('color: expected 6 or 8 hex digits');
  }
}
