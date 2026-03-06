/**
 * Fluent builder for Excel header/footer format strings.
 *
 * Excel headers/footers use special `&`-prefixed codes for positioning,
 * dynamic content, and font formatting. This builder provides a type-safe
 * API for constructing these format strings.
 *
 * @example
 * ```typescript
 * const hf = new HeaderFooterBuilder()
 *   .left(`Printed: ${HeaderFooterBuilder.date()}`)
 *   .center(HeaderFooterBuilder.bold('Confidential'))
 *   .right(`Page ${HeaderFooterBuilder.pageNumber()} of ${HeaderFooterBuilder.totalPages()}`)
 *   .build();
 * ```
 */
export class HeaderFooterBuilder {
  private parts: string[] = [];

  /**
   * Add left-aligned content.
   *
   * @param text - The text or format codes for the left section.
   * @returns `this` for chaining.
   */
  left(text: string): this {
    this.parts.push(`&L${text}`);
    return this;
  }

  /**
   * Add center-aligned content.
   *
   * @param text - The text or format codes for the center section.
   * @returns `this` for chaining.
   */
  center(text: string): this {
    this.parts.push(`&C${text}`);
    return this;
  }

  /**
   * Add right-aligned content.
   *
   * @param text - The text or format codes for the right section.
   * @returns `this` for chaining.
   */
  right(text: string): this {
    this.parts.push(`&R${text}`);
    return this;
  }

  /**
   * Build the final format string.
   *
   * @returns The complete header/footer format string for Excel.
   */
  build(): string {
    return this.parts.join('');
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
   * @param size - The font size in points.
   * @returns The font size format code string.
   */
  static fontSize(size: number): string {
    return `&${size}`;
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
   * @param rgb - The hex RGB color code (without `#`).
   * @returns The color format code string.
   */
  static color(rgb: string): string {
    return `&K${rgb}`;
  }
}
