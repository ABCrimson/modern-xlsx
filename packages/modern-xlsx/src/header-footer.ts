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

  /** Add left-aligned content. */
  left(text: string): this {
    this.parts.push(`&L${text}`);
    return this;
  }

  /** Add center-aligned content. */
  center(text: string): this {
    this.parts.push(`&C${text}`);
    return this;
  }

  /** Add right-aligned content. */
  right(text: string): this {
    this.parts.push(`&R${text}`);
    return this;
  }

  /** Build the final format string. */
  build(): string {
    return this.parts.join('');
  }

  // ---------------------------------------------------------------------------
  // Static code helpers — return format code strings for embedding
  // ---------------------------------------------------------------------------

  /** Current page number (`&P`). */
  static pageNumber(): string {
    return '&P';
  }

  /** Total number of pages (`&N`). */
  static totalPages(): string {
    return '&N';
  }

  /** Current date (`&D`). */
  static date(): string {
    return '&D';
  }

  /** Current time (`&T`). */
  static time(): string {
    return '&T';
  }

  /** File name without path (`&F`). */
  static fileName(): string {
    return '&F';
  }

  /** Sheet tab name (`&A`). */
  static sheetName(): string {
    return '&A';
  }

  /** File path (`&Z`). */
  static filePath(): string {
    return '&Z';
  }

  // ---------------------------------------------------------------------------
  // Font formatting helpers — wrap text with toggle codes
  // ---------------------------------------------------------------------------

  /** Wrap text in bold toggle (`&B...&B`). */
  static bold(text: string): string {
    return `&B${text}&B`;
  }

  /** Wrap text in italic toggle (`&I...&I`). */
  static italic(text: string): string {
    return `&I${text}&I`;
  }

  /** Wrap text in underline toggle (`&U...&U`). */
  static underline(text: string): string {
    return `&U${text}&U`;
  }

  /** Wrap text in strikethrough toggle (`&S...&S`). */
  static strikethrough(text: string): string {
    return `&S${text}&S`;
  }

  /** Set font size in points (e.g. `&12` for 12pt). */
  static fontSize(size: number): string {
    return `&${size}`;
  }

  /** Set font name (e.g. `&"Arial"`). */
  static fontName(name: string): string {
    return `&"${name}"`;
  }

  /** Set text color using hex RGB (e.g. `&KFF0000` for red). */
  static color(rgb: string): string {
    return `&K${rgb}`;
  }
}
