/**
 * Rich text builder — creates rich text runs for inline strings.
 *
 * Rich text allows mixed formatting within a single cell, where different
 * portions of text can have different fonts, colors, and styles.
 *
 * This is a Pro-only feature in SheetJS — available for free in modern-xlsx.
 */

import type { RichTextRun } from './types.js';

/**
 * Fluent builder for constructing rich text runs with mixed formatting.
 *
 * Each method appends a styled text segment and returns `this` for chaining.
 * Call {@link build} to produce the final `RichTextRun[]` array.
 *
 * @example
 * ```ts
 * const runs = new RichTextBuilder()
 *   .text('Hello ')
 *   .bold('World')
 *   .colored('!', 'FF0000')
 *   .build();
 * ws.cell('A1').richText = runs;
 * ```
 */
export class RichTextBuilder {
  private readonly runs: RichTextRun[] = [];

  /**
   * Add plain (unstyled) text.
   *
   * @param str - The text content.
   * @returns `this` for chaining.
   */
  text(str: string): this {
    this.runs.push({ text: str });
    return this;
  }

  /**
   * Add bold text.
   *
   * @param str - The text content to render bold.
   * @returns `this` for chaining.
   */
  bold(str: string): this {
    this.runs.push({ text: str, bold: true });
    return this;
  }

  /**
   * Add italic text.
   *
   * @param str - The text content to render italic.
   * @returns `this` for chaining.
   */
  italic(str: string): this {
    this.runs.push({ text: str, italic: true });
    return this;
  }

  /**
   * Add bold + italic text.
   *
   * @param str - The text content to render bold and italic.
   * @returns `this` for chaining.
   */
  boldItalic(str: string): this {
    this.runs.push({ text: str, bold: true, italic: true });
    return this;
  }

  /**
   * Add underlined text.
   *
   * @param str - The text content to underline.
   * @returns `this` for chaining.
   */
  underline(str: string): this {
    this.runs.push({ text: str, underline: true });
    return this;
  }

  /**
   * Add strikethrough text.
   *
   * @param str - The text content to strike through.
   * @returns `this` for chaining.
   */
  strike(str: string): this {
    this.runs.push({ text: str, strike: true });
    return this;
  }

  /**
   * Add colored text.
   *
   * @param str - The text content.
   * @param color - Hex RGB color code (e.g., `'FF0000'` for red).
   * @returns `this` for chaining.
   */
  colored(str: string, color: string): this {
    this.runs.push({ text: str, color });
    return this;
  }

  /**
   * Add text with custom styling options.
   *
   * @param str - The text content.
   * @param opts - Formatting options (bold, italic, underline, strike, fontName, fontSize, color).
   * @returns `this` for chaining.
   *
   * @example
   * ```ts
   * builder.styled('Custom', { fontName: 'Courier', fontSize: 16, color: '0000FF' });
   * ```
   */
  styled(
    str: string,
    opts: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strike?: boolean;
      fontName?: string;
      fontSize?: number;
      color?: string;
    },
  ): this {
    this.runs.push({ text: str, ...opts });
    return this;
  }

  /**
   * Build and return the array of rich text runs.
   *
   * @returns An immutable array of rich text runs ready for `cell.richText`.
   */
  build(): readonly RichTextRun[] {
    return Array.from(this.runs);
  }

  /**
   * Get the plain text content (all runs concatenated, no formatting).
   *
   * @returns The concatenated plain text of all runs.
   */
  plainText(): string {
    return this.runs.map((r) => r.text).join('');
  }
}
