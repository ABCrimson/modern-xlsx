/**
 * Rich text builder — creates rich text runs for inline strings.
 *
 * Rich text allows mixed formatting within a single cell, where different
 * portions of text can have different fonts, colors, and styles.
 *
 * This is a Pro-only feature in SheetJS — available for free in modern-xlsx.
 */

import type { RichTextRun } from './types.js';

export class RichTextBuilder {
  private readonly runs: RichTextRun[] = [];

  /** Add plain (unstyled) text. */
  text(str: string): this {
    this.runs.push({ text: str });
    return this;
  }

  /** Add bold text. */
  bold(str: string): this {
    this.runs.push({ text: str, bold: true });
    return this;
  }

  /** Add italic text. */
  italic(str: string): this {
    this.runs.push({ text: str, italic: true });
    return this;
  }

  /** Add bold + italic text. */
  boldItalic(str: string): this {
    this.runs.push({ text: str, bold: true, italic: true });
    return this;
  }

  /** Add colored text. */
  colored(str: string, color: string): this {
    this.runs.push({ text: str, color });
    return this;
  }

  /** Add text with custom styling options. */
  styled(
    str: string,
    opts: {
      bold?: boolean;
      italic?: boolean;
      fontName?: string;
      fontSize?: number;
      color?: string;
    },
  ): this {
    this.runs.push({ text: str, ...opts });
    return this;
  }

  /** Build and return the array of rich text runs. */
  build(): RichTextRun[] {
    return [...this.runs];
  }

  /** Get the plain text content (all runs concatenated). */
  plainText(): string {
    return this.runs.map((r) => r.text).join('');
  }
}
