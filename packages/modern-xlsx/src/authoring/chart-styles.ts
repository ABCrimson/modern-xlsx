/**
 * Colour palettes bundled for chart styles 1-8. `c:style` accepts 1-48; styles 9-48 have no
 * bundled palette.
 */
export const CHART_STYLE_PALETTES: ReadonlyMap<number, readonly string[]> = new Map([
  // Style 1: Office Default
  [1, ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']],
  // Style 2: Colorful
  [2, ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47']],
  // Style 3-6: Monochrome variants
  [3, ['4472C4', '5B9BD5', '8FAADC', 'B4C7E7', 'D6DCE4', 'BDD7EE']],
  [4, ['ED7D31', 'F4B183', 'F8CBAD', 'FCE4D6', 'FBE5D6', 'F8CBAD']],
  [5, ['A5A5A5', 'C0C0C0', 'D0D0D0', 'E0E0E0', 'EDEDED', 'F2F2F2']],
  [6, ['70AD47', '92D050', 'A9D18E', 'C5E0B4', 'E2EFDA', 'A9D18E']],
  // Style 7-8: darker variants
  [7, ['2F5496', 'BF8F00', '808080', 'D97600', '2E75B6', '548235']],
  [8, ['1F4E79', '997300', '595959', 'BF5B00', '1F75A6', '375623']],
]);

/**
 * Get the colour palette for a chart style ID.
 *
 * @param styleId - The chart style ID (1-8 have bundled palettes).
 * @returns The bundled palette for `styleId`.
 * @throws {RangeError} If no palette is bundled for `styleId` (only 1-8 are defined).
 */
export function getChartStylePalette(styleId: number): readonly string[] {
  const palette = CHART_STYLE_PALETTES.get(styleId);
  if (palette === undefined) {
    throw new RangeError(`no bundled palette for chart style ${styleId}; only 1-8 are defined`);
  }
  return palette;
}
