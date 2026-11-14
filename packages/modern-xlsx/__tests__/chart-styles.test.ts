import { describe, expect, it } from 'vitest';
import { CHART_STYLE_PALETTES, getChartStylePalette } from '../src/index.js';
import {
  CHART_STYLE_PALETTES as LITE_PALETTES,
  getChartStylePalette as liteGetChartStylePalette,
} from '../src/index-lite.js';

describe('chart style palettes', () => {
  it('style 1 is the Office default palette', () => {
    expect(getChartStylePalette(1)).toEqual([
      '4472C4',
      'ED7D31',
      'A5A5A5',
      'FFC000',
      '5B9BD5',
      '70AD47',
    ]);
  });

  it('bundles exactly styles 1-8', () => {
    expect([...CHART_STYLE_PALETTES.keys()]).toEqual([1, 2, 3, 4, 5, 6, 7, 8]);
  });

  it('every palette is non-empty uppercase 6-digit RGB hex', () => {
    for (const [, palette] of CHART_STYLE_PALETTES) {
      expect(palette.length).toBeGreaterThan(0);
      for (const colour of palette) expect(colour).toMatch(/^[0-9A-F]{6}$/);
    }
  });

  it('getChartStylePalette returns the map entry for every bundled style', () => {
    for (const [id, palette] of CHART_STYLE_PALETTES) {
      expect(getChartStylePalette(id)).toBe(palette);
    }
  });

  it('throws RangeError for c:style ids with no bundled palette (9 and 48)', () => {
    for (const id of [9, 48]) {
      expect(() => getChartStylePalette(id)).toThrow(RangeError);
    }
  });

  it('throws RangeError for ids outside 1-48 and non-integers', () => {
    for (const id of [0, 49, 1.5, Number.NaN]) {
      expect(() => getChartStylePalette(id)).toThrow(RangeError);
    }
  });

  it('index and index-lite share the same exports', () => {
    expect(liteGetChartStylePalette).toBe(getChartStylePalette);
    expect(LITE_PALETTES).toBe(CHART_STYLE_PALETTES);
  });
});
