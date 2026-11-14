import { describe, expect, it } from 'vitest';
import { formatCell } from '../src/index.js';

const ZEROS_101 = '0'.repeat(101);

/** Count the digits after the first '.' (before any non-digit suffix). */
function fractionDigits(text: string): number {
  return /\.(\d*)/.exec(text)?.[1]?.length ?? 0;
}

describe('format-cell decimal guard', () => {
  it('keeps 101 decimals in a plain number format', () => {
    const out = formatCell(1.5, `0.${ZEROS_101}`);
    expect(out).toBe(`1.5${'0'.repeat(100)}`);
    expect(fractionDigits(out)).toBe(101);
  });

  it('keeps 101 decimals in a grouped number format', () => {
    const out = formatCell(1234.5, `#,##0.${ZEROS_101}`);
    expect(out.startsWith('1,234.5')).toBe(true);
    expect(fractionDigits(out)).toBe(101);
  });

  it('keeps 101 decimals in a percent format', () => {
    const out = formatCell(0.5, `0.${ZEROS_101}%`);
    expect(out.endsWith('%')).toBe(true);
    expect(out.startsWith('50.')).toBe(true);
    expect(fractionDigits(out)).toBe(101);
  });

  it('keeps 101 decimals in a scientific format', () => {
    const out = formatCell(1234.5, `0.${ZEROS_101}E+00`);
    expect(out).toMatch(/^1\.2345\d+E\+03$/);
    expect(out.split('E')[0]?.split('.')[1]?.length).toBe(101);
  });
});

describe('format-cell single rounding path', () => {
  it('rounds 1.005 to 1.01 in a plain format, like the grouped one', () => {
    expect(formatCell(1.005, '0.00')).toBe('1.01');
    expect(formatCell(1.005, '#,##0.00')).toBe('1.01');
  });

  it('shows no sign for negative zero but keeps it for a tiny negative', () => {
    expect(formatCell(-0, '0.00')).toBe('0.00');
    expect(formatCell(-0, '#,##0')).toBe('0');
    expect(formatCell(-0.001, '0.00')).toBe('-0.00');
  });
});

describe('format-cell length-safe date tokens', () => {
  it('keeps token positions after a length-changing lower-case character', () => {
    expect(formatCell(45000.75, 'İ yyyy-mm-dd')).toBe('İ 2023-03-15');
    expect(formatCell(45000.75, 'İİyy')).toBe('İİ23');
  });

  it('matches fixed tokens case-insensitively', () => {
    expect(formatCell(45000.75, 'YYYY-MM-DD AM/PM')).toBe('2023-03-15 PM');
  });
});

/** The value `fn` returns and the milliseconds the call took. */
function timed<T>(fn: () => T): { result: T; ms: number } {
  const start = performance.now();
  const result = fn();
  return { result, ms: performance.now() - start };
}

describe('format-cell linear-time numeric parsing', () => {
  it('returns a 32,767-digit string with a rejecting tail unchanged without backtracking', () => {
    const text = `${'1'.repeat(32_767)}x`;
    const { result, ms } = timed(() => formatCell(text, '0.00'));
    expect(result).toBe(text);
    expect(ms).toBeLessThan(50);
  });

  it('scans a 32,767-digit condition with a rejecting tail without backtracking', () => {
    const { ms } = timed(() => formatCell(5, `[>${'1'.repeat(32_767)}x0`));
    expect(ms).toBeLessThan(50);
  });

  it('scans a closed 32,767-digit condition with a rejecting tail without backtracking', () => {
    const { ms } = timed(() => formatCell(5, `[>${'1'.repeat(32_767)}x]0`));
    expect(ms).toBeLessThan(50);
  });

  it('still accepts a trailing decimal point', () => {
    expect(formatCell('1.', '0.00')).toBe('1.00');
  });
});

describe('format-cell linear-time long format codes', () => {
  it('renders a date code of 200,001 tokens without a stack overflow', () => {
    const { result, ms } = timed(() => formatCell(45000.5, `d${'-x'.repeat(100_000)}`));
    expect(result).toBe(`15${'-x'.repeat(100_000)}`);
    expect(ms).toBeLessThan(500);
  });

  it('renders a date code of 1,000,001 tokens, past the worker-thread spread limit', () => {
    expect(formatCell(45000.5, `d${'-x'.repeat(500_000)}`)).toBe(`15${'-x'.repeat(500_000)}`);
  });

  it('classifies 100,000 consecutive commas in linear time', () => {
    const { result, ms } = timed(() => formatCell(5, `0${','.repeat(100_000)}x`));
    expect(result).toBe('0x');
    expect(ms).toBeLessThan(500);
  });

  it('marks 50,000 month runs in linear time', () => {
    const { result, ms } = timed(() => formatCell(45000.5, 'm-'.repeat(50_000)));
    expect(result).toBe('3-'.repeat(50_000));
    expect(ms).toBeLessThan(500);
  });
});
