/**
 * Excel's two number-to-text conversions.
 *
 * - {@link numberToExcelText}: the text a formula produces when it turns a
 *   number into a string (`&`, `LEN`, `LEFT`, `CONCATENATE`, ...).
 * - {@link formatGeneralNumber}: the text the `General` number format
 *   displays, limited to 11 characters.
 *
 * Both round with JavaScript's correctly rounded `toExponential`/`toFixed`
 * and write exponents the Excel way: uppercase `E`, explicit sign, at least
 * two digits (`1E+21`, `1.5E-09`). The rules were taken from Excel 16.0
 * (build 20527) on 2026-11-10: formula text from `=(x)&""`, General text
 * from a CSV export of General cells. LibreOffice's
 * `rtl::str::doubleToString` (`sal/rtl/strtmpl.hxx`,
 * `rtl_math_StringFormat_Automatic`) also rounds to 15 digits but switches
 * to E notation from a decimal exponent of +/-15, which Excel does not.
 *
 * @module number-text
 */

/** Significant digits Excel keeps in a number. */
const EXCEL_PRECISION = 15;

/** Longest plain (non-E) formula text, sign excluded. */
const FORMULA_TEXT_WIDTH = 20;

/** Longest General display text, sign excluded. */
const GENERAL_WIDTH = 11;

/** Significant digits of a positive number with trailing zeros removed, and its decimal exponent. */
interface DecimalDigits {
  /** Digits without leading or trailing zeros (`'0'` for zero). */
  readonly digits: string;
  /** Power of ten of the first digit. */
  readonly exp: number;
}

/** Round the positive finite `abs` to `significant` digits. */
function toDecimalDigits(abs: number, significant: number): DecimalDigits {
  const [mantissa = '0', exp = '0'] = abs.toExponential(significant - 1).split('e');
  return { digits: mantissa.replace('.', '').replace(/0+$/, '') || '0', exp: Number(exp) };
}

/** Write `d` as `d[.ddd]E+NN`, the exponent writer both conversions share. */
function writeScientific({ digits, exp }: DecimalDigits): string {
  const mantissa = digits.length > 1 ? `${digits[0]}.${digits.slice(1)}` : digits;
  return `${mantissa}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/** Write `d` positionally, without an exponent. */
function writePlain({ digits, exp }: DecimalDigits): string {
  if (exp < 0) return `0.${'0'.repeat(-exp - 1)}${digits}`;
  const intDigits = exp + 1;
  return digits.length <= intDigits
    ? digits.padEnd(intDigits, '0')
    : `${digits.slice(0, intDigits)}.${digits.slice(intDigits)}`;
}

/** Drop trailing fractional zeros and a trailing decimal point. */
function stripFractionZeros(text: string): string {
  return text.includes('.') ? text.replace(/\.?0+$/, '') : text;
}

/** Text for a value that is not a finite number (never an Excel cell value). */
function nonFiniteText(n: number): string {
  return Number.isNaN(n) ? '#NUM!' : '#DIV/0!';
}

/**
 * Convert a number to text the way an Excel formula does.
 *
 * Rounds to 15 significant digits, writes `-0` as `'0'`, and keeps the plain
 * form while it fits in 20 characters (sign excluded); longer values use E
 * notation: `0.1+0.2` gives `'0.3'`, `1/3` gives `'0.333333333333333'`,
 * `1e-10` gives `'0.0000000001'`, `1e20` gives `'1E+20'`. A non-finite input
 * gives `'#NUM!'` (NaN) or `'#DIV/0!'` (infinity).
 *
 * Known divergences from Excel 16.0 (build 20527, observed 2026-11-11), which
 * keeps fewer digits near the top of the double range:
 * `(1.23456789012345E+99)&""` is `'1.2345678901235E+99'` there (this gives
 * `'1.23456789012345E+99'`) and `9.99999999999999E+307&""` is `'1E+308'`
 * there (this gives `'9.99999999999999E+307'`).
 */
export function numberToExcelText(n: number): string {
  if (!Number.isFinite(n)) return nonFiniteText(n);
  if (n === 0) return '0';
  const decimal = toDecimalDigits(Math.abs(n), EXCEL_PRECISION);
  const plain = writePlain(decimal);
  const body = plain.length <= FORMULA_TEXT_WIDTH ? plain : writeScientific(decimal);
  return n < 0 ? `-${body}` : body;
}

/**
 * Format a number as Excel's `General` format displays it (11 characters,
 * sign excluded).
 *
 * The plain form keeps as many decimals as fit. E notation (6 significant
 * digits, 5 once the exponent has 3 digits) is used when the plain form does
 * not fit, or when it would round the value and show fewer significant
 * digits than E notation: `1/3` gives `'0.333333333'`, `123456789012345`
 * gives `'1.23457E+14'`, `0.0000123` gives `'0.0000123'` and `0.0000123456`
 * gives `'1.23456E-05'`. A non-finite input gives `'#NUM!'` (NaN) or
 * `'#DIV/0!'` (infinity).
 *
 * Known divergence from Excel 16.0 (build 20527, observed 2026-11-11): Excel
 * shows `1.23456789E+99` as `'1.2346E+99'` (5 significant digits with a
 * 2-digit exponent); this gives `'1.23457E+99'`.
 */
export function formatGeneralNumber(n: number): string {
  if (!Number.isFinite(n)) return nonFiniteText(n);
  if (n === 0) return '0';
  const abs = Math.abs(n);
  const exact = toDecimalDigits(abs, EXCEL_PRECISION);
  const intDigits = Math.max(exact.exp + 1, 1);
  const decimals = Math.max(GENERAL_WIDTH - intDigits - 1, 0);
  const plain = intDigits <= GENERAL_WIDTH ? stripFractionZeros(abs.toFixed(decimals)) : '';
  const plainSignificant = exact.exp >= 0 ? intDigits + decimals : decimals + exact.exp + 1;
  const sciSignificant = Math.abs(exact.exp) >= 100 ? 5 : 6;
  const usePlain =
    plain !== '' &&
    plain.length <= GENERAL_WIDTH &&
    (plainSignificant >= sciSignificant || plain === writePlain(exact));
  const body = usePlain ? plain : writeScientific(toDecimalDigits(abs, sciSignificant));
  return n < 0 ? `-${body}` : body;
}
