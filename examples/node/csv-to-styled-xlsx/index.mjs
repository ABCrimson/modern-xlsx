/**
 * CSV to Styled XLSX — modern-xlsx example
 *
 * Reads a CSV file, auto-detects column types (numbers, dates, strings,
 * booleans), and writes a styled XLSX workbook with:
 *   - Styled header row
 *   - Alternating row colors (zebra striping)
 *   - Type-appropriate number formats
 *   - Auto-calculated column widths
 *   - Frozen header row
 *   - Auto-filter on the header row
 */

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { columnToLetter, dateToSerial, initWasm, StyleBuilder, Workbook } from 'modern-xlsx';

// ---------------------------------------------------------------------------
// 1. Initialize WASM
// ---------------------------------------------------------------------------
await initWasm();

// ---------------------------------------------------------------------------
// 2. Parse CSV
// ---------------------------------------------------------------------------
const csvPath = process.argv[2] || resolve(import.meta.dirname, 'data.csv');
const csvText = readFileSync(csvPath, 'utf-8');

/**
 * Read a quoted field starting at the opening quote `src[start]`.
 * "" is a literal quote; CR and LF are kept. Returns the text and the index
 * just past the closing quote (or the end of input when it is unterminated).
 */
function readQuoted(src, start) {
  let value = '';
  let i = start + 1;
  while (i < src.length) {
    if (src[i] !== '"') {
      value += src[i++];
    } else if (src[i + 1] === '"') {
      value += '"';
      i += 2;
    } else {
      return [value, i + 1];
    }
  }
  return [value, i];
}

/** Drop one leading byte order mark (U+FEFF). */
function stripBom(text) {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** Remove trailing records that hold a single empty field, always keeping the first record. */
function dropTrailingEmptyRecords(records) {
  while (records.length > 1 && records.at(-1).length === 1 && records.at(-1)[0] === '') {
    records.pop();
  }
  return records;
}

/**
 * RFC 4180: quoted fields, "" escapes, CRLF/LF, newlines within quotes.
 * Single pass over the whole text. Unquoted fields are trimmed; quoted ones are
 * kept verbatim. For production use, consider a dedicated CSV library.
 */
function parseCsv(text) {
  const src = stripBom(text);
  const records = [];
  let record = [];
  let field = '';
  let quoted = false; // the current field was quoted (do not trim it)

  const atFieldStart = () => !quoted && field.trim() === ''; // a quote here opens a quoted field
  const endField = () => {
    record.push(quoted ? field : field.trim());
    field = '';
    quoted = false;
  };

  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"' && atFieldStart()) {
      [field, i] = readQuoted(src, i);
      quoted = true;
    } else if (ch === ',') {
      endField();
      i++;
    } else if (ch === '\n' || src.startsWith('\r\n', i)) {
      endField();
      records.push(record);
      record = [];
      i += ch === '\r' ? 2 : 1;
    } else {
      field += ch;
      i++;
    }
  }
  // Final record without a trailing newline; a trailing newline leaves nothing to add
  if (field !== '' || quoted || record.length > 0) {
    endField();
    records.push(record);
  }
  return dropTrailingEmptyRecords(records);
}

const parsed = parseCsv(csvText);
const [headerRow, ...dataRows] = parsed;

console.log(`Parsed ${csvPath}`);
console.log(`  Headers: ${headerRow.join(', ')}`);
console.log(`  Rows: ${dataRows.length}`);

// ---------------------------------------------------------------------------
// 3. Auto-detect column types
// ---------------------------------------------------------------------------

/** Date pattern: YYYY-MM-DD */
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Detect the predominant type for a column by sampling all values.
 * Returns: 'number' | 'date' | 'boolean' | 'string'
 */
function detectColumnType(values) {
  let numCount = 0;
  let dateCount = 0;
  let boolCount = 0;

  for (const v of values) {
    if (v === '' || v == null) continue;
    const lower = v.toLowerCase();
    if (lower === 'true' || lower === 'false') {
      boolCount++;
    } else if (DATE_RE.test(v)) {
      dateCount++;
    } else if (!Number.isNaN(Number(v)) && v !== '') {
      numCount++;
    }
  }

  const total = values.filter((v) => v !== '' && v != null).length;
  if (dateCount > total * 0.5) return 'date';
  if (boolCount > total * 0.5) return 'boolean';
  if (numCount > total * 0.5) return 'number';
  return 'string';
}

const colTypes = headerRow.map((_, colIdx) => {
  const colValues = dataRows.map((row) => row[colIdx] || '');
  return detectColumnType(colValues);
});

console.log(`  Types: ${colTypes.join(', ')}`);

// ---------------------------------------------------------------------------
// 4. Dates: YYYY-MM-DD strings become Excel serial numbers via the library's
//    dateToSerial() with calendar fields (never a Date, so the zone cannot shift the day)
// ---------------------------------------------------------------------------
// 5. Build the workbook
// ---------------------------------------------------------------------------
const wb = new Workbook();
const ws = wb.addSheet('Data');
const colCount = headerRow.length;

// --- Header style ---
const hdrStyle = new StyleBuilder()
  .font({ bold: true, size: 11, color: 'FFFFFF' })
  .fill({ pattern: 'solid', fgColor: '2D3748' })
  .alignment({ horizontal: 'center', vertical: 'center' })
  .border({
    bottom: { style: 'medium', color: '1A202C' },
  })
  .build(wb.styles);

// --- Columns holding any non-integer value get a decimal format ('#,##0.0#') ---
const fractionalCols = colTypes.map(
  (type, c) =>
    type === 'number' &&
    dataRows.some((row) => row[c] !== '' && row[c] != null && !Number.isInteger(Number(row[c]))),
);

// --- Body styles (even/odd for zebra striping) ---
function buildBodyStyle(type, isOdd, colIdx) {
  const sb = new StyleBuilder();

  // Zebra stripe fill
  if (isOdd) {
    sb.fill({ pattern: 'solid', fgColor: 'EDF2F7' });
  }

  // Thin border
  sb.border({
    left: { style: 'thin', color: 'CBD5E0' },
    right: { style: 'thin', color: 'CBD5E0' },
    top: { style: 'thin', color: 'CBD5E0' },
    bottom: { style: 'thin', color: 'CBD5E0' },
  });

  // Type-specific formatting
  switch (type) {
    case 'number':
      sb.numberFormat(fractionalCols[colIdx] ? '#,##0.0#' : '#,##0');
      sb.alignment({ horizontal: 'right' });
      break;
    case 'date':
      sb.numberFormat('yyyy-mm-dd');
      sb.alignment({ horizontal: 'center' });
      break;
    case 'boolean':
      sb.alignment({ horizontal: 'center' });
      break;
    default:
      sb.alignment({ horizontal: 'left' });
      break;
  }

  return sb.build(wb.styles);
}

// Pre-build style indices for each column x parity combination
const evenStyles = colTypes.map((t, c) => buildBodyStyle(t, false, c));
const oddStyles = colTypes.map((t, c) => buildBodyStyle(t, true, c));

// --- Write header row ---
headerRow.forEach((h, c) => {
  const cell = ws.cell(`${columnToLetter(c)}1`);
  cell.value = h;
  cell.styleIndex = hdrStyle;
});

// --- Write data rows ---
dataRows.forEach((row, r) => {
  const rowNum = r + 2;
  const isOdd = r % 2 === 1;

  row.forEach((raw, c) => {
    const ref = columnToLetter(c) + rowNum;
    const cell = ws.cell(ref);
    const type = colTypes[c];

    switch (type) {
      case 'number': {
        const num = Number(raw);
        cell.value = Number.isNaN(num) ? raw : num;
        break;
      }
      case 'date': {
        if (DATE_RE.test(raw)) {
          const [year, month, day] = raw.split('-').map(Number);
          cell.value = dateToSerial({ year, month, day });
        } else {
          cell.value = raw;
        }
        break;
      }
      case 'boolean': {
        const lower = (raw || '').toLowerCase();
        cell.value = lower === 'true';
        break;
      }
      default:
        cell.value = raw;
    }

    cell.styleIndex = isOdd ? oddStyles[c] : evenStyles[c];
  });
});

// --- Auto-calculate column widths ---
function estimateWidth(val) {
  const str = String(val ?? '');
  return Math.max(8, str.length + 3);
}

headerRow.forEach((h, c) => {
  let maxWidth = estimateWidth(h);
  for (const row of dataRows) {
    const w = estimateWidth(row[c]);
    if (w > maxWidth) maxWidth = w;
  }
  // Cap at reasonable width
  ws.setColumnWidth(c + 1, Math.min(maxWidth, 40));
});

// --- Freeze header row ---
ws.frozenPane = { rows: 1, cols: 0 };

// --- Auto-filter ---
ws.autoFilter = `A1:${columnToLetter(colCount - 1)}${dataRows.length + 1}`;

// ---------------------------------------------------------------------------
// 6. Write output
// ---------------------------------------------------------------------------
const outPath = csvPath.replace(/\.csv$/i, '.xlsx');
await wb.toFile(outPath);

console.log(`\nCreated ${outPath}`);
console.log(`  Columns: ${colCount}`);
console.log(`  Column types: ${colTypes.map((t, i) => `${headerRow[i]}=${t}`).join(', ')}`);
console.log(`  Features: zebra striping, auto-width, frozen header, auto-filter`);
