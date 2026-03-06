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
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, StyleBuilder, initWasm } from 'modern-xlsx';

// ---------------------------------------------------------------------------
// 1. Initialize WASM
// ---------------------------------------------------------------------------
await initWasm();

// ---------------------------------------------------------------------------
// 2. Parse CSV
// ---------------------------------------------------------------------------
const __dirname = dirname(fileURLToPath(import.meta.url));
const csvPath = process.argv[2] || resolve(__dirname, 'data.csv');
const csvText = readFileSync(csvPath, 'utf-8');

/**
 * Simple CSV parser — handles quoted fields and newlines within quotes.
 * For production use, consider a dedicated CSV library.
 */
function parseCsv(text) {
  const lines = text.trim().split('\n');
  return lines.map((line) => {
    const cells = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (ch === ',' && !inQuotes) {
        cells.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
    cells.push(current.trim());
    return cells;
  });
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
// 4. Convert date strings to Excel serial numbers
// ---------------------------------------------------------------------------

/**
 * Convert a YYYY-MM-DD string to an Excel serial date number (1900 system).
 */
function dateToSerial(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  // JavaScript Date months are 0-based
  const dt = new Date(y, m - 1, d);
  // Excel epoch: Jan 1, 1900 = serial 1 (with the Lotus 1-2-3 leap year bug)
  const epoch = new Date(1899, 11, 30); // Dec 30, 1899
  const diff = (dt.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
  return Math.round(diff);
}

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

// --- Body styles (even/odd for zebra striping) ---
function buildBodyStyle(type, isOdd) {
  const sb = new StyleBuilder();

  // Zebra stripe fill
  if (isOdd) {
    sb.fill({ pattern: 'solid', fgColor: 'EDF2F7' });
  }

  // Thin border
  sb.border({
    left:   { style: 'thin', color: 'CBD5E0' },
    right:  { style: 'thin', color: 'CBD5E0' },
    top:    { style: 'thin', color: 'CBD5E0' },
    bottom: { style: 'thin', color: 'CBD5E0' },
  });

  // Type-specific formatting
  switch (type) {
    case 'number':
      sb.numberFormat('#,##0');
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
const evenStyles = colTypes.map((t) => buildBodyStyle(t, false));
const oddStyles = colTypes.map((t) => buildBodyStyle(t, true));

// --- Write header row ---
headerRow.forEach((h, c) => {
  const cell = ws.cell(colLetter(c) + '1');
  cell.value = h;
  cell.styleIndex = hdrStyle;
});

// --- Write data rows ---
dataRows.forEach((row, r) => {
  const rowNum = r + 2;
  const isOdd = r % 2 === 1;

  row.forEach((raw, c) => {
    const ref = colLetter(c) + rowNum;
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
          cell.value = dateToSerial(raw);
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
ws.autoFilter = `A1:${colLetter(colCount - 1)}${dataRows.length + 1}`;

// ---------------------------------------------------------------------------
// 6. Write output
// ---------------------------------------------------------------------------
const outPath = csvPath.replace(/\.csv$/i, '.xlsx');
const buffer = await wb.toBuffer();
writeFileSync(outPath, buffer);

console.log(`\nCreated ${outPath}`);
console.log(`  Columns: ${colCount}`);
console.log(`  Column types: ${colTypes.map((t, i) => `${headerRow[i]}=${t}`).join(', ')}`);
console.log(`  Features: zebra striping, auto-width, frozen header, auto-filter`);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Convert 0-based column index to Excel letter (A, B, ..., Z, AA, ...). */
function colLetter(idx) {
  let result = '';
  let n = idx;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}
