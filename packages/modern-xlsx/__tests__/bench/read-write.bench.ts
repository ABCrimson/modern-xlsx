import { cpus } from 'node:os';
import {
  aoaToSheet,
  columnToLetter,
  initWasm,
  readBuffer,
  sheetAddAoa,
  sheetToCsv,
  sheetToJson,
  Workbook,
} from 'modern-xlsx';
import { afterAll, test } from 'vitest';
import XLSX, { version as sheetJsVersion } from 'xlsx';

// modern-xlsx vs SheetJS. This file runs under `pnpm bench` only (`vitest bench`, picked up by
// `benchmark.include`), never under `pnpm test`: `bench` is a test-context fixture in Vitest 5,
// and the numbers are machine-dependent, so nothing here asserts on timing. After the
// comparisons, afterAll prints the Performance block of packages/modern-xlsx/README.md: a dated
// caption sentence, the table with the unit in each row label, and the line on the ratio. The
// README holds that block as printed: __tests__/readme-examples.test.ts checks the README's
// table format and that this file and the README carry the same header, row labels and ratio line.
// The benches measure
// the built package: 'modern-xlsx' resolves to dist/index.mjs, which Node loads natively (see
// vitest.config.ts), and `pnpm bench` builds it first. That dist instance is initialised here;
// setup.ts initialises only the src one.
await initWasm();

/** Few, bounded iterations: the 100K-row cases take seconds per call. */
const RUN = { iterations: 5, time: 0, warmupIterations: 1, warmupTime: 0 } as const;

function generateLargeData(rows: number, cols: number): unknown[][] {
  const data: unknown[][] = [];
  const header: string[] = [];
  for (let c = 0; c < cols; c++) {
    header.push(`Col${c}`);
  }
  data.push(header);

  for (let r = 0; r < rows; r++) {
    const row: unknown[] = [];
    for (let c = 0; c < cols; c++) {
      if (c % 3 === 0) row.push(`String value ${r}-${c}`);
      else if (c % 3 === 1) row.push(r * 1000 + c);
      else row.push(r % 2 === 0);
    }
    data.push(row);
  }
  return data;
}

/** Populate a modern-xlsx worksheet from 2D data using the cell-by-cell API. */
function populateSheet(
  ws: ReturnType<Workbook['addSheet']>,
  data: unknown[][],
  colLetters: string[],
): void {
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
        ws.cell(`${colLetters[c]}${r + 1}`).value = val;
      }
    }
  }
}

function colLettersForCount(n: number): string[] {
  return Array.from({ length: n }, (_, i) => columnToLetter(i));
}

/** SheetJS bytes for a 2D array. Compressed (deflate), like modern-xlsx output: its default is a stored ZIP. */
function sheetJsBuffer(data: unknown[][]): Uint8Array {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Data');
  return new Uint8Array(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true }));
}

async function modernBuffer(data: unknown[][], cols: string[]): Promise<Uint8Array> {
  const wb = new Workbook();
  populateSheet(wb.addSheet('Data'), data, cols);
  return wb.toBuffer();
}

/** The README rows, in the order the README prints them (reads first). */
const ROW_ORDER = [
  'Read 10K rows (ms)',
  'Read 100K rows (ms)',
  'Write 10K rows (ms)',
  'File size, 10K rows (KB)',
] as const;

interface Row {
  label: (typeof ROW_ORDER)[number];
  modern: string;
  sheetJs: string;
  ratio: string;
}
const tableRows: Row[] = [];

const oneDecimal = new Intl.NumberFormat('en-US', {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const whole = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });

/** A mean latency in milliseconds, without its unit (the row label names it): `1,591.3`. */
function ms(v: number): string {
  return oneDecimal.format(v);
}

/** A byte count in KiB, without its unit (the row label names it): `1,245`. */
function kb(bytes: number): string {
  return whole.format(bytes / 1024);
}

/** SheetJS over modern-xlsx, so a value above 1 favours modern-xlsx: `3.25x`. */
function ratio(sheetJs: number, modern: number): string {
  return `${(sheetJs / modern).toFixed(2)}x`;
}

const COLS = colLettersForCount(10);

test('write 10K rows x 10 cols (cell-by-cell)', async ({ bench }) => {
  const data = generateLargeData(10_000, 10);
  const result = await bench.compare(
    bench('modern-xlsx', async () => {
      await modernBuffer(data, COLS);
    }),
    bench('SheetJS', () => {
      sheetJsBuffer(data);
    }),
    RUN,
  );
  const modern = result.get('modern-xlsx').latency.mean;
  const sheetJs = result.get('SheetJS').latency.mean;
  tableRows.push({
    label: 'Write 10K rows (ms)',
    modern: ms(modern),
    sheetJs: ms(sheetJs),
    ratio: ratio(sheetJs, modern),
  });

  const modernSize = (await modernBuffer(data, COLS)).length;
  const sheetJsSize = sheetJsBuffer(data).length;
  tableRows.push({
    label: 'File size, 10K rows (KB)',
    modern: kb(modernSize),
    sheetJs: kb(sheetJsSize),
    ratio: ratio(sheetJsSize, modernSize),
  });
});

test('read 10K rows x 10 cols', async ({ bench }) => {
  const input = sheetJsBuffer(generateLargeData(10_000, 10));
  const result = await bench.compare(
    bench('modern-xlsx', async () => {
      await readBuffer(input);
    }),
    bench('SheetJS', () => {
      XLSX.read(input, { type: 'buffer' });
    }),
    RUN,
  );
  const modern = result.get('modern-xlsx').latency.mean;
  const sheetJs = result.get('SheetJS').latency.mean;
  tableRows.push({
    label: 'Read 10K rows (ms)',
    modern: ms(modern),
    sheetJs: ms(sheetJs),
    ratio: ratio(sheetJs, modern),
  });
});

test('read 100K rows x 10 cols', async ({ bench }) => {
  const input = sheetJsBuffer(generateLargeData(100_000, 10));
  const result = await bench.compare(
    bench('modern-xlsx', async () => {
      await readBuffer(input);
    }),
    bench('SheetJS', () => {
      XLSX.read(input, { type: 'buffer' });
    }),
    RUN,
  );
  const modern = result.get('modern-xlsx').latency.mean;
  const sheetJs = result.get('SheetJS').latency.mean;
  tableRows.push({
    label: 'Read 100K rows (ms)',
    modern: ms(modern),
    sheetJs: ms(sheetJs),
    ratio: ratio(sheetJs, modern),
  });
});

test('write 100K rows x 10 cols via sheetAddAoa (batch API)', async ({ bench }) => {
  const data = generateLargeData(100_000, 10);
  await bench.compare(
    bench('modern-xlsx (sheetAddAoa + toBuffer)', async () => {
      const wb = new Workbook();
      sheetAddAoa(wb.addSheet('Data'), data, { origin: 'A1' });
      await wb.toBuffer();
    }),
    bench('SheetJS (aoa_to_sheet + write)', () => {
      sheetJsBuffer(data);
    }),
    RUN,
  );
});

test('aoaToSheet 50K rows x 5 cols', async ({ bench }) => {
  const data = generateLargeData(50_000, 5);
  await bench.compare(
    bench('modern-xlsx', () => {
      aoaToSheet(data);
    }),
    bench('SheetJS', () => {
      XLSX.utils.aoa_to_sheet(data);
    }),
  );
});

test('sheetToCsv 10K rows x 10 cols', async ({ bench }) => {
  const data = generateLargeData(10_000, 10);
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  populateSheet(ws, data, COLS);
  const xlsxWs = XLSX.utils.aoa_to_sheet(data);
  await bench.compare(
    bench('modern-xlsx', () => {
      sheetToCsv(ws);
    }),
    bench('SheetJS', () => {
      XLSX.utils.sheet_to_csv(xlsxWs);
    }),
  );
});

test('sheetToJson 10K rows x 10 cols', async ({ bench }) => {
  const data = generateLargeData(10_000, 10);
  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  populateSheet(ws, data, COLS);
  const xlsxWs = XLSX.utils.aoa_to_sheet(data);
  await bench.compare(
    bench('modern-xlsx', () => {
      sheetToJson(ws);
    }),
    bench('SheetJS', () => {
      XLSX.utils.sheet_to_json(xlsxWs);
    }),
  );
});

afterAll(() => {
  // os.cpus() names the CPU with trademark signs, as in 'Intel(R) Core(TM) Ultra 9 290HX Plus'.
  const cpu =
    cpus()[0]
      ?.model.replace(/\((?:R|TM)\)/gi, '')
      .replace(/\s+/g, ' ')
      .trim() ?? 'unknown CPU';
  const day = new Date().toISOString().slice(0, 10);
  const rows = tableRows.toSorted(
    (a, b) => ROW_ORDER.indexOf(a.label) - ROW_ORDER.indexOf(b.label),
  );
  const table = [
    ['10 columns', 'modern-xlsx', 'SheetJS', 'Ratio'],
    ...rows.map((r) => [r.label, r.modern, r.sheetJs, r.ratio]),
  ];
  // Padded as Biome formats the README (its lint scope includes Markdown): the label column
  // left-aligned, the figures right-aligned, every column as wide as its widest cell.
  const widths = table[0]?.map((_, c) => Math.max(...table.map((row) => row[c]?.length ?? 0)));
  const pad = (cells: string[]) =>
    `| ${cells.map((cell, c) => (c === 0 ? cell.padEnd(widths?.[c] ?? 0) : cell.padStart(widths?.[c] ?? 0))).join(' | ')} |`;
  const rule = `| ${widths?.map((w, c) => (c === 0 ? '-'.repeat(w) : `${'-'.repeat(w - 1)}:`)).join(' | ')} |`;
  const lines = [
    '',
    `Measured on ${day} (UTC) on ${cpu}, with Node ${process.version} and SheetJS ${sheetJsVersion}.`,
    '',
    pad(table[0] ?? []),
    rule,
    ...table.slice(1).map(pad),
    '',
    'Ratio is the SheetJS figure divided by the modern-xlsx figure: above 1x favours modern-xlsx.',
    '',
  ];
  console.log(lines.join('\n'));
});
