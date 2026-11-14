/**
 * The `modern-xlsx` command line, as a function: `run(argv, io)` resolves to the process
 * exit code instead of calling `process.exit`, so it can be tested in-process. `cli.ts` is
 * the thin executable wrapper around it.
 *
 * Exit codes: 0 success, 1 runtime failure, 2 usage error. Diagnostics go to stderr as
 * `modern-xlsx: <message>`, never as a stack trace.
 */

import { writeFile } from 'node:fs/promises';
import { parseArgs } from 'node:util';
import { initWasm, readFile, type Workbook } from './index.js';
import { sheetToCsv, sheetToJson } from './io/sheet-utils.js';

/** Where the CLI writes its text. */
export interface CliIO {
  stdout(s: string): void;
  stderr(s: string): void;
}

const USAGE = `Usage:
  modern-xlsx info <file.xlsx>           Show sheet names, row counts, dimensions
  modern-xlsx convert <in.xlsx> <out>    Convert to CSV or JSON
    --sheet <n>                          0-based sheet index (default: 0 for CSV, all sheets for JSON)
    --format csv|json                    Output format (default: json)
  modern-xlsx --help                     Show this help message
`;

/** Thrown for a usage error; `run` turns it into exit code 2. */
class UsageError extends Error {}

const OPTIONS = {
  help: { type: 'boolean', short: 'h', default: false },
  sheet: { type: 'string' },
  format: { type: 'string', default: 'json' },
} as const;

/**
 * Run the CLI.
 *
 * @param argv - The arguments after the executable, e.g. `process.argv.slice(2)`.
 * @param io - Sinks for stdout and stderr.
 * @returns The exit code: 0, 1 (runtime failure) or 2 (usage error).
 */
export async function run(argv: readonly string[], io: CliIO): Promise<number> {
  let parsed: ReturnType<typeof parse>;
  try {
    parsed = parse(argv);
  } catch (error) {
    return usageFailure(io, error instanceof Error ? error.message : String(error));
  }
  const { values, positionals } = parsed;
  const command = positionals[0];

  if (values.help || !command) {
    io.stdout(USAGE);
    return 0;
  }

  try {
    if (command === 'info') return await info(positionals, io);
    if (command === 'convert') return await convert(positionals, values, io);
    throw new UsageError(`unknown command '${command}'`);
  } catch (error) {
    if (error instanceof UsageError) return usageFailure(io, error.message);
    io.stderr(`modern-xlsx: ${error instanceof Error ? error.message : String(error)}\n`);
    return 1;
  }
}

function parse(argv: readonly string[]) {
  return parseArgs({ args: [...argv], allowPositionals: true, options: OPTIONS });
}

function usageFailure(io: CliIO, message: string): number {
  io.stderr(`modern-xlsx: ${message}\n\n${USAGE}`);
  return 2;
}

async function info(positionals: readonly string[], io: CliIO): Promise<number> {
  const file = positionals[1];
  if (!file) throw new UsageError('info needs <file.xlsx>');
  await initWasm();
  const wb = await readFile(file);
  io.stdout(`Sheets: ${wb.sheetCount}\n`);
  for (let i = 0; i < wb.sheetCount; i++) {
    const ws = wb.getSheetByIndex(i);
    if (!ws) continue;
    io.stdout(
      `  ${i}: "${wb.sheetNames[i]}" — ${ws.rowCount} rows, dimension: ${ws.dimension ?? 'unknown'}\n`,
    );
  }
  return 0;
}

async function convert(
  positionals: readonly string[],
  values: ReturnType<typeof parse>['values'],
  io: CliIO,
): Promise<number> {
  const [, file, output] = positionals;
  if (!file || !output) throw new UsageError('convert needs <in.xlsx> and <out>');

  const { format, sheetIndex } = convertOptions(values);

  await initWasm();
  const wb = await readFile(file);

  await writeFile(output, render(wb, format, sheetIndex));
  io.stdout(`Converted ${file} -> ${output} (${format})\n`);
  return 0;
}

function convertOptions(values: ReturnType<typeof parse>['values']): {
  format: 'csv' | 'json';
  sheetIndex: number | undefined;
} {
  const format = values.format ?? 'json';
  if (format !== 'csv' && format !== 'json') {
    throw new UsageError("--format must be 'csv' or 'json'");
  }
  if (values.sheet === undefined) return { format, sheetIndex: undefined };
  if (!/^\d+$/.test(values.sheet)) throw new UsageError('--sheet must be an integer >= 0');
  return { format, sheetIndex: Number(values.sheet) };
}

function render(wb: Workbook, format: 'csv' | 'json', sheetIndex: number | undefined): string {
  if (format === 'json' && sheetIndex === undefined) {
    const entries: [string, unknown[]][] = [];
    for (let i = 0; i < wb.sheetCount; i++) {
      const ws = wb.getSheetByIndex(i);
      if (ws) entries.push([wb.sheetNames[i] ?? `Sheet${i}`, sheetToJson(ws)]);
    }
    // fromEntries defines own properties, so a sheet named `__proto__` survives.
    return JSON.stringify(Object.fromEntries(entries), null, 2);
  }
  const idx = sheetIndex ?? 0;
  const ws = wb.getSheetByIndex(idx);
  if (!ws) throw new Error(`sheet index ${idx} not found`);
  return format === 'csv' ? sheetToCsv(ws) : JSON.stringify(sheetToJson(ws), null, 2);
}
