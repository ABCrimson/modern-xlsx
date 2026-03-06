#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import { parseArgs } from 'node:util';
import { readBuffer } from './index.js';
import { sheetToCsv, sheetToJson } from './utils.js';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    help: { type: 'boolean', short: 'h', default: false },
    sheet: { type: 'string' },
    format: { type: 'string', default: 'json' },
  },
});

const command = positionals[0];

if (!command || values.help) {
  console.log(`Usage:
  modern-xlsx info <file.xlsx>           Show sheet names, row counts, dimensions
  modern-xlsx convert <in.xlsx> <out>    Convert to CSV or JSON
    --sheet <n>                          0-based sheet index (default: 0, or all for JSON)
    --format csv|json                    Output format (default: json)
  modern-xlsx --help                     Show this help message`);
  process.exit(0);
}

if (command === 'info') {
  const file = positionals[1];
  if (!file) {
    console.error('Usage: modern-xlsx info <file.xlsx>');
    process.exit(1);
  }
  const data = new Uint8Array(await readFile(file));
  const wb = await readBuffer(data);
  console.log(`Sheets: ${wb.sheetCount}`);
  for (let i = 0; i < wb.sheetCount; i++) {
    const ws = wb.getSheetByIndex(i);
    if (!ws) continue;
    const name = wb.sheetNames[i];
    const rows = ws.rowCount;
    const dim = ws.dimension ?? 'unknown';
    console.log(`  ${i}: "${name}" — ${rows} rows, dimension: ${dim}`);
  }
  process.exit(0);
}

if (command === 'convert') {
  const file = positionals[1];
  const output = positionals[2];
  if (!file || !output) {
    console.error('Usage: modern-xlsx convert <in.xlsx> <out> [--sheet <n>] [--format csv|json]');
    process.exit(1);
  }

  const sheetIndex = values.sheet != null ? Number(values.sheet) : undefined;
  const format = values.format ?? 'json';

  const data = new Uint8Array(await readFile(file));
  const wb = await readBuffer(data);

  if (format === 'csv') {
    const idx = sheetIndex ?? 0;
    const ws = wb.getSheetByIndex(idx);
    if (!ws) {
      console.error(`Sheet index ${idx} not found`);
      process.exit(1);
    }
    const csv = sheetToCsv(ws);
    await writeFile(output, csv);
  } else {
    // JSON output
    if (sheetIndex !== undefined) {
      const ws = wb.getSheetByIndex(sheetIndex);
      if (!ws) {
        console.error(`Sheet index ${sheetIndex} not found`);
        process.exit(1);
      }
      const json = sheetToJson(ws);
      await writeFile(output, JSON.stringify(json, null, 2));
    } else {
      // Full workbook as JSON
      const result: Record<string, unknown[]> = {};
      for (let i = 0; i < wb.sheetCount; i++) {
        const ws = wb.getSheetByIndex(i);
        if (!ws) continue;
        const sheetName = wb.sheetNames[i] ?? `Sheet${i}`;
        result[sheetName] = sheetToJson(ws);
      }
      await writeFile(output, JSON.stringify(result, null, 2));
    }
  }
  console.log(`Converted ${file} -> ${output} (${format})`);
  process.exit(0);
}

console.error(`Unknown command: ${command}. Use --help for usage.`);
process.exit(1);
