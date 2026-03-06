#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';
import { argv, exit } from 'node:process';
import { readBuffer } from './index.js';
import { sheetToCsv, sheetToJson } from './utils.js';

const args = argv.slice(2);
const command = args[0];

if (!command || command === '--help' || command === '-h') {
  console.log(`Usage:
  modern-xlsx info <file.xlsx>           Show sheet names, row counts, dimensions
  modern-xlsx convert <in.xlsx> <out>    Convert to CSV or JSON
    --sheet <n>                          0-based sheet index (default: 0, or all for JSON)
    --format csv|json                    Output format (default: json)
  modern-xlsx --help                     Show this help message`);
  exit(0);
}

if (command === 'info') {
  const file = args[1];
  if (!file) {
    console.error('Usage: modern-xlsx info <file.xlsx>');
    exit(1);
  }
  const data = new Uint8Array(readFileSync(file));
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
  exit(0);
}

if (command === 'convert') {
  const file = args[1];
  const output = args[2];
  if (!file || !output) {
    console.error('Usage: modern-xlsx convert <in.xlsx> <out> [--sheet <n>] [--format csv|json]');
    exit(1);
  }

  let sheetIndex: number | undefined;
  let format = 'json';
  for (let i = 3; i < args.length; i++) {
    if (args[i] === '--sheet' && args[i + 1]) {
      sheetIndex = Number(args[++i]);
    }
    if (args[i] === '--format' && args[i + 1]) {
      format = args[++i] ?? 'json';
    }
  }

  const data = new Uint8Array(readFileSync(file));
  const wb = await readBuffer(data);

  if (format === 'csv') {
    const idx = sheetIndex ?? 0;
    const ws = wb.getSheetByIndex(idx);
    if (!ws) {
      console.error(`Sheet index ${idx} not found`);
      exit(1);
    }
    const csv = sheetToCsv(ws);
    writeFileSync(output, csv);
  } else {
    // JSON output
    if (sheetIndex !== undefined) {
      const ws = wb.getSheetByIndex(sheetIndex);
      if (!ws) {
        console.error(`Sheet index ${sheetIndex} not found`);
        exit(1);
      }
      const json = sheetToJson(ws);
      writeFileSync(output, JSON.stringify(json, null, 2));
    } else {
      // Full workbook as JSON
      const result: Record<string, unknown[]> = {};
      for (let i = 0; i < wb.sheetCount; i++) {
        const ws = wb.getSheetByIndex(i);
        if (!ws) continue;
        const sheetName = wb.sheetNames[i] ?? `Sheet${i}`;
        result[sheetName] = sheetToJson(ws);
      }
      writeFileSync(output, JSON.stringify(result, null, 2));
    }
  }
  console.log(`Converted ${file} -> ${output} (${format})`);
  exit(0);
}

console.error(`Unknown command: ${command}. Use --help for usage.`);
exit(1);
