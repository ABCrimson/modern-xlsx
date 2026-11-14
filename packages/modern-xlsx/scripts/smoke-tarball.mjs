// Smoke test for the packed npm tarball, run by the CI smoke job. Not shipped: the
// package.json `files` allowlist leaves scripts/ out.
//
// Usage: copy this file into a clean project that ran `npm install <modern-xlsx tgz>`,
// then `node smoke-tarball.mjs`. Each entry point ('modern-xlsx', 'modern-xlsx/lite')
// runs in its own Node process, so one entry's initialized WASM can never stand in for
// the other's. There it initializes WASM synchronously from the packed binary, writes,
// validates and reads back a one-cell workbook, and checks VERSION against the packed
// package.json. Then the packed CLI's `info` command reads the file the full entry wrote.
// It prints one `ok` line per check; any failure throws, so node exits non-zero.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const pkgDir = join(import.meta.dirname, 'node_modules', 'modern-xlsx');
const pkg = JSON.parse(readFileSync(join(pkgDir, 'package.json'), 'utf8'));

/** Entry point -> the WASM binary its loader expects, relative to the package root. */
const entries = {
  'modern-xlsx': 'wasm/modern_xlsx_wasm_bg.wasm',
  'modern-xlsx/lite': 'wasm-lite/modern_xlsx_wasm_bg.wasm',
};
const outFile = (name) => join(import.meta.dirname, `${name.replace('/', '-')}.xlsx`);

const entry = process.argv[2];
if (entry === undefined) {
  for (const name of Object.keys(entries)) {
    execFileSync(process.execPath, [import.meta.filename, name], { stdio: 'inherit' });
  }
  // The CLI initializes the full build with no-argument initWasm(), the Node default path.
  const cli = join(pkgDir, pkg.bin['modern-xlsx']);
  const out = execFileSync(process.execPath, [cli, 'info', outFile('modern-xlsx')], {
    encoding: 'utf8',
  });
  assert.match(out, /^Sheets: 1\n {2}0: "S"/, 'cli: info printed an unexpected summary');
  console.log(`ok modern-xlsx CLI info: ${out.split('\n')[1].trim()}`);
} else {
  const api = await import(entry);
  api.initWasmSync(readFileSync(join(pkgDir, entries[entry])));

  const wb = new api.Workbook();
  wb.addSheet('S').cell('A1').value = 'smoke';
  const buf = await wb.toBuffer();
  assert.ok(buf instanceof Uint8Array, `${entry}: toBuffer() did not return a Uint8Array`);
  assert.ok(buf.length > 0, `${entry}: toBuffer() returned no bytes`);
  assert.equal(wb.validate().isValid, true, `${entry}: validate() reported issues`);

  const wb2 = await api.readBuffer(buf);
  assert.deepEqual(wb2.sheetNames, ['S'], `${entry}: sheet names did not round-trip`);
  assert.equal(wb2.getSheet('S')?.cell('A1').value, 'smoke', `${entry}: A1 did not round-trip`);
  assert.equal(api.VERSION, pkg.version, `${entry}: VERSION differs from the packed package.json`);

  writeFileSync(outFile(entry), buf);
  console.log(`ok ${entry} ${api.VERSION}: ${buf.length} bytes written, validated and read back`);
}
