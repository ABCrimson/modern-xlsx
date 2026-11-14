#!/usr/bin/env node
/**
 * Runs the examples against a packed modern-xlsx tarball (the CI `npm-tarball` artifact,
 * never the registry). Each example is copied to a temp directory, its modern-xlsx
 * dependency is rewritten to file:<tgz> and installed there, so the checkout stays untouched.
 *
 * usage: node scripts/examples-smoke.mjs <node|edge> --tgz <path to modern-xlsx-<version>.tgz>
 *
 *   node  Runs `node index.mjs` in chart-dashboard, csv-to-styled-xlsx, encrypted-payroll and
 *         export-sales-report, re-opens each output with readBuffer and asserts its sheet
 *         names; the payroll file must open with its password and fail without it.
 *   edge  In cloudflare-worker: `wrangler types --check`, `tsc --noEmit`, then
 *         `wrangler dev --local` on port 8787 (polled until GET / answers, 60 s cap):
 *         /generate must return a workbook whose A1 is the greeting, and POST /convert must
 *         put the 27th header at AA1, answer 413 to a 2 MiB chunked body and 400 to '{'.
 *         The dev server is always stopped, pass or fail.
 *
 * Any failed command or assertion exits non-zero. The temp directory is removed on success.
 */
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { cpSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

const root = join(import.meta.dirname, '..');
const WINDOWS = process.platform === 'win32';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { tgz: { type: 'string' } },
});
const [mode] = positionals;
if (!['node', 'edge'].includes(mode) || values.tgz === undefined) {
  console.error('usage: node scripts/examples-smoke.mjs <node|edge> --tgz <modern-xlsx tgz>');
  process.exit(2);
}
const tgz = resolve(values.tgz).replaceAll('\\', '/');

/** The Node examples (name is the label, dir the path under examples/), the file each writes, and the sheets that file must hold. */
const NODE_EXAMPLES = [
  {
    name: 'chart-dashboard',
    dir: 'node/chart-dashboard',
    output: 'chart-dashboard.xlsx',
    sheets: ['Raw Data', 'Dashboard'],
  },
  {
    name: 'csv-to-styled-xlsx',
    dir: 'node/csv-to-styled-xlsx',
    output: 'data.xlsx',
    sheets: ['Data'],
  },
  {
    name: 'encrypted-payroll',
    dir: 'node/encrypted-payroll',
    output: 'payroll-encrypted.xlsx',
    sheets: ['Summary', 'Payroll Detail'],
    password: 'payroll2025',
  },
  {
    name: 'export-sales-report',
    dir: 'node/export-sales-report',
    output: 'sales-report.xlsx',
    sheets: ['Sales Report'],
  },
];

/** The Worker's GET /generate greeting in A1 of sheet "Generated". */
const WORKER_GREETING = 'Hello from Cloudflare!';
const WORKER_PORT = 8787;
const READY_TIMEOUT_MS = 60_000;

/** Quotes one argument for cmd.exe, which runs npm and npx (.cmd shims) on Windows. */
const quote = (arg) => (/^[\w@./:=+-]+$/.test(arg) ? arg : `"${arg}"`);

/** Runs a command to completion with inherited stdio; a non-zero exit throws. */
function run(cwd, command, ...args) {
  console.log(`$ ${[command, ...args].join(' ')}  (in ${basename(cwd)})`);
  const result = WINDOWS
    ? spawnSync([command, ...args].map(quote).join(' '), { cwd, stdio: 'inherit', shell: true })
    : spawnSync(command, args, { cwd, stdio: 'inherit' });
  if (result.error) throw result.error;
  assert.equal(result.status, 0, `${command} ${args.join(' ')} exited with ${result.status}`);
}

/** Copies examples/<group>/<name> to <base>/<name> and installs it with modern-xlsx from the tarball. */
function installExample(base, example) {
  const name = basename(example);
  const dir = join(base, name);
  cpSync(join(root, 'examples', example), dir, {
    recursive: true,
    filter: (src) => !['node_modules', '.wrangler'].includes(basename(src)),
  });
  run(dir, 'npm', 'pkg', 'set', `dependencies.modern-xlsx=file:${tgz}`);
  run(dir, 'npm', 'install', '--no-audit', '--no-fund');
  return dir;
}

/** Imports the modern-xlsx that `dir` installed and initializes its WASM. */
async function loadModernXlsx(dir) {
  const entry = createRequire(join(dir, 'package.json')).resolve('modern-xlsx');
  const api = await import(pathToFileURL(entry).href);
  await api.initWasm();
  return api;
}

async function smokeNode(base) {
  let api;
  for (const example of NODE_EXAMPLES) {
    const dir = installExample(base, example.dir);
    run(dir, 'node', 'index.mjs');
    api ??= await loadModernXlsx(dir);

    const bytes = readFileSync(join(dir, example.output));
    const options = example.password === undefined ? undefined : { password: example.password };
    const wb = await api.readBuffer(bytes, options);
    assert.deepEqual(wb.sheetNames, example.sheets, `${example.name}: sheet names`);
    if (example.password !== undefined) {
      await assert.rejects(
        api.readBuffer(bytes),
        { code: 'PASSWORD_PROTECTED' },
        `${example.name}: ${example.output} opened without its password`,
      );
    }
    console.log(`ok ${example.name}: ${example.output} [${wb.sheetNames.join(', ')}]`);
  }
}

/** Stops the dev server and everything it spawned (workerd). */
async function stopProcessTree(child) {
  if (child.exitCode !== null || child.signalCode !== null) return;
  const exited = new Promise((done) => child.once('exit', done));
  if (WINDOWS) {
    spawnSync('taskkill', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore' });
  } else {
    process.kill(-child.pid, 'SIGTERM');
  }
  const stopped = await Promise.race([
    exited.then(() => true),
    delay(10_000, false, { ref: false }),
  ]);
  if (!stopped && !WINDOWS) process.kill(-child.pid, 'SIGKILL');
}

/** Polls GET / until it answers 200; fails early if the dev server exits. */
async function waitForWorker(origin, child) {
  const deadline = Date.now() + READY_TIMEOUT_MS;
  while (Date.now() < deadline) {
    assert.equal(child.exitCode, null, `wrangler dev exited early with ${child.exitCode}`);
    try {
      const res = await fetch(`${origin}/`, { signal: AbortSignal.timeout(5_000) });
      if (res.ok) return res;
    } catch {
      // Not listening yet.
    }
    await delay(500);
  }
  assert.fail(`wrangler dev did not answer GET / within ${READY_TIMEOUT_MS / 1000} s`);
}

/** A request body of `size` bytes sent as a stream, so it goes out chunked with no Content-Length. */
function chunkedBody(size) {
  const chunk = new Uint8Array(64 * 1024).fill(0x20);
  let sent = 0;
  return new ReadableStream({
    pull(controller) {
      if (sent >= size) {
        controller.close();
        return;
      }
      controller.enqueue(chunk);
      sent += chunk.byteLength;
    },
  });
}

async function smokeEdge(base) {
  const dir = installExample(base, 'edge/cloudflare-worker');
  run(dir, 'npx', 'wrangler', 'types', '--check');
  run(dir, 'npx', 'tsc', '--noEmit');
  const api = await loadModernXlsx(dir);

  // The wrangler CLI runs directly under node (not through npx) so the process tree is ours
  // to stop: its own process group on POSIX, taskkill /t on Windows.
  const wranglerDir = join(dir, 'node_modules', 'wrangler');
  const wranglerManifest = JSON.parse(readFileSync(join(wranglerDir, 'package.json'), 'utf8'));
  const wranglerBin = join(wranglerDir, wranglerManifest.bin.wrangler);
  const origin = `http://127.0.0.1:${WORKER_PORT}`;
  console.log(
    `$ wrangler dev --local --ip 127.0.0.1 --port ${WORKER_PORT}  (in cloudflare-worker)`,
  );
  const wrangler = spawn(
    process.execPath,
    [wranglerBin, 'dev', '--local', '--ip', '127.0.0.1', '--port', String(WORKER_PORT)],
    {
      cwd: dir,
      detached: !WINDOWS,
      stdio: ['ignore', 'inherit', 'inherit'],
      env: { ...process.env, WRANGLER_SEND_METRICS: 'false' },
    },
  );

  try {
    const index = await (await waitForWorker(origin, wrangler)).json();
    assert.deepEqual(index.endpoints, ['/generate', '/convert'], 'GET /: endpoints');
    console.log('ok worker GET /');

    const generated = await fetch(`${origin}/generate`);
    assert.equal(generated.status, 200, 'GET /generate: status');
    const wb = await api.readBuffer(new Uint8Array(await generated.arrayBuffer()));
    assert.deepEqual(wb.sheetNames, ['Generated'], 'GET /generate: sheet names');
    assert.equal(wb.getSheet('Generated')?.cell('A1').value, WORKER_GREETING, 'GET /generate: A1');
    console.log(`ok worker GET /generate: A1 = ${WORKER_GREETING}`);

    const wide = Object.fromEntries(Array.from({ length: 27 }, (_, i) => [`k${i + 1}`, i + 1]));
    const converted = await fetch(`${origin}/convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify([wide]),
    });
    assert.equal(converted.status, 200, 'POST /convert (27 keys): status');
    const data = await api.readBuffer(new Uint8Array(await converted.arrayBuffer()));
    const sheet = data.getSheet('Data');
    assert.equal(sheet?.cell('AA1').value, 'k27', 'POST /convert (27 keys): AA1');
    assert.equal(sheet?.cell('AA2').value, 27, 'POST /convert (27 keys): AA2');
    console.log('ok worker POST /convert: 27th header at AA1');

    const tooLarge = await fetch(`${origin}/convert`, {
      method: 'POST',
      body: chunkedBody(2 * 1024 * 1024),
      duplex: 'half',
    });
    assert.equal(tooLarge.status, 413, 'POST /convert (2 MiB chunked): status');
    console.log('ok worker POST /convert: 2 MiB chunked body -> 413');

    const invalid = await fetch(`${origin}/convert`, { method: 'POST', body: '{' });
    assert.equal(invalid.status, 400, "POST /convert ('{'): status");
    console.log("ok worker POST /convert: '{' -> 400");
  } finally {
    await stopProcessTree(wrangler);
  }
}

const base = mkdtempSync(join(process.env.RUNNER_TEMP ?? tmpdir(), 'modern-xlsx-examples-'));
await (mode === 'node' ? smokeNode(base) : smokeEdge(base));
// Retries cover Windows, where a just-killed workerd can hold its files for a moment.
rmSync(base, { recursive: true, force: true, maxRetries: 10, retryDelay: 500 });
console.log(`ok examples-smoke ${mode}`);
