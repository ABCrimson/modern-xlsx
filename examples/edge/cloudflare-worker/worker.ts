/**
 * Cloudflare Worker — Generate XLSX files at the edge.
 *
 * Deploy:
 *   npx wrangler deploy
 *
 * Usage:
 *   GET /           → JSON info
 *   GET /generate   → Download generated XLSX
 *   POST /convert   → Convert a JSON array of objects to XLSX (body capped at 1 MiB; at most 10,000 rows and 256 columns)
 */

import { encodeCellRef, initWasmSync, Workbook } from 'modern-xlsx';
// wrangler imports a .wasm file as a compiled WebAssembly.Module; the package's
// "./wasm" export points at the binary.
import wasmModule from 'modern-xlsx/wasm';

initWasmSync(wasmModule);

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** Request-body cap for /convert, in bytes. */
const MAX_BODY_BYTES = 1_048_576;
/** Row cap for /convert. */
const MAX_ROWS = 10_000;
/** Column cap for /convert. */
const MAX_COLS = 256;

export default {
  async fetch(request): Promise<Response> {
    const url = new URL(request.url);

    switch (url.pathname) {
      case '/':
        return Response.json({
          service: 'modern-xlsx Cloudflare Worker',
          endpoints: ['/generate', '/convert'],
        });

      case '/generate':
        return handleGenerate();

      case '/convert':
        if (request.method !== 'POST') {
          return new Response('POST required', { status: 405 });
        }
        return handleConvert(request);

      default:
        return new Response('Not Found', { status: 404 });
    }
  },
} satisfies ExportedHandler;

async function handleGenerate(): Promise<Response> {
  const wb = new Workbook();
  const ws = wb.addSheet('Generated');
  ws.cell('A1').value = 'Hello from Cloudflare!';
  ws.cell('B1').value = 42;
  ws.cell('A2').value = 'Generated at';
  ws.cell('B2').value = new Date().toISOString();

  return new Response(await wb.toBuffer(), {
    headers: {
      'Content-Type': XLSX_TYPE,
      'Content-Disposition': 'attachment; filename="generated.xlsx"',
      'Cache-Control': 'no-cache',
    },
  });
}

/** Read the request body up to MAX_BODY_BYTES; `null` means the cap was exceeded. */
async function readCappedBody(request: Request): Promise<Uint8Array | null> {
  const declared = Number(request.headers.get('Content-Length'));
  if (declared > MAX_BODY_BYTES) return null;
  if (!request.body) return new Uint8Array();

  const reader = request.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_BODY_BYTES) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }

  const body = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    body.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return body;
}

async function handleConvert(request: Request): Promise<Response> {
  const body = await readCappedBody(request);
  if (body === null) {
    return Response.json({ error: `Body exceeds ${MAX_BODY_BYTES} bytes` }, { status: 413 });
  }

  let data: unknown;
  try {
    data = JSON.parse(new TextDecoder().decode(body));
  } catch {
    return Response.json({ error: 'Invalid JSON' }, { status: 400 });
  }
  if (!Array.isArray(data) || data.length === 0) {
    return Response.json({ error: 'Body must be a non-empty JSON array' }, { status: 400 });
  }
  const records = data as Record<string, unknown>[];
  if (records.length > MAX_ROWS) {
    return Response.json({ error: `More than ${MAX_ROWS} rows` }, { status: 413 });
  }

  // Headers come from the first object.
  const headers = Object.keys(records[0] ?? {});
  if (headers.length > MAX_COLS) {
    return Response.json({ error: `More than ${MAX_COLS} columns` }, { status: 413 });
  }

  const wb = new Workbook();
  const ws = wb.addSheet('Data');
  headers.forEach((h, ci) => {
    ws.cell(encodeCellRef(0, ci)).value = h;
  });
  records.forEach((row, ri) => {
    headers.forEach((h, ci) => {
      const val = row?.[h];
      ws.cell(encodeCellRef(ri + 1, ci)).value =
        typeof val === 'number' || typeof val === 'boolean' ? val : String(val ?? '');
    });
  });

  return new Response(await wb.toBuffer(), {
    headers: {
      'Content-Type': XLSX_TYPE,
      'Content-Disposition': 'attachment; filename="data.xlsx"',
    },
  });
}
