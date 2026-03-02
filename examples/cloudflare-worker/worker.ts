/**
 * Cloudflare Worker — Generate XLSX files at the edge.
 *
 * Deploy:
 *   npx wrangler deploy
 *
 * Usage:
 *   GET /           → JSON info
 *   GET /generate   → Download generated XLSX
 *   POST /convert   → Convert JSON body to XLSX
 */

import init, { read, write } from 'modern-xlsx/wasm/modern_xlsx_wasm.js';
// Import the WASM binary as an ArrayBuffer (Cloudflare Workers supports this)
import wasmModule from 'modern-xlsx/wasm/modern_xlsx_wasm_bg.wasm';

let initialized = false;

async function ensureInit(): Promise<void> {
  if (initialized) return;
  await init(wasmModule);
  initialized = true;
}

interface Env {}

export default {
  async fetch(request: Request, _env: Env): Promise<Response> {
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
} satisfies ExportedHandler<Env>;

async function handleGenerate(): Promise<Response> {
  await ensureInit();

  // Build a minimal workbook as JSON (matching WorkbookData shape)
  const workbookJson = JSON.stringify({
    sheets: [
      {
        name: 'Generated',
        worksheet: {
          rows: [
            {
              index: 1,
              cells: [
                { ref: 'A1', cellType: 'sharedString', value: 'Hello from Cloudflare!' },
                { ref: 'B1', cellType: 'number', value: '42' },
              ],
            },
            {
              index: 2,
              cells: [
                { ref: 'A2', cellType: 'sharedString', value: 'Generated at' },
                { ref: 'B2', cellType: 'sharedString', value: new Date().toISOString() },
              ],
            },
          ],
          mergeCells: [],
          conditionalFormatting: [],
          dataValidations: [],
          hyperlinks: [],
          columns: [],
        },
      },
    ],
    styles: {
      fonts: [{}],
      fills: [{ patternType: 'none' }, { patternType: 'gray125' }],
      borders: [{}],
      cellXfs: [{}],
      numFmts: [],
      cellStyles: [],
      dxfs: [],
    },
  });

  const xlsxBytes = write(workbookJson);

  return new Response(xlsxBytes, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="generated.xlsx"',
      'Cache-Control': 'no-cache',
    },
  });
}

async function handleConvert(request: Request): Promise<Response> {
  await ensureInit();

  const data = await request.json<Record<string, unknown>[]>();
  if (!Array.isArray(data) || data.length === 0) {
    return Response.json({ error: 'Body must be a non-empty JSON array' }, { status: 400 });
  }

  // Extract headers from first object
  const headers = Object.keys(data[0]);

  const rows = [
    {
      index: 1,
      cells: headers.map((h, i) => ({
        ref: String.fromCharCode(65 + i) + '1',
        cellType: 'sharedString' as const,
        value: h,
      })),
    },
    ...data.map((row, ri) => ({
      index: ri + 2,
      cells: headers.map((h, ci) => {
        const val = row[h];
        const ref = String.fromCharCode(65 + ci) + (ri + 2);
        if (typeof val === 'number') {
          return { ref, cellType: 'number' as const, value: String(val) };
        }
        if (typeof val === 'boolean') {
          return { ref, cellType: 'boolean' as const, value: val ? '1' : '0' };
        }
        return { ref, cellType: 'sharedString' as const, value: String(val ?? '') };
      }),
    })),
  ];

  const workbookJson = JSON.stringify({
    sheets: [
      {
        name: 'Data',
        worksheet: {
          rows,
          mergeCells: [],
          conditionalFormatting: [],
          dataValidations: [],
          hyperlinks: [],
          columns: [],
        },
      },
    ],
    styles: {
      fonts: [{}],
      fills: [{ patternType: 'none' }, { patternType: 'gray125' }],
      borders: [{}],
      cellXfs: [{}],
      numFmts: [],
      cellStyles: [],
      dxfs: [],
    },
  });

  const xlsxBytes = write(workbookJson);

  return new Response(xlsxBytes, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="data.xlsx"',
    },
  });
}
