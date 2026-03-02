/**
 * Deno Deploy — Generate XLSX files with modern-xlsx.
 *
 * Deploy:
 *   deployctl deploy --project=my-xlsx main.ts
 *
 * Local dev:
 *   deno run --allow-net --allow-read main.ts
 */

import init, { write } from 'npm:modern-xlsx/wasm/modern_xlsx_wasm.js';

let initialized = false;

async function ensureInit(): Promise<void> {
  if (initialized) return;
  // Deno supports WASM init from URL or file
  await init();
  initialized = true;
}

Deno.serve({ port: 8000 }, async (request: Request): Promise<Response> => {
  const url = new URL(request.url);

  if (url.pathname === '/') {
    return Response.json({
      service: 'modern-xlsx Deno Deploy',
      runtime: `Deno ${Deno.version.deno}`,
      endpoints: ['/generate'],
    });
  }

  if (url.pathname === '/generate') {
    await ensureInit();

    const workbookJson = JSON.stringify({
      sheets: [
        {
          name: 'Deno Sheet',
          worksheet: {
            rows: [
              {
                index: 1,
                cells: [
                  { ref: 'A1', cellType: 'sharedString', value: 'Hello from Deno Deploy!' },
                  { ref: 'B1', cellType: 'number', value: '2026' },
                ],
              },
              {
                index: 2,
                cells: [
                  { ref: 'A2', cellType: 'sharedString', value: 'Runtime' },
                  { ref: 'B2', cellType: 'sharedString', value: `Deno ${Deno.version.deno}` },
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
        'Content-Disposition': 'attachment; filename="deno-generated.xlsx"',
      },
    });
  }

  return new Response('Not Found', { status: 404 });
});
