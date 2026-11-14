/**
 * Deno Deploy — Generate XLSX files with modern-xlsx.
 *
 * Deploy:
 *   deno deploy   (Deno's built-in `deno deploy` subcommand; run from examples/edge/deno-deploy)
 *
 * Local dev:
 *   deno run --allow-net --allow-read main.ts
 */

import { initWasm, Workbook } from 'modern-xlsx';

await initWasm();

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
    const wb = new Workbook();
    const ws = wb.addSheet('Deno Sheet');
    ws.cell('A1').value = 'Hello from Deno Deploy!';
    ws.cell('B1').value = 2026;
    ws.cell('A2').value = 'Runtime';
    ws.cell('B2').value = `Deno ${Deno.version.deno}`;

    return new Response(await wb.toBuffer(), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="deno-generated.xlsx"',
      },
    });
  }

  return new Response('Not Found', { status: 404 });
});
