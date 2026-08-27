# Examples

Runnable, self-contained demonstrations of modern-xlsx across runtimes, deploy targets, and frameworks. Each example pins the published npm package (no local build required).

## Full Applications (Node.js)

Run with `npm install && node index.mjs` inside the example directory.

| Example | What it shows |
|---------|---------------|
| [export-sales-report](./export-sales-report/) | Professional styled report — merged title, zebra striping, number formats, formulas, auto-filter, frozen header |
| [chart-dashboard](./chart-dashboard/) | Multi-sheet workbook with four embedded charts (bar, line ×2, pie) and a KPI row |
| [csv-to-styled-xlsx](./csv-to-styled-xlsx/) | CSV ingestion with per-column type detection, then styled output with auto widths |
| [encrypted-payroll](./encrypted-payroll/) | AES-256 Agile encryption plus sheet and workbook protection layers |

## Runtimes & Deploy Targets

| Example | Runtime | Run |
|---------|---------|-----|
| [bun-test](./bun-test/) | Bun ≥ 1.0 | `bun install && bun run test.ts` |
| [deno-test](./deno-test/) | Deno ≥ 1.40 | `deno task test` |
| [deno-deploy](./deno-deploy/) | Deno Deploy | `deno task dev` locally, `deployctl deploy` to ship |
| [cloudflare-worker](./cloudflare-worker/) | Cloudflare Workers | `npm install && npm run dev` (wrangler), `npm run deploy` to ship |
| [service-worker](./service-worker/) | Browser Service Worker | Register `sw.js` from a page: `navigator.serviceWorker.register('/sw.js', { type: 'module' })` |
| [demo-site](./demo-site/) | Browser (CDN, no bundler) | Open `index.html` or serve the directory with any static server |

## Framework Integration Snippets

Drop-in source files (no per-example install) showing the idiomatic WASM-init-plus-helpers pattern for each framework:

| Example | Framework | Files |
|---------|-----------|-------|
| [react](./react/) | React | `useXlsx.ts` hook + `ExcelExport.tsx` component |
| [vue](./vue/) | Vue 3 | `useXlsx.ts` composable + `ExcelExport.vue` component |
| [svelte](./svelte/) | Svelte 5 | `xlsx.svelte.ts` rune + `ExcelExport.svelte` component |
| [angular](./angular/) | Angular | `xlsx.service.ts` injectable service + `excel-export.component.ts` component |

## Related Documentation

- [Usage examples & recipes](../docs/examples.md) — copy-paste snippets for every major API
- [Migrating from SheetJS](../docs/migration-from-sheetjs.md) · [Migrating from ExcelJS](../docs/migration-from-exceljs.md)
- [Table layout engine guide](../docs/guide/tables.md) · [Barcode & QR guide](../docs/guide/barcodes.md)
- [Feature comparison vs SheetJS](../docs/FEATURE-COMPARISON.md)
