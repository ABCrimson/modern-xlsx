# Examples

Runnable demonstrations of modern-xlsx across runtimes, deploy targets, and frameworks. Each example depends on modern-xlsx at exactly the version in `packages/modern-xlsx/package.json`; `scripts/sync-example-pins.mjs` keeps the pins in step and CI checks it (`node scripts/sync-example-pins.mjs --check`).

## Tested versions

CI runs the Node examples, the Bun and Deno smoke tests, the Cloudflare Worker and the framework type checks against the tarball packed from this tree (`scripts/examples-smoke.mjs`, `scripts/examples-typecheck.mjs`; the `examples` and `examples-types` jobs in `.github/workflows/ci.yml`), not against the registry. The versions come from:

- Node: the `node-version` of the `examples` job in `.github/workflows/ci.yml`
- Bun and Deno: `BUN_VERSION` and `DENO_VERSION` in the `env` of that job
- Framework type checks: `PINS` in `scripts/examples-typecheck.mjs`
- wrangler: `examples/edge/cloudflare-worker/package.json`

## Full Applications (Node.js)

Run with `npm install && node index.mjs` inside the example directory.

| Example                                            | What it shows                                                                                                   |
| -------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| [export-sales-report](./node/export-sales-report/) | Professional styled report — merged title, zebra striping, number formats, formulas, auto-filter, frozen header |
| [chart-dashboard](./node/chart-dashboard/)         | Multi-sheet workbook with four embedded charts (bar, line ×2, pie) and a KPI row                                |
| [csv-to-styled-xlsx](./node/csv-to-styled-xlsx/)   | CSV ingestion with per-column type detection, then styled output with auto widths                               |
| [encrypted-payroll](./node/encrypted-payroll/)     | AES-256 Agile encryption plus sheet and workbook protection layers                                              |

## Runtimes & Deploy Targets

| Example                                        | Runtime                   | Run                                                                                                                                                                                 |
| ---------------------------------------------- | ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [runtimes](./runtimes/)                        | Bun, Deno                 | Bun: `bun install && bun run test` (the `test` script runs `bun test ./bun.test.ts`); Deno: `deno task test` (`deno.json`)                                                          |
| [deno-deploy](./edge/deno-deploy/)             | Deno Deploy               | `deno task dev` locally, `deno deploy` from `examples/edge/deno-deploy` to ship                                                                                                     |
| [cloudflare-worker](./edge/cloudflare-worker/) | Cloudflare Workers        | `npm install && npm run dev` (wrangler, configured in `wrangler.jsonc`), `npm run deploy` to ship; `npm run typecheck` checks types                                                 |
| [service-worker](./edge/service-worker/)       | Browser Service Worker    | Register `sw.js` from a page as a classic worker: `navigator.serviceWorker.register('/sw.js')`, with no module option, because `sw.js` loads the IIFE bundle with `importScripts()` |
| [demo-site](./browser/demo-site/)              | Browser (CDN, no bundler) | Open `index.html` or serve the directory with any static server                                                                                                                     |

## Framework Integration Snippets

Drop-in source files (no per-example install) showing the idiomatic WASM-init-plus-helpers pattern for each framework:

| Example                          | Framework | Files                                                                        |
| -------------------------------- | --------- | ---------------------------------------------------------------------------- |
| [react](./frameworks/react/)     | React     | `useXlsx.ts` hook + `ExcelExport.tsx` component                              |
| [vue](./frameworks/vue/)         | Vue 3     | `useXlsx.ts` composable + `ExcelExport.vue` component                        |
| [svelte](./frameworks/svelte/)   | Svelte 5  | `xlsx.svelte.ts` rune + `ExcelExport.svelte` component                       |
| [angular](./frameworks/angular/) | Angular   | `xlsx.service.ts` injectable service + `excel-export.component.ts` component |

## Related Documentation

- [Usage examples & recipes](../docs/examples.md) — copy-paste snippets for every major API
- [Migrating from SheetJS](../docs/migration-from-sheetjs.md) · [Migrating from ExcelJS](../docs/migration-from-exceljs.md)
- [Table layout engine guide](../docs/guide/tables.md) · [Barcode & QR guide](../docs/guide/barcodes.md)
- [Feature comparison](../docs/FEATURE-COMPARISON.md)
