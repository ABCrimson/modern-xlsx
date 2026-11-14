import { initWasm, readBuffer, Workbook } from 'modern-xlsx';
import { test } from 'vitest';
import { fillGrid } from '../helpers/fixtures.js';

// Write and read throughput at 1K and 10K rows x 5 columns. Runs under `pnpm bench` only
// and measures the built package (see read-write.bench.ts); the correctness and size-scaling
// checks stay in large-workbook.test.ts.
await initWasm();

const RUN = { iterations: 5, time: 0, warmupIterations: 1, warmupTime: 0 } as const;

function populate(rows: number): Workbook {
  const wb = new Workbook();
  fillGrid(wb.addSheet('Bench'), rows, 5);
  return wb;
}

for (const rows of [1000, 10_000]) {
  const label = rows === 1000 ? '1K' : '10K';

  test(`${label} rows x 5 cols: write`, async ({ bench }) => {
    // The sheet is built once; each iteration serializes it.
    const wb = populate(rows);
    await bench(`${label} write`, async () => {
      await wb.toBuffer();
    }).run(RUN);
  });

  test(`${label} rows x 5 cols: read`, async ({ bench }) => {
    const buffer = await populate(rows).toBuffer();
    await bench(`${label} read`, async () => {
      await readBuffer(buffer);
    }).run(RUN);
  });
}
