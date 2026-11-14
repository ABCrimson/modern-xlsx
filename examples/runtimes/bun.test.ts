import { expect, test } from 'bun:test';
import { initWasm, readBuffer, StyleBuilder, Workbook } from 'modern-xlsx';

test('round-trips values and a bold style through toBuffer and readBuffer', async () => {
  await initWasm();
  const wb = new Workbook();
  const ws = wb.addSheet('Test');
  ws.cell('A1').value = 'Hello';
  ws.cell('B1').value = 42;
  ws.cell('A1').styleIndex = new StyleBuilder().font({ bold: true }).build(wb.styles);

  const bytes = await wb.toBuffer();
  const wb2 = await readBuffer(bytes);
  const ws2 = wb2.getSheet('Test');
  expect(ws2?.cell('A1').value).toBe('Hello');
  expect(ws2?.cell('B1').value).toBe(42);

  const xfIndex = ws2?.cell('A1').styleIndex;
  expect(xfIndex).not.toBeNull();
  const xf = wb2.styles.cellXfs[xfIndex ?? -1];
  expect(wb2.styles.fonts[xf?.fontId ?? -1]?.bold).toBe(true);
});
