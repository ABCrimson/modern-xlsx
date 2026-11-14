import { assertEquals, assertNotEquals } from '@std/assert';
import { initWasm, readBuffer, StyleBuilder, Workbook } from 'modern-xlsx';

Deno.test('round-trips values and a bold style through toBuffer and readBuffer', async () => {
  await initWasm();
  const wb = new Workbook();
  const ws = wb.addSheet('Test');
  ws.cell('A1').value = 'Hello';
  ws.cell('B1').value = 42;
  ws.cell('A1').styleIndex = new StyleBuilder().font({ bold: true }).build(wb.styles);

  const bytes = await wb.toBuffer();
  const wb2 = await readBuffer(bytes);
  const ws2 = wb2.getSheet('Test');
  assertEquals(ws2?.cell('A1').value, 'Hello');
  assertEquals(ws2?.cell('B1').value, 42);

  const xfIndex = ws2?.cell('A1').styleIndex ?? null;
  assertNotEquals(xfIndex, null);
  const xf = wb2.styles.cellXfs[xfIndex ?? -1];
  assertEquals(wb2.styles.fonts[xf?.fontId ?? -1]?.bold, true);
});
