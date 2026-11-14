import { describe, expect, it } from 'vitest';
import XLSX from 'xlsx';
import { readBuffer, StreamingXlsxWriter, Workbook } from '../src/index.js';

/**
 * The TypeScript defaults (`new Workbook()`), Rust `Styles::default_styles()`
 * (used when xl/styles.xml is absent) and the streaming writer's
 * MINIMAL_STYLES_XML are three copies of one default (the streaming copy also
 * writes the explicit cellStyleXfs/xfId pair); these tests fail if they drift.
 */
describe('default styles stay in sync', () => {
  it('reader fallback for a missing xl/styles.xml equals the TS defaults', async () => {
    const wb = new Workbook();
    wb.addSheet('S').cell('A1').value = 1;
    const buf = await wb.toBuffer();

    const cfb = XLSX.CFB.read(Buffer.from(buf), { type: 'buffer' });
    XLSX.CFB.utils.cfb_del(cfb, '/xl/styles.xml');
    const stripped = new Uint8Array(XLSX.CFB.write(cfb, { type: 'buffer', fileType: 'zip' }));

    const wb2 = await readBuffer(stripped);
    expect(wb2.toJSON().styles).toEqual(new Workbook().toJSON().styles);
  });

  it('streaming writer minimal styles equal the TS defaults', async () => {
    const writer = StreamingXlsxWriter.create();
    writer.startSheet('S');
    writer.writeRow([{ value: '1' }]);
    const wb2 = await readBuffer(await writer.finish());
    // MINIMAL_STYLES_XML spells out Excel's default cell style: one cellStyleXfs
    // entry and xfId="0" on the cell format. The reader keeps both, so they are
    // the only difference from the TS defaults.
    const defaults = new Workbook().toJSON().styles;
    expect(wb2.toJSON().styles).toEqual({
      ...defaults,
      cellStyleXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
      cellXfs: defaults.cellXfs.map((xf) => ({ ...xf, xfId: 0 })),
    });
  });
});
