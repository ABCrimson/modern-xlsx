import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';
import { readZipText } from './helpers/zip.js';

const CTRL_PROP = 'xl/ctrlProps/ctrlProp1.xml';
const CT_CTRL_PROPS = 'application/vnd.ms-excel.controlproperties+xml';
const OVERRIDE = `<Override PartName="/${CTRL_PROP}" ContentType="${CT_CTRL_PROPS}"/>`;

describe('preservedContentTypes', () => {
  it('a seeded workbook declares the recorded type of a preserved part', async () => {
    const wb = new Workbook({
      preservedEntries: { [CTRL_PROP]: Array.from(new TextEncoder().encode('<formControlPr/>')) },
      preservedContentTypes: { [CTRL_PROP]: CT_CTRL_PROPS },
    });
    wb.addSheet('Sheet1');

    const bytes = await wb.toBuffer();
    expect(readZipText(bytes)['[Content_Types].xml']).toContain(OVERRIDE);

    // The reader records the type again, so a second round trip keeps it.
    const back = await readBuffer(bytes);
    expect(readZipText(await back.toBuffer())['[Content_Types].xml']).toContain(OVERRIDE);
  });
});
