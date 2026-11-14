import { describe, expect, it } from 'vitest';
import { readBuffer, Workbook } from '../src/index.js';

describe('docProperties extended fields', () => {
  it('appVersion, hyperlinkBase and revision survive a round trip', async () => {
    const wb = new Workbook();
    wb.addSheet('S').cell('A1').value = 1;
    wb.docProperties = {
      appVersion: '16.0300',
      hyperlinkBase: 'https://example.com/base/',
      revision: '7',
    };

    const wb2 = await readBuffer(await wb.toBuffer());
    expect(wb2.docProperties).toStrictEqual({
      appVersion: '16.0300',
      hyperlinkBase: 'https://example.com/base/',
      revision: '7',
    });
  });
});
