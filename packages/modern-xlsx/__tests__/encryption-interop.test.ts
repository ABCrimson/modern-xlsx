import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readBuffer } from '../src/index.js';
import { PASSWORD_PROTECTED } from '../src/model/errors.js';

const PASSWORD = 'Password1234_';

function fixture(dir: 'msoffcrypto' | 'legacy', name: string): Uint8Array {
  return readFileSync(
    new URL(`../../../crates/modern-xlsx-core/tests/fixtures/${dir}/${name}`, import.meta.url),
  );
}

describe('Encryption interop: files from other implementations', () => {
  it('reads the msoffcrypto Agile fixture to the same sheets and cells as the plain file', async () => {
    const plain = await readBuffer(fixture('msoffcrypto', 'example.xlsx'));
    const decrypted = await readBuffer(fixture('msoffcrypto', 'example_password.xlsx'), {
      password: PASSWORD,
    });
    expect(plain.sheetNames.length).toBeGreaterThan(0);
    expect(decrypted.sheetNames).toEqual(plain.sheetNames);
    expect(decrypted.toJSON().sheets).toEqual(plain.toJSON().sheets);
  });

  it('rejects a wrong password with PASSWORD_PROTECTED', async () => {
    await expect(
      readBuffer(fixture('msoffcrypto', 'example_password.xlsx'), { password: 'nope' }),
    ).rejects.toMatchObject({
      name: 'ModernXlsxError',
      code: PASSWORD_PROTECTED,
      message: 'Password protected: Incorrect password.',
    });
  });

  it('reads the modern-xlsx 1.2.1 legacy fixture to the same values', async () => {
    const plain = await readBuffer(fixture('msoffcrypto', 'example.xlsx'));
    const legacy = await readBuffer(fixture('legacy', 'modern-xlsx-1.2.1-agile.xlsx'), {
      password: PASSWORD,
    });
    expect(legacy.sheetNames).toEqual(plain.sheetNames);
    expect(legacy.toJSON().sheets).toEqual(plain.toJSON().sheets);
  });
});
