import { fileURLToPath } from 'node:url';
import { playwright } from '@vitest/browser-playwright';
import { defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: {
    // Tests run src/ in Chromium: '#runtime' is the browser twin.
    alias: { '#runtime': fileURLToPath(new URL('./src/runtime-browser.ts', import.meta.url)) },
  },
  test: {
    include: ['__tests__/browser.test.ts', '__tests__/*.browser.test.ts'],
    browser: {
      enabled: true,
      provider: playwright(),
      instances: [{ browser: 'chromium' }],
    },
    testTimeout: 15_000,
  },
});
