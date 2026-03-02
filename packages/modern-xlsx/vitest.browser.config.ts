import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['__tests__/browser.test.ts'],
    browser: {
      enabled: true,
      provider: 'playwright',
      instances: [
        { browser: 'chromium' },
      ],
    },
    testTimeout: 15_000,
  },
});
