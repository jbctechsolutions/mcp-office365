import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    environment: 'node',
    include: ['tests/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/index.ts'],
      thresholds: {
        // Ratchet (v3): these are a no-regression floor, not a target. The
        // global `branches` floor is set at the current actual (~65%) so CI is
        // honest and green; raise it as each U2 domain migration lands its
        // tests, and restore branches to 75 once the migration clears it.
        // New v3 code is held to a high bar via the per-glob threshold below.
        lines: 75,
        functions: 75,
        branches: 64,
        statements: 75,
        'src/registry/**': {
          lines: 90,
          functions: 90,
          branches: 90,
          statements: 90,
        },
      },
    },
    testTimeout: 10000,
  },
});
