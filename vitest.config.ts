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
      // Excluded: src/index.ts is the entrypoint + (shrinking) legacy dispatch.
      // The *-graph.ts / *-apple.ts files are extracted dispatch bodies — thin
      // per-backend param-mapping glue over methods already covered by the
      // repository/tools/integration test suites; they inherit index.ts's
      // exclusion so a pure relocation doesn't read as a coverage regression.
      // (Backfilling dedicated handler tests for them is a tracked follow-up.)
      exclude: ['src/index.ts', 'src/tools/*-graph.ts'],
      thresholds: {
        // Ratchet (v3): these are a no-regression floor, not a target. The
        // global `branches` floor tracks the current actual so CI is honest and
        // green; raise it as each domain lands its tests, and restore branches
        // to 75 once the migration clears it. New v3 code is held to a high bar
        // via the per-glob threshold below. Floor lowered 64 -> 63 when the
        // AppleScript backend (well-tested code) was removed, which shifted the
        // whole-repo branch baseline down ~0.6% without any new-code regression.
        lines: 75,
        functions: 75,
        branches: 63,
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
    // Setup hooks (SQLite fixture builds for the AppleScript integration tests)
    // occasionally exceed the 10s default on slow/contended CI runners
    // (notably Windows + Node 20), producing flaky "Hook timed out" failures.
    // The hooks are not slow by design; give them headroom without masking a
    // genuine hang.
    hookTimeout: 30000,
  },
});
