/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * When the better-sqlite3 native module cannot load at all (Node ABI mismatch
 * after a version switch, missing build), the in-memory fallback is backed by
 * the same module — attempting it just rethrows the dlopen error and the
 * server died with a raw stack. StateStore.open must instead surface one
 * actionable error naming the remediation steps.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const dlopenError = Object.assign(
  new Error(
    "The module '/x/better_sqlite3.node' was compiled against a different Node.js version using NODE_MODULE_VERSION 127. This version of Node.js requires NODE_MODULE_VERSION 147.",
  ),
  { code: 'ERR_DLOPEN_FAILED' },
);

vi.mock('better-sqlite3', () => ({
  default: class {
    constructor() {
      throw dlopenError;
    }
  },
}));

describe('StateStore.open with an unloadable native module', () => {
  let dir: string;
  let legacyDir: string;

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'mcp-state-'));
    legacyDir = mkdtempSync(join(tmpdir(), 'mcp-legacy-'));
  });

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true });
    rmSync(legacyDir, { recursive: true, force: true });
  });

  it('throws one actionable error instead of crashing on the in-memory fallback', async () => {
    const { StateStore } = await import('../../../src/state/store.js');
    const warnings: string[] = [];

    expect(() => StateStore.open({ dir, legacyDir, warn: (m) => warnings.push(m) })).toThrowError(
      /better-sqlite3 native module failed to load[\s\S]*npm rebuild better-sqlite3[\s\S]*npx cache/,
    );
    // It must not have pretended to degrade — the fallback is impossible here.
    expect(warnings.some((w) => w.includes('running in-memory'))).toBe(false);
  });

  it('names the running Node version and preserves the original error as cause', async () => {
    const { StateStore } = await import('../../../src/state/store.js');

    let thrown: unknown;
    try {
      StateStore.open({ dir, legacyDir, warn: () => {} });
    } catch (e) {
      thrown = e;
    }
    expect(thrown).toBeInstanceOf(Error);
    expect((thrown as Error).message).toContain(process.version);
    expect((thrown as Error).cause).toBe(dlopenError);
  });
});
