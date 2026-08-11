/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Regression guard for the recurring better-sqlite3 ABI mismatch (#88, #95, and
 * again on 2026-08-10): the compiled binding must match the Node that loads it.
 * `scripts/ensure-native.mjs` makes that self-correcting on dev/CI runs, but only
 * while it stays wired into the lifecycle hooks — an unwired guard is no guard.
 *
 * The hooks are deliberately dev/CI-only. `scripts/` is not in the package's
 * published `files`, so wiring `prestart` or `postinstall` would break consumers
 * of the npm package with a missing-file error. That constraint is asserted here
 * too, because it is the non-obvious part someone would otherwise "fix".
 */

import { describe, it, expect } from 'vitest';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

// Join OS paths rather than interpolating one into a `file://` string: a repo
// checked out under a path containing #, ? or % would otherwise produce a URL
// that parses as fragment/query/escape and reads the wrong file.
const repoRoot = fileURLToPath(new URL('../../', import.meta.url));
const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8')) as {
  scripts: Record<string, string>;
  files: string[];
};

describe('better-sqlite3 ABI auto-heal stays wired', () => {
  it('the heal script exists', () => {
    expect(existsSync(join(repoRoot, 'scripts', 'ensure-native.mjs'))).toBe(true);
  });

  it.each(['pretest', 'prestart:dev'])('%s runs the heal script', (hook) => {
    expect(pkg.scripts[hook]).toContain('ensure-native.mjs');
  });

  it('is reachable manually', () => {
    expect(pkg.scripts['ensure-native']).toContain('ensure-native.mjs');
  });

  it('is not wired to hooks that ship to npm consumers', () => {
    // scripts/ is not published, so these would fail on an installed package.
    expect(pkg.files).not.toContain('scripts');
    for (const hook of ['postinstall', 'prestart', 'prepare']) {
      expect(pkg.scripts[hook] ?? '').not.toContain('ensure-native.mjs');
    }
  });
});
