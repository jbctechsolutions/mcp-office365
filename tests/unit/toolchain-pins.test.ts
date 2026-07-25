/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Regression guard for issue #99: the toolchain that builds and releases the
 * server must pin a Node that is an active LTS — even-major AND not newer than
 * the current LTS line. Pinning a Current release (e.g. Node 26, as #88 did,
 * before 26 entered LTS) ships/compiles a native ABI that no LTS consumer can
 * load, which is what broke installed builds. Note parity alone is insufficient:
 * 26 is even but was *Current*, so the max-LTS ceiling below is what actually
 * catches the #99 regression. The CI test matrix still exercises newer/odd
 * majors for forward-compat — those are intentionally NOT checked here; only the
 * single-version *build/release* pins are.
 *
 * When a newer even-major actually enters LTS (e.g. Node 26 in Oct 2026), bump
 * MAX_LTS_MAJOR deliberately in the same change that repins — that conscious
 * edit is the gate this guard is protecting.
 */

import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

/** Highest Node major that is an active LTS line today. Bump on purpose. */
const MAX_LTS_MAJOR = 24;

const repoRoot = fileURLToPath(new URL('../../', import.meta.url));

/** Extracts the major from the first `node-version:`-style pin in the text. */
function nodeMajor(relPath: string, pattern: RegExp): number {
  const text = readFileSync(new URL(relPath, `file://${repoRoot}`), 'utf8');
  const match = pattern.exec(text);
  if (match?.[1] == null) {
    throw new Error(`no Node version pin found in ${relPath}`);
  }
  return Number(match[1]);
}

const PINS: ReadonlyArray<[label: string, relPath: string, pattern: RegExp]> = [
  ['.tool-versions', '.tool-versions', /nodejs\s+(\d+)/],
  ['release.yml', '.github/workflows/release.yml', /node-version:\s*'?(\d+)/],
  ['integration.yml', '.github/workflows/integration.yml', /node-version:\s*'?(\d+)/],
];

describe('toolchain Node pins are even-major LTS (#99)', () => {
  it.each(PINS)('%s pins an active-LTS Node', (_label, relPath, pattern) => {
    const major = nodeMajor(relPath, pattern);
    expect(Number.isInteger(major)).toBe(true);
    expect(major).toBeGreaterThanOrEqual(20);
    // Even major === an LTS line (odd majors never become LTS).
    expect(major % 2).toBe(0);
    // ...and not newer than the current LTS. A Current even-major (e.g. 26
    // before it enters LTS) is exactly the #99 regression. Bump MAX_LTS_MAJOR
    // deliberately when the newer line actually reaches LTS.
    expect(major).toBeLessThanOrEqual(MAX_LTS_MAJOR);
  });
});
