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
import { join } from 'node:path';

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
  // The Dockerfile was absent from this list until #109, which is exactly how
  // it kept a Node 26 base image through #99's LTS repin — and it is the
  // artifact that actually ships to the remote connector.
  ['Dockerfile', 'Dockerfile', /FROM node:(\d+)/],
];

/**
 * Lowest Node major the package supports. Since v5.0.0 this is driven by a
 * stdlib API requirement, not only LTS policy: the durable store uses
 * `node:sqlite`, which is unflagged and warning-free from Node 24 (#108).
 * Keep this distinct from MAX_LTS_MAJOR — the two constraints move for
 * different reasons and should not be conflated when the LTS ceiling rises.
 */
const MIN_SUPPORTED_MAJOR = 24;

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

  it('every Dockerfile stage pins the same Node major', () => {
    // Builder and runtime drifting apart is how a compiled artifact built in
    // one stage becomes unloadable in the other. No native module ships today
    // (#108 removed the last one), but the invariant is cheap to hold.
    const text = readFileSync(new URL('Dockerfile', `file://${repoRoot}`), 'utf8');
    const majors = [...text.matchAll(/FROM node:(\d+)/g)].map((m) => Number(m[1]));
    expect(majors.length).toBeGreaterThan(0);
    expect(new Set(majors).size).toBe(1);
  });
});

describe('supported Node floor matches the node:sqlite requirement (#108)', () => {
  const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8')) as {
    engines: { node: string };
  };

  it('engines.node floor is at least the version node:sqlite needs', () => {
    const floor = Number(/(\d+)/.exec(pkg.engines.node)?.[1]);
    expect(floor).toBeGreaterThanOrEqual(MIN_SUPPORTED_MAJOR);
  });

  it('the CI matrix exercises nothing below the supported floor', () => {
    const text = readFileSync(join(repoRoot, '.github/workflows/test.yml'), 'utf8');
    const matrix = /node-version:\s*\[([^\]]+)\]/.exec(text)?.[1] ?? '';
    const majors = [...matrix.matchAll(/(\d+)\.x/g)].map((m) => Number(m[1]));
    expect(majors.length).toBeGreaterThan(0);
    for (const major of majors) {
      expect(major).toBeGreaterThanOrEqual(MIN_SUPPORTED_MAJOR);
    }
  });
});
