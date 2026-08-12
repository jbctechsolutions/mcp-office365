/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * The entry point refuses to start on a Node too old for `node:sqlite`, with a
 * message that says what to do (#76). Without it the failure is a raw
 * `ERR_UNKNOWN_BUILTIN_MODULE` stack that an MCP host renders as an unexplained
 * "failed to connect" — the same illegible-startup-failure class #77 fixed for
 * the native binding, reintroduced by the v5.0.0 driver swap.
 */

import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { nodeFloorError, MIN_NODE_MAJOR } from '../../src/node-floor.js';

describe('node floor guard (#76)', () => {
  it.each(['20.20.2', '22.11.0', '22.13.0', '22.15.0', '18.0.0'])(
    'refuses to start on Node %s',
    (version) => {
      const message = nodeFloorError(version);
      expect(message).not.toBeNull();
      expect(message).toContain(version);
      expect(message).toContain(`requires Node.js ${MIN_NODE_MAJOR}`);
    },
  );

  it.each(['24.0.0', '24.18.0', '26.2.0', '30.1.0'])('starts on Node %s', (version) => {
    expect(nodeFloorError(version)).toBeNull();
  });

  it('names all three remedies, including the version consumers can fall back to', () => {
    const message = nodeFloorError('20.20.2') ?? '';
    expect(message).toContain('Upgrade to Node.js');
    expect(message).toContain('absolute path');
    // The fallback matters most: a Node 20 user has a supported option today.
    expect(message).toContain('@jbctechsolutions/mcp-office365@4');
  });

  it('does not claim node:sqlite is absent, which is untrue for 22.13+', () => {
    // 22.13 and 22.15 do provide it — experimentally, with a launch warning.
    // Overstating the reason would send someone chasing a broken install.
    const message = nodeFloorError('22.15.0') ?? '';
    expect(message).toContain('either lacks it entirely or exposes it only as an experimental');
  });

  it('starts rather than refusing when the version string is unreadable', () => {
    // A parse failure is not evidence of an old runtime. Refusing here would
    // strand a working install over an unexpected version format.
    expect(nodeFloorError('')).toBeNull();
    expect(nodeFloorError('not-a-version')).toBeNull();
  });

  it('matches the floor declared in engines.node', () => {
    const repoRoot = fileURLToPath(new URL('../../', import.meta.url));
    const pkg = JSON.parse(readFileSync(join(repoRoot, 'package.json'), 'utf8')) as {
      engines: { node: string };
    };
    const declared = Number(/(\d+)/.exec(pkg.engines.node)?.[1]);
    expect(MIN_NODE_MAJOR).toBe(declared);
  });
});
