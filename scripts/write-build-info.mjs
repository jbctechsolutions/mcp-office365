/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Stamps dist/build-info.json at build time so the server reports the version
 * it was BUILT from, not whatever package.json says at runtime. A stale dist
 * next to a newer package.json (npm-link dev setups) otherwise misreports its
 * version, which masks "you are running an old build" during debugging.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join, dirname } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

writeFileSync(
  join(root, 'dist', 'build-info.json'),
  JSON.stringify({ version: pkg.version, builtAt: new Date().toISOString() }, null, 2) + '\n',
);
