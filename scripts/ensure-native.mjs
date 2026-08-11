#!/usr/bin/env node
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Auto-heals a better-sqlite3 native ABI mismatch before dev/test runs.
 *
 * The recurring failure (#88, #95, and again on 2026-08-10) is always the same
 * shape: the Node that compiled the binding is not the Node now loading it, so
 * `dlopen` rejects it with a NODE_MODULE_VERSION mismatch. On a dev machine the
 * usual trigger is a `.tool-versions` repin — #99 moved the pin from 26 to 24
 * LTS — with no reinstall in between, leaving `node_modules` built for the old
 * ABI. #99 made that failure legible; this makes it self-correcting.
 *
 * Rebuilding is safe and idempotent: if the binding already matches, this exits
 * immediately without touching anything.
 *
 * Deliberately NOT wired to `prestart` or `postinstall` — `scripts/` is not in
 * the package's published `files`, so those hooks would break consumers of the
 * npm package. This is a dev/CI-only guard.
 */

import { execSync } from 'node:child_process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);

// Any binding that will not load is rebuild-fixable, not just a version-number
// mismatch: a truncated or wrong-arch artifact ("slice is not valid mach-o file")
// and an absent one ("Could not locate the bindings file", which is what an
// --ignore-scripts install leaves behind) both heal the same way. Note that
// ERR_DLOPEN_FAILED arrives on err.code, and the bindings-package miss carries
// no code at all — so classification reads code and message together.
const REBUILDABLE =
  /NODE_MODULE_VERSION|ERR_DLOPEN_FAILED|compiled against a different Node|invalid ELF|mach-o|Could not locate the bindings file/i;

// The package itself is absent (no node_modules at all). `npm install` is the
// right fix and compiles the binding for this Node on its own — distinct from
// the package being present with its binding missing, which is rebuildable.
const NOT_INSTALLED = /Cannot find module '?better-sqlite3|MODULE_NOT_FOUND/i;

/**
 * Returns the load error, or null when the binding loads cleanly.
 *
 * Must construct a Database, not merely require the module: better-sqlite3
 * dlopens the binding lazily on first construction, so a plain `require` of a
 * package with a broken binding still succeeds. That is exactly why the
 * original crash surfaced inside `StateStore.open` rather than at import.
 */
function loadError() {
  try {
    const Database = require('better-sqlite3');
    new Database(':memory:').close();
    return null;
  } catch (err) {
    return err;
  }
}

const err = loadError();
if (err == null) {
  process.exit(0);
}

const message = `${err?.code ?? ''} ${err?.message ?? err}`;

// Dependencies simply are not installed yet — `npm install` is the right fix,
// and it will compile the binding for this Node on its own. Not our problem.
if (NOT_INSTALLED.test(message)) {
  process.exit(0);
}

if (!REBUILDABLE.test(message)) {
  console.error('better-sqlite3 failed to load, and a rebuild will not fix it:');
  console.error(message);
  process.exit(1);
}

console.error(
  `better-sqlite3's native binding will not load under Node ${process.version} ` +
  `(ABI ${process.versions.modules}) — rebuilding...`,
);

try {
  execSync('npm rebuild better-sqlite3', { stdio: 'inherit' });
} catch {
  console.error('`npm rebuild better-sqlite3` failed. Try `rm -rf node_modules && npm install`.');
  process.exit(1);
}

const after = loadError();
if (after != null) {
  console.error('Rebuild completed but the binding still will not load:');
  console.error(String(after?.message ?? after));
  console.error('Try `rm -rf node_modules && npm install` under the pinned Node in .tool-versions.');
  process.exit(1);
}

console.error(`better-sqlite3 rebuilt for Node ${process.version}.`);
