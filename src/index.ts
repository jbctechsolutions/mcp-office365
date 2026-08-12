#!/usr/bin/env node
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Executable entry point: check the runtime, then load the server.
 *
 * This file exists to be tiny. The durable state store imports `node:sqlite`,
 * which does not exist before Node 22.13 — and ESM resolves an entire module
 * graph before executing any of it, so a version check placed anywhere inside
 * the server would never run. Node would reject the graph first and print a raw
 * `ERR_UNKNOWN_BUILTIN_MODULE` stack, which an MCP host surfaces as nothing
 * more useful than "failed to connect" (#76).
 *
 * Hence: nothing is imported statically except the check itself, which imports
 * nothing. The real entry loads via `import()` afterwards.
 *
 * Why check at all when `engines.node` already declares `>=24`? Because npm
 * only *warns* on an unsatisfied `engines` by default — it refuses the install
 * only under `engine-strict=true`. An install on Node 20 or 22 therefore
 * succeeds and fails later, at launch, which is the worst place to find out.
 */

import { nodeFloorError } from './node-floor.js';

const floorError = nodeFloorError(process.versions.node);
if (floorError !== null) {
  process.stderr.write(floorError);
  // Set the code and let the process end on its own rather than calling
  // process.exit(). Node writes to stderr asynchronously when it is a pipe —
  // which it always is under an MCP host — so exiting here can terminate
  // before the buffer drains and truncate the very message this check exists
  // to deliver.
  process.exitCode = 1;
} else {
  // Dynamic on purpose — see the note above. A static import would be resolved
  // before the check above ever runs.
  const { main } = await import('./server.js');

  main().catch((error: unknown) => {
    // Same reasoning as the floor message above: process.exit() can terminate
    // before an asynchronous stderr drains, and this is the output that
    // explains why the server died. Set the code and let the process end.
    console.error('Fatal error:', error);
    process.exitCode = 1;
  });
}
