/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Runtime-floor check for the executable entry (#76).
 *
 * Kept as its own module for two reasons: it must import nothing (the entry
 * point cannot pull in anything that reaches `node:sqlite` before the check
 * runs), and it must be testable without executing the entry point, which
 * would start a server.
 */

/** Lowest Node major this build supports; mirrors `engines.node`. */
export const MIN_NODE_MAJOR = 24;

/**
 * Returns an actionable error message when `version` cannot run this build, or
 * null when it can.
 *
 * @param version a `process.versions.node`-style string, e.g. `'22.15.0'`
 */
export function nodeFloorError(version: string): string | null {
  const major = Number.parseInt(version.split('.')[0] ?? '', 10);
  // An unparseable version is not evidence of an old runtime — let it through
  // rather than refusing to start on a version string we failed to read.
  if (Number.isNaN(major) || major >= MIN_NODE_MAJOR) return null;

  return (
    `mcp-office365 requires Node.js ${MIN_NODE_MAJOR} or newer — this is Node.js ${version}.\n` +
    `\n` +
    // Deliberately covers both sub-24 cases without claiming the wrong one:
    // below 22.13 the module is absent entirely, and 22.13+ exposes it only as
    // an experimental feature that warns on every launch.
    `The durable state store uses the built-in node:sqlite module. Node.js below\n` +
    `${MIN_NODE_MAJOR} either lacks it entirely or exposes it only as an experimental feature\n` +
    `that warns on every launch. Nothing is wrong with your install.\n` +
    `\n` +
    `Fix one of:\n` +
    `  - Upgrade to Node.js ${MIN_NODE_MAJOR}+ and relaunch\n` +
    `  - Point your MCP client at a Node.js ${MIN_NODE_MAJOR}+ binary (an absolute path\n` +
    `    avoids version-manager shims the client cannot see)\n` +
    `  - Pin the previous major, which supports Node.js 20+:\n` +
    `      npm install @jbctechsolutions/mcp-office365@4\n`
  );
}
