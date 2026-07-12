/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/** U6 per-user entitlements: resolution, allow/exclude filtering, drift, reload. */

import { afterEach, describe, expect, it } from 'vitest';
import { mkdtempSync, rmSync, writeFileSync, utimesSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  DEFAULT_TOOL_SURFACE,
  createEntitlementResolver,
} from '../../../src/remote/entitlements.js';
import { ToolRegistry } from '../../../src/registry/index.js';
import { allToolDefinitions } from '../../../src/registry/all-tools.js';

const dirs: string[] = [];
function tmpConfig(obj: unknown): string {
  const dir = mkdtempSync(join(tmpdir(), 'mcp-ent-'));
  dirs.push(dir);
  const p = join(dir, 'entitlements.json');
  writeFileSync(p, JSON.stringify(obj));
  return p;
}
afterEach(() => {
  while (dirs.length > 0) {
    const d = dirs.pop();
    if (d != null) rmSync(d, { recursive: true, force: true });
  }
});

describe('createEntitlementResolver (U6)', () => {
  it('gives an unconfigured user the pinned default surface', () => {
    const r = createEntitlementResolver();
    const s = r.resolve('anyone', 'graph');
    expect(s.allow).toBe(DEFAULT_TOOL_SURFACE);
    // Default excludes shared-mailbox and mail-rules tools.
    expect(s.allow).not.toContain('list_mail_rules');
    expect(s.allow).not.toContain('list_shared_mailbox_emails');
    // ...and includes the core v1 surface.
    expect(s.allow).toContain('list_folders');
    expect(s.allow).toContain('create_planner_task');
  });

  it('grants full access (no allow-list) to a fullAccess user', () => {
    const p = tmpConfig({ users: { joel: { fullAccess: true } } });
    const s = createEntitlementResolver(p).resolve('joel', 'graph');
    expect(s.allow).toBeUndefined();
  });

  it('honors a per-user explicit allow-list and exclusions', () => {
    const p = tmpConfig({
      users: { u1: { allow: ['list_folders', 'get_email'] }, u2: { exclude: ['send_email'] } },
    });
    const r = createEntitlementResolver(p);
    expect(r.resolve('u1', 'graph').allow).toEqual(['list_folders', 'get_email']);
    const u2 = r.resolve('u2', 'graph');
    expect(u2.allow).toBe(DEFAULT_TOOL_SURFACE);
    expect(u2.exclude).toEqual(['send_email']);
  });

  it('reloads on config change (edit takes effect on next resolve)', () => {
    const p = tmpConfig({ users: { u1: { allow: ['list_folders'] } } });
    const r = createEntitlementResolver(p);
    expect(r.resolve('u1', 'graph').allow).toEqual(['list_folders']);
    writeFileSync(p, JSON.stringify({ users: { u1: { allow: ['get_email'] } } }));
    utimesSync(p, new Date(), new Date(Date.now() + 1000)); // bump mtime
    expect(r.resolve('u1', 'graph').allow).toEqual(['get_email']);
  });

  it('fails safe to the default surface when the file is missing', () => {
    const r = createEntitlementResolver('/no/such/entitlements.json');
    expect(r.resolve('u1', 'graph').allow).toBe(DEFAULT_TOOL_SURFACE);
  });

  it('fails safe (never full surface) on malformed JSON', () => {
    const dir = mkdtempSync(join(tmpdir(), 'mcp-ent-'));
    dirs.push(dir);
    const p = join(dir, 'entitlements.json');
    writeFileSync(p, '{ not valid json');
    const s = createEntitlementResolver(p).resolve('u1', 'graph');
    expect(s.allow).toBe(DEFAULT_TOOL_SURFACE); // default, not undefined/full
  });

  it('rejects non-string allow/exclude elements (fails safe to default)', () => {
    const p = tmpConfig({ users: { u1: { allow: ['list_folders', 42] } } });
    // parse throws → last-good (empty) → default surface, never the bad allow.
    expect(createEntitlementResolver(p).resolve('u1', 'graph').allow).toBe(DEFAULT_TOOL_SURFACE);
  });

  it('honors an explicit empty allow-list (deny all), not the default', () => {
    const p = tmpConfig({ users: { u1: { allow: [] } } });
    expect(createEntitlementResolver(p).resolve('u1', 'graph').allow).toEqual([]);
  });

  it('applies exclude even for a fullAccess user', () => {
    const p = tmpConfig({ users: { joel: { fullAccess: true, exclude: ['send_email'] } } });
    const s = createEntitlementResolver(p).resolve('joel', 'graph');
    expect(s.allow).toBeUndefined();
    expect(s.exclude).toEqual(['send_email']);
  });
});

describe('DEFAULT_TOOL_SURFACE registry contract', () => {
  it('every pinned tool name still exists in the registry (drift guard)', () => {
    const registry = new ToolRegistry();
    registry.register(allToolDefinitions());
    const missing = DEFAULT_TOOL_SURFACE.filter((name) => !registry.has(name));
    expect(missing).toEqual([]);
  });

  it('the pinned surface excludes shared-mailbox, mail-rules, downloads, and photos', () => {
    const bad = DEFAULT_TOOL_SURFACE.filter(
      (n) => /^download_|_photo$/.test(n) || /mail_rule/.test(n) || /shared_mailbox|shared_calendar/.test(n),
    );
    expect(bad).toEqual([]);
  });
});

describe('registry allow/exclude filtering (U6)', () => {
  const registry = new ToolRegistry();
  registry.register(allToolDefinitions());

  it('allow-list mode exposes exactly the listed tools', () => {
    const tools = registry.listTools({ backend: 'graph', allow: ['list_folders', 'get_email'] });
    expect(tools.map((t) => t.name).sort()).toEqual(['get_email', 'list_folders']);
  });

  it('exclude removes a tool even from the full surface', () => {
    const full = registry.listTools({ backend: 'graph' });
    const filtered = registry.listTools({ backend: 'graph', exclude: ['send_email'] });
    expect(full.some((t) => t.name === 'send_email')).toBe(true);
    expect(filtered.some((t) => t.name === 'send_email')).toBe(false);
  });

  it('exclude wins over allow', () => {
    const tools = registry.listTools({
      backend: 'graph',
      allow: ['list_folders', 'send_email'],
      exclude: ['send_email'],
    });
    expect(tools.map((t) => t.name)).toEqual(['list_folders']);
  });

  it('composes the --preset outer bound with the allow-list (intersection)', () => {
    // allow spans mail + planner; a process-wide --preset mail cap must exclude
    // the planner tool even though it is allow-listed.
    const tools = registry.listTools({
      backend: 'graph',
      presets: ['mail'],
      allow: ['list_folders', 'list_plans'],
    });
    const names = tools.map((t) => t.name);
    expect(names).toContain('list_folders'); // mail-preset tool
    expect(names).not.toContain('list_plans'); // planner tool, outside the cap
  });

  it('a tool outside the allow-list is not exposed (dispatch invariant)', () => {
    expect(registry.isExposed('send_email', { backend: 'graph', allow: ['list_folders'] })).toBe(
      false,
    );
    expect(registry.isExposed('list_folders', { backend: 'graph', allow: ['list_folders'] })).toBe(
      true,
    );
  });
});
