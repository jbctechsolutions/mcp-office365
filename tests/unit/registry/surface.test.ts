/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * U10 surface tests against the REAL assembled registry (allToolDefinitions),
 * verifying `--preset` and `--read-only` produce the expected ListTools surface.
 */

import { describe, it, expect } from 'vitest';
import { ToolRegistry } from '../../../src/registry/registry.js';
import { allToolDefinitions } from '../../../src/registry/all-tools.js';

function graphRegistry(): ToolRegistry {
  const registry = new ToolRegistry();
  registry.register(allToolDefinitions());
  return registry;
}

describe('registry surface (U10)', () => {
  it('default (no flags) exposes the full graph surface — no shrink on upgrade', () => {
    const names = graphRegistry()
      .listTools({ backend: 'graph' })
      .map((t) => t.name);
    // Sanity floor — the full v3 graph surface is large (~218 tools).
    expect(names.length).toBeGreaterThan(200);
    expect(names).toContain('list_emails');
    expect(names).toContain('list_events');
    expect(names).toContain('list_teams');
  });

  it('--preset mail,calendar exposes the union and drops other domains', () => {
    const names = graphRegistry()
      .listTools({ backend: 'graph', presets: ['mail', 'calendar'] })
      .map((t) => t.name);
    expect(names).toContain('list_emails');
    expect(names).toContain('list_events');
    // A Teams tool is not in the mail/calendar union.
    expect(names).not.toContain('list_teams');
    expect(names).not.toContain('list_plans');
  });

  it('--read-only exposes zero destructive tools and zero prepare_/confirm_ tools', () => {
    const registry = graphRegistry();
    const readOnly = registry.listTools({ backend: 'graph', readOnly: true });
    const names = readOnly.map((t) => t.name);

    expect(names.some((n) => n.startsWith('prepare_'))).toBe(false);
    expect(names.some((n) => n.startsWith('confirm_'))).toBe(false);

    // Cross-check against the definitions: no destructive tool survives.
    for (const name of names) {
      expect(registry.get(name)?.destructive ?? false).toBe(false);
    }
    // ...and a known destructive tool is gone.
    expect(names).not.toContain('confirm_send_email');
    // A plain read tool is still present.
    expect(names).toContain('list_emails');
  });

  it('every read-only-surface tool advertises readOnlyHint:true', () => {
    const readOnly = graphRegistry().listTools({ backend: 'graph', readOnly: true });
    for (const tool of readOnly) {
      expect(tool.annotations?.readOnlyHint).toBe(true);
    }
  });

  it('every tool in the full surface carries annotations', () => {
    const tools = graphRegistry().listTools({ backend: 'graph' });
    const missing = tools.filter((t) => t.annotations == null).map((t) => t.name);
    expect(missing).toEqual([]);
  });
});
