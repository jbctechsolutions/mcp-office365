/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

// Use vi.hoisted so mocks are available when vi.mock factory runs (hoisted)
const {
  mockGetAccessToken,
  mockIsAuthenticated,
  mockGetAccount,
  mockSignOut,
  mockGetTokenCacheFile,
} = vi.hoisted(() => ({
  mockGetAccessToken: vi.fn(),
  mockIsAuthenticated: vi.fn(),
  mockGetAccount: vi.fn(),
  mockSignOut: vi.fn(),
  mockGetTokenCacheFile: vi.fn(() => '/home/user/.mcp-office365/tokens.json'),
}));

vi.mock('../../src/graph/index.js', () => ({
  getAccessToken: mockGetAccessToken,
  isAuthenticated: mockIsAuthenticated,
  getAccount: mockGetAccount,
  signOut: mockSignOut,
  getTokenCacheFile: mockGetTokenCacheFile,
}));

import {
  handleAuthCommand,
  parseCliCommand,
  parseServerOptions,
  parseServeOptions,
  VALID_PRESETS,
  createAuthMutex,
} from '../../src/cli.js';

describe('CLI Auth', () => {
  beforeEach(() => {
    vi.resetAllMocks();
  });

  describe('auth (no flags)', () => {
    it('runs device code flow and prints success', async () => {
      mockGetAccessToken.mockResolvedValue('test-token');
      mockGetAccount.mockResolvedValue({ username: 'user@example.com' });

      const output: string[] = [];
      const exit = await handleAuthCommand([], (msg) => output.push(msg));

      expect(mockGetAccessToken).toHaveBeenCalledOnce();
      expect(exit).toBe(0);
      expect(output.some(line => line.includes('user@example.com'))).toBe(true);
    });

    it('returns exit code 1 on auth failure', async () => {
      mockGetAccessToken.mockRejectedValue(new Error('Auth failed'));

      const output: string[] = [];
      const exit = await handleAuthCommand([], (msg) => output.push(msg));

      expect(exit).toBe(1);
      expect(output.some(line => line.includes('failed'))).toBe(true);
    });
  });

  describe('auth --status', () => {
    it('prints authenticated status when tokens exist', async () => {
      mockIsAuthenticated.mockResolvedValue(true);
      mockGetAccount.mockResolvedValue({ username: 'user@example.com' });

      const output: string[] = [];
      const exit = await handleAuthCommand(['--status'], (msg) => output.push(msg));

      expect(exit).toBe(0);
      expect(output.some(line => line.includes('Authenticated'))).toBe(true);
      expect(output.some(line => line.includes('user@example.com'))).toBe(true);
    });

    it('prints not authenticated when no tokens', async () => {
      mockIsAuthenticated.mockResolvedValue(false);

      const output: string[] = [];
      const exit = await handleAuthCommand(['--status'], (msg) => output.push(msg));

      expect(exit).toBe(1);
      expect(output.some(line => line.includes('Not authenticated'))).toBe(true);
    });
  });

  describe('auth --logout', () => {
    it('signs out and prints confirmation', async () => {
      mockSignOut.mockResolvedValue(undefined);

      const output: string[] = [];
      const exit = await handleAuthCommand(['--logout'], (msg) => output.push(msg));

      expect(mockSignOut).toHaveBeenCalledOnce();
      expect(exit).toBe(0);
      expect(output.some(line => line.includes('Signed out'))).toBe(true);
    });

    it('returns exit code 1 on signout failure', async () => {
      mockSignOut.mockRejectedValue(new Error('Signout failed'));

      const output: string[] = [];
      const exit = await handleAuthCommand(['--logout'], (msg) => output.push(msg));

      expect(exit).toBe(1);
    });
  });
});

describe('parseCliCommand', () => {
  it('returns null for no args (MCP server mode)', () => {
    expect(parseCliCommand([])).toBeNull();
  });

  it('returns auth command with no flags', () => {
    expect(parseCliCommand(['auth'])).toEqual({ command: 'auth', flags: [] });
  });

  it('returns auth command with --status flag', () => {
    expect(parseCliCommand(['auth', '--status'])).toEqual({ command: 'auth', flags: ['--status'] });
  });

  it('returns auth command with --logout flag', () => {
    expect(parseCliCommand(['auth', '--logout'])).toEqual({ command: 'auth', flags: ['--logout'] });
  });

  it('returns null for unknown commands', () => {
    expect(parseCliCommand(['unknown'])).toBeNull();
  });

  it('parses serve command with its flags', () => {
    expect(parseCliCommand(['serve'])).toEqual({ command: 'serve', flags: [] });
    expect(parseCliCommand(['serve', '--port', '8080'])).toEqual({
      command: 'serve',
      flags: ['--port', '8080'],
    });
  });
});

describe('parseServeOptions (U3)', () => {
  it('defaults to loopback:3000', () => {
    expect(parseServeOptions([])).toEqual({ host: '127.0.0.1', port: 3000 });
  });

  it('parses --host and --port (space and = forms)', () => {
    expect(parseServeOptions(['--host', '0.0.0.0', '--port', '8080'])).toEqual({
      host: '0.0.0.0',
      port: 8080,
    });
    expect(parseServeOptions(['--host=0.0.0.0', '--port=8080'])).toEqual({
      host: '0.0.0.0',
      port: 8080,
    });
  });

  it('ignores preset/read-only flags (parsed separately)', () => {
    expect(parseServeOptions(['--read-only', '--preset', 'mail', '--port', '9000'])).toEqual({
      host: '127.0.0.1',
      port: 9000,
    });
  });

  it('rejects an invalid or missing port', () => {
    expect(() => parseServeOptions(['--port', 'abc'])).toThrow(/integer 1-65535/);
    expect(() => parseServeOptions(['--port', '0'])).toThrow(/integer 1-65535/);
    expect(() => parseServeOptions(['--port', '70000'])).toThrow(/integer 1-65535/);
    expect(() => parseServeOptions(['--port'])).toThrow(/integer 1-65535/);
  });

  it('rejects a missing host value', () => {
    expect(() => parseServeOptions(['--host'])).toThrow(/--host requires/);
    expect(() => parseServeOptions(['--host', '--port'])).toThrow(/--host requires/);
  });
});

describe('createAuthMutex', () => {
  it('only runs the auth function once for concurrent calls', async () => {
    const authFn = vi.fn().mockResolvedValue(undefined);
    const mutex = createAuthMutex(authFn);

    // Call 3 times concurrently
    const results = await Promise.all([mutex(), mutex(), mutex()]);

    expect(authFn).toHaveBeenCalledOnce();
    expect(results).toEqual([undefined, undefined, undefined]);
  });

  it('propagates errors to all waiters', async () => {
    const authFn = vi.fn().mockRejectedValue(new Error('Auth failed'));
    const mutex = createAuthMutex(authFn);

    const results = await Promise.allSettled([mutex(), mutex(), mutex()]);

    expect(authFn).toHaveBeenCalledOnce();
    expect(results.every(r => r.status === 'rejected')).toBe(true);
  });

  it('allows retry after failure', async () => {
    const authFn = vi.fn()
      .mockRejectedValueOnce(new Error('First attempt failed'))
      .mockResolvedValueOnce(undefined);
    const mutex = createAuthMutex(authFn);

    // First call fails
    await expect(mutex()).rejects.toThrow('First attempt failed');

    // Second call succeeds (new attempt)
    await expect(mutex()).resolves.toBeUndefined();

    expect(authFn).toHaveBeenCalledTimes(2);
  });
});

describe('parseServerOptions (U10)', () => {
  it('defaults to the full surface with read-only off when no flags', () => {
    expect(parseServerOptions([])).toEqual({ readOnly: false, confirmMode: 'token' });
  });

  it('treats "all" (or --preset all) as the full surface', () => {
    expect(parseServerOptions(['--preset', 'all'])).toEqual({ readOnly: false, confirmMode: 'token' });
    // all wins even when combined with a specific preset
    expect(parseServerOptions(['--preset', 'all,mail'])).toEqual({ readOnly: false, confirmMode: 'token' });
  });

  it('parses a comma-separated preset list', () => {
    expect(parseServerOptions(['--preset', 'mail,calendar'])).toEqual({
      readOnly: false,
      confirmMode: 'token',
      presets: ['mail', 'calendar'],
    });
  });

  it('supports --preset=<names> and repeated --preset flags', () => {
    expect(parseServerOptions(['--preset=mail', '--preset', 'tasks'])).toEqual({
      readOnly: false,
      confirmMode: 'token',
      presets: ['mail', 'tasks'],
    });
  });

  it('parses --read-only', () => {
    expect(parseServerOptions(['--read-only'])).toEqual({ readOnly: true, confirmMode: 'token' });
    expect(parseServerOptions(['--read-only', '--preset', 'mail'])).toEqual({
      readOnly: true,
      confirmMode: 'token',
      presets: ['mail'],
    });
  });

  it('throws with the valid list on an unknown preset', () => {
    expect(() => parseServerOptions(['--preset', 'mail,nope'])).toThrow(/Unknown preset\(s\): nope/);
    expect(() => parseServerOptions(['--preset', 'nope'])).toThrow(/Valid presets:/);
  });

  it('validates co-listed names even when "all" is present (no silent swallow)', () => {
    // `all` must not short-circuit past validation of a typo'd sibling.
    expect(() => parseServerOptions(['--preset', 'all,bogus'])).toThrow(/Unknown preset\(s\): bogus/);
    expect(() => parseServerOptions(['--preset', 'mial,all'])).toThrow(/Unknown preset\(s\): mial/);
  });

  it('throws when --preset has no value', () => {
    expect(() => parseServerOptions(['--preset'])).toThrow(/requires a comma-separated list/);
    expect(() => parseServerOptions(['--preset', '--read-only'])).toThrow(/requires/);
  });

  it('throws (does not fail open to full surface) on an empty/whitespace preset value', () => {
    // Regression: these previously collapsed to the FULL surface silently.
    expect(() => parseServerOptions(['--preset', ''])).toThrow(/requires/);
    expect(() => parseServerOptions(['--preset', '   '])).toThrow(/requires/);
    expect(() => parseServerOptions(['--preset', ',,,'])).toThrow(/requires/);
    expect(() => parseServerOptions(['--preset='])).toThrow(/requires/);
  });

  it('ignores unknown args (runner-injected argv)', () => {
    expect(parseServerOptions(['--inspect', 'foo', '--read-only'])).toEqual({ readOnly: true, confirmMode: 'token' });
  });

  it('defaults confirmMode to "token" (U11)', () => {
    expect(parseServerOptions([]).confirmMode).toBe('token');
  });

  it('parses --confirm elicit and --confirm=elicit (U11)', () => {
    expect(parseServerOptions(['--confirm', 'elicit']).confirmMode).toBe('elicit');
    expect(parseServerOptions(['--confirm=elicit']).confirmMode).toBe('elicit');
    expect(parseServerOptions(['--confirm', 'token']).confirmMode).toBe('token');
    expect(parseServerOptions(['--read-only', '--confirm', 'elicit'])).toEqual({
      readOnly: true,
      confirmMode: 'elicit',
    });
  });

  it('throws on an invalid or missing --confirm value (U11)', () => {
    expect(() => parseServerOptions(['--confirm', 'nope'])).toThrow(/--confirm requires one of: token, elicit/);
    expect(() => parseServerOptions(['--confirm'])).toThrow(/--confirm requires one of: token, elicit/);
    expect(() => parseServerOptions(['--confirm='])).toThrow(/--confirm requires one of: token, elicit/);
  });

  it('every valid preset name is accepted', () => {
    for (const p of VALID_PRESETS) {
      expect(parseServerOptions(['--preset', p]).presets).toEqual([p]);
    }
  });
});
