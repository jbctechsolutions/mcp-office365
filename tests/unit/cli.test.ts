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
  mockGetTokenCacheFile: vi.fn(() => '/home/user/.outlook-mcp/tokens.json'),
}));

vi.mock('../../src/graph/index.js', () => ({
  getAccessToken: mockGetAccessToken,
  isAuthenticated: mockIsAuthenticated,
  getAccount: mockGetAccount,
  signOut: mockSignOut,
  getTokenCacheFile: mockGetTokenCacheFile,
}));

import { handleAuthCommand, parseCliCommand, createAuthMutex } from '../../src/cli.js';

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
