/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for signed-in account identity (U5 / D7).
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import type { AccountInfo } from '@azure/msal-node';

const { mockGetAccount } = vi.hoisted(() => ({
  mockGetAccount: vi.fn<() => Promise<AccountInfo | null>>(),
}));
vi.mock('../../../../src/graph/auth/device-code-flow.js', () => ({
  getAccount: mockGetAccount,
}));

import {
  DEFAULT_ACCOUNT_ID,
  resolveAccountId,
  currentAccountId,
  resetAccountId,
} from '../../../../src/graph/auth/account-id.js';

function account(homeAccountId: string): AccountInfo {
  return {
    homeAccountId,
    environment: 'login.microsoftonline.com',
    tenantId: 'tenant',
    username: 'user@example.com',
    localAccountId: 'local',
  } as AccountInfo;
}

describe('graph/auth/account-id', () => {
  beforeEach(() => {
    resetAccountId();
    mockGetAccount.mockReset();
  });

  it('resolves to the MSAL homeAccountId when signed in', async () => {
    mockGetAccount.mockResolvedValue(account('oid.tid'));
    await expect(resolveAccountId()).resolves.toBe('oid.tid');
    expect(currentAccountId()).toBe('oid.tid');
  });

  it('falls back to DEFAULT_ACCOUNT_ID when unauthenticated', async () => {
    mockGetAccount.mockResolvedValue(null);
    await expect(resolveAccountId()).resolves.toBe(DEFAULT_ACCOUNT_ID);
    expect(currentAccountId()).toBe(DEFAULT_ACCOUNT_ID);
  });

  it('currentAccountId is the fallback before any resolve', () => {
    expect(currentAccountId()).toBe(DEFAULT_ACCOUNT_ID);
  });

  it('does not memoize the fallback — picks up the real id after sign-in', async () => {
    mockGetAccount.mockResolvedValueOnce(null);
    await expect(resolveAccountId()).resolves.toBe(DEFAULT_ACCOUNT_ID);

    mockGetAccount.mockResolvedValue(account('oid.tid'));
    await expect(resolveAccountId()).resolves.toBe('oid.tid');
    expect(currentAccountId()).toBe('oid.tid');
  });

  it('memoizes the real id — a second resolve does not re-query MSAL', async () => {
    mockGetAccount.mockResolvedValue(account('oid.tid'));
    await resolveAccountId();
    await resolveAccountId();
    expect(mockGetAccount).toHaveBeenCalledTimes(1);
  });

  it('treats an empty homeAccountId as unresolved (fallback, uncached)', async () => {
    mockGetAccount.mockResolvedValue(account(''));
    await expect(resolveAccountId()).resolves.toBe(DEFAULT_ACCOUNT_ID);
    expect(currentAccountId()).toBe(DEFAULT_ACCOUNT_ID);
  });

  it('resetAccountId clears the memo so a new account re-resolves', async () => {
    mockGetAccount.mockResolvedValue(account('first.tid'));
    await resolveAccountId();
    expect(currentAccountId()).toBe('first.tid');

    resetAccountId();
    mockGetAccount.mockResolvedValue(account('second.tid'));
    await expect(resolveAccountId()).resolves.toBe('second.tid');
  });
});
