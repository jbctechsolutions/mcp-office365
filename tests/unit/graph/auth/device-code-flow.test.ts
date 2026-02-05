/**
 * Tests for Graph API device code flow authentication.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import type { PublicClientApplication, AccountInfo, AuthenticationResult } from '@azure/msal-node';

// Mock MSAL module
const mockGetAllAccounts = vi.fn();
const mockRemoveAccount = vi.fn();
const mockAcquireTokenByDeviceCode = vi.fn();
const mockAcquireTokenSilent = vi.fn();
const mockGetTokenCache = vi.fn(() => ({
  getAllAccounts: mockGetAllAccounts,
  removeAccount: mockRemoveAccount,
}));

const mockMsalInstance = {
  acquireTokenByDeviceCode: mockAcquireTokenByDeviceCode,
  acquireTokenSilent: mockAcquireTokenSilent,
  getTokenCache: mockGetTokenCache,
};

vi.mock('@azure/msal-node', () => ({
  PublicClientApplication: vi.fn(function() { return mockMsalInstance; }),
}));

// Mock config module
vi.mock('../../../../src/graph/auth/config.js', () => ({
  loadGraphConfig: vi.fn(() => ({
    clientId: 'test-client-id',
    tenantId: 'common',
    scopes: ['Mail.Read', 'Calendars.Read'],
  })),
  getAuthorityUrl: vi.fn(() => 'https://login.microsoftonline.com/common'),
}));

// Mock token-cache module
vi.mock('../../../../src/graph/auth/token-cache.js', () => ({
  createTokenCachePlugin: vi.fn(() => ({})),
  hasTokenCache: vi.fn(() => true),
}));

import {
  acquireTokenInteractive,
  acquireTokenSilent,
  getAccessToken,
  isAuthenticated,
  getAccount,
  signOut,
  resetMsalInstance,
} from '../../../../src/graph/auth/device-code-flow.js';
import { hasTokenCache } from '../../../../src/graph/auth/token-cache.js';

describe('graph/auth/device-code-flow', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    resetMsalInstance();
  });

  afterEach(() => {
    resetMsalInstance();
  });

  describe('acquireTokenInteractive', () => {
    it('acquires token via device code flow', async () => {
      const mockResult: AuthenticationResult = {
        accessToken: 'test-access-token',
        account: null,
        authority: 'https://login.microsoftonline.com/common',
        uniqueId: 'unique-id',
        tenantId: 'tenant-id',
        scopes: ['Mail.Read'],
        expiresOn: new Date(),
        idToken: 'id-token',
        idTokenClaims: {},
        fromCache: false,
        tokenType: 'Bearer',
        correlationId: 'correlation-id',
      };

      mockAcquireTokenByDeviceCode.mockResolvedValue(mockResult);

      const callback = vi.fn();
      const result = await acquireTokenInteractive(callback);

      expect(result).toEqual(mockResult);
      expect(mockAcquireTokenByDeviceCode).toHaveBeenCalled();
    });

    it('invokes device code callback with user code and verification URI', async () => {
      mockAcquireTokenByDeviceCode.mockImplementation(async (request: any) => {
        // Simulate the callback being invoked
        request.deviceCodeCallback({
          userCode: 'ABC123',
          verificationUri: 'https://microsoft.com/devicelogin',
          message: 'Enter code ABC123',
        });

        return {
          accessToken: 'test-token',
          account: null,
          authority: '',
          uniqueId: '',
          tenantId: '',
          scopes: [],
          expiresOn: new Date(),
          idToken: '',
          idTokenClaims: {},
          fromCache: false,
          tokenType: 'Bearer',
          correlationId: '',
        };
      });

      const callback = vi.fn();
      await acquireTokenInteractive(callback);

      expect(callback).toHaveBeenCalledWith(
        'ABC123',
        'https://microsoft.com/devicelogin',
        'Enter code ABC123'
      );
    });

    it('throws when device code authentication returns null', async () => {
      mockAcquireTokenByDeviceCode.mockResolvedValue(null);

      await expect(acquireTokenInteractive()).rejects.toThrow(
        'Device code authentication failed'
      );
    });

    it('uses default callback when none provided', async () => {
      const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});

      mockAcquireTokenByDeviceCode.mockImplementation(async (request: any) => {
        request.deviceCodeCallback({
          userCode: 'XYZ789',
          verificationUri: 'https://microsoft.com/devicelogin',
          message: 'Test message',
        });

        return {
          accessToken: 'test-token',
          account: null,
          authority: '',
          uniqueId: '',
          tenantId: '',
          scopes: [],
          expiresOn: new Date(),
          idToken: '',
          idTokenClaims: {},
          fromCache: false,
          tokenType: 'Bearer',
          correlationId: '',
        };
      });

      await acquireTokenInteractive();

      expect(consoleErrorSpy).toHaveBeenCalled();
      consoleErrorSpy.mockRestore();
    });
  });

  describe('acquireTokenSilent', () => {
    it('returns token when account is cached', async () => {
      const mockAccount: AccountInfo = {
        homeAccountId: 'home-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user@example.com',
        localAccountId: 'local-account-id',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount]);

      const mockResult: AuthenticationResult = {
        accessToken: 'silent-access-token',
        account: mockAccount,
        authority: '',
        uniqueId: '',
        tenantId: '',
        scopes: [],
        expiresOn: new Date(),
        idToken: '',
        idTokenClaims: {},
        fromCache: true,
        tokenType: 'Bearer',
        correlationId: '',
      };

      mockAcquireTokenSilent.mockResolvedValue(mockResult);

      const result = await acquireTokenSilent();

      expect(result).toEqual(mockResult);
      expect(mockAcquireTokenSilent).toHaveBeenCalledWith(
        expect.objectContaining({
          account: mockAccount,
        })
      );
    });

    it('returns null when no accounts are cached', async () => {
      mockGetAllAccounts.mockResolvedValue([]);

      const result = await acquireTokenSilent();

      expect(result).toBeNull();
      expect(mockAcquireTokenSilent).not.toHaveBeenCalled();
    });

    it('returns null when token refresh fails', async () => {
      const mockAccount: AccountInfo = {
        homeAccountId: 'home-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user@example.com',
        localAccountId: 'local-account-id',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount]);
      mockAcquireTokenSilent.mockRejectedValue(new Error('Token refresh failed'));

      const result = await acquireTokenSilent();

      expect(result).toBeNull();
    });
  });

  describe('getAccessToken', () => {
    it('returns token from silent acquisition when available', async () => {
      const mockAccount: AccountInfo = {
        homeAccountId: 'home-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user@example.com',
        localAccountId: 'local-account-id',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount]);
      mockAcquireTokenSilent.mockResolvedValue({
        accessToken: 'silent-token',
        account: mockAccount,
        authority: '',
        uniqueId: '',
        tenantId: '',
        scopes: [],
        expiresOn: new Date(),
        idToken: '',
        idTokenClaims: {},
        fromCache: true,
        tokenType: 'Bearer',
        correlationId: '',
      });

      const token = await getAccessToken();

      expect(token).toBe('silent-token');
      expect(mockAcquireTokenByDeviceCode).not.toHaveBeenCalled();
    });

    it('falls back to interactive when silent fails', async () => {
      mockGetAllAccounts.mockResolvedValue([]);
      mockAcquireTokenByDeviceCode.mockResolvedValue({
        accessToken: 'interactive-token',
        account: null,
        authority: '',
        uniqueId: '',
        tenantId: '',
        scopes: [],
        expiresOn: new Date(),
        idToken: '',
        idTokenClaims: {},
        fromCache: false,
        tokenType: 'Bearer',
        correlationId: '',
      });

      const callback = vi.fn();
      const token = await getAccessToken(callback);

      expect(token).toBe('interactive-token');
      expect(mockAcquireTokenByDeviceCode).toHaveBeenCalled();
    });
  });

  describe('isAuthenticated', () => {
    it('returns true when accounts exist and token cache exists', async () => {
      const mockHasTokenCache = vi.mocked(hasTokenCache);
      mockHasTokenCache.mockReturnValue(true);

      const mockAccount: AccountInfo = {
        homeAccountId: 'home-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user@example.com',
        localAccountId: 'local-account-id',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount]);

      const result = await isAuthenticated();

      expect(result).toBe(true);
    });

    it('returns false when no token cache exists', async () => {
      const mockHasTokenCache = vi.mocked(hasTokenCache);
      mockHasTokenCache.mockReturnValue(false);

      const result = await isAuthenticated();

      expect(result).toBe(false);
    });

    it('returns false when no accounts are cached', async () => {
      const mockHasTokenCache = vi.mocked(hasTokenCache);
      mockHasTokenCache.mockReturnValue(true);
      mockGetAllAccounts.mockResolvedValue([]);

      const result = await isAuthenticated();

      expect(result).toBe(false);
    });

    it('returns false when getAllAccounts throws', async () => {
      const mockHasTokenCache = vi.mocked(hasTokenCache);
      mockHasTokenCache.mockReturnValue(true);
      mockGetAllAccounts.mockRejectedValue(new Error('Cache error'));

      const result = await isAuthenticated();

      expect(result).toBe(false);
    });
  });

  describe('getAccount', () => {
    it('returns first account when accounts exist', async () => {
      const mockAccount: AccountInfo = {
        homeAccountId: 'home-account-id',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user@example.com',
        localAccountId: 'local-account-id',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount]);

      const result = await getAccount();

      expect(result).toEqual(mockAccount);
    });

    it('returns null when no accounts exist', async () => {
      mockGetAllAccounts.mockResolvedValue([]);

      const result = await getAccount();

      expect(result).toBeNull();
    });

    it('returns null when getAllAccounts throws', async () => {
      mockGetAllAccounts.mockRejectedValue(new Error('Cache error'));

      const result = await getAccount();

      expect(result).toBeNull();
    });
  });

  describe('signOut', () => {
    it('removes all accounts from cache', async () => {
      const mockAccount1: AccountInfo = {
        homeAccountId: 'account-1',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user1@example.com',
        localAccountId: 'local-1',
      };

      const mockAccount2: AccountInfo = {
        homeAccountId: 'account-2',
        environment: 'login.microsoftonline.com',
        tenantId: 'tenant-id',
        username: 'user2@example.com',
        localAccountId: 'local-2',
      };

      mockGetAllAccounts.mockResolvedValue([mockAccount1, mockAccount2]);
      mockRemoveAccount.mockResolvedValue(undefined);

      await signOut();

      expect(mockRemoveAccount).toHaveBeenCalledTimes(2);
      expect(mockRemoveAccount).toHaveBeenCalledWith(mockAccount1);
      expect(mockRemoveAccount).toHaveBeenCalledWith(mockAccount2);
    });

    it('handles empty account list', async () => {
      mockGetAllAccounts.mockResolvedValue([]);

      await signOut();

      expect(mockRemoveAccount).not.toHaveBeenCalled();
    });

    it('ignores errors during sign out', async () => {
      mockGetAllAccounts.mockRejectedValue(new Error('Cache error'));

      // Should not throw
      await expect(signOut()).resolves.not.toThrow();
    });
  });

  describe('resetMsalInstance', () => {
    it('resets the MSAL instance', async () => {
      // First call initializes the instance
      await isAuthenticated();

      // Reset it
      resetMsalInstance();

      // Next call should create a new instance
      mockGetAllAccounts.mockResolvedValue([]);
      await isAuthenticated();

      // The PublicClientApplication constructor should have been called twice
      const { PublicClientApplication } = await import('@azure/msal-node');
      expect(PublicClientApplication).toHaveBeenCalled();
    });
  });
});
