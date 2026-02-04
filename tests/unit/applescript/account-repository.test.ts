import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../../../src/applescript/executor.js', () => ({
  executeAppleScriptOrThrow: vi.fn(),
}));

vi.mock('../../../src/applescript/parser.js', () => ({
  parseAccounts: vi.fn(),
  parseDefaultAccountId: vi.fn(),
  parseFoldersWithAccount: vi.fn(),
}));

vi.mock('../../../src/applescript/account-scripts.js', () => ({
  LIST_ACCOUNTS: 'mock-list-accounts-script',
  GET_DEFAULT_ACCOUNT: 'mock-get-default-script',
  listMailFoldersByAccounts: vi.fn(() => 'mock-folders-script'),
}));

import { AccountRepository, createAccountRepository } from '../../../src/applescript/account-repository.js';
import { executeAppleScriptOrThrow } from '../../../src/applescript/executor.js';
import { parseAccounts, parseDefaultAccountId, parseFoldersWithAccount } from '../../../src/applescript/parser.js';
import { listMailFoldersByAccounts as listMailFoldersByAccountsScript } from '../../../src/applescript/account-scripts.js';

const mockedExecute = vi.mocked(executeAppleScriptOrThrow);
const mockedParseAccounts = vi.mocked(parseAccounts);
const mockedParseDefaultAccountId = vi.mocked(parseDefaultAccountId);
const mockedParseFoldersWithAccount = vi.mocked(parseFoldersWithAccount);
const mockedListMailFoldersScript = vi.mocked(listMailFoldersByAccountsScript);

describe('AccountRepository', () => {
  let repo: AccountRepository;

  beforeEach(() => {
    vi.clearAllMocks();
    repo = new AccountRepository();
  });

  describe('listAccounts', () => {
    it('calls executor with LIST_ACCOUNTS script and parses result', () => {
      const mockAccounts = [{ id: 1, name: 'Work', email: 'test@example.com', type: 'exchange' as const }];
      mockedExecute.mockReturnValue('raw output');
      mockedParseAccounts.mockReturnValue(mockAccounts);

      const result = repo.listAccounts();

      expect(mockedExecute).toHaveBeenCalledWith('mock-list-accounts-script');
      expect(mockedParseAccounts).toHaveBeenCalledWith('raw output');
      expect(result).toEqual(mockAccounts);
    });
  });

  describe('getDefaultAccountId', () => {
    it('calls executor with GET_DEFAULT_ACCOUNT script and parses result', () => {
      mockedExecute.mockReturnValue('id=42');
      mockedParseDefaultAccountId.mockReturnValue(42);

      const result = repo.getDefaultAccountId();

      expect(mockedExecute).toHaveBeenCalledWith('mock-get-default-script');
      expect(mockedParseDefaultAccountId).toHaveBeenCalledWith('id=42');
      expect(result).toBe(42);
    });

    it('returns null when parser returns null', () => {
      mockedExecute.mockReturnValue('error=No accounts found');
      mockedParseDefaultAccountId.mockReturnValue(null);

      const result = repo.getDefaultAccountId();
      expect(result).toBeNull();
    });
  });

  describe('listMailFoldersByAccounts', () => {
    it('returns empty array for empty account IDs', () => {
      const result = repo.listMailFoldersByAccounts([]);

      expect(result).toEqual([]);
      expect(mockedExecute).not.toHaveBeenCalled();
    });

    it('generates script and parses result for non-empty account IDs', () => {
      const mockFolders = [
        { id: 10, name: 'Inbox', unreadCount: 5, messageCount: 100, accountId: 1 },
      ];
      mockedListMailFoldersScript.mockReturnValue('generated-script');
      mockedExecute.mockReturnValue('raw folders');
      mockedParseFoldersWithAccount.mockReturnValue(mockFolders);

      const result = repo.listMailFoldersByAccounts([1, 2]);

      expect(mockedListMailFoldersScript).toHaveBeenCalledWith([1, 2]);
      expect(mockedExecute).toHaveBeenCalledWith('generated-script');
      expect(mockedParseFoldersWithAccount).toHaveBeenCalledWith('raw folders');
      expect(result).toEqual(mockFolders);
    });
  });
});

describe('createAccountRepository', () => {
  it('returns an AccountRepository instance', () => {
    const repo = createAccountRepository();
    expect(repo).toBeInstanceOf(AccountRepository);
  });
});
