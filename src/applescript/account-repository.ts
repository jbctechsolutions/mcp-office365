/**
 * Repository for Outlook account operations using AppleScript.
 */

import { executeAppleScriptOrThrow } from './executor.js';
import { LIST_ACCOUNTS, GET_DEFAULT_ACCOUNT, listMailFoldersByAccounts } from './account-scripts.js';
import {
  parseAccounts,
  parseDefaultAccountId,
  parseFoldersWithAccount,
  type AppleScriptAccountRow,
  type AppleScriptFolderWithAccountRow,
} from './parser.js';

// =============================================================================
// Account Repository Interface
// =============================================================================

export interface IAccountRepository {
  /**
   * Lists all Exchange accounts configured in Outlook.
   */
  listAccounts(): AppleScriptAccountRow[];

  /**
   * Gets the ID of the default account.
   * Returns null if no default is set or no accounts exist.
   */
  getDefaultAccountId(): number | null;

  /**
   * Lists mail folders for specific accounts.
   */
  listMailFoldersByAccounts(accountIds: number[]): AppleScriptFolderWithAccountRow[];
}

// =============================================================================
// Implementation
// =============================================================================

export class AccountRepository implements IAccountRepository {
  listAccounts(): AppleScriptAccountRow[] {
    const output = executeAppleScriptOrThrow(LIST_ACCOUNTS);
    return parseAccounts(output);
  }

  getDefaultAccountId(): number | null {
    const output = executeAppleScriptOrThrow(GET_DEFAULT_ACCOUNT);
    return parseDefaultAccountId(output);
  }

  listMailFoldersByAccounts(accountIds: number[]): AppleScriptFolderWithAccountRow[] {
    if (accountIds.length === 0) {
      return [];
    }
    const script = listMailFoldersByAccounts(accountIds);
    const output = executeAppleScriptOrThrow(script);
    return parseFoldersWithAccount(output);
  }
}

/**
 * Creates an account repository instance.
 */
export function createAccountRepository(): IAccountRepository {
  return new AccountRepository();
}
