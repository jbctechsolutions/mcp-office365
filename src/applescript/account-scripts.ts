/**
 * AppleScript template strings for Outlook account operations.
 *
 * Outputs data in delimiter-based format for reliable parsing.
 */

import { DELIMITERS } from './scripts.js';

// =============================================================================
// Account Scripts
// =============================================================================

/**
 * Lists all accounts in Outlook (Exchange, IMAP, POP).
 * Returns: id, name, email address, type for each account
 */
export const LIST_ACCOUNTS = `
tell application "Microsoft Outlook"
  set output to ""

  -- Get Exchange accounts
  try
    set exchangeAccounts to every exchange account
    repeat with acc in exchangeAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "exchange"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  -- Get IMAP accounts
  try
    set imapAccounts to every imap account
    repeat with acc in imapAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "imap"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  -- Get POP accounts
  try
    set popAccounts to every pop account
    repeat with acc in popAccounts
      try
        set accId to id of acc
        set accName to name of acc
        set accEmail to email address of acc
        set accType to "pop"
        set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & accId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & accName & "${DELIMITERS.FIELD}email${DELIMITERS.EQUALS}" & accEmail & "${DELIMITERS.FIELD}type${DELIMITERS.EQUALS}" & accType
      end try
    end repeat
  end try

  return output
end tell
`;

/**
 * Gets the default account in Outlook.
 * Returns: id of the default account, or first account if no default is set
 */
export const GET_DEFAULT_ACCOUNT = `
tell application "Microsoft Outlook"
  try
    -- Try to get the default account
    set defaultAcc to default account
    set accId to id of defaultAcc
    return "id${DELIMITERS.EQUALS}" & accId
  on error
    -- Fallback: return first available account (try exchange, then imap, then pop)
    try
      set firstAcc to first exchange account
      set accId to id of firstAcc
      return "id${DELIMITERS.EQUALS}" & accId
    on error
      try
        set firstAcc to first imap account
        set accId to id of firstAcc
        return "id${DELIMITERS.EQUALS}" & accId
      on error
        try
          set firstAcc to first pop account
          set accId to id of firstAcc
          return "id${DELIMITERS.EQUALS}" & accId
        on error
          return "error${DELIMITERS.EQUALS}No accounts found"
        end try
      end try
    end try
  end try
end tell
`;

/**
 * Lists mail folders for specific accounts (all types: Exchange, IMAP, POP).
 *
 * @param accountIds - Array of account IDs to query
 */
export function listMailFoldersByAccounts(accountIds: number[]): string {
  const accountFilter = accountIds.map(id => `id ${id}`).join(' or id ');

  return `
tell application "Microsoft Outlook"
  set output to ""

  -- Get folders from Exchange accounts
  try
    set targetAccounts to (every exchange account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  -- Get folders from IMAP accounts
  try
    set targetAccounts to (every imap account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  -- Get folders from POP accounts
  try
    set targetAccounts to (every pop account whose ${accountFilter})
    repeat with acc in targetAccounts
      set accId to id of acc
      set allFolders to mail folders of acc
      repeat with f in allFolders
        try
          set fId to id of f
          set fName to name of f
          set uCount to unread count of f
          set mCount to 0
          try
            set mCount to count of messages of f
          end try
          set output to output & "${DELIMITERS.RECORD}id${DELIMITERS.EQUALS}" & fId & "${DELIMITERS.FIELD}name${DELIMITERS.EQUALS}" & fName & "${DELIMITERS.FIELD}unreadCount${DELIMITERS.EQUALS}" & uCount & "${DELIMITERS.FIELD}messageCount${DELIMITERS.EQUALS}" & mCount & "${DELIMITERS.FIELD}accountId${DELIMITERS.EQUALS}" & accId
        end try
      end repeat
    end repeat
  end try

  return output
end tell
`;
}
