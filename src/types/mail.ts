/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail-related type definitions.
 */

/**
 * Special folder types in Outlook for Mac.
 */
export const SpecialFolderType = {
  Inbox: 1,
  Outbox: 2,
  Calendar: 4,
  Sent: 8,
  Deleted: 9,
  Drafts: 10,
  Junk: 12,
} as const;

export type SpecialFolderTypeValue =
  (typeof SpecialFolderType)[keyof typeof SpecialFolderType];

/**
 * Email priority levels.
 */
export const Priority = {
  High: 1,
  Normal: 3,
  Low: 5,
} as const;

export type PriorityValue = (typeof Priority)[keyof typeof Priority];

/**
 * Email flag status.
 */
export const FlagStatus = {
  None: 0,
  Flagged: 1,
  Completed: 2,
} as const;

export type FlagStatusValue = (typeof FlagStatus)[keyof typeof FlagStatus];

/**
 * Mail folder with message counts.
 */
export interface Folder {
  readonly id: number;
  readonly name: string;
  readonly parentId: number | null;
  readonly specialType: number;
  readonly folderType: number;
  readonly accountId: number;
  readonly messageCount: number;
  readonly unreadCount: number;
}

/**
 * Email summary for list views.
 */
export interface EmailSummary {
  readonly id: number;
  readonly folderId: number;
  readonly subject: string | null;
  readonly sender: string | null;
  readonly senderAddress: string | null;
  readonly preview: string | null;
  readonly isRead: boolean;
  readonly timeReceived: string | null;
  readonly timeSent: string | null;
  readonly hasAttachment: boolean;
  readonly priority: PriorityValue;
  readonly flagStatus: FlagStatusValue;
  readonly categories: readonly string[];
}

/**
 * Full email details including body content.
 */
export interface Email extends EmailSummary {
  readonly recipients: string | null;
  readonly displayTo: string | null;
  readonly toAddresses: string | null;
  readonly ccAddresses: string | null;
  readonly size: number;
  readonly messageId: string | null;
  readonly conversationId: number | null;
  readonly body: string | null;
  readonly htmlBody: string | null;
}

/**
 * Unread count result.
 */
export interface UnreadCount {
  readonly total: number;
  readonly byFolder?: Record<number, number>;
}
