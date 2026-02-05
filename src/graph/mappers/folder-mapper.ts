/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph MailFolder and Calendar types to FolderRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { FolderRow } from '../../database/repository.js';
import { hashStringToNumber } from './utils.js';

/**
 * Maps a Graph MailFolder to a FolderRow.
 */
export function mapMailFolderToRow(folder: MicrosoftGraph.MailFolder): FolderRow {
  return {
    id: hashStringToNumber(folder.id ?? ''),
    name: folder.displayName ?? null,
    parentId: folder.parentFolderId != null ? hashStringToNumber(folder.parentFolderId) : null,
    specialType: 0, // Graph doesn't expose special type directly
    folderType: 1, // Mail folder
    accountId: 1, // Default account
    messageCount: folder.totalItemCount ?? 0,
    unreadCount: folder.unreadItemCount ?? 0,
  };
}

/**
 * Maps a Graph Calendar to a FolderRow (calendars use FolderRow structure).
 */
export function mapCalendarToFolderRow(calendar: MicrosoftGraph.Calendar): FolderRow {
  return {
    id: hashStringToNumber(calendar.id ?? ''),
    name: calendar.name ?? null,
    parentId: null,
    specialType: 0,
    folderType: 2, // Calendar folder
    accountId: 1,
    messageCount: 0,
    unreadCount: 0,
  };
}

/**
 * Maps a Graph TodoTaskList to a FolderRow.
 */
export function mapTaskListToFolderRow(taskList: MicrosoftGraph.TodoTaskList): FolderRow {
  return {
    id: hashStringToNumber(taskList.id ?? ''),
    name: taskList.displayName ?? null,
    parentId: null,
    specialType: 0,
    folderType: 3, // Task folder
    accountId: 1,
    messageCount: 0,
    unreadCount: 0,
  };
}
