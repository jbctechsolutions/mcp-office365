/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph MailFolder and Calendar types to FolderRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { FolderRow } from '../../database/repository.js';
import { mintSelfEncoded } from '../../ids/token.js';

/**
 * Maps a Graph MailFolder to a FolderRow.
 */
export function mapMailFolderToRow(folder: MicrosoftGraph.MailFolder): FolderRow {
  const graphId = folder.id ?? '';
  const parentGraphId = folder.parentFolderId;
  return {
    // Durable self-encoding fd_ token carrying the immutable Graph folder id (U5).
    id: graphId.length > 0 ? mintSelfEncoded('folder', graphId) : '',
    name: folder.displayName ?? null,
    parentId: parentGraphId != null && parentGraphId.length > 0 ? mintSelfEncoded('folder', parentGraphId) : null,
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
  const graphId = calendar.id ?? '';
  return {
    // Durable self-encoding fd_ token carrying the immutable Graph calendar id (U5).
    id: graphId.length > 0 ? mintSelfEncoded('folder', graphId) : '',
    name: calendar.name ?? null,
    parentId: null,
    specialType: 0,
    folderType: 2, // Calendar folder
    accountId: 1,
    messageCount: 0,
    unreadCount: 0,
  };
}
