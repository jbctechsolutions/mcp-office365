/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Maps Microsoft Graph Message type to EmailRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { EmailRow } from '../../database/repository.js';
import { mintSelfEncoded } from '../../ids/token.js';
import {
  hashStringToNumber,
  isoToTimestamp,
  importanceToPriority,
  flagStatusToNumber,
  extractDisplayName,
  extractEmailAddress,
  formatRecipients,
  formatRecipientAddresses,
  createGraphContentPath,
} from './utils.js';

/**
 * Maps a Graph Message to an EmailRow.
 */
export function mapMessageToEmailRow(
  message: MicrosoftGraph.Message,
  folderId?: string
): EmailRow {
  const messageId = message.id ?? '';
  const parentFolderId = folderId ?? message.parentFolderId;

  // Type assertions needed due to Graph API's NullableOption types
  // which are incompatible with exactOptionalPropertyTypes
  const from = message.from as { emailAddress?: { address?: string; name?: string } } | null | undefined;
  const toRecipients = message.toRecipients as Array<{ emailAddress?: { address?: string; name?: string } }> | null | undefined;
  const ccRecipients = message.ccRecipients as Array<{ emailAddress?: { address?: string; name?: string } }> | null | undefined;
  const flag = message.flag as { flagStatus?: string } | null | undefined;

  return {
    // Durable self-encoding em_ token carrying the immutable Graph message id (U5).
    id: messageId.length > 0 ? mintSelfEncoded('message', messageId) : '',
    // Durable self-encoding fd_ token carrying the immutable Graph folder id (U5).
    folderId: parentFolderId != null ? mintSelfEncoded('folder', parentFolderId) : '',
    subject: message.subject ?? null,
    sender: extractDisplayName(from),
    senderAddress: extractEmailAddress(from),
    recipients: formatRecipients(toRecipients),
    displayTo: formatRecipients(toRecipients),
    toAddresses: formatRecipientAddresses(toRecipients),
    ccAddresses: formatRecipientAddresses(ccRecipients),
    preview: message.bodyPreview ?? null,
    isRead: message.isRead === true ? 1 : 0,
    timeReceived: isoToTimestamp(message.receivedDateTime),
    timeSent: isoToTimestamp(message.sentDateTime),
    hasAttachment: message.hasAttachments === true ? 1 : 0,
    size: 0, // Not available in Graph API message response
    priority: importanceToPriority(message.importance),
    flagStatus: flagStatusToNumber(flag),
    categories: message.categories != null && message.categories.length > 0
      ? Buffer.from(message.categories.join(','), 'utf-8')
      : null,
    messageId: message.internetMessageId ?? null,
    conversationId: message.conversationId != null ? hashStringToNumber(message.conversationId) : null,
    dataFilePath: createGraphContentPath('email', messageId),
  };
}
