/**
 * Maps Microsoft Graph Message type to EmailRow.
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import type { EmailRow } from '../../database/repository.js';
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
    id: hashStringToNumber(messageId),
    folderId: parentFolderId != null ? hashStringToNumber(parentFolderId) : 0,
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
