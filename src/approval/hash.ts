/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Target hashing for approval tokens.
 *
 * Generates hashes of critical properties to detect if a target
 * has been modified between prepare and confirm.
 */

import { createHash } from 'node:crypto';

/**
 * Creates a hash of an email's critical properties.
 * Used to detect if the email changed between prepare and confirm.
 */
export function hashEmailForApproval(email: {
  id: number;
  subject: string | null;
  folderId: number;
  timeReceived: number | null;
}): string {
  return createHash('sha256')
    .update(`${email.id}:${email.subject ?? ''}:${email.folderId}:${email.timeReceived ?? 0}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash of a folder's critical properties.
 * Used to detect if the folder changed between prepare and confirm.
 */
export function hashFolderForApproval(folder: {
  id: number;
  name: string | null;
  messageCount: number;
}): string {
  return createHash('sha256')
    .update(`${folder.id}:${folder.name ?? ''}:${folder.messageCount}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash of a draft's critical properties for send approval.
 * Used to detect if the draft changed between prepare and confirm.
 */
export function hashDraftForSend(draft: {
  id: number;
  subject: string | null;
  recipientCount: number;
}): string {
  return createHash('sha256')
    .update(`${draft.id}:${draft.subject ?? ''}:${draft.recipientCount}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash for a direct send (compose-and-send) approval.
 * Captures the recipient counts that define the send scope.
 */
export function hashDirectSendForApproval(params: {
  subject: string;
  toCount: number;
  ccCount: number;
  bccCount: number;
}): string {
  return createHash('sha256')
    .update(`${params.subject}:${params.toCount}:${params.ccCount}:${params.bccCount}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash for a reply approval.
 * Captures the original message ID, comment length, and reply-all flag.
 */
export function hashReplyForApproval(params: {
  originalId: number;
  commentLength: number;
  replyAll: boolean;
}): string {
  return createHash('sha256')
    .update(`${params.originalId}:${params.commentLength}:${params.replyAll}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash for a forward approval.
 * Captures the original message ID and the number of forward recipients.
 */
export function hashForwardForApproval(params: {
  originalId: number;
  recipientCount: number;
}): string {
  return createHash('sha256')
    .update(`${params.originalId}:${params.recipientCount}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash of a calendar event's critical properties.
 * Used to detect if the event changed between prepare and confirm.
 */
export function hashEventForApproval(event: {
  id: number;
  subject: string | null;
  startDateTime: string | null;
}): string {
  return createHash('sha256')
    .update(`${event.id}:${event.subject ?? ''}:${event.startDateTime ?? ''}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash of a contact's critical properties.
 * Used to detect if the contact changed between prepare and confirm.
 */
export function hashContactForApproval(contact: {
  id: number;
  displayName: string | null;
  emailAddress: string | null;
}): string {
  return createHash('sha256')
    .update(`${contact.id}:${contact.displayName ?? ''}:${contact.emailAddress ?? ''}`)
    .digest('hex')
    .slice(0, 16);
}

/**
 * Creates a hash of a task's critical properties.
 * Used to detect if the task changed between prepare and confirm.
 */
export function hashTaskForApproval(task: {
  taskId: string;
  title: string | null;
  listId: string;
}): string {
  return createHash('sha256')
    .update(`${task.taskId}:${task.title ?? ''}:${task.listId}`)
    .digest('hex')
    .slice(0, 16);
}
