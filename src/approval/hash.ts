/**
 * Target hashing for approval tokens.
 *
 * Generates hashes of critical properties to detect if a target
 * (email or folder) has been modified between prepare and confirm.
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
