/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph type mappers.
 *
 * Exports functions to convert Graph API types to internal row types.
 */

export {
  mapMailFolderToRow,
  mapCalendarToFolderRow,
  mapTaskListToFolderRow,
} from './folder-mapper.js';

export { mapMessageToEmailRow } from './email-mapper.js';

export { mapEventToEventRow } from './event-mapper.js';

export { mapContactToContactRow } from './contact-mapper.js';

export { mapTaskToTaskRow, type TodoTaskWithList } from './task-mapper.js';

export {
  hashStringToNumber,
  isoToTimestamp,
  dateTimeTimeZoneToTimestamp,
  importanceToPriority,
  flagStatusToNumber,
  extractEmailAddress,
  extractDisplayName,
  formatRecipients,
  formatRecipientAddresses,
  createGraphContentPath,
} from './utils.js';
