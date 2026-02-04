/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Utility functions and error classes.
 */

export {
  APPLE_EPOCH_OFFSET,
  appleTimestampToIso,
  appleTimestampToDate,
  isoToAppleTimestamp,
  dateToAppleTimestamp,
} from './dates.js';

export {
  ErrorCode,
  OutlookMcpError,
  DatabaseNotFoundError,
  DatabaseLockedError,
  DatabaseError,
  ContentFileNotFoundError,
  ContentParseError,
  ValidationError,
  NotFoundError,
  isOutlookMcpError,
  wrapError,
} from './errors.js';
