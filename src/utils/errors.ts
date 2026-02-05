/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Custom error classes for the Outlook MCP server.
 */

/**
 * Error codes for categorizing errors.
 */
export const ErrorCode = {
  DATABASE_NOT_FOUND: 'DATABASE_NOT_FOUND',
  DATABASE_LOCKED: 'DATABASE_LOCKED',
  DATABASE_ERROR: 'DATABASE_ERROR',
  CONTENT_FILE_NOT_FOUND: 'CONTENT_FILE_NOT_FOUND',
  CONTENT_PARSE_ERROR: 'CONTENT_PARSE_ERROR',
  VALIDATION_ERROR: 'VALIDATION_ERROR',
  NOT_FOUND: 'NOT_FOUND',
  OUTLOOK_NOT_RUNNING: 'OUTLOOK_NOT_RUNNING',
  APPLESCRIPT_PERMISSION_DENIED: 'APPLESCRIPT_PERMISSION_DENIED',
  APPLESCRIPT_TIMEOUT: 'APPLESCRIPT_TIMEOUT',
  APPLESCRIPT_ERROR: 'APPLESCRIPT_ERROR',
  GRAPH_AUTH_REQUIRED: 'GRAPH_AUTH_REQUIRED',
  GRAPH_RATE_LIMITED: 'GRAPH_RATE_LIMITED',
  GRAPH_PERMISSION_DENIED: 'GRAPH_PERMISSION_DENIED',
  GRAPH_ERROR: 'GRAPH_ERROR',
  ATTACHMENT_NOT_FOUND: 'ATTACHMENT_NOT_FOUND',
  MAIL_SEND_ERROR: 'MAIL_SEND_ERROR',
  RECURRING_EVENT_ERROR: 'RECURRING_EVENT_ERROR',
  APPROVAL_EXPIRED: 'APPROVAL_EXPIRED',
  APPROVAL_INVALID: 'APPROVAL_INVALID',
  TARGET_CHANGED: 'TARGET_CHANGED',
} as const;

export type ErrorCode = (typeof ErrorCode)[keyof typeof ErrorCode];

/**
 * Base class for all Outlook MCP errors.
 */
export abstract class OutlookMcpError extends Error {
  abstract readonly code: ErrorCode;

  constructor(message: string) {
    super(message);
    this.name = this.constructor.name;
    // Maintains proper stack trace for where error was thrown
    Error.captureStackTrace(this, this.constructor);
  }
}

/**
 * Thrown when the Outlook database file cannot be found.
 */
export class DatabaseNotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.DATABASE_NOT_FOUND;

  constructor(path: string) {
    super(
      `Outlook database not found at: ${path}. ` +
        'Make sure Outlook for Mac has been opened at least once.'
    );
  }
}

/**
 * Thrown when the Outlook database is locked by another process.
 */
export class DatabaseLockedError extends OutlookMcpError {
  readonly code = ErrorCode.DATABASE_LOCKED;

  constructor() {
    super(
      'Outlook database is locked. ' +
        'This may happen during a sync operation. Please try again in a few seconds.'
    );
  }
}

/**
 * Thrown for general database errors.
 */
export class DatabaseError extends OutlookMcpError {
  readonly code = ErrorCode.DATABASE_ERROR;

  constructor(message: string, readonly cause?: Error) {
    super(message);
  }
}

/**
 * Thrown when an olk15 content file cannot be found.
 */
export class ContentFileNotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.CONTENT_FILE_NOT_FOUND;

  constructor(path: string) {
    super(`Content file not found: ${path}`);
  }
}

/**
 * Thrown when an olk15 file cannot be parsed.
 */
export class ContentParseError extends OutlookMcpError {
  readonly code = ErrorCode.CONTENT_PARSE_ERROR;

  constructor(path: string, readonly cause?: Error) {
    super(`Failed to parse content file: ${path}`);
  }
}

/**
 * Thrown for input validation errors.
 */
export class ValidationError extends OutlookMcpError {
  readonly code = ErrorCode.VALIDATION_ERROR;

  constructor(message: string) {
    super(message);
  }
}

/**
 * Thrown when a requested resource is not found.
 */
export class NotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.NOT_FOUND;

  constructor(resourceType: string, id: number | string) {
    super(`${resourceType} with ID ${id} not found`);
  }
}

/**
 * Type guard to check if an error is an OutlookMcpError.
 */
export function isOutlookMcpError(error: unknown): error is OutlookMcpError {
  return error instanceof OutlookMcpError;
}

/**
 * Wraps an unknown error in an OutlookMcpError if needed.
 */
export function wrapError(error: unknown, defaultMessage: string): OutlookMcpError {
  if (isOutlookMcpError(error)) {
    return error;
  }

  if (error instanceof Error) {
    return new DatabaseError(error.message, error);
  }

  return new DatabaseError(defaultMessage);
}

// =============================================================================
// AppleScript Errors
// =============================================================================

/**
 * Thrown when Outlook is not running and needs to be.
 */
export class OutlookNotRunningError extends OutlookMcpError {
  readonly code = ErrorCode.OUTLOOK_NOT_RUNNING;

  constructor() {
    super(
      'Microsoft Outlook is not running. ' +
        'Please start Outlook and try again.'
    );
  }
}

/**
 * Thrown when AppleScript automation permission is denied.
 */
export class AppleScriptPermissionError extends OutlookMcpError {
  readonly code = ErrorCode.APPLESCRIPT_PERMISSION_DENIED;

  constructor() {
    super(
      'Automation permission denied for Microsoft Outlook. ' +
        'Please grant access in System Settings > Privacy & Security > Automation.'
    );
  }
}

/**
 * Thrown when AppleScript execution times out.
 */
export class AppleScriptTimeoutError extends OutlookMcpError {
  readonly code = ErrorCode.APPLESCRIPT_TIMEOUT;

  constructor(operation: string) {
    super(
      `AppleScript operation timed out: ${operation}. ` +
        'This may happen with large data sets. Try reducing the limit.'
    );
  }
}

/**
 * Thrown for general AppleScript errors.
 */
export class AppleScriptError extends OutlookMcpError {
  readonly code = ErrorCode.APPLESCRIPT_ERROR;

  constructor(message: string, readonly cause?: Error) {
    super(message);
  }
}

// =============================================================================
// Microsoft Graph API Errors
// =============================================================================

/**
 * Thrown when Microsoft Graph authentication is required.
 */
export class GraphAuthRequiredError extends OutlookMcpError {
  readonly code = ErrorCode.GRAPH_AUTH_REQUIRED;

  constructor() {
    super(
      'Microsoft Graph authentication required. ' +
        'Please authenticate using the device code flow.'
    );
  }
}

/**
 * Thrown when the Graph API rate limit is exceeded.
 */
export class GraphRateLimitedError extends OutlookMcpError {
  readonly code = ErrorCode.GRAPH_RATE_LIMITED;
  readonly retryAfter: number | undefined;

  constructor(retryAfter?: number) {
    super(
      'Microsoft Graph API rate limit exceeded. ' +
        (retryAfter != null ? `Retry after ${retryAfter} seconds.` : 'Please try again later.')
    );
    this.retryAfter = retryAfter;
  }
}

/**
 * Thrown when the Graph API denies access due to permissions.
 */
export class GraphPermissionDeniedError extends OutlookMcpError {
  readonly code = ErrorCode.GRAPH_PERMISSION_DENIED;

  constructor(scope?: string) {
    super(
      'Microsoft Graph permission denied. ' +
        (scope != null
          ? `The application needs the '${scope}' permission.`
          : 'Please check your application permissions in Azure AD.')
    );
  }
}

/**
 * Thrown for general Microsoft Graph API errors.
 */
export class GraphError extends OutlookMcpError {
  readonly code = ErrorCode.GRAPH_ERROR;

  constructor(message: string, readonly cause?: Error) {
    super(message);
  }
}

// =============================================================================
// Event Management and Email Errors
// =============================================================================

/**
 * Thrown when an attachment file cannot be found.
 */
export class AttachmentNotFoundError extends OutlookMcpError {
  readonly code = ErrorCode.ATTACHMENT_NOT_FOUND;

  constructor(path: string) {
    super(`Attachment file not found: ${path}. Please check the file path exists.`);
  }
}

/**
 * Thrown when sending an email fails.
 */
export class MailSendError extends OutlookMcpError {
  readonly code = ErrorCode.MAIL_SEND_ERROR;

  constructor(reason: string) {
    super(`Failed to send email: ${reason}`);
  }
}

/**
 * Thrown when there's an error with recurring events.
 */
export class RecurringEventError extends OutlookMcpError {
  readonly code = ErrorCode.RECURRING_EVENT_ERROR;

  constructor(message: string) {
    super(message);
  }
}

// =============================================================================
// Approval Errors
// =============================================================================

/**
 * Thrown when an approval token has expired.
 */
export class ApprovalExpiredError extends OutlookMcpError {
  readonly code = ErrorCode.APPROVAL_EXPIRED;

  constructor() {
    super(
      'Approval token has expired. Please prepare the operation again.'
    );
  }
}

/**
 * Thrown when an approval token is invalid.
 */
export class ApprovalInvalidError extends OutlookMcpError {
  readonly code = ErrorCode.APPROVAL_INVALID;

  constructor(reason: string) {
    super(`Invalid approval token: ${reason}`);
  }
}

/**
 * Thrown when the target has been modified since the approval was generated.
 */
export class TargetChangedError extends OutlookMcpError {
  readonly code = ErrorCode.TARGET_CHANGED;

  constructor() {
    super(
      'The target has been modified since the approval was generated. ' +
        'Please prepare the operation again.'
    );
  }
}
