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
  GRAPH_RATE_LIMITED: 'GRAPH_RATE_LIMITED',
  GRAPH_PERMISSION_DENIED: 'GRAPH_PERMISSION_DENIED',
  GRAPH_ERROR: 'GRAPH_ERROR',
  // D10 transport error envelope codes (mapped at the single dispatch chokepoint).
  AUTH_EXPIRED: 'AUTH_EXPIRED',
  THROTTLED: 'THROTTLED',
  GRAPH_UNAVAILABLE: 'GRAPH_UNAVAILABLE',
  ATTACHMENT_NOT_FOUND: 'ATTACHMENT_NOT_FOUND',
  ATTACHMENT_TOO_LARGE: 'ATTACHMENT_TOO_LARGE',
  ATTACHMENT_SAVE_ERROR: 'ATTACHMENT_SAVE_ERROR',
  MAIL_SEND_ERROR: 'MAIL_SEND_ERROR',
  RECURRING_EVENT_ERROR: 'RECURRING_EVENT_ERROR',
  APPROVAL_EXPIRED: 'APPROVAL_EXPIRED',
  APPROVAL_INVALID: 'APPROVAL_INVALID',
  TARGET_CHANGED: 'TARGET_CHANGED',
  READ_ONLY_MODE: 'READ_ONLY_MODE',
  // Durable-ID resolution vocabulary (U5 / D1, D2, D4, D7).
  ID_UNKNOWN: 'ID_UNKNOWN',
  ID_STALE: 'ID_STALE',
  ID_COLLISION: 'ID_COLLISION',
  ID_FOREIGN_ACCOUNT: 'ID_FOREIGN_ACCOUNT',
  ID_ENTITY_MISMATCH: 'ID_ENTITY_MISMATCH',
  NUMERIC_ID_UNSUPPORTED: 'NUMERIC_ID_UNSUPPORTED',
} as const;

export type ErrorCode = (typeof ErrorCode)[keyof typeof ErrorCode];

/** Per-error hints carried into the D10 envelope. */
export interface ErrorMeta {
  /** True when the caller can reasonably retry the same operation. */
  retriable?: boolean;
  /** Actionable next step for the caller (agent or human). */
  suggestion?: string;
}

/**
 * Base class for all Outlook MCP errors.
 */
export abstract class OutlookMcpError extends Error {
  abstract readonly code: ErrorCode;
  readonly retriable: boolean;
  readonly suggestion: string | undefined;

  constructor(message: string, meta?: ErrorMeta) {
    super(message);
    this.name = this.constructor.name;
    this.retriable = meta?.retriable ?? false;
    this.suggestion = meta?.suggestion;
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
// Microsoft Graph API Errors
// =============================================================================

/**
 * Thrown when Microsoft Graph authentication is required or a cached session has
 * expired mid-run. Uses the D10 `AUTH_EXPIRED` code (the older
 * `GRAPH_AUTH_REQUIRED` was never thrown and has been removed) so raw-401
 * mapping and this typed path agree on one code for "you need to (re-)authenticate".
 */
export class GraphAuthRequiredError extends OutlookMcpError {
  readonly code = ErrorCode.AUTH_EXPIRED;

  constructor(reason: 'not_authenticated' | 'session_expired' = 'not_authenticated') {
    const lead =
      reason === 'session_expired'
        ? 'Microsoft Graph session expired.'
        : 'Microsoft Graph authentication required.';
    super(`${lead} Run: npx @jbctechsolutions/mcp-office365 auth`, {
      retriable: false,
      suggestion: 'Run `npx @jbctechsolutions/mcp-office365 auth` to (re)authenticate.',
    });
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
        (retryAfter != null ? `Retry after ${retryAfter} seconds.` : 'Please try again later.'),
      {
        retriable: true,
        suggestion:
          retryAfter != null
            ? `Wait ${retryAfter}s and retry.`
            : 'Wait a few seconds and retry.',
      }
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
 * Thrown when an attachment exceeds the size limit.
 */
export class AttachmentTooLargeError extends OutlookMcpError {
  readonly code = ErrorCode.ATTACHMENT_TOO_LARGE;

  constructor(name: string, sizeBytes: number, maxBytes: number) {
    super(
      `Attachment "${name}" is ${Math.round(sizeBytes / 1024 / 1024)}MB ` +
        `which exceeds the maximum size of ${Math.round(maxBytes / 1024 / 1024)}MB.`
    );
  }
}

/**
 * Thrown when saving an attachment to disk fails.
 */
export class AttachmentSaveError extends OutlookMcpError {
  readonly code = ErrorCode.ATTACHMENT_SAVE_ERROR;

  constructor(name: string, reason: string) {
    super(`Failed to save attachment "${name}": ${reason}`);
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

/**
 * Thrown when a destructive/write tool is invoked while the server runs in
 * `--read-only` mode (D13). Read-only filters these tools out of the surface,
 * but a client can still call one by name — this is the runtime guard.
 */
export class ReadOnlyModeError extends OutlookMcpError {
  readonly code = ErrorCode.READ_ONLY_MODE;

  constructor(toolName: string) {
    super(
      `Tool "${toolName}" is not available in read-only mode.`,
      { retriable: false, suggestion: 'Restart the server without --read-only to use write tools.' },
    );
  }
}

// =============================================================================
// Durable-ID resolution errors (U5)
// =============================================================================

/** A durable-ID token could not be resolved (unknown / cold composite token). */
export class IdUnknownError extends OutlookMcpError {
  readonly code = ErrorCode.ID_UNKNOWN;

  constructor(token: string, hint = 'Re-list the parent collection to get a fresh ID.') {
    super(`Unknown or unresolvable ID: ${token}.`, { retriable: false, suggestion: hint });
  }
}

/** A resolved ID no longer points at a live object (mutable ID drifted). */
export class IdStaleError extends OutlookMcpError {
  readonly code = ErrorCode.ID_STALE;

  constructor(token: string) {
    super(`ID is stale and could not be re-resolved: ${token}.`, {
      retriable: false,
      suggestion: 'Re-list the collection to obtain a current ID.',
    });
  }
}

/** Two distinct keys minted the same composite token (D1a) — never mis-resolve. */
export class IdCollisionError extends OutlookMcpError {
  readonly code = ErrorCode.ID_COLLISION;

  constructor(token: string) {
    super(`ID token collision detected for ${token}; refusing to resolve ambiguously.`, {
      retriable: false,
      suggestion: 'Re-list the collection to obtain a fresh ID.',
    });
  }
}

/** A token minted under a different signed-in account (D7). */
export class IdForeignAccountError extends OutlookMcpError {
  readonly code = ErrorCode.ID_FOREIGN_ACCOUNT;

  constructor(token: string) {
    super(`ID ${token} belongs to a different account than the one signed in.`, {
      retriable: false,
      suggestion: 'Re-list the collection while signed in as the owning account.',
    });
  }
}

/** A durable token for one entity kind was passed where another was expected. */
export class IdEntityMismatchError extends OutlookMcpError {
  readonly code = ErrorCode.ID_ENTITY_MISMATCH;

  constructor(token: string, expected: string, actual: string) {
    super(`ID ${token} is a ${actual} token, but a ${expected} ID was expected.`, {
      retriable: false,
      suggestion: `Pass a ${expected} ID (re-list the ${expected} collection to get one).`,
    });
  }
}

/** A legacy numeric (v2 hash) ID was passed on the Graph backend (D4). */
export class NumericIdUnsupportedError extends OutlookMcpError {
  readonly code = ErrorCode.NUMERIC_ID_UNSUPPORTED;

  constructor(id: number) {
    super(`Numeric ID ${id} is not supported on the Graph backend (v2 hash IDs are lossy).`, {
      retriable: false,
      suggestion: 'Re-list the collection to obtain a current durable ID.',
    });
  }
}

// =============================================================================
// D10 — Typed error envelope
// =============================================================================

/**
 * The stable, machine-readable shape returned for every tool failure. Ends the
 * ad-hoc `GRAPH_ERROR:`/`DATABASE_ERROR:` string-prefix inconsistency: callers
 * get a stable `code`, a human `message`, whether a retry could help
 * (`retriable`), and an actionable `suggestion` when one exists.
 */
export interface ErrorEnvelope {
  code: ErrorCode;
  message: string;
  retriable: boolean;
  suggestion?: string;
}

/** Shape of a Microsoft Graph SDK error (has a numeric HTTP `statusCode`). */
interface GraphSdkErrorLike {
  statusCode: number;
  code?: string;
  message?: string;
}

function isGraphSdkError(error: unknown): error is GraphSdkErrorLike {
  return (
    typeof error === 'object' &&
    error !== null &&
    typeof (error as { statusCode?: unknown }).statusCode === 'number'
  );
}

/** Maps a raw Graph SDK / HTTP status to a stable envelope. */
function graphStatusToEnvelope(status: number, message: string): ErrorEnvelope {
  switch (status) {
    case 401:
      return {
        code: ErrorCode.AUTH_EXPIRED,
        message,
        retriable: false,
        suggestion: 'Session expired. Run `npx @jbctechsolutions/mcp-office365 auth` to re-authenticate.',
      };
    case 403:
      return {
        code: ErrorCode.GRAPH_PERMISSION_DENIED,
        message,
        retriable: false,
        suggestion: 'Check the app permissions/scopes granted in Azure AD.',
      };
    case 404:
      return { code: ErrorCode.NOT_FOUND, message, retriable: false };
    case 400:
      return { code: ErrorCode.VALIDATION_ERROR, message, retriable: false };
    case 429:
      return {
        code: ErrorCode.THROTTLED,
        message,
        retriable: true,
        suggestion: 'Rate limited. Retry after the Retry-After interval.',
      };
    case 502:
    case 503:
    case 504:
      // The transport auto-retries these (see shouldRetryGraphRequest), so
      // retriable:true reflects real coverage, not just caller advice.
      return {
        code: ErrorCode.GRAPH_UNAVAILABLE,
        message,
        retriable: true,
        suggestion: 'Microsoft Graph is temporarily unavailable. Retry shortly.',
      };
    default:
      if (status >= 500) {
        // Other 5xx (500/501/505…) are NOT auto-retried by the transport, so
        // retriable stays false to keep the envelope honest with D5.
        return {
          code: ErrorCode.GRAPH_UNAVAILABLE,
          message,
          retriable: false,
          suggestion: 'Microsoft Graph returned a server error.',
        };
      }
      return { code: ErrorCode.GRAPH_ERROR, message, retriable: false };
  }
}

/** True when `code` is a member of the stable {@link ErrorCode} vocabulary. */
export function isKnownErrorCode(code: string): code is ErrorCode {
  return (Object.values(ErrorCode) as string[]).includes(code);
}

/**
 * True when `value` already has the {@link ErrorEnvelope} shape *with a known
 * code*. Requiring a real {@link ErrorCode} (not just any string) means an
 * unrelated JSON payload that happens to carry `code`/`message`/`retriable`
 * fields is not mistaken for an envelope by {@link ensureErrorEnvelopeText}.
 */
export function isErrorEnvelope(value: unknown): value is ErrorEnvelope {
  const candidate = value as Partial<Record<keyof ErrorEnvelope, unknown>>;
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof candidate.code === 'string' &&
    isKnownErrorCode(candidate.code) &&
    typeof candidate.message === 'string' &&
    typeof candidate.retriable === 'boolean' &&
    (candidate.suggestion === undefined || typeof candidate.suggestion === 'string')
  );
}

/**
 * Normalizes the text of a tool's error result into a D10 envelope JSON string.
 * Thrown failures already funnel through {@link toErrorEnvelope} at the dispatch
 * chokepoint, but many handlers return `{ isError: true, text: '…' }` directly
 * (not-found, approval-token mismatches, etc.). Passing those through here keeps
 * a single stable error shape across *every* failure path. Text that is already
 * a valid envelope is returned unchanged (idempotent — never double-wraps).
 */
export function ensureErrorEnvelopeText(text: string): string {
  try {
    if (isErrorEnvelope(JSON.parse(text))) {
      return text;
    }
  } catch {
    // not JSON — fall through and wrap
  }
  const envelope: ErrorEnvelope = {
    code: ErrorCode.GRAPH_ERROR,
    message: text,
    retriable: false,
  };
  return JSON.stringify(envelope, null, 2);
}

/**
 * Single mapping point (D10): converts any thrown value into a stable
 * {@link ErrorEnvelope}. Typed {@link OutlookMcpError}s carry their own
 * code/retriable/suggestion; raw Graph SDK errors are mapped by HTTP status;
 * everything else becomes a non-retriable `GRAPH_ERROR`.
 */
export function toErrorEnvelope(error: unknown): ErrorEnvelope {
  if (isOutlookMcpError(error)) {
    const envelope: ErrorEnvelope = {
      code: error.code,
      message: error.message,
      retriable: error.retriable,
    };
    if (error.suggestion !== undefined) {
      envelope.suggestion = error.suggestion;
    }
    return envelope;
  }

  if (isGraphSdkError(error)) {
    const message =
      typeof error.message === 'string' && error.message.length > 0
        ? error.message
        : `Microsoft Graph request failed (HTTP ${error.statusCode}).`;
    return graphStatusToEnvelope(error.statusCode, message);
  }

  if (error instanceof Error) {
    return { code: ErrorCode.GRAPH_ERROR, message: error.message, retriable: false };
  }

  return { code: ErrorCode.GRAPH_ERROR, message: 'An unknown error occurred.', retriable: false };
}
