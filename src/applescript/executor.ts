/**
 * AppleScript execution utilities.
 *
 * Provides functions to execute AppleScript commands via osascript
 * and handle errors, timeouts, and permission issues.
 */

import { execSync, type ExecSyncOptionsWithStringEncoding } from 'node:child_process';

// =============================================================================
// Constants
// =============================================================================

/**
 * Default timeout for AppleScript execution in milliseconds.
 */
const DEFAULT_TIMEOUT_MS = 30000;

/**
 * Error messages for common AppleScript failures.
 */
const ERROR_PATTERNS = {
  notRunning: /not running|application isn't running/i,
  permissionDenied: /not authorized|permission denied|assistive access/i,
  timeout: /timed out|timeout/i,
  handlerFailed: /AppleEvent handler failed/i,
} as const;

// =============================================================================
// Types
// =============================================================================

/**
 * Result of an AppleScript execution.
 */
export interface AppleScriptResult {
  readonly success: boolean;
  readonly output: string;
  readonly error?: string;
}

/**
 * Options for AppleScript execution.
 */
export interface ExecuteOptions {
  readonly timeoutMs?: number;
}

// =============================================================================
// Error Detection
// =============================================================================

/**
 * Determines the type of error from an error message.
 */
function categorizeError(errorMessage: string): 'not_running' | 'permission_denied' | 'timeout' | 'unknown' {
  if (ERROR_PATTERNS.notRunning.test(errorMessage)) {
    return 'not_running';
  }
  if (ERROR_PATTERNS.permissionDenied.test(errorMessage)) {
    return 'permission_denied';
  }
  if (ERROR_PATTERNS.timeout.test(errorMessage)) {
    return 'timeout';
  }
  return 'unknown';
}

// =============================================================================
// Script Escaping
// =============================================================================

/**
 * Escapes a string for safe inclusion in AppleScript.
 * Handles quotes and backslashes.
 */
export function escapeForAppleScript(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"');
}

// =============================================================================
// Execution Functions
// =============================================================================

/**
 * Executes an AppleScript and returns the result.
 *
 * @param script - The AppleScript code to execute
 * @param options - Execution options
 * @returns The result of execution
 */
export function executeAppleScript(script: string, options: ExecuteOptions = {}): AppleScriptResult {
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;

  const execOptions: ExecSyncOptionsWithStringEncoding = {
    encoding: 'utf8',
    timeout: timeoutMs,
    maxBuffer: 50 * 1024 * 1024, // 50MB for large results
    stdio: ['pipe', 'pipe', 'pipe'],
  };

  try {
    // Execute via osascript
    const output = execSync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`, execOptions);
    return {
      success: true,
      output: output.trim(),
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stderr = (error as { stderr?: Buffer | string })?.stderr;
    const stderrText = stderr instanceof Buffer ? stderr.toString('utf8') : (stderr ?? '');

    const fullError = `${errorMessage}\n${stderrText}`.trim();
    const errorType = categorizeError(fullError);

    return {
      success: false,
      output: '',
      error: fullError,
    };
  }
}

/**
 * Executes an AppleScript and throws on failure.
 *
 * @param script - The AppleScript code to execute
 * @param options - Execution options
 * @returns The output string
 * @throws Error if execution fails
 */
export function executeAppleScriptOrThrow(script: string, options: ExecuteOptions = {}): string {
  const result = executeAppleScript(script, options);
  if (!result.success) {
    const errorType = categorizeError(result.error ?? '');
    throw new AppleScriptExecutionError(result.error ?? 'Unknown error', errorType);
  }
  return result.output;
}

// =============================================================================
// Outlook-Specific Functions
// =============================================================================

/**
 * Checks if Microsoft Outlook is currently running.
 */
export function isOutlookRunning(): boolean {
  const script = `
tell application "System Events"
  set isRunning to (name of processes) contains "Microsoft Outlook"
  return isRunning
end tell
`;

  const result = executeAppleScript(script);
  return result.success && result.output.toLowerCase() === 'true';
}

/**
 * Launches Microsoft Outlook if not already running.
 */
export function launchOutlook(): boolean {
  const script = `
tell application "Microsoft Outlook"
  launch
end tell
return "launched"
`;

  const result = executeAppleScript(script, { timeoutMs: 10000 });
  return result.success;
}

/**
 * Gets the version of Microsoft Outlook.
 */
export function getOutlookVersion(): string | null {
  const script = `
tell application "Microsoft Outlook"
  return version
end tell
`;

  const result = executeAppleScript(script);
  return result.success ? result.output : null;
}

// =============================================================================
// Error Classes
// =============================================================================

/**
 * Error thrown when AppleScript execution fails.
 */
export class AppleScriptExecutionError extends Error {
  readonly errorType: 'not_running' | 'permission_denied' | 'timeout' | 'unknown';

  constructor(message: string, errorType: 'not_running' | 'permission_denied' | 'timeout' | 'unknown') {
    super(message);
    this.name = 'AppleScriptExecutionError';
    this.errorType = errorType;
  }
}
