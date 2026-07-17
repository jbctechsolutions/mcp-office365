/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Remote-mode write/destructive audit recorder (U8, R16).
 *
 * Attaches at the single CallTool chokepoint (`createServer`) and records one
 * row per non-read tool call: who (Entra oid/tid), which tool, the target
 * resource, the prepare/confirm phase + linkage, and the outcome. It is present
 * only for authenticated remote requests — stdio and unauthenticated loopback
 * pass a null recorder and the chokepoint runs unchanged.
 *
 * Two guarantees shape the design:
 * - **Fail-closed for `confirm_*`.** A confirm executes a client-tenant mutation
 *   (send/delete), so it must never proceed unrecorded: the audit row is written
 *   *before* dispatch and, if that write fails, the call aborts with a retriable
 *   {@link AuditUnavailableError}. Prepares and non-two-phase writes fail *open*
 *   (recorded after the fact; a write failure is warned, not fatal).
 * - **No secrets.** Only oid/tid identify the caller; targets are extracted from
 *   id-shaped params only (never subject/body/content), and no auth token
 *   material is ever stored.
 */

import type { AuditPhase, AuditStore } from '../state/store.js';
import type { ToolResult } from '../registry/types.js';
import { AuditUnavailableError } from '../utils/errors.js';

/** The acting user's identity for an audit row. */
export interface AuditIdentity {
  readonly oid: string;
  readonly tid: string;
}

/** The tool facts the recorder needs, read from the registry definition. */
export interface AuditToolInfo {
  /** MCP tool name. */
  readonly name: string;
  /** True when the tool advertises `readOnlyHint` (skipped — not a write). */
  readonly readOnly: boolean;
  /**
   * A `prepare_*` tool's declarative token extractor (`onElicit.collectTokenIds`),
   * used to link a prepare row to its confirm row by the minted approval-token id.
   * Absent for tools that don't opt into elicitation.
   */
  readonly collectTokenIds?: (result: ToolResult) => string[];
}

/** The outcome of the dispatched call, supplied to {@link PendingAudit.finish}. */
export interface AuditFinish {
  readonly ok: boolean;
  readonly errorCode?: string | null;
  /** The tool result — used to recover a prepare's minted token for linkage. */
  readonly result?: ToolResult;
}

/** Handle returned by {@link AuditRecorder.begin}; finalized after dispatch. */
export interface PendingAudit {
  finish(outcome: AuditFinish): void;
}

/** Records write/destructive tool calls for authenticated remote requests. */
export interface AuditRecorder {
  /**
   * Called before dispatch. Returns a handle to finalize after the tool runs, or
   * null when the call is not audited (a read-only tool). Throws
   * {@link AuditUnavailableError} when a `confirm_*` row cannot be reserved — the
   * caller must abort without dispatching.
   */
  begin(tool: AuditToolInfo, args: unknown): PendingAudit | null;
}

/** Maximum stored target length — a coarse cap; targets are id-shaped, so small. */
const MAX_TARGET_LEN = 500;

/** Matches id-shaped param keys (`id`, `email_id`, `channel_id`, …). */
const ID_KEY = /(?:^|_)id$/i;

/** Link-key param names carried by `confirm_*` tools. */
const LINK_KEYS = ['approval_token', 'token_id'] as const;

function classifyPhase(name: string): AuditPhase {
  if (name.startsWith('confirm_')) return 'confirm';
  if (name.startsWith('prepare_')) return 'prepare';
  return 'write';
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

/** A scalar param safe to store as an identifier (string/number). */
function scalarId(v: unknown): string | null {
  if (typeof v === 'string' && v.length > 0) return v;
  if (typeof v === 'number' && Number.isFinite(v)) return String(v);
  return null;
}

/**
 * Extracts the approval-token id linking a confirm (or batch confirm) back to
 * its prepare, from the call arguments. Content-free by construction.
 */
function extractLinkKeyFromArgs(args: unknown): string | null {
  if (!isRecord(args)) return null;
  for (const key of LINK_KEYS) {
    const v = scalarId(args[key]);
    if (v != null) return v;
  }
  // Batch confirm: an array of { token_id, ... } pairs.
  const arr = args['confirmations'];
  if (Array.isArray(arr)) {
    const ids = arr
      .map((e) => (isRecord(e) ? scalarId(e['token_id']) : null))
      .filter((v): v is string => v != null);
    if (ids.length > 0) return ids.join(',');
  }
  return null;
}

/**
 * Builds a best-effort target string from id-shaped params only. Never reads
 * content fields (subject/body/message/…) — they simply aren't id-shaped, so the
 * allow-by-key-shape rule excludes them by construction. Approval/token link
 * fields are excluded (they're recorded as the link key, not the target).
 */
function extractTarget(args: unknown): string | null {
  try {
    const ids: Record<string, string> = {};
    const collect = (obj: Record<string, unknown>): void => {
      for (const [key, value] of Object.entries(obj)) {
        if ((LINK_KEYS as readonly string[]).includes(key)) continue;
        if (!ID_KEY.test(key)) continue;
        const scalar = scalarId(value);
        if (scalar != null) ids[key] = scalar;
      }
    };
    if (isRecord(args)) {
      collect(args);
      // One level into arrays of objects (e.g. batch confirmations' email ids).
      for (const value of Object.values(args)) {
        if (Array.isArray(value)) {
          for (const el of value) {
            if (isRecord(el)) collect(el);
          }
        }
      }
    }
    const keys = Object.keys(ids);
    if (keys.length === 0) return null;
    const text = JSON.stringify(ids);
    return text.length > MAX_TARGET_LEN ? `${text.slice(0, MAX_TARGET_LEN)}…` : text;
  } catch {
    return null;
  }
}

/** Options for {@link createAuditRecorder}. */
export interface AuditRecorderOptions {
  /** Warning sink for fail-open write failures (defaults to stderr). */
  warn?: (message: string) => void;
  /** Clock, for deterministic tests. */
  now?: () => number;
}

/**
 * Builds a recorder bound to a store and the request's identity. One recorder
 * per authenticated request; `begin` is called once per tool call.
 */
export function createAuditRecorder(
  store: AuditStore,
  identity: AuditIdentity,
  options: AuditRecorderOptions = {},
): AuditRecorder {
  const warn =
    options.warn ?? ((msg: string): void => void process.stderr.write(`${msg}\n`));
  const now = options.now;

  return {
    begin(tool: AuditToolInfo, args: unknown): PendingAudit | null {
      if (tool.readOnly) {
        return null; // reads are not audited
      }
      const phase = classifyPhase(tool.name);
      const target = extractTarget(args);
      const base = {
        oid: identity.oid,
        tid: identity.tid,
        tool: tool.name,
        phase,
        target,
        ...(now != null ? { createdAt: now() } : {}),
      };

      if (phase === 'confirm') {
        // Fail-closed: reserve the row BEFORE the mutation runs. A failure here
        // aborts the call — the client-tenant action must not proceed unaudited.
        const linkKey = extractLinkKeyFromArgs(args);
        let id: number;
        try {
          id = store.recordAudit({ ...base, linkKey, outcome: 'pending' });
        } catch {
          throw new AuditUnavailableError();
        }
        return {
          finish(outcome: AuditFinish): void {
            // Best-effort: the mutation already happened and the reservation row
            // exists, so a failed outcome-update is warned, not fatal.
            try {
              store.updateAuditOutcome(id, outcome.ok ? 'ok' : 'error', outcome.errorCode ?? null);
            } catch (e) {
              warn(auditWarn('finalize', tool.name, e));
            }
          },
        };
      }

      // prepare / write: fail-open. Record once, after the call, with the final
      // outcome. A prepare recovers its minted token from the result for linkage.
      return {
        finish(outcome: AuditFinish): void {
          const linkKey =
            phase === 'prepare' && tool.collectTokenIds != null && outcome.result != null
              ? (safeCollect(tool.collectTokenIds, outcome.result)[0] ?? null)
              : extractLinkKeyFromArgs(args);
          try {
            store.recordAudit({
              ...base,
              linkKey,
              outcome: outcome.ok ? 'ok' : 'error',
              errorCode: outcome.errorCode ?? null,
            });
          } catch (e) {
            warn(auditWarn('record', tool.name, e));
          }
        },
      };
    },
  };
}

/** Runs a prepare's token extractor without letting a bad extractor throw. */
function safeCollect(fn: (r: ToolResult) => string[], result: ToolResult): string[] {
  try {
    return fn(result);
  } catch {
    return [];
  }
}

function auditWarn(op: string, tool: string, error: unknown): string {
  const reason = error instanceof Error ? error.message : String(error);
  return `[mcp-office365] audit ${op} failed for ${tool} (${reason}); continuing (write already applied).`;
}
