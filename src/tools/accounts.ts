/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Account tools (v3 registry-driven architecture, U2 — dual backend).
 *
 * `list_accounts` is served by the AppleScript account repository in both
 * backends. AccountsTools wraps that repository and is registered on both the
 * Graph and AppleScript toolset bags, so a single registry handler resolves the
 * active backend's instance.
 */

import { z } from 'zod';
import type { IAccountRepository } from '../applescript/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset, requireAppleScriptToolset } from '../registry/context.js';
import type { ToolDefinition, ToolResult } from '../registry/types.js';

declare module '../registry/types.js' {
  interface GraphToolsets {
    accounts: AccountsTools;
  }
  interface AppleScriptToolsets {
    accounts: AccountsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListAccountsInput = z.strictObject({});

// =============================================================================
// Accounts Tools
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Account tools backed by the AppleScript account repository.
 */
export class AccountsTools {
  constructor(private readonly accountRepository: IAccountRepository) {}

  listAccounts(): ToolResult {
    const accounts = this.accountRepository.listAccounts();
    return jsonResult({
      accounts: accounts.map((acc) => ({
        id: acc.id,
        name: acc.name,
        email: acc.email,
        type: acc.type,
      })),
    });
  }
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2 — dual backend)
// =============================================================================

/**
 * Registry tool definitions for the accounts domain.
 */
export function accountsToolDefinitions(): ToolDefinition[] {
  return [
    defineTool({
      name: 'list_accounts',
      description: 'List all Exchange accounts configured in Outlook with their details',
      input: ListAccountsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: [],
      backends: ['graph', 'applescript'],
      handler: (ctx) =>
        (ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'accounts')
          : requireAppleScriptToolset(ctx, 'accounts')
        ).listAccounts(),
    }),
  ];
}
