/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Aggregated tool definitions across all registry-migrated domains.
 *
 * This is the single list the server registers and the contract harness
 * iterates. As U2 migrates each domain, add its `*ToolDefinitions()` here and
 * the harness covers it automatically — no per-domain test wiring.
 */

import type { ToolDefinition } from './types.js';
import { mailRulesToolDefinitions } from '../tools/mail-rules.js';

export function allToolDefinitions(): ToolDefinition[] {
  return [
    ...mailRulesToolDefinitions(),
    // U2: append each migrated domain's definitions here.
  ];
}
