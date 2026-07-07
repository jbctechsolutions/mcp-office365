/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * `defineTool` — typed helper for authoring a tool definition.
 *
 * Preserves the inference link between the Zod input schema and the handler's
 * `params` argument, so a domain module declares its schema once and the
 * handler is type-checked against it.
 */

import type { z } from 'zod';
import type { ToolDefinition } from './types.js';

export function defineTool<S extends z.ZodType>(def: ToolDefinition<S>): ToolDefinition<S> {
  return def;
}
