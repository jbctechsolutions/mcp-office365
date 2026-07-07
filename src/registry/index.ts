/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tool registry barrel.
 */

export type {
  Backend,
  Preset,
  ToolAnnotations,
  ToolResult,
  ToolContext,
  ToolDefinition,
  GraphToolsets,
  AppleScriptToolsets,
} from './types.js';
export { ToolRegistry, toInputSchema } from './registry.js';
export type { SurfaceOptions } from './registry.js';
export { defineTool } from './define-tool.js';
