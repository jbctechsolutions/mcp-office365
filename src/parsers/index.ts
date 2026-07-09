/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Parsers for Outlook content.
 *
 * Re-exports all parsers for convenient importing.
 */

export {
  stripHtml,
  containsHtml,
  extractPlainText,
  type StripHtmlOptions,
} from './html-stripper.js';
