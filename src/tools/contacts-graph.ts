/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Graph-backend contact tools (v3 registry-driven architecture, U2 — dual
 * backend). Holds the contact logic that previously lived inline in the
 * `handleGraphToolCall` switch, so the registry handlers stay thin and branch
 * on `ctx.backend`.
 */

import type { GraphRepository } from '../graph/repository.js';
import type { GraphContentReaders } from '../graph/content-readers.js';
import type { ContactRow } from '../database/repository.js';
import type { ToolResult } from '../registry/types.js';
import type { ListContactsParams, SearchContactsParams, GetContactParams } from './contacts.js';

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Transforms a Graph contact row to the summary shape returned by the graph
 * backend's contact tools.
 */
export function transformContactRow(row: ContactRow): {
  id: number;
  displayName: string | null;
  sortName: string | null;
} {
  return {
    id: row.id,
    displayName: row.displayName,
    sortName: row.sortName,
  };
}

/**
 * Graph contact tools. Each method mirrors the extracted inline graph case body
 * and returns an MCP `ToolResult`.
 */
export class GraphContactsTools {
  constructor(
    private readonly repository: GraphRepository,
    private readonly contentReaders: GraphContentReaders
  ) {}

  async listContacts(params: ListContactsParams): Promise<ToolResult> {
    const contacts = params.folder_id != null
      ? await this.repository.listContactsInFolderAsync(params.folder_id, params.limit)
      : await this.repository.listContactsAsync(params.limit, params.offset);
    return jsonResult({ contacts: contacts.map(transformContactRow) });
  }

  async searchContacts(params: SearchContactsParams): Promise<ToolResult> {
    const contacts = await this.repository.searchContactsAsync(params.query, params.limit);
    return jsonResult({ contacts: contacts.map(transformContactRow) });
  }

  async getContact(params: GetContactParams): Promise<ToolResult> {
    const contact = await this.repository.getContactAsync(params.contact_id);
    if (contact == null) {
      return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
    }

    const details = await this.contentReaders.contact.readContactDetailsAsync(contact.dataFilePath);
    return jsonResult({ ...transformContactRow(contact), ...details });
  }
}
