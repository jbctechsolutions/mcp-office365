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
import type { ApprovalTokenManager } from '../approval/index.js';
import { hashContactForApproval } from '../approval/index.js';
import type { ToolResult } from '../registry/types.js';
import type {
  ListContactsParams,
  SearchContactsParams,
  GetContactParams,
  CreateContactParams,
  UpdateContactParams,
  PrepareDeleteContactParams,
  ConfirmDeleteContactParams,
  GetContactPhotoParams,
  SetContactPhotoParams,
} from './contacts.js';

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
    private readonly contentReaders: GraphContentReaders,
    private readonly tokenManager: ApprovalTokenManager
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

  async createContact(params: CreateContactParams): Promise<ToolResult> {
    const numericId = await this.repository.createContactAsync({
      ...(params.given_name != null ? { given_name: params.given_name } : {}),
      ...(params.surname != null ? { surname: params.surname } : {}),
      ...(params.email != null ? { email: params.email } : {}),
      ...(params.phone != null ? { phone: params.phone } : {}),
      ...(params.mobile_phone != null ? { mobile_phone: params.mobile_phone } : {}),
      ...(params.company != null ? { company: params.company } : {}),
      ...(params.job_title != null ? { job_title: params.job_title } : {}),
      ...(params.street_address != null ? { street_address: params.street_address } : {}),
      ...(params.city != null ? { city: params.city } : {}),
      ...(params.state != null ? { state: params.state } : {}),
      ...(params.postal_code != null ? { postal_code: params.postal_code } : {}),
      ...(params.country != null ? { country: params.country } : {}),
    });
    return jsonResult({
      id: numericId,
      given_name: params.given_name ?? null,
      surname: params.surname ?? null,
      email: params.email ?? null,
      status: 'created',
    });
  }

  async updateContact(params: UpdateContactParams): Promise<ToolResult> {
    const updates: Record<string, unknown> = {};
    if (params.given_name != null) updates.givenName = params.given_name;
    if (params.surname != null) updates.surname = params.surname;
    if (params.email != null) updates.emailAddresses = [{ address: params.email }];
    if (params.phone != null) updates.businessPhones = [params.phone];
    if (params.mobile_phone != null) updates.mobilePhone = params.mobile_phone;
    if (params.company != null) updates.companyName = params.company;
    if (params.job_title != null) updates.jobTitle = params.job_title;
    if (params.street_address != null || params.city != null || params.state != null || params.postal_code != null || params.country != null) {
      const address: Record<string, string> = {};
      if (params.street_address != null) address.street = params.street_address;
      if (params.city != null) address.city = params.city;
      if (params.state != null) address.state = params.state;
      if (params.postal_code != null) address.postalCode = params.postal_code;
      if (params.country != null) address.countryOrRegion = params.country;
      updates.businessAddress = address;
    }
    await this.repository.updateContactAsync(params.contact_id, updates);
    return { content: [{ type: 'text', text: `Successfully updated contact ${params.contact_id}` }] };
  }

  async prepareDeleteContact(params: PrepareDeleteContactParams): Promise<ToolResult> {
    const contact = await this.repository.getContactAsync(params.contact_id);
    if (contact == null) {
      return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
    }

    const graphId = this.repository.getGraphId('contact', params.contact_id);
    const graphContact = graphId != null ? await this.repository.getClient().getContact(graphId) : null;
    const hash = hashContactForApproval({
      id: params.contact_id,
      displayName: graphContact?.displayName ?? null,
      emailAddress: graphContact?.emailAddresses?.[0]?.address ?? null,
    });

    const token = this.tokenManager.generateToken({
      operation: 'delete_contact',
      targetType: 'contact',
      targetId: params.contact_id,
      targetHash: hash,
    });

    return jsonResult({
      token_id: token.tokenId,
      expires_at: new Date(token.expiresAt).toISOString(),
      contact: transformContactRow(contact),
      action: 'This contact will be permanently deleted.',
    });
  }

  async confirmDeleteContact(params: ConfirmDeleteContactParams): Promise<ToolResult> {
    const graphId = this.repository.getGraphId('contact', params.contact_id);
    const graphContact = graphId != null ? await this.repository.getClient().getContact(graphId) : null;
    const currentHash = hashContactForApproval({
      id: params.contact_id,
      displayName: graphContact?.displayName ?? null,
      emailAddress: graphContact?.emailAddresses?.[0]?.address ?? null,
    });

    const validation = this.tokenManager.consumeToken(params.token_id, 'delete_contact', params.contact_id);
    if (!validation.valid) {
      const errorMessages: Record<string, string> = {
        NOT_FOUND: 'Token not found or already used',
        EXPIRED: 'Token has expired. Please call prepare_delete_contact again.',
        OPERATION_MISMATCH: 'Token was not generated for delete_contact',
        TARGET_MISMATCH: 'Token was generated for a different contact',
        ALREADY_CONSUMED: 'Token has already been used',
      };
      return {
        content: [{ type: 'text', text: errorMessages[validation.error ?? ''] ?? 'Invalid token' }],
        isError: true,
      };
    }

    if (validation.token!.targetHash !== currentHash) {
      return {
        content: [{ type: 'text', text: 'Contact has changed since prepare was called. Please call prepare_delete_contact again.' }],
        isError: true,
      };
    }

    await this.repository.deleteContactAsync(params.contact_id);
    return { content: [{ type: 'text', text: `Successfully deleted contact ${params.contact_id}` }] };
  }

  async getContactPhoto(params: GetContactPhotoParams): Promise<ToolResult> {
    const result = await this.repository.getContactPhotoAsync(params.contact_id);
    return jsonResult({ success: true, file_path: result.filePath, content_type: result.contentType });
  }

  async setContactPhoto(params: SetContactPhotoParams): Promise<ToolResult> {
    await this.repository.setContactPhotoAsync(params.contact_id, params.file_path);
    return jsonResult({ success: true, message: 'Contact photo updated' });
  }
}
