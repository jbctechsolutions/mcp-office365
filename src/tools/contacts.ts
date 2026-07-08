/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Contact-related MCP tools.
 *
 * Provides tools for listing and searching contacts.
 */

import { z } from 'zod';
import type { IRepository, ContactRow } from '../database/repository.js';
import type { ContactSummary, Contact, ContactTypeValue } from '../types/index.js';
import { defineTool } from '../registry/define-tool.js';
import { requireGraphToolset, requireAppleScriptToolset } from '../registry/context.js';
import type { ToolContext, ToolDefinition, ToolResult } from '../registry/types.js';
import type { GraphContactsTools } from './contacts-graph.js';

// Contacts are a dual-backend domain: the AppleScript backend serves them via
// ContactsTools; the Graph backend serves them via GraphContactsTools.
declare module '../registry/types.js' {
  interface GraphToolsets {
    contactsGraph: GraphContactsTools;
  }
  interface AppleScriptToolsets {
    contacts: ContactsTools;
  }
}

// =============================================================================
// Input Schemas
// =============================================================================

export const ListContactsInput = z.strictObject({
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of contacts to return (1-100)'),
  offset: z.number().int().min(0).default(0).describe('Number of contacts to skip'),
  folder_id: z.number().int().positive().optional().describe('Filter contacts by contact folder ID'),
});

export const SearchContactsInput = z.strictObject({
  query: z.string().min(1).describe('Search query for contact names'),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum number of contacts to return (1-100)'),
});

export const GetContactInput = z.strictObject({
  contact_id: z.number().int().positive().describe('The contact ID to retrieve'),
});

// Contact write schemas (Graph API)
export const CreateContactInput = z.strictObject({
  given_name: z.string().optional(),
  surname: z.string().optional(),
  email: z.string().email().optional(),
  phone: z.string().optional(),
  mobile_phone: z.string().optional(),
  company: z.string().optional(),
  job_title: z.string().optional(),
  street_address: z.string().optional(),
  city: z.string().optional(),
  state: z.string().optional(),
  postal_code: z.string().optional(),
  country: z.string().optional(),
});

export const UpdateContactInput = z.strictObject({
  contact_id: z.number().int().positive(),
  given_name: z.string().optional(),
  surname: z.string().optional(),
  email: z.string().email().optional(),
  phone: z.string().optional(),
  mobile_phone: z.string().optional(),
  company: z.string().optional(),
  job_title: z.string().optional(),
  street_address: z.string().optional(),
  city: z.string().optional(),
  state: z.string().optional(),
  postal_code: z.string().optional(),
  country: z.string().optional(),
});

export const PrepareDeleteContactInput = z.strictObject({
  contact_id: z.number().int().positive(),
});

export const ConfirmDeleteContactInput = z.strictObject({
  token_id: z.uuid(),
  contact_id: z.number().int().positive(),
});

export const GetContactPhotoInput = z.strictObject({
  contact_id: z.number().int().positive().describe('Contact ID'),
});

export const SetContactPhotoInput = z.strictObject({
  contact_id: z.number().int().positive().describe('Contact ID'),
  file_path: z.string().describe('Path to the photo file (JPEG or PNG)'),
});

// =============================================================================
// Type Definitions
// =============================================================================

export type ListContactsParams = z.infer<typeof ListContactsInput>;
export type SearchContactsParams = z.infer<typeof SearchContactsInput>;
export type GetContactParams = z.infer<typeof GetContactInput>;
export type CreateContactParams = z.infer<typeof CreateContactInput>;
export type UpdateContactParams = z.infer<typeof UpdateContactInput>;
export type PrepareDeleteContactParams = z.infer<typeof PrepareDeleteContactInput>;
export type ConfirmDeleteContactParams = z.infer<typeof ConfirmDeleteContactInput>;
export type GetContactPhotoParams = z.infer<typeof GetContactPhotoInput>;
export type SetContactPhotoParams = z.infer<typeof SetContactPhotoInput>;

// =============================================================================
// Content Reader Interface
// =============================================================================

/**
 * Interface for reading contact content from data files.
 */
export interface IContactContentReader {
  /**
   * Reads contact details from the given data file path.
   */
  readContactDetails(dataFilePath: string | null): ContactDetails | null;
}

/**
 * Contact details from content file.
 */
export interface ContactDetails {
  readonly firstName: string | null;
  readonly lastName: string | null;
  readonly middleName: string | null;
  readonly nickname: string | null;
  readonly company: string | null;
  readonly jobTitle: string | null;
  readonly department: string | null;
  readonly emails: readonly { type: string; address: string }[];
  readonly phones: readonly { type: string; number: string }[];
  readonly addresses: readonly {
    type: string;
    street: string | null;
    city: string | null;
    state: string | null;
    postalCode: string | null;
    country: string | null;
  }[];
  readonly notes: string | null;
}

/**
 * Default contact content reader that returns null.
 */
export const nullContactContentReader: IContactContentReader = {
  readContactDetails: (): ContactDetails | null => null,
};

// =============================================================================
// Transformers
// =============================================================================

/**
 * Transforms a database contact row to ContactSummary.
 */
function transformContactSummary(row: ContactRow): ContactSummary {
  return {
    id: row.id,
    folderId: row.folderId,
    displayName: row.displayName,
    sortName: row.sortName,
    contactType: (row.contactType ?? 0) as ContactTypeValue,
  };
}

/**
 * Transforms a database contact row to full Contact.
 */
function transformContact(row: ContactRow, details: ContactDetails | null): Contact {
  const summary = transformContactSummary(row);

  return {
    ...summary,
    firstName: details?.firstName ?? null,
    lastName: details?.lastName ?? null,
    middleName: details?.middleName ?? null,
    nickname: details?.nickname ?? null,
    company: details?.company ?? null,
    jobTitle: details?.jobTitle ?? null,
    department: details?.department ?? null,
    emails: details?.emails?.map((e) => ({ type: e.type as 'work' | 'home' | 'other', address: e.address })) ?? [],
    phones: details?.phones?.map((p) => ({ type: p.type as 'work' | 'home' | 'mobile' | 'fax' | 'other', number: p.number })) ?? [],
    addresses: details?.addresses?.map((a) => ({
      type: a.type as 'work' | 'home' | 'other',
      street: a.street,
      city: a.city,
      state: a.state,
      postalCode: a.postalCode,
      country: a.country,
    })) ?? [],
    notes: details?.notes ?? null,
  };
}

// =============================================================================
// Contacts Tools Class
// =============================================================================

/**
 * Contacts tools implementation with dependency injection.
 */
export class ContactsTools {
  constructor(
    private readonly repository: IRepository,
    private readonly contentReader: IContactContentReader = nullContactContentReader
  ) {}

  /**
   * Lists contacts with pagination.
   */
  listContacts(params: ListContactsParams): ContactSummary[] {
    const { limit, offset } = params;
    const rows = this.repository.listContacts(limit, offset);
    return rows.map(transformContactSummary);
  }

  /**
   * Searches contacts by name.
   */
  searchContacts(params: SearchContactsParams): ContactSummary[] {
    const { query, limit } = params;
    const rows = this.repository.searchContacts(query, limit);
    return rows.map(transformContactSummary);
  }

  /**
   * Gets a single contact by ID.
   */
  getContact(params: GetContactParams): Contact | null {
    const { contact_id } = params;

    const row = this.repository.getContact(contact_id);
    if (row == null) {
      return null;
    }

    const details = this.contentReader.readContactDetails(row.dataFilePath);
    return transformContact(row, details);
  }
}

/**
 * Creates contacts tools with the given repository.
 */
export function createContactsTools(
  repository: IRepository,
  contentReader: IContactContentReader = nullContactContentReader
): ContactsTools {
  return new ContactsTools(repository, contentReader);
}

// =============================================================================
// Registry Definitions (v3 registry-driven architecture, U2 — dual backend)
// =============================================================================

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

/**
 * Registry tool definitions for the contacts domain. Each handler branches on
 * the active backend: Graph delegates to GraphContactsTools (which returns MCP
 * content directly); AppleScript delegates to ContactsTools (which returns raw
 * objects, wrapped here to match the pre-registry dispatch behavior exactly).
 */
export function contactsToolDefinitions(): ToolDefinition[] {
  return [
    defineTool({
      name: 'list_contacts',
      description: 'List contacts with pagination',
      input: ListContactsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'contactsGraph').listContacts(params)
          : jsonResult(requireAppleScriptToolset(ctx, 'contacts').listContacts(params)),
    }),
    defineTool({
      name: 'search_contacts',
      description: 'Search contacts by name',
      input: SearchContactsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph', 'applescript'],
      handler: (ctx, params) =>
        ctx.backend === 'graph'
          ? requireGraphToolset(ctx, 'contactsGraph').searchContacts(params)
          : jsonResult(requireAppleScriptToolset(ctx, 'contacts').searchContacts(params)),
    }),
    defineTool({
      name: 'get_contact',
      description: 'Get contact details',
      input: GetContactInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph', 'applescript'],
      handler: (ctx: ToolContext, params): Promise<ToolResult> | ToolResult => {
        if (ctx.backend === 'graph') {
          return requireGraphToolset(ctx, 'contactsGraph').getContact(params);
        }
        const result = requireAppleScriptToolset(ctx, 'contacts').getContact(params);
        if (result == null) {
          return { content: [{ type: 'text', text: 'Contact not found' }], isError: true };
        }
        return jsonResult(result);
      },
    }),
    defineTool({
      name: 'create_contact',
      description: 'Create a new contact in Outlook. All fields are optional but at least one should be provided.',
      input: CreateContactInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').createContact(params),
    }),
    defineTool({
      name: 'update_contact',
      description: 'Update an existing contact. Only specified fields will be updated.',
      input: UpdateContactInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').updateContact(params),
    }),
    defineTool({
      name: 'prepare_delete_contact',
      description: 'Prepare to delete a contact. Returns a preview and approval token. Call confirm_delete_contact to execute.',
      input: PrepareDeleteContactInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: true,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').prepareDeleteContact(params),
    }),
    defineTool({
      name: 'confirm_delete_contact',
      description: 'Confirm deletion of a contact using a token from prepare_delete_contact',
      input: ConfirmDeleteContactInput,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      destructive: true,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').confirmDeleteContact(params),
    }),
    defineTool({
      name: 'get_contact_photo',
      description: 'Download a contact\'s photo to a local file (Graph API)',
      input: GetContactPhotoInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').getContactPhoto(params),
    }),
    defineTool({
      name: 'set_contact_photo',
      description: 'Set or update a contact\'s photo from a local file (JPEG or PNG) (Graph API)',
      input: SetContactPhotoInput,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').setContactPhoto(params),
    }),
  ];
}
