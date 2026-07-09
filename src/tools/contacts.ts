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
import { defineTool } from '../registry/define-tool.js';
import { tokenIdLink } from '../registry/elicit-links.js';
import { requireGraphToolset } from '../registry/context.js';
import type { ToolDefinition } from '../registry/types.js';
import type { GraphContactsTools } from './contacts-graph.js';

// Contacts are served by GraphContactsTools.
declare module '../registry/types.js' {
  interface GraphToolsets {
    contactsGraph: GraphContactsTools;
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
  folder_id: z.string().min(1).optional().describe('Filter contacts by contact folder ID'),
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

// A contact id accepts either a durable `ct_…` token (Graph backend, U5) or a
// numeric id (AppleScript/SQLite backend, D4). The resolver rejects a numeric id
// on Graph with NUMERIC_ID_UNSUPPORTED.
const ContactIdSchema = z.union([z.string().min(1), z.number().int().positive()]);

export const GetContactInput = z.strictObject({
  contact_id: ContactIdSchema.describe('The contact ID to retrieve'),
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
  contact_id: ContactIdSchema,
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
  contact_id: ContactIdSchema,
});

export const ConfirmDeleteContactInput = z.strictObject({
  token_id: z.uuid(),
  contact_id: ContactIdSchema,
});

export const GetContactPhotoInput = z.strictObject({
  contact_id: ContactIdSchema.describe('Contact ID'),
});

export const SetContactPhotoInput = z.strictObject({
  contact_id: ContactIdSchema.describe('Contact ID'),
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

// =============================================================================
// Registry Definitions (v3 registry-driven architecture)
// =============================================================================

/**
 * Registry tool definitions for the contacts domain. Each handler delegates to
 * GraphContactsTools, which returns MCP content directly.
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
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').listContacts(params),
    }),
    defineTool({
      name: 'search_contacts',
      description: 'Search contacts by name',
      input: SearchContactsInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').searchContacts(params),
    }),
    defineTool({
      name: 'get_contact',
      description: 'Get contact details',
      input: GetContactInput,
      annotations: { readOnlyHint: true },
      destructive: false,
      presets: ['contacts'],
      backends: ['graph'],
      handler: (ctx, params) => requireGraphToolset(ctx, 'contactsGraph').getContact(params),
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
      onElicit: tokenIdLink('confirm_delete_contact', ['contact_id']),
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
