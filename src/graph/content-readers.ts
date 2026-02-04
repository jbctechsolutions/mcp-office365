/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph API content readers.
 *
 * Implements content reader interfaces for fetching detailed content
 * directly from the Graph API.
 *
 * Path format: "graph-{type}:{id}" where:
 * - type: email, event, contact, task
 * - id: Graph API string ID
 */

import type { IContentReader } from '../tools/mail.js';
import type { IEventContentReader, EventDetails } from '../tools/calendar.js';
import type { IContactContentReader, ContactDetails } from '../tools/contacts.js';
import type { ITaskContentReader, TaskDetails } from '../tools/tasks.js';
import type { INoteContentReader, NoteDetails } from '../tools/notes.js';
import { GraphClient } from './client/index.js';
import type { DeviceCodeCallback } from './auth/index.js';

// ===========================================================================
// Path Format Constants
// ===========================================================================

/**
 * Prefix for Graph email content paths.
 * Format: "graph-email:messageId"
 */
export const GRAPH_EMAIL_PATH_PREFIX = 'graph-email:';

/**
 * Prefix for Graph event content paths.
 * Format: "graph-event:eventId"
 */
export const GRAPH_EVENT_PATH_PREFIX = 'graph-event:';

/**
 * Prefix for Graph contact content paths.
 * Format: "graph-contact:contactId"
 */
export const GRAPH_CONTACT_PATH_PREFIX = 'graph-contact:';

/**
 * Prefix for Graph task content paths.
 * Format: "graph-task:taskListId:taskId"
 */
export const GRAPH_TASK_PATH_PREFIX = 'graph-task:';

// ===========================================================================
// Utility Functions
// ===========================================================================

/**
 * Extracts the ID from a Graph content path.
 */
function extractId(path: string | null, prefix: string): string | null {
  if (path == null || !path.startsWith(prefix)) {
    return null;
  }
  return path.substring(prefix.length);
}

/**
 * Parses task path to extract taskListId and taskId.
 * Format: "graph-task:taskListId:taskId"
 */
function parseTaskPath(path: string | null): { taskListId: string; taskId: string } | null {
  const id = extractId(path, GRAPH_TASK_PATH_PREFIX);
  if (id == null) {
    return null;
  }

  const parts = id.split(':');
  if (parts.length !== 2 || parts[0] == null || parts[1] == null) {
    return null;
  }

  return { taskListId: parts[0], taskId: parts[1] };
}

// ===========================================================================
// Email Content Reader
// ===========================================================================

/**
 * Graph API email content reader.
 */
export class GraphEmailContentReader implements IContentReader {
  private readonly client: GraphClient;

  constructor(client: GraphClient) {
    this.client = client;
  }

  readEmailBody(_dataFilePath: string | null): string | null {
    // Sync method - needs async handling
    // For now, return null and rely on async version
    return null;
  }

  async readEmailBodyAsync(dataFilePath: string | null): Promise<string | null> {
    const messageId = extractId(dataFilePath, GRAPH_EMAIL_PATH_PREFIX);
    if (messageId == null) {
      return null;
    }

    try {
      const message = await this.client.getMessage(messageId);
      if (message == null) {
        return null;
      }

      // Return HTML content if available, otherwise plain text
      if (message.body?.contentType === 'html') {
        return message.body.content ?? null;
      }
      return message.body?.content ?? null;
    } catch {
      return null;
    }
  }
}

// ===========================================================================
// Event Content Reader
// ===========================================================================

/**
 * Graph API event content reader.
 */
export class GraphEventContentReader implements IEventContentReader {
  private readonly client: GraphClient;

  constructor(client: GraphClient) {
    this.client = client;
  }

  readEventDetails(_dataFilePath: string | null): EventDetails | null {
    // Sync method - needs async handling
    return null;
  }

  async readEventDetailsAsync(dataFilePath: string | null): Promise<EventDetails | null> {
    const eventId = extractId(dataFilePath, GRAPH_EVENT_PATH_PREFIX);
    if (eventId == null) {
      return null;
    }

    try {
      const event = await this.client.getEvent(eventId);
      if (event == null) {
        return null;
      }

      // Type assertion for response status due to NullableOption types
      const attendees = (event.attendees ?? []).map((a) => {
        const responseStatus = a.status as { response?: string } | null | undefined;
        return {
          email: a.emailAddress?.address ?? '',
          name: a.emailAddress?.name ?? null,
          status: mapAttendeeStatus(responseStatus?.response),
        };
      });

      return {
        title: event.subject ?? null,
        location: event.location?.displayName ?? null,
        description: event.body?.content ?? null,
        organizer: event.organizer?.emailAddress?.name ?? event.organizer?.emailAddress?.address ?? null,
        attendees,
      };
    } catch {
      return null;
    }
  }
}

/**
 * Maps Graph attendee response status to our status type.
 */
function mapAttendeeStatus(
  response: string | undefined
): 'unknown' | 'accepted' | 'declined' | 'tentative' {
  switch (response?.toLowerCase()) {
    case 'accepted':
      return 'accepted';
    case 'declined':
      return 'declined';
    case 'tentativelyaccepted':
    case 'tentative':
      return 'tentative';
    default:
      return 'unknown';
  }
}

// ===========================================================================
// Contact Content Reader
// ===========================================================================

/**
 * Graph API contact content reader.
 */
export class GraphContactContentReader implements IContactContentReader {
  private readonly client: GraphClient;

  constructor(client: GraphClient) {
    this.client = client;
  }

  readContactDetails(_dataFilePath: string | null): ContactDetails | null {
    // Sync method - needs async handling
    return null;
  }

  async readContactDetailsAsync(dataFilePath: string | null): Promise<ContactDetails | null> {
    const contactId = extractId(dataFilePath, GRAPH_CONTACT_PATH_PREFIX);
    if (contactId == null) {
      return null;
    }

    try {
      const contact = await this.client.getContact(contactId);
      if (contact == null) {
        return null;
      }

      // Build emails array
      const emails: { type: string; address: string }[] = (contact.emailAddresses ?? [])
        .filter((e) => e.address != null)
        .map((e) => ({
          type: e.name ?? 'work',
          address: e.address!,
        }));

      // Build phones array
      const phones: { type: string; number: string }[] = [];
      for (const phone of contact.homePhones ?? []) {
        phones.push({ type: 'home', number: phone });
      }
      for (const phone of contact.businessPhones ?? []) {
        phones.push({ type: 'work', number: phone });
      }
      if (contact.mobilePhone != null) {
        phones.push({ type: 'mobile', number: contact.mobilePhone });
      }

      // Build addresses array
      const addresses: {
        type: string;
        street: string | null;
        city: string | null;
        state: string | null;
        postalCode: string | null;
        country: string | null;
      }[] = [];

      if (contact.homeAddress != null) {
        addresses.push({
          type: 'home',
          street: contact.homeAddress.street ?? null,
          city: contact.homeAddress.city ?? null,
          state: contact.homeAddress.state ?? null,
          postalCode: contact.homeAddress.postalCode ?? null,
          country: contact.homeAddress.countryOrRegion ?? null,
        });
      }

      if (contact.businessAddress != null) {
        addresses.push({
          type: 'work',
          street: contact.businessAddress.street ?? null,
          city: contact.businessAddress.city ?? null,
          state: contact.businessAddress.state ?? null,
          postalCode: contact.businessAddress.postalCode ?? null,
          country: contact.businessAddress.countryOrRegion ?? null,
        });
      }

      return {
        firstName: contact.givenName ?? null,
        lastName: contact.surname ?? null,
        middleName: contact.middleName ?? null,
        nickname: contact.nickName ?? null,
        company: contact.companyName ?? null,
        jobTitle: contact.jobTitle ?? null,
        department: contact.department ?? null,
        emails,
        phones,
        addresses,
        notes: contact.personalNotes ?? null,
      };
    } catch {
      return null;
    }
  }
}

// ===========================================================================
// Task Content Reader
// ===========================================================================

/**
 * Graph API task content reader.
 */
export class GraphTaskContentReader implements ITaskContentReader {
  private readonly client: GraphClient;

  constructor(client: GraphClient) {
    this.client = client;
  }

  readTaskDetails(_dataFilePath: string | null): TaskDetails | null {
    // Sync method - needs async handling
    return null;
  }

  async readTaskDetailsAsync(dataFilePath: string | null): Promise<TaskDetails | null> {
    const taskInfo = parseTaskPath(dataFilePath);
    if (taskInfo == null) {
      return null;
    }

    try {
      const task = await this.client.getTask(taskInfo.taskListId, taskInfo.taskId);
      if (task == null) {
        return null;
      }

      return {
        body: task.body?.content ?? null,
        completedDate: task.completedDateTime?.dateTime ?? null,
        reminderDate: task.reminderDateTime?.dateTime ?? null,
        categories: [], // Graph tasks don't have categories in the same way
      };
    } catch {
      return null;
    }
  }
}

// ===========================================================================
// Note Content Reader (NOT SUPPORTED)
// ===========================================================================

/**
 * Graph API note content reader.
 *
 * Note: Microsoft Graph does not have an API for Outlook Notes.
 * This reader always returns null.
 */
export class GraphNoteContentReader implements INoteContentReader {
  readNoteDetails(_dataFilePath: string | null): NoteDetails | null {
    // Microsoft Graph does not support Outlook Notes
    return null;
  }

  readNoteDetailsAsync(_dataFilePath: string | null): Promise<NoteDetails | null> {
    // Microsoft Graph does not support Outlook Notes
    return Promise.resolve(null);
  }
}

// ===========================================================================
// Factory Functions
// ===========================================================================

/**
 * All Graph content readers bundled together.
 */
export interface GraphContentReaders {
  readonly email: GraphEmailContentReader;
  readonly event: GraphEventContentReader;
  readonly contact: GraphContactContentReader;
  readonly task: GraphTaskContentReader;
  readonly note: GraphNoteContentReader;
}

/**
 * Creates all Graph content readers.
 */
export function createGraphContentReaders(
  deviceCodeCallback?: DeviceCodeCallback
): GraphContentReaders {
  const client = new GraphClient(deviceCodeCallback);

  return {
    email: new GraphEmailContentReader(client),
    event: new GraphEventContentReader(client),
    contact: new GraphContactContentReader(client),
    task: new GraphTaskContentReader(client),
    note: new GraphNoteContentReader(),
  };
}

/**
 * Creates Graph content readers with an existing client.
 */
export function createGraphContentReadersWithClient(
  client: GraphClient
): GraphContentReaders {
  return {
    email: new GraphEmailContentReader(client),
    event: new GraphEventContentReader(client),
    contact: new GraphContactContentReader(client),
    task: new GraphTaskContentReader(client),
    note: new GraphNoteContentReader(),
  };
}
