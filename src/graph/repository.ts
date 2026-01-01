/**
 * Microsoft Graph API repository implementation.
 *
 * Implements the IRepository interface using Microsoft Graph API
 * for data access instead of AppleScript or SQLite.
 */

import type {
  IRepository,
  FolderRow,
  EmailRow,
  EventRow,
  ContactRow,
  TaskRow,
  NoteRow,
} from '../database/repository.js';
import { GraphClient } from './client/index.js';
import {
  mapMailFolderToRow,
  mapCalendarToFolderRow,
  mapMessageToEmailRow,
  mapEventToEventRow,
  mapContactToContactRow,
  mapTaskToTaskRow,
  hashStringToNumber,
} from './mappers/index.js';
import type { DeviceCodeCallback } from './auth/index.js';

/**
 * Cache for mapping numeric IDs back to Graph string IDs.
 */
interface IdCache {
  folders: Map<number, string>;
  messages: Map<number, string>;
  events: Map<number, string>;
  contacts: Map<number, string>;
  tasks: Map<number, { taskListId: string; taskId: string }>;
}

/**
 * Repository implementation using Microsoft Graph API.
 *
 * Provides read-only access to Outlook data via the Graph API.
 */
export class GraphRepository implements IRepository {
  private readonly client: GraphClient;
  private readonly idCache: IdCache = {
    folders: new Map(),
    messages: new Map(),
    events: new Map(),
    contacts: new Map(),
    tasks: new Map(),
  };

  constructor(deviceCodeCallback?: DeviceCodeCallback) {
    this.client = new GraphClient(deviceCodeCallback);
  }

  // ===========================================================================
  // Folders
  // ===========================================================================

  listFolders(): FolderRow[] {
    // Note: Graph API is async, but IRepository interface is sync
    // We need to use a sync wrapper or change the interface
    // For now, we'll throw and require the async version
    throw new Error('Use listFoldersAsync() for Graph repository');
  }

  async listFoldersAsync(): Promise<FolderRow[]> {
    const folders = await this.client.listMailFolders();

    // Update ID cache
    for (const folder of folders) {
      if (folder.id != null) {
        const numericId = hashStringToNumber(folder.id);
        this.idCache.folders.set(numericId, folder.id);
      }
    }

    return folders.map(mapMailFolderToRow);
  }

  getFolder(id: number): FolderRow | undefined {
    throw new Error('Use getFolderAsync() for Graph repository');
  }

  async getFolderAsync(id: number): Promise<FolderRow | undefined> {
    const graphId = this.idCache.folders.get(id);
    if (graphId == null) {
      // Try to find it by listing all folders
      await this.listFoldersAsync();
      const refreshedGraphId = this.idCache.folders.get(id);
      if (refreshedGraphId == null) {
        return undefined;
      }
      const folder = await this.client.getMailFolder(refreshedGraphId);
      return folder != null ? mapMailFolderToRow(folder) : undefined;
    }

    const folder = await this.client.getMailFolder(graphId);
    return folder != null ? mapMailFolderToRow(folder) : undefined;
  }

  // ===========================================================================
  // Emails
  // ===========================================================================

  listEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    throw new Error('Use listEmailsAsync() for Graph repository');
  }

  async listEmailsAsync(folderId: number, limit: number, offset: number): Promise<EmailRow[]> {
    const graphFolderId = this.idCache.folders.get(folderId);
    if (graphFolderId == null) {
      // Refresh folder cache
      await this.listFoldersAsync();
      const refreshedId = this.idCache.folders.get(folderId);
      if (refreshedId == null) {
        return [];
      }
      return this.listEmailsWithGraphId(refreshedId, limit, offset);
    }

    return this.listEmailsWithGraphId(graphFolderId, limit, offset);
  }

  private async listEmailsWithGraphId(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    const messages = await this.client.listMessages(folderId, limit, offset);

    // Update ID cache
    for (const message of messages) {
      if (message.id != null) {
        const numericId = hashStringToNumber(message.id);
        this.idCache.messages.set(numericId, message.id);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  listUnreadEmails(folderId: number, limit: number, offset: number): EmailRow[] {
    throw new Error('Use listUnreadEmailsAsync() for Graph repository');
  }

  async listUnreadEmailsAsync(folderId: number, limit: number, offset: number): Promise<EmailRow[]> {
    const graphFolderId = this.idCache.folders.get(folderId);
    if (graphFolderId == null) {
      await this.listFoldersAsync();
      const refreshedId = this.idCache.folders.get(folderId);
      if (refreshedId == null) {
        return [];
      }
      return this.listUnreadEmailsWithGraphId(refreshedId, limit, offset);
    }

    return this.listUnreadEmailsWithGraphId(graphFolderId, limit, offset);
  }

  private async listUnreadEmailsWithGraphId(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    const messages = await this.client.listUnreadMessages(folderId, limit, offset);

    for (const message of messages) {
      if (message.id != null) {
        const numericId = hashStringToNumber(message.id);
        this.idCache.messages.set(numericId, message.id);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  searchEmails(query: string, limit: number): EmailRow[] {
    throw new Error('Use searchEmailsAsync() for Graph repository');
  }

  async searchEmailsAsync(query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessages(query, limit);

    for (const message of messages) {
      if (message.id != null) {
        const numericId = hashStringToNumber(message.id);
        this.idCache.messages.set(numericId, message.id);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m));
  }

  searchEmailsInFolder(folderId: number, query: string, limit: number): EmailRow[] {
    throw new Error('Use searchEmailsInFolderAsync() for Graph repository');
  }

  async searchEmailsInFolderAsync(folderId: number, query: string, limit: number): Promise<EmailRow[]> {
    const graphFolderId = this.idCache.folders.get(folderId);
    if (graphFolderId == null) {
      await this.listFoldersAsync();
      const refreshedId = this.idCache.folders.get(folderId);
      if (refreshedId == null) {
        return [];
      }
      return this.searchEmailsInFolderWithGraphId(refreshedId, query, limit);
    }

    return this.searchEmailsInFolderWithGraphId(graphFolderId, query, limit);
  }

  private async searchEmailsInFolderWithGraphId(folderId: string, query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessagesInFolder(folderId, query, limit);

    for (const message of messages) {
      if (message.id != null) {
        const numericId = hashStringToNumber(message.id);
        this.idCache.messages.set(numericId, message.id);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  getEmail(id: number): EmailRow | undefined {
    throw new Error('Use getEmailAsync() for Graph repository');
  }

  async getEmailAsync(id: number): Promise<EmailRow | undefined> {
    const graphId = this.idCache.messages.get(id);
    if (graphId == null) {
      return undefined;
    }

    const message = await this.client.getMessage(graphId);
    return message != null ? mapMessageToEmailRow(message) : undefined;
  }

  getUnreadCount(): number {
    throw new Error('Use getUnreadCountAsync() for Graph repository');
  }

  async getUnreadCountAsync(): Promise<number> {
    const folders = await this.client.listMailFolders();
    return folders.reduce((sum, f) => sum + (f.unreadItemCount ?? 0), 0);
  }

  getUnreadCountByFolder(folderId: number): number {
    throw new Error('Use getUnreadCountByFolderAsync() for Graph repository');
  }

  async getUnreadCountByFolderAsync(folderId: number): Promise<number> {
    const graphId = this.idCache.folders.get(folderId);
    if (graphId == null) {
      await this.listFoldersAsync();
      const refreshedId = this.idCache.folders.get(folderId);
      if (refreshedId == null) {
        return 0;
      }
      const folder = await this.client.getMailFolder(refreshedId);
      return folder?.unreadItemCount ?? 0;
    }

    const folder = await this.client.getMailFolder(graphId);
    return folder?.unreadItemCount ?? 0;
  }

  // ===========================================================================
  // Calendar
  // ===========================================================================

  listCalendars(): FolderRow[] {
    throw new Error('Use listCalendarsAsync() for Graph repository');
  }

  async listCalendarsAsync(): Promise<FolderRow[]> {
    const calendars = await this.client.listCalendars();

    for (const calendar of calendars) {
      if (calendar.id != null) {
        const numericId = hashStringToNumber(calendar.id);
        this.idCache.folders.set(numericId, calendar.id);
      }
    }

    return calendars.map(mapCalendarToFolderRow);
  }

  listEvents(limit: number): EventRow[] {
    throw new Error('Use listEventsAsync() for Graph repository');
  }

  async listEventsAsync(limit: number): Promise<EventRow[]> {
    const events = await this.client.listEvents(limit);

    for (const event of events) {
      if (event.id != null) {
        const numericId = hashStringToNumber(event.id);
        this.idCache.events.set(numericId, event.id);
      }
    }

    return events.map((e) => mapEventToEventRow(e));
  }

  listEventsByFolder(folderId: number, limit: number): EventRow[] {
    throw new Error('Use listEventsByFolderAsync() for Graph repository');
  }

  async listEventsByFolderAsync(folderId: number, limit: number): Promise<EventRow[]> {
    const graphCalendarId = this.idCache.folders.get(folderId);
    if (graphCalendarId == null) {
      return this.listEventsAsync(limit);
    }

    const events = await this.client.listEvents(limit, graphCalendarId);

    for (const event of events) {
      if (event.id != null) {
        const numericId = hashStringToNumber(event.id);
        this.idCache.events.set(numericId, event.id);
      }
    }

    return events.map((e) => mapEventToEventRow(e, graphCalendarId));
  }

  listEventsByDateRange(startDate: number, endDate: number, limit: number): EventRow[] {
    throw new Error('Use listEventsByDateRangeAsync() for Graph repository');
  }

  async listEventsByDateRangeAsync(startDate: number, endDate: number, limit: number): Promise<EventRow[]> {
    const start = new Date(startDate * 1000);
    const end = new Date(endDate * 1000);

    const events = await this.client.listEvents(limit, undefined, start, end);

    for (const event of events) {
      if (event.id != null) {
        const numericId = hashStringToNumber(event.id);
        this.idCache.events.set(numericId, event.id);
      }
    }

    return events.map((e) => mapEventToEventRow(e));
  }

  getEvent(id: number): EventRow | undefined {
    throw new Error('Use getEventAsync() for Graph repository');
  }

  async getEventAsync(id: number): Promise<EventRow | undefined> {
    const graphId = this.idCache.events.get(id);
    if (graphId == null) {
      return undefined;
    }

    const event = await this.client.getEvent(graphId);
    return event != null ? mapEventToEventRow(event) : undefined;
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  listContacts(limit: number, offset: number): ContactRow[] {
    throw new Error('Use listContactsAsync() for Graph repository');
  }

  async listContactsAsync(limit: number, offset: number): Promise<ContactRow[]> {
    const contacts = await this.client.listContacts(limit, offset);

    for (const contact of contacts) {
      if (contact.id != null) {
        const numericId = hashStringToNumber(contact.id);
        this.idCache.contacts.set(numericId, contact.id);
      }
    }

    return contacts.map(mapContactToContactRow);
  }

  searchContacts(query: string, limit: number): ContactRow[] {
    throw new Error('Use searchContactsAsync() for Graph repository');
  }

  async searchContactsAsync(query: string, limit: number): Promise<ContactRow[]> {
    const contacts = await this.client.searchContacts(query, limit);

    for (const contact of contacts) {
      if (contact.id != null) {
        const numericId = hashStringToNumber(contact.id);
        this.idCache.contacts.set(numericId, contact.id);
      }
    }

    return contacts.map(mapContactToContactRow);
  }

  getContact(id: number): ContactRow | undefined {
    throw new Error('Use getContactAsync() for Graph repository');
  }

  async getContactAsync(id: number): Promise<ContactRow | undefined> {
    const graphId = this.idCache.contacts.get(id);
    if (graphId == null) {
      return undefined;
    }

    const contact = await this.client.getContact(graphId);
    return contact != null ? mapContactToContactRow(contact) : undefined;
  }

  // ===========================================================================
  // Tasks
  // ===========================================================================

  listTasks(limit: number, offset: number): TaskRow[] {
    throw new Error('Use listTasksAsync() for Graph repository');
  }

  async listTasksAsync(limit: number, offset: number): Promise<TaskRow[]> {
    const tasks = await this.client.listAllTasks(limit, offset, true);

    for (const task of tasks) {
      if (task.id != null && task.taskListId != null) {
        const numericId = hashStringToNumber(task.id);
        this.idCache.tasks.set(numericId, { taskListId: task.taskListId, taskId: task.id });
      }
    }

    return tasks.map(mapTaskToTaskRow);
  }

  listIncompleteTasks(limit: number, offset: number): TaskRow[] {
    throw new Error('Use listIncompleteTasksAsync() for Graph repository');
  }

  async listIncompleteTasksAsync(limit: number, offset: number): Promise<TaskRow[]> {
    const tasks = await this.client.listAllTasks(limit, offset, false);

    for (const task of tasks) {
      if (task.id != null && task.taskListId != null) {
        const numericId = hashStringToNumber(task.id);
        this.idCache.tasks.set(numericId, { taskListId: task.taskListId, taskId: task.id });
      }
    }

    return tasks.map(mapTaskToTaskRow);
  }

  searchTasks(query: string, limit: number): TaskRow[] {
    throw new Error('Use searchTasksAsync() for Graph repository');
  }

  async searchTasksAsync(query: string, limit: number): Promise<TaskRow[]> {
    const tasks = await this.client.searchTasks(query, limit);

    for (const task of tasks) {
      if (task.id != null && task.taskListId != null) {
        const numericId = hashStringToNumber(task.id);
        this.idCache.tasks.set(numericId, { taskListId: task.taskListId, taskId: task.id });
      }
    }

    return tasks.map(mapTaskToTaskRow);
  }

  getTask(id: number): TaskRow | undefined {
    throw new Error('Use getTaskAsync() for Graph repository');
  }

  async getTaskAsync(id: number): Promise<TaskRow | undefined> {
    const taskInfo = this.idCache.tasks.get(id);
    if (taskInfo == null) {
      return undefined;
    }

    const task = await this.client.getTask(taskInfo.taskListId, taskInfo.taskId);
    if (task == null) {
      return undefined;
    }

    return mapTaskToTaskRow({ ...task, taskListId: taskInfo.taskListId });
  }

  // ===========================================================================
  // Notes (NOT SUPPORTED)
  // ===========================================================================

  listNotes(limit: number, offset: number): NoteRow[] {
    // Microsoft Graph does not have an API for Outlook Notes
    return [];
  }

  async listNotesAsync(limit: number, offset: number): Promise<NoteRow[]> {
    // Microsoft Graph does not have an API for Outlook Notes
    return [];
  }

  getNote(id: number): NoteRow | undefined {
    // Microsoft Graph does not have an API for Outlook Notes
    return undefined;
  }

  async getNoteAsync(id: number): Promise<NoteRow | undefined> {
    // Microsoft Graph does not have an API for Outlook Notes
    return undefined;
  }

  // ===========================================================================
  // Utility Methods
  // ===========================================================================

  /**
   * Gets the Graph client instance for direct access if needed.
   */
  getClient(): GraphClient {
    return this.client;
  }

  /**
   * Gets the Graph string ID from a numeric ID.
   */
  getGraphId(type: 'folder' | 'message' | 'event' | 'contact', numericId: number): string | undefined {
    switch (type) {
      case 'folder':
        return this.idCache.folders.get(numericId);
      case 'message':
        return this.idCache.messages.get(numericId);
      case 'event':
        return this.idCache.events.get(numericId);
      case 'contact':
        return this.idCache.contacts.get(numericId);
    }
  }

  /**
   * Gets task info from a numeric ID.
   */
  getTaskInfo(numericId: number): { taskListId: string; taskId: string } | undefined {
    return this.idCache.tasks.get(numericId);
  }
}

/**
 * Creates a Microsoft Graph API repository.
 */
export function createGraphRepository(deviceCodeCallback?: DeviceCodeCallback): GraphRepository {
  return new GraphRepository(deviceCodeCallback);
}
