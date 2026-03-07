/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

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
import { downloadAttachment, getDownloadDir } from './attachments.js';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Cache for mapping numeric IDs back to Graph string IDs.
 */
interface IdCache {
  folders: Map<number, string>;
  messages: Map<number, string>;
  conversations: Map<number, string>;
  events: Map<number, string>;
  contacts: Map<number, string>;
  tasks: Map<number, { taskListId: string; taskId: string }>;
  taskLists: Map<number, string>;
  attachments: Map<number, { messageId: string; attachmentId: string }>;
  rules: Map<number, string>;
  contactFolders: Map<number, string>;
  categories: Map<number, string>;
  focusedOverrides: Map<number, string>;
}

/**
 * Repository implementation using Microsoft Graph API.
 *
 * Provides read-only access to Outlook data via the Graph API.
 */
export class GraphRepository implements IRepository {
  private readonly client: GraphClient;
  private readonly deltaLinks: Map<number, string> = new Map();
  private readonly idCache: IdCache = {
    folders: new Map(),
    messages: new Map(),
    conversations: new Map(),
    events: new Map(),
    contacts: new Map(),
    tasks: new Map(),
    taskLists: new Map(),
    attachments: new Map(),
    rules: new Map(),
    contactFolders: new Map(),
    categories: new Map(),
    focusedOverrides: new Map(),
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

  getFolder(_id: number): FolderRow | undefined {
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

  listEmails(_folderId: number, _limit: number, _offset: number): EmailRow[] {
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
      if (message.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(message.conversationId), message.conversationId);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  listUnreadEmails(_folderId: number, _limit: number, _offset: number): EmailRow[] {
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
      if (message.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(message.conversationId), message.conversationId);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  searchEmails(_query: string, _limit: number): EmailRow[] {
    throw new Error('Use searchEmailsAsync() for Graph repository');
  }

  async searchEmailsAsync(query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessages(query, limit);

    for (const message of messages) {
      if (message.id != null) {
        const numericId = hashStringToNumber(message.id);
        this.idCache.messages.set(numericId, message.id);
      }
      if (message.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(message.conversationId), message.conversationId);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m));
  }

  searchEmailsInFolder(_folderId: number, _query: string, _limit: number): EmailRow[] {
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
      if (message.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(message.conversationId), message.conversationId);
      }
    }

    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  /**
   * Advanced search using raw KQL query syntax.
   */
  async searchEmailsAdvancedAsync(query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessagesKql(query, limit);
    for (const msg of messages) {
      if (msg.id != null) {
        this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
      }
      if (msg.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(msg.conversationId), msg.conversationId);
      }
    }
    return messages.map((m) => mapMessageToEmailRow(m));
  }

  /**
   * Advanced search in a specific folder using raw KQL query syntax.
   */
  async searchEmailsAdvancedInFolderAsync(folderId: number, query: string, limit: number): Promise<EmailRow[]> {
    const graphFolderId = this.idCache.folders.get(folderId);
    if (graphFolderId == null) {
      throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    }
    const messages = await this.client.searchMessagesKqlInFolder(graphFolderId, query, limit);
    for (const msg of messages) {
      if (msg.id != null) {
        this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
      }
      if (msg.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(msg.conversationId), msg.conversationId);
      }
    }
    return messages.map((m) => mapMessageToEmailRow(m));
  }

  async checkNewEmailsAsync(folderId: number): Promise<{ emails: EmailRow[]; isInitialSync: boolean }> {
    const graphFolderId = this.idCache.folders.get(folderId);
    if (graphFolderId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const existingDeltaLink = this.deltaLinks.get(folderId);
    const isInitialSync = existingDeltaLink == null;

    const { messages, deltaLink } = await this.client.getMessagesDelta(
      graphFolderId,
      existingDeltaLink
    );

    if (deltaLink) {
      this.deltaLinks.set(folderId, deltaLink);
    }

    for (const msg of messages) {
      if (msg.id != null) {
        this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
      }
      if (msg.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(msg.conversationId), msg.conversationId);
      }
    }

    const activeMessages = messages.filter((m) => !(m as any)['@removed']);
    return {
      emails: activeMessages.map((m) => mapMessageToEmailRow(m)),
      isInitialSync,
    };
  }

  getEmail(_id: number): EmailRow | undefined {
    throw new Error('Use getEmailAsync() for Graph repository');
  }

  /**
   * Populates the message ID cache by listing messages from mail folders.
   * Used as a fallback when getEmailAsync is called with an ID not yet in cache
   * (e.g. after server restart or when list_emails/search_emails wasn't called first).
   */
  private async refreshMessageCacheForGetEmail(targetId: number): Promise<boolean> {
    let folders: FolderRow[];
    try {
      folders = await this.listFoldersAsync();
    } catch {
      return false;
    }
    if (folders.length === 0) return false;
    const MESSAGE_LIMIT_PER_FOLDER = 100;
    const MAX_FOLDERS_TO_SCAN = 15;
    for (let i = 0; i < Math.min(folders.length, MAX_FOLDERS_TO_SCAN); i++) {
      const folder = folders[i]!;
      const graphFolderId = this.idCache.folders.get(folder.id);
      if (graphFolderId == null) continue;
      let messages: Array<{ id?: string | null; conversationId?: string | null }>;
      try {
        messages = await this.client.listMessages(graphFolderId, MESSAGE_LIMIT_PER_FOLDER, 0);
      } catch {
        continue;
      }
      if (!Array.isArray(messages)) continue;
      for (const message of messages) {
        if (message.id != null) {
          const numericId = hashStringToNumber(message.id);
          this.idCache.messages.set(numericId, message.id);
          if (numericId === targetId) return true;
        }
        if (message.conversationId != null) {
          this.idCache.conversations.set(hashStringToNumber(message.conversationId), message.conversationId);
        }
      }
    }
    return this.idCache.messages.has(targetId);
  }

  async getEmailAsync(id: number): Promise<EmailRow | undefined> {
    let graphId = this.idCache.messages.get(id);
    if (graphId == null) {
      const found = await this.refreshMessageCacheForGetEmail(id);
      if (found) graphId = this.idCache.messages.get(id) ?? undefined;
    }
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

  getUnreadCountByFolder(_folderId: number): number {
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
  // Conversation / Thread
  // ===========================================================================

  /**
   * Lists all messages in a conversation thread.
   *
   * Looks up the message to get its conversationId, resolves the Graph string
   * conversationId from cache, then queries for all messages with that ID.
   */
  async listConversationAsync(messageId: number, limit: number): Promise<EmailRow[]> {
    const email = await this.getEmailAsync(messageId);
    if (email == null) throw new Error(`Message ID ${messageId} not found`);
    if (email.conversationId == null) throw new Error('Message has no conversation ID');

    const graphConversationId = this.idCache.conversations.get(email.conversationId);
    if (graphConversationId == null) throw new Error('Conversation ID not found in cache. Try fetching the email first to populate the cache.');

    const messages = await this.client.listConversationMessages(graphConversationId, limit);
    for (const msg of messages) {
      if (msg.id != null) {
        this.idCache.messages.set(hashStringToNumber(msg.id), msg.id);
      }
      if (msg.conversationId != null) {
        this.idCache.conversations.set(hashStringToNumber(msg.conversationId), msg.conversationId);
      }
    }
    return messages.map((m) => mapMessageToEmailRow(m));
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

  listEvents(_limit: number): EventRow[] {
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

  listEventsByFolder(_folderId: number, _limit: number): EventRow[] {
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

  searchEvents(_query: string | null, _startDate: string | null, _endDate: string | null, _limit: number): EventRow[] {
    throw new Error('Use searchEventsAsync() for Graph repository');
  }

  async searchEventsAsync(query: string | null, startDate: string | null, endDate: string | null, limit: number): Promise<EventRow[]> {
    // Graph doesn't have direct event search, so we filter client-side
    const start = startDate != null ? new Date(startDate) : undefined;
    const end = endDate != null ? new Date(endDate) : undefined;

    const events = await this.client.listEvents(1000, undefined, start, end);

    for (const event of events) {
      if (event.id != null) {
        const numericId = hashStringToNumber(event.id);
        this.idCache.events.set(numericId, event.id);
      }
    }

    let rows = events.map((e) => mapEventToEventRow(e));

    // Filter by title client-side if query provided
    if (query != null) {
      const queryLower = query.toLowerCase();
      rows = rows.filter((row) => {
        // EventRow doesn't have title, so we need to check the original event
        const originalEvent = events.find((e) => e.id != null && hashStringToNumber(e.id) === row.id);
        const subject = originalEvent?.subject?.toLowerCase() ?? '';
        return subject.includes(queryLower);
      });
    }

    return rows.slice(0, limit);
  }

  listEventsByDateRange(_startDate: number, _endDate: number, _limit: number): EventRow[] {
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

  getEvent(_id: number): EventRow | undefined {
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

  async listEventInstancesAsync(
    eventId: number,
    startDate: string,
    endDate: string
  ): Promise<EventRow[]> {
    const graphId = this.idCache.events.get(eventId);
    if (graphId == null) {
      throw new Error(`Event ID ${eventId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    }

    const instances = await this.client.listEventInstances(graphId, startDate, endDate);
    for (const inst of instances) {
      if (inst.id != null) {
        this.idCache.events.set(hashStringToNumber(inst.id), inst.id);
      }
    }
    return instances.map((e) => mapEventToEventRow(e));
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  listContacts(_limit: number, _offset: number): ContactRow[] {
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

  searchContacts(_query: string, _limit: number): ContactRow[] {
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

  getContact(_id: number): ContactRow | undefined {
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
  // Contact Folders
  // ===========================================================================

  async listContactFoldersAsync(): Promise<Array<{ id: number; name: string; parentFolderId: string | null }>> {
    const folders = await this.client.listContactFolders();
    return folders.map((folder) => {
      const graphId = folder.id!;
      const numericId = hashStringToNumber(graphId);
      this.idCache.contactFolders.set(numericId, graphId);
      return {
        id: numericId,
        name: folder.displayName ?? '',
        parentFolderId: folder.parentFolderId ?? null,
      };
    });
  }

  async createContactFolderAsync(name: string): Promise<number> {
    const created = await this.client.createContactFolder(name);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.contactFolders.set(numericId, graphId);
    return numericId;
  }

  async deleteContactFolderAsync(folderId: number): Promise<void> {
    const graphId = this.idCache.contactFolders.get(folderId);
    if (graphId == null) throw new Error(`Contact folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteContactFolder(graphId);
    this.idCache.contactFolders.delete(folderId);
  }

  async listContactsInFolderAsync(folderId: number, limit: number = 100): Promise<ContactRow[]> {
    const graphId = this.idCache.contactFolders.get(folderId);
    if (graphId == null) throw new Error(`Contact folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    const contacts = await this.client.listContactsInFolder(graphId, limit);
    return contacts.map((c) => {
      if (c.id != null) {
        this.idCache.contacts.set(hashStringToNumber(c.id), c.id);
      }
      return mapContactToContactRow(c);
    });
  }

  // ===========================================================================
  // Contact Photos
  // ===========================================================================

  async getContactPhotoAsync(contactId: number): Promise<{ filePath: string; contentType: string }> {
    const graphId = this.idCache.contacts.get(contactId);
    if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const photoData = await this.client.getContactPhoto(graphId);
    const downloadDir = getDownloadDir();
    const filePath = path.join(downloadDir, `contact-${contactId}-photo.jpg`);
    fs.writeFileSync(filePath, Buffer.from(photoData));
    return { filePath, contentType: 'image/jpeg' };
  }

  async setContactPhotoAsync(contactId: number, filePath: string): Promise<void> {
    const graphId = this.idCache.contacts.get(contactId);
    if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const photoData = fs.readFileSync(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const contentType = ext === '.png' ? 'image/png' : 'image/jpeg';
    await this.client.setContactPhoto(graphId, photoData, contentType);
  }

  // ===========================================================================
  // Tasks
  // ===========================================================================

  listTasks(_limit: number, _offset: number): TaskRow[] {
    throw new Error('Use listTasksAsync() for Graph repository');
  }

  async listTasksAsync(limit: number, offset: number): Promise<TaskRow[]> {
    const tasks = await this.client.listAllTasks(limit, offset, true);

    for (const task of tasks) {
      if (task.id != null && task.taskListId != null) {
        const numericId = hashStringToNumber(task.id);
        this.idCache.tasks.set(numericId, { taskListId: task.taskListId, taskId: task.id });
        const listNumericId = hashStringToNumber(task.taskListId);
        this.idCache.taskLists.set(listNumericId, task.taskListId);
      }
    }

    return tasks.map(mapTaskToTaskRow);
  }

  listIncompleteTasks(_limit: number, _offset: number): TaskRow[] {
    throw new Error('Use listIncompleteTasksAsync() for Graph repository');
  }

  async listIncompleteTasksAsync(limit: number, offset: number): Promise<TaskRow[]> {
    const tasks = await this.client.listAllTasks(limit, offset, false);

    for (const task of tasks) {
      if (task.id != null && task.taskListId != null) {
        const numericId = hashStringToNumber(task.id);
        this.idCache.tasks.set(numericId, { taskListId: task.taskListId, taskId: task.id });
        const listNumericId = hashStringToNumber(task.taskListId);
        this.idCache.taskLists.set(listNumericId, task.taskListId);
      }
    }

    return tasks.map(mapTaskToTaskRow);
  }

  searchTasks(_query: string, _limit: number): TaskRow[] {
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

  getTask(_id: number): TaskRow | undefined {
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

  async listTaskListsAsync(): Promise<Array<{ id: number; name: string; isDefault: boolean }>> {
    const lists = await this.client.listTaskLists();
    return lists.map((list) => {
      const graphId = list.id!;
      const numericId = hashStringToNumber(graphId);
      this.idCache.taskLists.set(numericId, graphId);
      return {
        id: numericId,
        name: list.displayName ?? '',
        isDefault: list.wellknownListName === 'defaultList',
      };
    });
  }

  // ===========================================================================
  // Notes (NOT SUPPORTED)
  // ===========================================================================

  listNotes(_limit: number, _offset: number): NoteRow[] {
    // Microsoft Graph does not have an API for Outlook Notes
    return [];
  }

  listNotesAsync(_limit: number, _offset: number): Promise<NoteRow[]> {
    // Microsoft Graph does not have an API for Outlook Notes
    return Promise.resolve([]);
  }

  getNote(_id: number): NoteRow | undefined {
    // Microsoft Graph does not have an API for Outlook Notes
    return undefined;
  }

  getNoteAsync(_id: number): Promise<NoteRow | undefined> {
    // Microsoft Graph does not have an API for Outlook Notes
    return Promise.resolve(undefined);
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
   * Returns the Graph client (satisfies IMailSendRepository).
   */
  getGraphClient(): GraphClient {
    return this.client;
  }

  /**
   * Returns the Graph string ID for a cached draft numeric ID (satisfies IMailSendRepository).
   */
  getGraphIdForDraft(draftId: number): string | undefined {
    return this.idCache.messages.get(draftId);
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

  /**
   * Gets the Graph string ID for a task list from a numeric ID.
   */
  getTaskListGraphId(numericId: number): string | undefined {
    return this.idCache.taskLists.get(numericId);
  }

  // ===========================================================================
  // Write Operations (Async)
  // ===========================================================================

  // Sync versions throw — use async versions from index.ts handler
  moveEmail(_emailId: number, _destinationFolderId: number): void {
    throw new Error('Use moveEmailAsync() for Graph repository');
  }
  deleteEmail(_emailId: number): void {
    throw new Error('Use deleteEmailAsync() for Graph repository');
  }
  archiveEmail(_emailId: number): void {
    throw new Error('Use archiveEmailAsync() for Graph repository');
  }
  junkEmail(_emailId: number): void {
    throw new Error('Use junkEmailAsync() for Graph repository');
  }
  markEmailRead(_emailId: number, _isRead: boolean): void {
    throw new Error('Use markEmailReadAsync() for Graph repository');
  }
  setEmailFlag(_emailId: number, _flagStatus: number): void {
    throw new Error('Use setEmailFlagAsync() for Graph repository');
  }
  setEmailCategories(_emailId: number, _categories: string[]): void {
    throw new Error('Use setEmailCategoriesAsync() for Graph repository');
  }
  setEmailImportance(_emailId: number, _importance: string): void {
    throw new Error('Use setEmailImportanceAsync() for Graph repository');
  }
  createFolder(_name: string, _parentFolderId?: number): FolderRow {
    throw new Error('Use createFolderAsync() for Graph repository');
  }
  deleteFolder(_folderId: number): void {
    throw new Error('Use deleteFolderAsync() for Graph repository');
  }
  renameFolder(_folderId: number, _newName: string): void {
    throw new Error('Use renameFolderAsync() for Graph repository');
  }
  moveFolder(_folderId: number, _destinationParentId: number): void {
    throw new Error('Use moveFolderAsync() for Graph repository');
  }
  emptyFolder(_folderId: number): void {
    throw new Error('Use emptyFolderAsync() for Graph repository');
  }

  // Async implementations

  async moveEmailAsync(emailId: number, destinationFolderId: number): Promise<void> {
    const graphMessageId = this.idCache.messages.get(emailId);
    const graphFolderId = this.idCache.folders.get(destinationFolderId);
    if (graphMessageId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    if (graphFolderId == null) throw new Error(`Folder ID ${destinationFolderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.moveMessage(graphMessageId, graphFolderId);
  }

  async deleteEmailAsync(emailId: number): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteMessage(graphId);
  }

  async archiveEmailAsync(emailId: number): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.archiveMessage(graphId);
  }

  async junkEmailAsync(emailId: number): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.junkMessage(graphId);
  }

  async markEmailReadAsync(emailId: number, isRead: boolean): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateMessage(graphId, { isRead });
  }

  async setEmailFlagAsync(emailId: number, flagStatus: number): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    const flagStatusMap: Record<number, string> = {
      0: 'notFlagged',
      1: 'flagged',
      2: 'complete',
    };
    await this.client.updateMessage(graphId, {
      flag: { flagStatus: flagStatusMap[flagStatus] ?? 'notFlagged' },
    });
  }

  async setEmailCategoriesAsync(emailId: number, categories: string[]): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateMessage(graphId, { categories });
  }

  async setEmailImportanceAsync(emailId: number, importance: string): Promise<void> {
    const graphId = this.idCache.messages.get(emailId);
    if (graphId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateMessage(graphId, { importance });
  }

  async createFolderAsync(name: string, parentFolderId?: number): Promise<FolderRow> {
    const graphParentId = parentFolderId != null
      ? this.idCache.folders.get(parentFolderId)
      : undefined;

    const folder = await this.client.createMailFolder(name, graphParentId ?? undefined);

    // Update cache with new folder
    if (folder.id != null) {
      const numericId = hashStringToNumber(folder.id);
      this.idCache.folders.set(numericId, folder.id);
    }

    return mapMailFolderToRow(folder);
  }

  async deleteFolderAsync(folderId: number): Promise<void> {
    const graphId = this.idCache.folders.get(folderId);
    if (graphId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteMailFolder(graphId);
    this.idCache.folders.delete(folderId);
  }

  async renameFolderAsync(folderId: number, newName: string): Promise<void> {
    const graphId = this.idCache.folders.get(folderId);
    if (graphId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.renameMailFolder(graphId, newName);
  }

  async moveFolderAsync(folderId: number, destinationParentId: number): Promise<void> {
    const graphFolderId = this.idCache.folders.get(folderId);
    const graphParentId = this.idCache.folders.get(destinationParentId);
    if (graphFolderId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    if (graphParentId == null) throw new Error(`Parent folder ID ${destinationParentId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.moveMailFolder(graphFolderId, graphParentId);
  }

  async emptyFolderAsync(folderId: number): Promise<void> {
    const graphId = this.idCache.folders.get(folderId);
    if (graphId == null) throw new Error(`Folder ID ${folderId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.emptyMailFolder(graphId);
  }

  // ===========================================================================
  // Draft & Send Operations (Async)
  // ===========================================================================

  /**
   * Creates a new draft message.
   *
   * Converts email address strings to Recipient objects, calls the Graph client,
   * adds the returned draft to idCache.messages, and returns its numeric ID.
   */
  async createDraftAsync(params: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    to?: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<{ numericId: number; graphId: string }> {
    const toRecipients = (params.to ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const ccRecipients = (params.cc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const bccRecipients = (params.bcc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));

    const draft = await this.client.createDraft({
      subject: params.subject,
      body: { contentType: params.bodyType, content: params.body },
      toRecipients,
      ccRecipients,
      bccRecipients,
    });

    const graphId = draft.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.messages.set(numericId, graphId);
    return { numericId, graphId };
  }

  /**
   * Updates an existing draft message.
   *
   * Looks up the Graph string ID from idCache.messages, then calls the client.
   */
  async updateDraftAsync(draftId: number, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.idCache.messages.get(draftId);
    if (graphId == null) throw new Error(`Message ID ${draftId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateDraft(graphId, updates);
  }

  /**
   * Lists draft messages.
   *
   * Uses the well-known 'drafts' folder name directly with the Graph API.
   */
  async listDraftsAsync(limit: number, offset: number): Promise<EmailRow[]> {
    return this.listEmailsWithGraphId('drafts', limit, offset);
  }

  /**
   * Sends an existing draft message.
   *
   * Looks up the Graph string ID from idCache.messages, then calls the client.
   */
  async sendDraftAsync(draftId: number): Promise<void> {
    const graphId = this.idCache.messages.get(draftId);
    if (graphId == null) throw new Error(`Message ID ${draftId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.sendDraft(graphId);
  }

  /**
   * Sends a new email directly without creating a draft first.
   *
   * Converts email address strings to Recipient objects and calls the client.
   */
  async sendMailAsync(params: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    to: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<void> {
    const toRecipients = params.to.map(addr => ({
      emailAddress: { address: addr },
    }));
    const ccRecipients = (params.cc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));
    const bccRecipients = (params.bcc ?? []).map(addr => ({
      emailAddress: { address: addr },
    }));

    await this.client.sendMail({
      subject: params.subject,
      body: { contentType: params.bodyType, content: params.body },
      toRecipients,
      ccRecipients,
      bccRecipients,
    });
  }

  /**
   * Replies to a message (or replies all).
   *
   * Looks up the Graph string ID from idCache.messages, then calls the client.
   */
  async replyMessageAsync(messageId: number, comment: string, replyAll: boolean): Promise<void> {
    const graphId = this.idCache.messages.get(messageId);
    if (graphId == null) throw new Error(`Message ID ${messageId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.replyMessage(graphId, comment, replyAll);
  }

  /**
   * Forwards a message to specified recipients.
   *
   * Looks up the Graph string ID from idCache.messages, converts recipient
   * email strings to Recipient objects, then calls the client.
   */
  async forwardMessageAsync(messageId: number, toRecipients: string[], comment?: string): Promise<void> {
    const graphId = this.idCache.messages.get(messageId);
    if (graphId == null) throw new Error(`Message ID ${messageId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    const recipients = toRecipients.map(addr => ({
      emailAddress: { address: addr },
    }));
    await this.client.forwardMessage(graphId, recipients, comment);
  }

  // ===========================================================================
  // Reply/Forward as Draft Operations (Async)
  // ===========================================================================

  /**
   * Creates a reply (or reply-all) draft for a message.
   *
   * Looks up the Graph string ID from idCache.messages, creates the draft
   * via the client, caches the new draft ID, and optionally updates the body.
   *
   * @returns The numeric and graph IDs of the new draft.
   */
  async replyAsDraftAsync(
    messageId: number,
    replyAll = false,
    comment?: string,
    bodyType: string = 'text',
  ): Promise<{ numericId: number; graphId: string }> {
    const graphMessageId = this.idCache.messages.get(messageId);
    if (graphMessageId == null) throw new Error(`Message ID ${messageId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const draft = replyAll
      ? await this.client.createReplyAllDraft(graphMessageId)
      : await this.client.createReplyDraft(graphMessageId);

    const graphId = draft.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.messages.set(numericId, graphId);

    if (comment != null) {
      await this.client.updateDraft(graphId, {
        body: { contentType: bodyType, content: comment },
      });
    }

    return { numericId, graphId };
  }

  /**
   * Creates a forward draft for a message.
   *
   * Looks up the Graph string ID from idCache.messages, creates the draft
   * via the client, caches the new draft ID, and optionally updates the
   * recipients and body.
   *
   * @returns The numeric and graph IDs of the new draft.
   */
  async forwardAsDraftAsync(
    messageId: number,
    toRecipients?: string[],
    comment?: string,
    bodyType: string = 'text',
  ): Promise<{ numericId: number; graphId: string }> {
    const graphMessageId = this.idCache.messages.get(messageId);
    if (graphMessageId == null) throw new Error(`Message ID ${messageId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const draft = await this.client.createForwardDraft(graphMessageId);

    const graphId = draft.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.messages.set(numericId, graphId);

    const updates: Record<string, unknown> = {};
    if (toRecipients != null && toRecipients.length > 0) {
      updates.toRecipients = toRecipients.map(addr => ({
        emailAddress: { address: addr },
      }));
    }
    if (comment != null) {
      updates.body = { contentType: bodyType, content: comment };
    }
    if (Object.keys(updates).length > 0) {
      await this.client.updateDraft(graphId, updates);
    }

    return { numericId, graphId };
  }

  // ---------------------------------------------------------------------------
  // Calendar Scheduling
  // ---------------------------------------------------------------------------

  async getScheduleAsync(params: {
    emailAddresses: string[];
    startTime: string;
    endTime: string;
    availabilityViewInterval?: number;
  }): Promise<unknown[]> {
    return await this.client.getSchedule({
      schedules: params.emailAddresses,
      startTime: { dateTime: params.startTime, timeZone: 'UTC' },
      endTime: { dateTime: params.endTime, timeZone: 'UTC' },
      availabilityViewInterval: params.availabilityViewInterval ?? 30,
    });
  }

  async findMeetingTimesAsync(params: {
    attendees: string[];
    durationMinutes: number;
    startTime?: string;
    endTime?: string;
    maxCandidates?: number;
  }): Promise<unknown> {
    const hours = Math.floor(params.durationMinutes / 60);
    const minutes = params.durationMinutes % 60;
    const meetingDuration = `PT${hours}H${minutes}M`;

    const attendees = params.attendees.map(addr => ({
      emailAddress: { address: addr },
      type: 'required' as const,
    }));

    const request: {
      attendees: Array<{ emailAddress: { address: string }; type: string }>;
      meetingDuration: string;
      maxCandidates: number;
      timeConstraint?: {
        timeslots: Array<{
          start: { dateTime: string; timeZone: string };
          end: { dateTime: string; timeZone: string };
        }>;
      };
    } = {
      attendees,
      meetingDuration,
      maxCandidates: params.maxCandidates ?? 5,
    };

    if (params.startTime != null && params.endTime != null) {
      request.timeConstraint = {
        timeslots: [{
          start: { dateTime: params.startTime, timeZone: 'UTC' },
          end: { dateTime: params.endTime, timeZone: 'UTC' },
        }],
      };
    }

    return await this.client.findMeetingTimes(request);
  }

  // ===========================================================================
  // Attachment Operations (Async)
  // ===========================================================================

  /**
   * Lists attachments for a given email.
   *
   * Looks up the Graph message ID from idCache.messages, calls
   * client.listAttachments, hashes each attachment ID to a numeric key,
   * and caches it in idCache.attachments with { messageId, attachmentId }.
   *
   * @returns Array of attachment metadata objects.
   */
  async listAttachmentsAsync(emailId: number): Promise<Array<{
    id: number;
    name: string;
    size: number;
    contentType: string;
    isInline: boolean;
  }>> {
    const graphMessageId = this.idCache.messages.get(emailId);
    if (graphMessageId == null) throw new Error(`Message ID ${emailId} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const attachments = await this.client.listAttachments(graphMessageId);

    return attachments.map((att) => {
      const attId = att.id ?? '';
      const numericId = hashStringToNumber(attId);
      this.idCache.attachments.set(numericId, {
        messageId: graphMessageId,
        attachmentId: attId,
      });
      return {
        id: numericId,
        name: att.name ?? '',
        size: att.size ?? 0,
        contentType: att.contentType ?? 'application/octet-stream',
        isInline: (att as { isInline?: boolean }).isInline ?? false,
      };
    });
  }

  /**
   * Downloads an attachment for a given email.
   *
   * Looks up { messageId, attachmentId } from idCache.attachments,
   * then delegates to the downloadAttachment helper which fetches
   * the content and writes it to disk.
   *
   * @returns Metadata about the downloaded file including its local path.
   */
  async downloadAttachmentAsync(
    attachmentId: number,
  ): Promise<{ filePath: string; name: string; size: number; contentType: string }> {
    const cached = this.idCache.attachments.get(attachmentId);
    if (cached == null) throw new Error(`Attachment ID ${attachmentId} not found in cache. Call list_attachments first.`);

    return downloadAttachment(this.client, cached.messageId, cached.attachmentId);
  }

  // ===========================================================================
  // Calendar Write Operations (Async)
  // ===========================================================================

  /**
   * Creates a new calendar event.
   *
   * Builds a Graph API event object from the given params, calls
   * client.createEvent(), adds the result to idCache.events, and
   * returns the numeric ID.
   */
  async createEventAsync(params: {
    subject: string;
    start: string;
    end: string;
    timezone?: string;
    location?: string;
    body?: string;
    bodyType?: 'text' | 'html';
    attendees?: Array<{ email: string; name?: string; type?: 'required' | 'optional' }>;
    isAllDay?: boolean;
    recurrence?: {
      pattern: {
        type: 'daily' | 'weekly' | 'monthly' | 'yearly';
        interval: number;
        daysOfWeek?: string[];
      };
      range: {
        type: 'endDate' | 'noEnd' | 'numbered';
        startDate: string;
        endDate?: string;
        numberOfOccurrences?: number;
      };
    };
    calendarId?: number;
  }): Promise<number> {
    const tz = params.timezone ?? Intl.DateTimeFormat().resolvedOptions().timeZone;

    const graphEvent: Record<string, unknown> = {
      subject: params.subject,
      start: { dateTime: params.start, timeZone: tz },
      end: { dateTime: params.end, timeZone: tz },
    };

    if (params.isAllDay === true) {
      graphEvent.isAllDay = true;
    }

    if (params.location != null) {
      graphEvent.location = { displayName: params.location };
    }

    if (params.body != null) {
      graphEvent.body = {
        contentType: params.bodyType ?? 'text',
        content: params.body,
      };
    }

    if (params.attendees != null && params.attendees.length > 0) {
      graphEvent.attendees = params.attendees.map((a) => ({
        emailAddress: { address: a.email, name: a.name },
        type: a.type ?? 'required',
      }));
    }

    if (params.recurrence != null) {
      graphEvent.recurrence = params.recurrence;
    }

    const graphCalendarId = params.calendarId != null
      ? this.idCache.folders.get(params.calendarId)
      : undefined;

    const created = await this.client.createEvent(graphEvent, graphCalendarId);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.events.set(numericId, graphId);
    return numericId;
  }

  /**
   * Updates an existing calendar event.
   *
   * Looks up the Graph string ID from idCache.events, then calls
   * client.updateEvent(). Throws if the event is not cached.
   */
  async updateEventAsync(eventId: number, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.idCache.events.get(eventId);
    if (graphId == null) throw new Error(`Event ID ${eventId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateEvent(graphId, updates);
  }

  /**
   * Deletes a calendar event.
   *
   * Looks up the Graph string ID from idCache.events, calls
   * client.deleteEvent(), and removes the entry from idCache.
   * Throws if the event is not cached.
   */
  async deleteEventAsync(eventId: number): Promise<void> {
    const graphId = this.idCache.events.get(eventId);
    if (graphId == null) throw new Error(`Event ID ${eventId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteEvent(graphId);
    this.idCache.events.delete(eventId);
  }

  /**
   * Responds to a calendar event invitation.
   *
   * Looks up the Graph string ID from idCache.events, then calls
   * client.respondToEvent(). Throws if the event is not cached.
   */
  async respondToEventAsync(
    eventId: number,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): Promise<void> {
    const graphId = this.idCache.events.get(eventId);
    if (graphId == null) throw new Error(`Event ID ${eventId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.respondToEvent(graphId, response, sendResponse, comment);
  }

  // ===========================================================================
  // Contact Write Operations (Async)
  // ===========================================================================

  /**
   * Creates a new contact.
   *
   * Maps snake_case input fields to Graph API camelCase fields, calls
   * client.createContact(), caches the resulting ID, and returns a numeric ID.
   */
  async createContactAsync(params: {
    given_name?: string;
    surname?: string;
    email?: string;
    phone?: string;
    mobile_phone?: string;
    company?: string;
    job_title?: string;
    street_address?: string;
    city?: string;
    state?: string;
    postal_code?: string;
    country?: string;
  }): Promise<number> {
    const graphContact: Record<string, unknown> = {};
    if (params.given_name != null) graphContact.givenName = params.given_name;
    if (params.surname != null) graphContact.surname = params.surname;
    if (params.email != null) graphContact.emailAddresses = [{ address: params.email }];
    if (params.phone != null) graphContact.businessPhones = [params.phone];
    if (params.mobile_phone != null) graphContact.mobilePhone = params.mobile_phone;
    if (params.company != null) graphContact.companyName = params.company;
    if (params.job_title != null) graphContact.jobTitle = params.job_title;

    // Build address only if any address field is present
    if (params.street_address != null || params.city != null || params.state != null || params.postal_code != null || params.country != null) {
      const address: Record<string, string> = {};
      if (params.street_address != null) address.street = params.street_address;
      if (params.city != null) address.city = params.city;
      if (params.state != null) address.state = params.state;
      if (params.postal_code != null) address.postalCode = params.postal_code;
      if (params.country != null) address.countryOrRegion = params.country;
      graphContact.businessAddress = address;
    }

    const created = await this.client.createContact(graphContact);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.contacts.set(numericId, graphId);
    return numericId;
  }

  /**
   * Updates an existing contact.
   *
   * Looks up the Graph string ID from idCache.contacts, then calls
   * client.updateContact(). Throws if the contact is not cached.
   */
  async updateContactAsync(contactId: number, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.idCache.contacts.get(contactId);
    if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateContact(graphId, updates);
  }

  /**
   * Deletes a contact.
   *
   * Looks up the Graph string ID from idCache.contacts, calls
   * client.deleteContact(), and removes the entry from idCache.
   * Throws if the contact is not cached.
   */
  async deleteContactAsync(contactId: number): Promise<void> {
    const graphId = this.idCache.contacts.get(contactId);
    if (graphId == null) throw new Error(`Contact ID ${contactId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteContact(graphId);
    this.idCache.contacts.delete(contactId);
  }

  // ===========================================================================
  // Task Write Operations (Async)
  // ===========================================================================

  /**
   * Creates a new task in a task list.
   *
   * Looks up the Graph task list ID from idCache.taskLists, builds a
   * Graph API task object from the given params, calls client.createTask(),
   * caches the resulting ID, and returns a numeric ID.
   */
  async createTaskAsync(params: {
    title: string;
    task_list_id: number;
    body?: string;
    body_type?: 'text' | 'html';
    due_date?: string;
    importance?: 'low' | 'normal' | 'high';
    reminder_date?: string;
    recurrence?: {
      pattern: 'daily' | 'weekly' | 'monthly' | 'yearly';
      interval?: number | undefined;
      days_of_week?: string[] | undefined;
      day_of_month?: number | undefined;
      range_type: 'endDate' | 'noEnd' | 'numbered';
      start_date: string;
      end_date?: string | undefined;
      occurrences?: number | undefined;
    } | undefined;
  }): Promise<number> {
    const graphListId = this.idCache.taskLists.get(params.task_list_id);
    if (graphListId == null) throw new Error(`Task list ID ${params.task_list_id} not found in cache. Try searching for or listing the item first to refresh the cache.`);

    const graphTask: Record<string, unknown> = {
      title: params.title,
    };

    if (params.body != null) {
      graphTask.body = {
        contentType: params.body_type ?? 'text',
        content: params.body,
      };
    }

    if (params.due_date != null) {
      graphTask.dueDateTime = {
        dateTime: params.due_date,
        timeZone: 'UTC',
      };
    }

    if (params.importance != null) {
      graphTask.importance = params.importance;
    }

    if (params.reminder_date != null) {
      graphTask.isReminderOn = true;
      graphTask.reminderDateTime = {
        dateTime: params.reminder_date,
        timeZone: 'UTC',
      };
    }

    if (params.recurrence != null) {
      (graphTask as any).recurrence = {
        pattern: {
          type: params.recurrence.pattern,
          interval: params.recurrence.interval ?? 1,
          ...(params.recurrence.days_of_week != null ? { daysOfWeek: params.recurrence.days_of_week } : {}),
          ...(params.recurrence.day_of_month != null ? { dayOfMonth: params.recurrence.day_of_month } : {}),
        },
        range: {
          type: params.recurrence.range_type,
          startDate: params.recurrence.start_date,
          ...(params.recurrence.end_date != null ? { endDate: params.recurrence.end_date } : {}),
          ...(params.recurrence.occurrences != null ? { numberOfOccurrences: params.recurrence.occurrences } : {}),
        },
      };
    }

    const created = await this.client.createTask(graphListId, graphTask);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.tasks.set(numericId, { taskListId: graphListId, taskId: graphId });
    return numericId;
  }

  /**
   * Updates an existing task.
   *
   * Looks up the Graph task info from idCache.tasks, then calls
   * client.updateTask(). Throws if the task is not cached.
   */
  async updateTaskAsync(taskId: number, updates: Record<string, unknown>): Promise<void> {
    const taskInfo = this.idCache.tasks.get(taskId);
    if (taskInfo == null) throw new Error(`Task ID ${taskId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateTask(taskInfo.taskListId, taskInfo.taskId, updates);
  }

  /**
   * Marks a task as completed.
   *
   * Convenience method that calls updateTaskAsync with status: 'completed'
   * and the current time as completedDateTime.
   */
  async completeTaskAsync(taskId: number): Promise<void> {
    await this.updateTaskAsync(taskId, {
      status: 'completed',
      completedDateTime: {
        dateTime: new Date().toISOString(),
        timeZone: 'UTC',
      },
    });
  }

  /**
   * Deletes a task.
   *
   * Looks up the Graph task info from idCache.tasks, calls
   * client.deleteTask(), and removes the entry from idCache.
   * Throws if the task is not cached.
   */
  async deleteTaskAsync(taskId: number): Promise<void> {
    const taskInfo = this.idCache.tasks.get(taskId);
    if (taskInfo == null) throw new Error(`Task ID ${taskId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteTask(taskInfo.taskListId, taskInfo.taskId);
    this.idCache.tasks.delete(taskId);
  }

  /**
   * Creates a new task list.
   *
   * Calls client.createTaskList(), caches the resulting ID in
   * idCache.taskLists, and returns a numeric ID.
   */
  async createTaskListAsync(displayName: string): Promise<number> {
    const created = await this.client.createTaskList(displayName);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.taskLists.set(numericId, graphId);
    return numericId;
  }

  /**
   * Renames a task list.
   */
  async renameTaskListAsync(listId: number, name: string): Promise<void> {
    const graphId = this.idCache.taskLists.get(listId);
    if (graphId == null) throw new Error(`Task list ID ${listId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.updateTaskList(graphId, { displayName: name });
  }

  /**
   * Deletes a task list.
   */
  async deleteTaskListAsync(listId: number): Promise<void> {
    const graphId = this.idCache.taskLists.get(listId);
    if (graphId == null) throw new Error(`Task list ID ${listId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteTaskList(graphId);
    this.idCache.taskLists.delete(listId);
  }

  // ===========================================================================
  // Mail Rules (Async)
  // ===========================================================================

  /**
   * Lists all inbox mail rules.
   */
  async listMailRulesAsync(): Promise<Array<{ id: number; displayName: string; sequence: number; isEnabled: boolean; conditions: unknown; actions: unknown }>> {
    const rules = await this.client.listMailRules();
    return rules.map((rule) => {
      const graphId = rule.id!;
      const numericId = hashStringToNumber(graphId);
      this.idCache.rules.set(numericId, graphId);
      return {
        id: numericId,
        displayName: rule.displayName ?? '',
        sequence: rule.sequence ?? 0,
        isEnabled: rule.isEnabled ?? true,
        conditions: rule.conditions ?? {},
        actions: rule.actions ?? {},
      };
    });
  }

  /**
   * Creates a new inbox mail rule.
   */
  async createMailRuleAsync(rule: Record<string, unknown>): Promise<number> {
    const created = await this.client.createMailRule(rule);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.rules.set(numericId, graphId);
    return numericId;
  }

  /**
   * Deletes an inbox mail rule.
   */
  async deleteMailRuleAsync(ruleId: number): Promise<void> {
    const graphId = this.idCache.rules.get(ruleId);
    if (graphId == null) throw new Error(`Rule ID ${ruleId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteMailRule(graphId);
    this.idCache.rules.delete(ruleId);
  }

  // ===========================================================================
  // Automatic Replies (Out of Office)
  // ===========================================================================

  /**
   * Gets the current automatic replies (OOF) settings.
   */
  async getAutomaticRepliesAsync(): Promise<{
    status: string;
    externalAudience: string;
    internalReplyMessage: string;
    externalReplyMessage: string;
    scheduledStartDateTime: string | null;
    scheduledEndDateTime: string | null;
  }> {
    const settings = await this.client.getAutomaticReplies();
    return {
      status: (settings as any).status ?? 'disabled',
      externalAudience: (settings as any).externalAudience ?? 'none',
      internalReplyMessage: (settings as any).internalReplyMessage ?? '',
      externalReplyMessage: (settings as any).externalReplyMessage ?? '',
      scheduledStartDateTime: (settings as any).scheduledStartDateTime?.dateTime ?? null,
      scheduledEndDateTime: (settings as any).scheduledEndDateTime?.dateTime ?? null,
    };
  }

  /**
   * Sets the automatic replies (OOF) settings.
   */
  async setAutomaticRepliesAsync(params: {
    status: 'disabled' | 'alwaysEnabled' | 'scheduled';
    externalAudience?: 'none' | 'contactsOnly' | 'all';
    internalReplyMessage?: string;
    externalReplyMessage?: string;
    scheduledStartDateTime?: string;
    scheduledEndDateTime?: string;
  }): Promise<void> {
    const settings: Record<string, unknown> = { status: params.status };
    if (params.externalAudience != null) settings['externalAudience'] = params.externalAudience;
    if (params.internalReplyMessage != null) settings['internalReplyMessage'] = params.internalReplyMessage;
    if (params.externalReplyMessage != null) settings['externalReplyMessage'] = params.externalReplyMessage;
    if (params.scheduledStartDateTime != null) settings['scheduledStartDateTime'] = { dateTime: params.scheduledStartDateTime, timeZone: 'UTC' };
    if (params.scheduledEndDateTime != null) settings['scheduledEndDateTime'] = { dateTime: params.scheduledEndDateTime, timeZone: 'UTC' };
    await this.client.setAutomaticReplies(settings);
  }

  // ===========================================================================
  // Mailbox Settings
  // ===========================================================================

  /**
   * Gets the current mailbox settings (language, time zone, formats, working hours).
   */
  async getMailboxSettingsAsync(): Promise<{
    language: string | null;
    timeZone: string | null;
    dateFormat: string | null;
    timeFormat: string | null;
    workingHours: unknown | null;
  }> {
    const settings = await this.client.getMailboxSettings();
    return {
      language: (settings as any).language?.locale ?? null,
      timeZone: (settings as any).timeZone ?? null,
      dateFormat: (settings as any).dateFormat ?? null,
      timeFormat: (settings as any).timeFormat ?? null,
      workingHours: (settings as any).workingHours ?? null,
    };
  }

  /**
   * Updates mailbox settings (language, time zone, date/time formats).
   */
  async updateMailboxSettingsAsync(params: {
    language?: string;
    timeZone?: string;
    dateFormat?: string;
    timeFormat?: string;
  }): Promise<void> {
    const settings: Record<string, unknown> = {};
    if (params.language != null) settings['language'] = { locale: params.language };
    if (params.timeZone != null) settings['timeZone'] = params.timeZone;
    if (params.dateFormat != null) settings['dateFormat'] = params.dateFormat;
    if (params.timeFormat != null) settings['timeFormat'] = params.timeFormat;
    await this.client.updateMailboxSettings(settings);
  }

  // ===========================================================================
  // Master Categories
  // ===========================================================================

  /**
   * Lists all master categories.
   */
  async listCategoriesAsync(): Promise<Array<{ id: number; name: string; color: string }>> {
    const categories = await this.client.listMasterCategories();
    return categories.map((cat) => {
      const graphId = cat.id!;
      const numericId = hashStringToNumber(graphId);
      this.idCache.categories.set(numericId, graphId);
      return {
        id: numericId,
        name: cat.displayName ?? '',
        color: cat.color ?? 'none',
      };
    });
  }

  /**
   * Creates a new master category.
   */
  async createCategoryAsync(name: string, color: string): Promise<number> {
    const created = await this.client.createMasterCategory(name, color);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.categories.set(numericId, graphId);
    return numericId;
  }

  /**
   * Deletes a master category.
   */
  async deleteCategoryAsync(categoryId: number): Promise<void> {
    const graphId = this.idCache.categories.get(categoryId);
    if (graphId == null) throw new Error(`Category ID ${categoryId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteMasterCategory(graphId);
    this.idCache.categories.delete(categoryId);
  }

  // ===========================================================================
  // Focused Inbox Overrides
  // ===========================================================================

  /**
   * Lists all focused inbox overrides.
   */
  async listFocusedOverridesAsync(): Promise<Array<{ id: number; senderAddress: string; classifyAs: string }>> {
    const overrides = await this.client.listFocusedOverrides();
    return overrides.map((o) => {
      const graphId = o.id!;
      const numericId = hashStringToNumber(graphId);
      this.idCache.focusedOverrides.set(numericId, graphId);
      return {
        id: numericId,
        senderAddress: o.senderEmailAddress?.address ?? '',
        classifyAs: o.classifyAs ?? '',
      };
    });
  }

  /**
   * Creates a focused inbox override.
   */
  async createFocusedOverrideAsync(senderAddress: string, classifyAs: 'focused' | 'other'): Promise<number> {
    const created = await this.client.createFocusedOverride(senderAddress, classifyAs);
    const graphId = created.id!;
    const numericId = hashStringToNumber(graphId);
    this.idCache.focusedOverrides.set(numericId, graphId);
    return numericId;
  }

  /**
   * Deletes a focused inbox override.
   */
  async deleteFocusedOverrideAsync(overrideId: number): Promise<void> {
    const graphId = this.idCache.focusedOverrides.get(overrideId);
    if (graphId == null) throw new Error(`Focused override ID ${overrideId} not found in cache. Try searching for or listing the item first to refresh the cache.`);
    await this.client.deleteFocusedOverride(graphId);
    this.idCache.focusedOverrides.delete(overrideId);
  }

  /**
   * Gets the Graph string ID for a folder from the cache.
   */
  getFolderGraphId(folderId: number): string | undefined {
    return this.idCache.folders.get(folderId);
  }
}

/**
 * Creates a Microsoft Graph API repository.
 */
export function createGraphRepository(deviceCodeCallback?: DeviceCodeCallback): GraphRepository {
  return new GraphRepository(deviceCodeCallback);
}
