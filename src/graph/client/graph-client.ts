/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph API client wrapper.
 *
 * Provides a typed interface to the Graph API with:
 * - Automatic token management
 * - Response caching
 * - Pagination support
 * - Error handling
 */

import 'isomorphic-fetch';
import { Client, type PageCollection } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { getAccessToken, type DeviceCodeCallback } from '../auth/index.js';
import { ResponseCache, CacheTTL, createCacheKey } from './cache.js';

/**
 * Graph client wrapper with caching and token management.
 */
export class GraphClient {
  private client: Client | null = null;
  private readonly cache = new ResponseCache();
  private readonly deviceCodeCallback: DeviceCodeCallback | undefined;

  constructor(deviceCodeCallback?: DeviceCodeCallback) {
    this.deviceCodeCallback = deviceCodeCallback;
  }

  /**
   * Gets or creates the Graph client instance.
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  private async getClient(): Promise<Client> {
    if (this.client == null) {
      this.client = Client.init({
        // eslint-disable-next-line @typescript-eslint/no-misused-promises
        authProvider: async (done) => {
          try {
            const token = await getAccessToken(this.deviceCodeCallback);
            done(null, token);
          } catch (error) {
            done(error as Error, null);
          }
        },
      });
    }
    return this.client;
  }

  /**
   * Clears the response cache.
   */
  clearCache(): void {
    this.cache.clear();
  }

  // ===========================================================================
  // Mail Folders
  // ===========================================================================

  /**
   * Lists all mail folders.
   */
  async listMailFolders(): Promise<MicrosoftGraph.MailFolder[]> {
    const cacheKey = createCacheKey('listMailFolders');
    const cached = this.cache.get<MicrosoftGraph.MailFolder[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();
    const result: MicrosoftGraph.MailFolder[] = [];

    // Get top-level folders with pagination
    let response = await client
      .api('/me/mailFolders')
      .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
      .top(100)
      .get() as PageCollection;

    result.push(...(response.value as MicrosoftGraph.MailFolder[]));

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      result.push(...(response.value as MicrosoftGraph.MailFolder[]));
    }

    // Also get child folders (one level deep)
    for (const folder of [...result]) {
      try {
        const children = await client
          .api(`/me/mailFolders/${folder.id}/childFolders`)
          .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
          .get() as PageCollection;

        result.push(...(children.value as MicrosoftGraph.MailFolder[]));
      } catch {
        // Some folders may not have children or may not be accessible
      }
    }

    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Gets a specific mail folder by ID.
   */
  async getMailFolder(folderId: string): Promise<MicrosoftGraph.MailFolder | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/mailFolders/${folderId}`)
        .select('id,displayName,parentFolderId,totalItemCount,unreadItemCount')
        .get() as MicrosoftGraph.MailFolder;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Messages (Emails)
  // ===========================================================================

  /**
   * Lists messages in a folder with pagination.
   */
  async listMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const cacheKey = createCacheKey('listMessages', folderId, limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Message[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Message[];
    this.cache.set(cacheKey, result, CacheTTL.EMAILS);
    return result;
  }

  /**
   * Lists unread messages in a folder.
   */
  async listUnreadMessages(
    folderId: string,
    limit: number = 50,
    skip: number = 0
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .filter('isRead eq false')
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .orderby('receivedDateTime desc')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages across all folders.
   */
  async searchMessages(query: string, limit: number = 50): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/messages')
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId,parentFolderId')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Searches messages in a specific folder.
   */
  async searchMessagesInFolder(
    folderId: string,
    query: string,
    limit: number = 50
  ): Promise<MicrosoftGraph.Message[]> {
    const client = await this.getClient();

    const response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .search(`"${query}"`)
      .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,bodyPreview,conversationId,internetMessageId')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Message[];
  }

  /**
   * Gets a specific message with full body.
   */
  async getMessage(messageId: string): Promise<MicrosoftGraph.Message | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/messages/${messageId}`)
        .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,flag,body,bodyPreview,conversationId,internetMessageId,parentFolderId')
        .get() as MicrosoftGraph.Message;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Calendars
  // ===========================================================================

  /**
   * Lists all calendars.
   */
  async listCalendars(): Promise<MicrosoftGraph.Calendar[]> {
    const cacheKey = createCacheKey('listCalendars');
    const cached = this.cache.get<MicrosoftGraph.Calendar[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/calendars')
      .select('id,name,color,isDefaultCalendar,canEdit')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Calendar[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  // ===========================================================================
  // Events
  // ===========================================================================

  /**
   * Lists events with optional date range.
   */
  async listEvents(
    limit: number = 50,
    calendarId?: string,
    startDate?: Date,
    endDate?: Date
  ): Promise<MicrosoftGraph.Event[]> {
    const client = await this.getClient();

    // If date range provided, use calendarView
    if (startDate != null && endDate != null) {
      const baseUrl = calendarId != null
        ? `/me/calendars/${calendarId}/calendarView`
        : '/me/calendarView';

      const response = await client
        .api(baseUrl)
        .query({
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
        })
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .orderby('start/dateTime')
        .top(limit)
        .get() as PageCollection;

      return response.value as MicrosoftGraph.Event[];
    }

    // Otherwise, get upcoming events
    const baseUrl = calendarId != null
      ? `/me/calendars/${calendarId}/events`
      : '/me/events';

    const response = await client
      .api(baseUrl)
      .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
      .orderby('start/dateTime')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Event[];
  }

  /**
   * Gets a specific event.
   */
  async getEvent(eventId: string): Promise<MicrosoftGraph.Event | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/events/${eventId}`)
        .select('id,subject,start,end,location,isAllDay,organizer,attendees,body,recurrence,iCalUId')
        .get() as MicrosoftGraph.Event;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  /**
   * Lists contacts with pagination.
   */
  async listContacts(limit: number = 50, skip: number = 0): Promise<MicrosoftGraph.Contact[]> {
    const cacheKey = createCacheKey('listContacts', limit, skip);
    const cached = this.cache.get<MicrosoftGraph.Contact[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .orderby('displayName')
      .top(limit)
      .skip(skip)
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.Contact[];
    this.cache.set(cacheKey, result, CacheTTL.CONTACTS);
    return result;
  }

  /**
   * Searches contacts by display name.
   */
  async searchContacts(query: string, limit: number = 50): Promise<MicrosoftGraph.Contact[]> {
    const client = await this.getClient();

    const response = await client
      .api('/me/contacts')
      .filter(`contains(displayName,'${query}')`)
      .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
      .top(limit)
      .get() as PageCollection;

    return response.value as MicrosoftGraph.Contact[];
  }

  /**
   * Gets a specific contact.
   */
  async getContact(contactId: string): Promise<MicrosoftGraph.Contact | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/contacts/${contactId}`)
        .select('id,displayName,givenName,surname,middleName,nickName,companyName,jobTitle,department,emailAddresses,homePhones,businessPhones,mobilePhone,homeAddress,businessAddress,personalNotes')
        .get() as MicrosoftGraph.Contact;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Tasks (Microsoft To Do)
  // ===========================================================================

  /**
   * Lists task lists.
   */
  async listTaskLists(): Promise<MicrosoftGraph.TodoTaskList[]> {
    const cacheKey = createCacheKey('listTaskLists');
    const cached = this.cache.get<MicrosoftGraph.TodoTaskList[]>(cacheKey);
    if (cached != null) {
      return cached;
    }

    const client = await this.getClient();

    const response = await client
      .api('/me/todo/lists')
      .select('id,displayName,isOwner,isShared,wellknownListName')
      .get() as PageCollection;

    const result = response.value as MicrosoftGraph.TodoTaskList[];
    this.cache.set(cacheKey, result, CacheTTL.FOLDERS);
    return result;
  }

  /**
   * Lists tasks in a task list.
   */
  async listTasks(
    taskListId: string,
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<MicrosoftGraph.TodoTask[]> {
    const client = await this.getClient();

    let api = client
      .api(`/me/todo/lists/${taskListId}/tasks`)
      .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
      .top(limit)
      .skip(skip);

    if (!includeCompleted) {
      api = api.filter("status ne 'completed'");
    }

    const response = await api.get() as PageCollection;
    return response.value as MicrosoftGraph.TodoTask[];
  }

  /**
   * Lists all tasks across all task lists.
   */
  async listAllTasks(
    limit: number = 50,
    skip: number = 0,
    includeCompleted: boolean = true
  ): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const taskLists = await this.listTaskLists();
    const allTasks: Array<MicrosoftGraph.TodoTask & { taskListId: string }> = [];

    for (const list of taskLists) {
      if (list.id == null) continue;

      const tasks = await this.listTasks(list.id, 100, 0, includeCompleted);

      for (const task of tasks) {
        allTasks.push({ ...task, taskListId: list.id });
      }
    }

    // Sort by due date, then slice for pagination
    allTasks.sort((a, b) => {
      if (a.dueDateTime == null && b.dueDateTime == null) return 0;
      if (a.dueDateTime == null) return 1;
      if (b.dueDateTime == null) return -1;
      return new Date(a.dueDateTime.dateTime ?? '').getTime() - new Date(b.dueDateTime.dateTime ?? '').getTime();
    });

    return allTasks.slice(skip, skip + limit);
  }

  /**
   * Gets a specific task.
   */
  async getTask(taskListId: string, taskId: string): Promise<MicrosoftGraph.TodoTask | null> {
    const client = await this.getClient();

    try {
      return await client
        .api(`/me/todo/lists/${taskListId}/tasks/${taskId}`)
        .select('id,title,status,importance,dueDateTime,completedDateTime,body,createdDateTime,lastModifiedDateTime,isReminderOn,reminderDateTime')
        .get() as MicrosoftGraph.TodoTask;
    } catch {
      return null;
    }
  }

  // ===========================================================================
  // Write Operations
  // ===========================================================================

  /**
   * Moves a message to a different folder.
   */
  async moveMessage(messageId: string, destinationFolderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: destinationFolderId });
    this.cache.clear(); // Invalidate cache after mutation
  }

  /**
   * Deletes a message (moves to Deleted Items).
   */
  async deleteMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    // Move to deletedItems well-known folder
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'deleteditems' });
    this.cache.clear();
  }

  /**
   * Archives a message (moves to Archive folder).
   */
  async archiveMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'archive' });
    this.cache.clear();
  }

  /**
   * Moves a message to the Junk folder.
   */
  async junkMessage(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/move`)
      .post({ destinationId: 'junkemail' });
    this.cache.clear();
  }

  /**
   * Updates message properties (read status, flag, categories).
   */
  async updateMessage(messageId: string, updates: Record<string, unknown>): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}`)
      .patch(updates);
    this.cache.clear();
  }

  /**
   * Creates a new mail folder.
   */
  async createMailFolder(
    displayName: string,
    parentFolderId?: string
  ): Promise<MicrosoftGraph.MailFolder> {
    const client = await this.getClient();
    const url = parentFolderId != null
      ? `/me/mailFolders/${parentFolderId}/childFolders`
      : '/me/mailFolders';

    const result = await client
      .api(url)
      .post({ displayName }) as MicrosoftGraph.MailFolder;
    this.cache.clear();
    return result;
  }

  /**
   * Deletes a mail folder.
   */
  async deleteMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .delete();
    this.cache.clear();
  }

  /**
   * Renames a mail folder.
   */
  async renameMailFolder(folderId: string, newName: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}`)
      .patch({ displayName: newName });
    this.cache.clear();
  }

  /**
   * Moves a mail folder to a new parent.
   */
  async moveMailFolder(folderId: string, destinationParentId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/mailFolders/${folderId}/move`)
      .post({ destinationId: destinationParentId });
    this.cache.clear();
  }

  /**
   * Deletes all messages in a folder.
   */
  async emptyMailFolder(folderId: string): Promise<void> {
    const client = await this.getClient();
    // Get all messages in the folder
    let response = await client
      .api(`/me/mailFolders/${folderId}/messages`)
      .select('id')
      .top(100)
      .get() as PageCollection;

    // Delete each message
    for (const message of response.value as MicrosoftGraph.Message[]) {
      if (message.id != null) {
        await client
          .api(`/me/messages/${message.id}/move`)
          .post({ destinationId: 'deleteditems' });
      }
    }

    // Handle pagination
    while (response['@odata.nextLink'] != null) {
      response = await client.api(response['@odata.nextLink']).get() as PageCollection;
      for (const message of response.value as MicrosoftGraph.Message[]) {
        if (message.id != null) {
          await client
            .api(`/me/messages/${message.id}/move`)
            .post({ destinationId: 'deleteditems' });
        }
      }
    }

    this.cache.clear();
  }

  // ===========================================================================
  // Draft & Send Operations
  // ===========================================================================

  /**
   * Creates a new draft message.
   */
  async createDraft(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients?: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
    isDraft?: boolean;
  }): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api('/me/messages')
      .post(message) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Updates an existing draft message.
   */
  async updateDraft(messageId: string, updates: Record<string, unknown>): Promise<MicrosoftGraph.Message> {
    const client = await this.getClient();
    const result = await client
      .api(`/me/messages/${messageId}`)
      .patch(updates) as MicrosoftGraph.Message;
    this.cache.clear();
    return result;
  }

  /**
   * Sends an existing draft message.
   */
  async sendDraft(messageId: string): Promise<void> {
    const client = await this.getClient();
    await client
      .api(`/me/messages/${messageId}/send`)
      .post(null);
    this.cache.clear();
  }

  /**
   * Sends a new email directly without creating a draft.
   */
  async sendMail(message: {
    subject: string;
    body: MicrosoftGraph.ItemBody;
    toRecipients: MicrosoftGraph.Recipient[];
    ccRecipients?: MicrosoftGraph.Recipient[];
    bccRecipients?: MicrosoftGraph.Recipient[];
  }): Promise<void> {
    const client = await this.getClient();
    await client
      .api('/me/sendMail')
      .post({ message });
    this.cache.clear();
  }

  /**
   * Replies to a message, or replies to all recipients.
   */
  async replyMessage(messageId: string, comment: string, replyAll: boolean): Promise<void> {
    const client = await this.getClient();
    const action = replyAll ? 'replyAll' : 'reply';
    await client
      .api(`/me/messages/${messageId}/${action}`)
      .post({ comment });
    this.cache.clear();
  }

  /**
   * Forwards a message to specified recipients.
   */
  async forwardMessage(
    messageId: string,
    toRecipients: MicrosoftGraph.Recipient[],
    comment?: string
  ): Promise<void> {
    const client = await this.getClient();
    const body: { toRecipients: MicrosoftGraph.Recipient[]; comment?: string } = { toRecipients };
    if (comment != null) {
      body.comment = comment;
    }
    await client
      .api(`/me/messages/${messageId}/forward`)
      .post(body);
    this.cache.clear();
  }

  // ===========================================================================
  // Tasks (Microsoft To Do) - continued
  // ===========================================================================

  /**
   * Searches tasks by title.
   */
  async searchTasks(query: string, limit: number = 50): Promise<Array<MicrosoftGraph.TodoTask & { taskListId: string }>> {
    const allTasks = await this.listAllTasks(1000, 0, true);

    const queryLower = query.toLowerCase();
    const matched = allTasks.filter(
      (task) => task.title?.toLowerCase().includes(queryLower) ?? false
    );

    return matched.slice(0, limit);
  }
}

/**
 * Creates a new Graph client instance.
 */
export function createGraphClient(deviceCodeCallback?: DeviceCodeCallback): GraphClient {
  return new GraphClient(deviceCodeCallback);
}
