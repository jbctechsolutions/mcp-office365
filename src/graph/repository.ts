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
import type { BatchRequest } from './client/batch.js';
import {
  mapMailFolderToRow,
  mapCalendarToFolderRow,
  mapMessageToEmailRow,
  mapEventToEventRow,
  mapContactToContactRow,
  mapTaskToTaskRow,
} from './mappers/index.js';
import type { DeviceCodeCallback } from './auth/index.js';
import { currentAccountId } from './auth/index.js';
import { resolveId } from '../ids/resolver.js';
import { registerComposite } from '../ids/mint.js';
import { mintSelfEncoded, type EntityType } from '../ids/token.js';
import { IdUnknownError, isGraphSdkError } from '../utils/errors.js';
import type { StateStore } from '../state/store.js';
import type { CompiledSearch } from '../search/compiler.js';
import { downloadAttachment, getDownloadDir } from './attachments.js';
import { buildPlannerTaskMessagePayload } from './planner-task-message-payload.js';
import type { PlanVisualizationData } from '../visualization/types.js';
import * as fs from 'fs';
import * as path from 'path';
import { createHash } from 'node:crypto';

/**
 * Repository implementation using Microsoft Graph API.
 *
 * Provides read-only access to Outlook data via the Graph API.
 */
export class GraphRepository implements IRepository {
  private readonly client: GraphClient;
  private readonly store: StateStore | undefined;
  private readonly accountId: () => string;
  private readonly deltaLinks: Map<string, string> = new Map();

  constructor(
    deviceCodeCallback?: DeviceCodeCallback,
    store?: StateStore,
    accountId: () => string = currentAccountId,
  ) {
    this.client = new GraphClient(deviceCodeCallback);
    this.store = store;
    this.accountId = accountId;
  }

  /**
   * Resolves any id a tool receives — a durable token, a raw Graph id, or a
   * legacy numeric id — to a live Graph id (U5). Self-encoding tokens decode with
   * no store; alias-backed tokens need {@link store}; a numeric id on Graph is
   * unsupported (D4). Throws a typed error on failure.
   */
  private toGraphId(id: string, expectedEntityType?: EntityType): string {
    return resolveId(id, this.accountId(), this.store, expectedEntityType).graphId;
  }

  /**
   * Mints an alias-backed token for a single-Graph-id entity (e.g. team, chat)
   * and records token → graphId in the alias table so it resolves on later calls.
   * Alias-backed (not self-encoding) so the token is short and account-scoped: a
   * cold store yields ID_UNKNOWN rather than leaking a decodable tenant-global id.
   */
  private mintAlias(entityType: EntityType, graphId: string): string {
    if (this.store == null) {
      // Production always has a store (index.ts) and degraded mode still yields a
      // non-null in-memory one, so this is only reachable from a store-less
      // embedding. Fail loudly rather than mint a token nothing can resolve.
      throw new IdUnknownError(
        `${entityType} (durable state store unavailable)`,
        'Alias-backed ids require the durable state store; construct the repository with one.',
      );
    }
    return registerComposite(this.store, {
      entityType,
      parts: { id: graphId },
      graphId,
      accountId: this.accountId(),
    });
  }

  /**
   * Mints an alias-backed token for a composite entity whose Graph URL needs a
   * tuple of ids (e.g. channel {teamId, channelId}). The tuple is the canonical
   * key AND is stored JSON-encoded as the resolved value, so {@link toGraphParts}
   * can recover every field.
   */
  private mintAliasComposite(entityType: EntityType, parts: Readonly<Record<string, string>>): string {
    if (this.store == null) {
      throw new IdUnknownError(`<no store: cannot mint ${entityType}>`);
    }
    return registerComposite(this.store, {
      entityType,
      parts,
      graphId: JSON.stringify(parts),
      accountId: this.accountId(),
    });
  }

  /**
   * Resolves a composite token to its identifying tuple. The alias row stores the
   * tuple JSON-encoded; a raw (non-token) string can't carry a tuple, so anything
   * that doesn't decode to a JSON object carrying every required key is an unusable
   * id (ID_UNKNOWN). Generic over the key set so callers get a precisely-typed
   * result — every requested field is a guaranteed non-empty string.
   */
  private toGraphParts<K extends string>(
    id: string,
    entityType: EntityType,
    keys: readonly K[],
  ): Record<K, string> {
    const raw = this.toGraphId(id, entityType);
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      throw new IdUnknownError(String(id));
    }
    if (parsed == null || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new IdUnknownError(String(id));
    }
    const obj = parsed as Record<string, unknown>;
    const out = {} as Record<K, string>;
    for (const key of keys) {
      const value = obj[key];
      if (typeof value !== 'string' || value.length === 0) {
        throw new IdUnknownError(String(id));
      }
      out[key] = value;
    }
    return out;
  }

  // ===========================================================================
  // Cache Resolvers (auto-fetch parent on cache miss)
  // ===========================================================================

  private async resolveTeamId(teamId: string): Promise<string> {
    // tm_ tokens resolve from the alias store. On a cold miss (never listed this
    // session, or a lost store) re-list — listTeamsAsync deterministically
    // re-mints and re-stores the same token — then resolve again.
    try {
      return this.toGraphId(teamId, 'team');
    } catch (e) {
      if (e instanceof IdUnknownError) {
        await this.listTeamsAsync();
        return this.toGraphId(teamId, 'team');
      }
      throw e;
    }
  }

  private async resolvePlanId(planId: string): Promise<string> {
    // pl_ tokens resolve from the alias store. On a cold miss (never listed this
    // session, or a lost store) re-list — listPlansAsync deterministically
    // re-mints and re-stores the same token — then resolve again.
    try {
      return this.toGraphId(planId, 'plan');
    } catch (e) {
      if (e instanceof IdUnknownError) {
        await this.listPlansAsync();
        return this.toGraphId(planId, 'plan');
      }
      throw e;
    }
  }

  /** True for an HTTP 412 (Precondition Failed) — the `If-Match` etag we sent no
   * longer matches the entity's current etag. */
  private isPreconditionFailed(e: unknown): boolean {
    return isGraphSdkError(e) && e.statusCode === 412;
  }

  /** Extracts the OData etag from a fetched entity, defaulting to `''` when absent. */
  private extractEtag(entity: unknown): string {
    return ((entity as Record<string, unknown>)['@odata.etag'] as string | undefined) ?? '';
  }

  /**
   * Fetches a fresh etag immediately before a write (U5b-5 — the Planner etag is
   * mutable and per-sub-resource, so it can't ride in the durable token or a
   * cache) and retries the write once with a re-fetched etag on a 412, covering
   * the narrow race between the fetch and the write.
   *
   * CONCURRENCY SEMANTIC (deliberate, last-writer-wins): because MCP tool calls
   * are stateless, the etag the caller observed at an earlier `get_*` cannot be
   * carried into a later `update_*`. We therefore re-read the etag at write time,
   * which means a concurrent edit landing between the caller's read and their
   * write is NOT detected — the write overwrites it. This is the intended
   * trade-off (it eliminates the spurious 412s the old cached-etag path produced
   * for a lone editor); Planner writes are not cross-read conflict-protected.
   */
  private async withFreshEtag<T>(
    fetchEtag: () => Promise<string>,
    write: (etag: string) => Promise<T>,
  ): Promise<T> {
    // An empty If-Match is never a valid write intent — fail loudly rather than
    // send `If-Match: ''` (which would 400/412 confusingly, or on a lenient
    // endpoint become an unconditional overwrite).
    const fetch = async (): Promise<string> => {
      const etag = await fetchEtag();
      if (etag.length === 0) {
        throw new Error('Cannot perform a conditional write: the entity returned no @odata.etag.');
      }
      return etag;
    };
    const etag = await fetch();
    try {
      return await write(etag);
    } catch (e) {
      if (this.isPreconditionFailed(e)) {
        return await write(await fetch());
      }
      throw e;
    }
  }

  private async resolveChatId(chatId: string): Promise<string> {
    try {
      return this.toGraphId(chatId, 'chat');
    } catch (e) {
      if (e instanceof IdUnknownError) {
        await this.listChatsAsync();
        return this.toGraphId(chatId, 'chat');
      }
      throw e;
    }
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
    // Rows carry self-encoding fd_ tokens (the mapper mints them) — no cache.
    const folders = await this.client.listMailFolders();
    return folders.map(mapMailFolderToRow);
  }

  getFolder(_id: string): FolderRow | undefined {
    throw new Error('Use getFolderAsync() for Graph repository');
  }

  async getFolderAsync(id: string): Promise<FolderRow | undefined> {
    const graphId = this.toGraphId(id, 'folder');
    const folder = await this.client.getMailFolder(graphId);
    return folder != null ? mapMailFolderToRow(folder) : undefined;
  }

  // ===========================================================================
  // Emails
  // ===========================================================================

  listEmails(_folderId: number, _limit: number, _offset: number): EmailRow[] {
    throw new Error('Use listEmailsAsync() for Graph repository');
  }

  async listEmailsAsync(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    const graphFolderId = this.toGraphId(folderId, 'folder');
    return this.listEmailsWithGraphId(graphFolderId, limit, offset);
  }

  private async listEmailsWithGraphId(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    // Rows carry self-encoding em_ tokens (the mapper mints them) — no cache.
    const messages = await this.client.listMessages(folderId, limit, offset);
    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  listUnreadEmails(_folderId: number, _limit: number, _offset: number): EmailRow[] {
    throw new Error('Use listUnreadEmailsAsync() for Graph repository');
  }

  async listUnreadEmailsAsync(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    const graphFolderId = this.toGraphId(folderId, 'folder');
    return this.listUnreadEmailsWithGraphId(graphFolderId, limit, offset);
  }

  private async listUnreadEmailsWithGraphId(folderId: string, limit: number, offset: number): Promise<EmailRow[]> {
    const messages = await this.client.listUnreadMessages(folderId, limit, offset);
    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  searchEmails(_query: string, _limit: number): EmailRow[] {
    throw new Error('Use searchEmailsAsync() for Graph repository');
  }

  async searchEmailsAsync(query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessages(query, limit);
    return messages.map((m) => mapMessageToEmailRow(m));
  }

  searchEmailsInFolder(_folderId: number, _query: string, _limit: number): EmailRow[] {
    throw new Error('Use searchEmailsInFolderAsync() for Graph repository');
  }

  async searchEmailsInFolderAsync(folderId: string, query: string, limit: number): Promise<EmailRow[]> {
    const graphFolderId = this.toGraphId(folderId, 'folder');
    return this.searchEmailsInFolderWithGraphId(graphFolderId, query, limit);
  }

  private async searchEmailsInFolderWithGraphId(folderId: string, query: string, limit: number): Promise<EmailRow[]> {
    const messages = await this.client.searchMessagesInFolder(folderId, query, limit);
    return messages.map((m) => mapMessageToEmailRow(m, folderId));
  }

  /**
   * Structured advanced search (U7 / D9). Runs a compiled query on the correct
   * Graph mechanism ($filter / quoted $search / /search/query), then caches IDs
   * and maps to EmailRow[] exactly like the raw-KQL path it replaces.
   */
  async searchEmailsStructuredAsync(compiled: CompiledSearch, limit: number): Promise<EmailRow[]> {
    const messages = await this.runStructuredSearch(compiled, limit);
    return messages.map((m) => mapMessageToEmailRow(m));
  }

  private runStructuredSearch(
    compiled: CompiledSearch,
    limit: number,
  ): ReturnType<GraphClient['searchMessagesFilter']> {
    switch (compiled.mechanism) {
      case 'filter':
        return this.client.searchMessagesFilter(compiled.filter, limit);
      case 'search':
        return this.client.searchMessagesSearchValue(compiled.search, limit);
      case 'searchQuery':
        return this.client.searchMessagesQuery(compiled.kql, limit);
    }
  }

  async checkNewEmailsAsync(folderId: string): Promise<{ emails: EmailRow[]; isInitialSync: boolean }> {
    const graphFolderId = this.toGraphId(folderId, 'folder');

    // Key the delta cursor by the RESOLVED Graph id, not the input string: a
    // folder can be addressed by its fd_ token or its raw Graph id, and both
    // must share one cursor — otherwise the second form triggers a fresh initial
    // sync and re-reports already-seen mail as new.
    const existingDeltaLink = this.deltaLinks.get(graphFolderId);
    const isInitialSync = existingDeltaLink == null;

    const { messages, deltaLink } = await this.client.getMessagesDelta(
      graphFolderId,
      existingDeltaLink
    );

    if (deltaLink) {
      this.deltaLinks.set(graphFolderId, deltaLink);
    }

    const activeMessages = messages.filter((m) => (m as unknown as Record<string, unknown>)['@removed'] == null);
    return {
      emails: activeMessages.map((m) => mapMessageToEmailRow(m)),
      isInitialSync,
    };
  }

  getEmail(_id: string): EmailRow | undefined {
    throw new Error('Use getEmailAsync() for Graph repository');
  }

  async getEmailAsync(id: string): Promise<EmailRow | undefined> {
    const graphId = this.toGraphId(id, 'message');
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

  async getUnreadCountByFolderAsync(folderId: string): Promise<number> {
    const graphId = this.toGraphId(folderId, 'folder');
    const folder = await this.client.getMailFolder(graphId);
    return folder?.unreadItemCount ?? 0;
  }

  // ===========================================================================
  // Conversation / Thread
  // ===========================================================================

  /**
   * Lists all messages in a conversation thread.
   *
   * Resolves the message id to its Graph id (durable `em_` token or raw Graph
   * id), fetches the message to read its raw Graph conversationId, then queries
   * for all messages with that ID. No cache required.
   */
  async listConversationAsync(messageId: string, limit: number): Promise<EmailRow[]> {
    const graphId = this.toGraphId(messageId, 'message');
    const message = await this.client.getMessage(graphId);
    if (message == null) throw new Error('Message not found');
    const convId = message.conversationId;
    if (convId == null) throw new Error('Message has no conversation ID');

    const messages = await this.client.listConversationMessages(convId, limit);
    return messages.map((m) => mapMessageToEmailRow(m));
  }

  // ===========================================================================
  // Calendar
  // ===========================================================================

  listCalendars(): FolderRow[] {
    throw new Error('Use listCalendarsAsync() for Graph repository');
  }

  async listCalendarsAsync(): Promise<FolderRow[]> {
    // Rows carry self-encoding fd_ tokens (the mapper mints them) — no cache.
    const calendars = await this.client.listCalendars();
    return calendars.map(mapCalendarToFolderRow);
  }

  listEvents(_limit: number): EventRow[] {
    throw new Error('Use listEventsAsync() for Graph repository');
  }

  async listEventsAsync(limit: number): Promise<EventRow[]> {
    // Rows carry self-encoding ev_ tokens (the mapper mints them) — no cache.
    const events = await this.client.listEvents(limit);
    return events.map((e) => mapEventToEventRow(e));
  }

  listEventsByFolder(_folderId: number, _limit: number): EventRow[] {
    throw new Error('Use listEventsByFolderAsync() for Graph repository');
  }

  async listEventsByFolderAsync(folderId: string, limit: number): Promise<EventRow[]> {
    const graphCalendarId = this.toGraphId(folderId, 'folder');
    const events = await this.client.listEvents(limit, graphCalendarId);
    return events.map((e) => mapEventToEventRow(e, graphCalendarId));
  }

  searchEvents(_query: string | null, _startDate: string | null, _endDate: string | null, _limit: number): EventRow[] {
    throw new Error('Use searchEventsAsync() for Graph repository');
  }

  async searchEventsAsync(query: string | null, startDate: string | null, endDate: string | null, limit: number): Promise<EventRow[]> {
    // Graph doesn't have direct event search, so we filter client-side on the
    // Graph events (by subject) before mapping to rows.
    const start = startDate != null ? new Date(startDate) : undefined;
    const end = endDate != null ? new Date(endDate) : undefined;

    const events = await this.client.listEvents(1000, undefined, start, end);

    const matched = query != null
      ? events.filter((e) => (e.subject?.toLowerCase() ?? '').includes(query.toLowerCase()))
      : events;

    return matched.slice(0, limit).map((e) => mapEventToEventRow(e));
  }

  listEventsByDateRange(_startDate: number, _endDate: number, _limit: number): EventRow[] {
    throw new Error('Use listEventsByDateRangeAsync() for Graph repository');
  }

  async listEventsByDateRangeAsync(startDate: number, endDate: number, limit: number): Promise<EventRow[]> {
    const start = new Date(startDate * 1000);
    const end = new Date(endDate * 1000);

    const events = await this.client.listEvents(limit, undefined, start, end);
    return events.map((e) => mapEventToEventRow(e));
  }

  getEvent(_id: string): EventRow | undefined {
    throw new Error('Use getEventAsync() for Graph repository');
  }

  async getEventAsync(id: string): Promise<EventRow | undefined> {
    const graphId = this.toGraphId(id, 'event');
    const event = await this.client.getEvent(graphId);
    return event != null ? mapEventToEventRow(event) : undefined;
  }

  /** Resolves an event id (durable `ev_` token or raw Graph id) to its Graph id. */
  getEventGraphId(id: string): string {
    return this.toGraphId(id, 'event');
  }

  async listEventInstancesAsync(
    eventId: string,
    startDate: string,
    endDate: string
  ): Promise<EventRow[]> {
    const graphId = this.toGraphId(eventId, 'event');
    const instances = await this.client.listEventInstances(graphId, startDate, endDate);
    return instances.map((e) => mapEventToEventRow(e));
  }

  // ===========================================================================
  // Contacts
  // ===========================================================================

  listContacts(_limit: number, _offset: number): ContactRow[] {
    throw new Error('Use listContactsAsync() for Graph repository');
  }

  async listContactsAsync(limit: number, offset: number): Promise<ContactRow[]> {
    // Rows carry self-encoding `ct_` tokens (the mapper mints them), so there is
    // no numeric-id cache to populate — the token decodes to the Graph id.
    const contacts = await this.client.listContacts(limit, offset);
    return contacts.map(mapContactToContactRow);
  }

  searchContacts(_query: string, _limit: number): ContactRow[] {
    throw new Error('Use searchContactsAsync() for Graph repository');
  }

  async searchContactsAsync(query: string, limit: number): Promise<ContactRow[]> {
    const contacts = await this.client.searchContacts(query, limit);
    return contacts.map(mapContactToContactRow);
  }

  getContact(_id: string): ContactRow | undefined {
    throw new Error('Use getContactAsync() for Graph repository');
  }

  async getContactAsync(id: string): Promise<ContactRow | undefined> {
    const graphId = this.toGraphId(id, 'contact');
    const contact = await this.client.getContact(graphId);
    return contact != null ? mapContactToContactRow(contact) : undefined;
  }

  /** Resolves a contact id (durable `ct_` token or raw Graph id) to its Graph id. */
  getContactGraphId(id: string): string {
    return this.toGraphId(id, 'contact');
  }

  // ===========================================================================
  // Contact Folders
  // ===========================================================================

  async listContactFoldersAsync(): Promise<Array<{ id: string; name: string; parentFolderId: string | null }>> {
    const folders = await this.client.listContactFolders();
    return folders.map((folder) => {
      const graphId = folder.id!;
      return {
        id: this.mintAlias('contactFolder', graphId),
        name: folder.displayName ?? '',
        parentFolderId: folder.parentFolderId ?? null,
      };
    });
  }

  async createContactFolderAsync(name: string): Promise<string> {
    const created = await this.client.createContactFolder(name);
    return this.mintAlias('contactFolder', created.id!);
  }

  async deleteContactFolderAsync(folderId: string): Promise<void> {
    const graphId = this.toGraphId(folderId, 'contactFolder');
    await this.client.deleteContactFolder(graphId);
  }

  async listContactsInFolderAsync(folderId: string, limit: number = 100): Promise<ContactRow[]> {
    const graphId = this.toGraphId(folderId, 'contactFolder');
    const contacts = await this.client.listContactsInFolder(graphId, limit);
    return contacts.map(mapContactToContactRow);
  }

  // ===========================================================================
  // Contact Photos
  // ===========================================================================

  async getContactPhotoAsync(contactId: string): Promise<{ filePath: string; contentType: string }> {
    const graphId = this.toGraphId(contactId, 'contact');

    const photoData = await this.client.getContactPhoto(graphId);
    const downloadDir = getDownloadDir();
    const filePath = path.join(downloadDir, `contact-${createHash('sha1').update(graphId).digest('hex').slice(0, 16)}-photo.jpg`);
    fs.writeFileSync(filePath, Buffer.from(photoData));
    return { filePath, contentType: 'image/jpeg' };
  }

  async setContactPhotoAsync(contactId: string, filePath: string): Promise<void> {
    const graphId = this.toGraphId(contactId, 'contact');

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

    return tasks
      .filter((task) => task.id != null && task.taskListId != null)
      .map((task) => {
        const listTok = this.mintAlias('taskList', task.taskListId);
        const taskTok = this.mintAliasComposite('task', { taskListId: task.taskListId, taskId: task.id! });
        return mapTaskToTaskRow({ ...task }, taskTok, listTok);
      });
  }

  listIncompleteTasks(_limit: number, _offset: number): TaskRow[] {
    throw new Error('Use listIncompleteTasksAsync() for Graph repository');
  }

  async listIncompleteTasksAsync(limit: number, offset: number): Promise<TaskRow[]> {
    const tasks = await this.client.listAllTasks(limit, offset, false);

    return tasks
      .filter((task) => task.id != null && task.taskListId != null)
      .map((task) => {
        const listTok = this.mintAlias('taskList', task.taskListId);
        const taskTok = this.mintAliasComposite('task', { taskListId: task.taskListId, taskId: task.id! });
        return mapTaskToTaskRow({ ...task }, taskTok, listTok);
      });
  }

  searchTasks(_query: string, _limit: number): TaskRow[] {
    throw new Error('Use searchTasksAsync() for Graph repository');
  }

  async searchTasksAsync(query: string, limit: number): Promise<TaskRow[]> {
    const tasks = await this.client.searchTasks(query, limit);

    // Tasks without a taskListId can't form a {taskListId, taskId} composite
    // token — skip them, matching the current cache-miss guard.
    return tasks
      .filter((task) => task.id != null && task.taskListId != null)
      .map((task) => {
        const listTok = this.mintAlias('taskList', task.taskListId);
        const taskTok = this.mintAliasComposite('task', { taskListId: task.taskListId, taskId: task.id! });
        return mapTaskToTaskRow({ ...task }, taskTok, listTok);
      });
  }

  getTask(_id: string): TaskRow | undefined {
    throw new Error('Use getTaskAsync() for Graph repository');
  }

  async getTaskAsync(id: string): Promise<TaskRow | undefined> {
    let taskInfo: { taskListId: string; taskId: string };
    try {
      taskInfo = this.toGraphParts(id, 'task', ['taskListId', 'taskId']);
    } catch {
      return undefined;
    }

    const task = await this.client.getTask(taskInfo.taskListId, taskInfo.taskId);
    if (task == null) {
      return undefined;
    }

    return mapTaskToTaskRow(
      { ...task, taskListId: taskInfo.taskListId },
      String(id),
      this.mintAlias('taskList', taskInfo.taskListId),
    );
  }

  async listTaskListsAsync(): Promise<Array<{ id: string; name: string; isDefault: boolean }>> {
    const lists = await this.client.listTaskLists();
    return lists.map((list) => {
      const graphId = list.id!;
      return {
        id: this.mintAlias('taskList', graphId),
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
   * Resolves a draft id (durable `em_` token or raw Graph id) to its Graph id
   * (satisfies IMailSendRepository). Drafts are messages.
   */
  getGraphIdForDraft(draftId: string): string {
    return this.toGraphId(draftId, 'message');
  }

  /**
   * Gets task info (Graph taskListId/taskId) from a durable `td_` token.
   */
  getTaskInfo(id: string): { taskListId: string; taskId: string } | undefined {
    try {
      return this.toGraphParts(id, 'task', ['taskListId', 'taskId']);
    } catch {
      return undefined;
    }
  }

  /**
   * Gets the Graph string ID for a task list from a durable `tl_` token.
   */
  getTaskListGraphId(id: string): string | undefined {
    try {
      return this.toGraphId(id, 'taskList');
    } catch {
      return undefined;
    }
  }

  // ===========================================================================
  // Write Operations (Async)
  // ===========================================================================

  // Sync versions throw — use async versions from index.ts handler
  moveEmail(_emailId: string, _destinationFolderId: string): void {
    throw new Error('Use moveEmailAsync() for Graph repository');
  }
  deleteEmail(_emailId: string): void {
    throw new Error('Use deleteEmailAsync() for Graph repository');
  }
  archiveEmail(_emailId: string): void {
    throw new Error('Use archiveEmailAsync() for Graph repository');
  }
  junkEmail(_emailId: string): void {
    throw new Error('Use junkEmailAsync() for Graph repository');
  }
  markEmailRead(_emailId: string, _isRead: boolean): void {
    throw new Error('Use markEmailReadAsync() for Graph repository');
  }
  setEmailFlag(_emailId: string, _flagStatus: number): void {
    throw new Error('Use setEmailFlagAsync() for Graph repository');
  }
  setEmailCategories(_emailId: string, _categories: string[]): void {
    throw new Error('Use setEmailCategoriesAsync() for Graph repository');
  }
  setEmailImportance(_emailId: string, _importance: string): void {
    throw new Error('Use setEmailImportanceAsync() for Graph repository');
  }
  createFolder(_name: string, _parentFolderId?: string): FolderRow {
    throw new Error('Use createFolderAsync() for Graph repository');
  }
  deleteFolder(_folderId: string): void {
    throw new Error('Use deleteFolderAsync() for Graph repository');
  }
  renameFolder(_folderId: string, _newName: string): void {
    throw new Error('Use renameFolderAsync() for Graph repository');
  }
  moveFolder(_folderId: string, _destinationParentId: string): void {
    throw new Error('Use moveFolderAsync() for Graph repository');
  }
  emptyFolder(_folderId: string): void {
    throw new Error('Use emptyFolderAsync() for Graph repository');
  }

  // Async implementations

  async moveEmailAsync(emailId: string, destinationFolderId: string): Promise<void> {
    const graphMessageId = this.toGraphId(emailId, 'message');
    const graphFolderId = this.toGraphId(destinationFolderId, 'folder');
    await this.client.moveMessage(graphMessageId, graphFolderId);
  }

  async deleteEmailAsync(emailId: string): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.deleteMessage(graphId);
  }

  async archiveEmailAsync(emailId: string): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.archiveMessage(graphId);
  }

  async junkEmailAsync(emailId: string): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.junkMessage(graphId);
  }

  async markEmailReadAsync(emailId: string, isRead: boolean): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.updateMessage(graphId, { isRead });
  }

  async setEmailFlagAsync(emailId: string, flagStatus: number): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    const flagStatusMap: Record<number, string> = {
      0: 'notFlagged',
      1: 'flagged',
      2: 'complete',
    };
    await this.client.updateMessage(graphId, {
      flag: { flagStatus: flagStatusMap[flagStatus] ?? 'notFlagged' },
    });
  }

  async setEmailCategoriesAsync(emailId: string, categories: string[]): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.updateMessage(graphId, { categories });
  }

  async setEmailImportanceAsync(emailId: string, importance: string): Promise<void> {
    const graphId = this.toGraphId(emailId, 'message');
    await this.client.updateMessage(graphId, { importance });
  }

  async createFolderAsync(name: string, parentFolderId?: string): Promise<FolderRow> {
    const graphParentId = parentFolderId != null
      ? this.toGraphId(parentFolderId, 'folder')
      : undefined;

    const folder = await this.client.createMailFolder(name, graphParentId);

    return mapMailFolderToRow(folder);
  }

  async deleteFolderAsync(folderId: string): Promise<void> {
    const graphId = this.toGraphId(folderId, 'folder');
    await this.client.deleteMailFolder(graphId);
  }

  async renameFolderAsync(folderId: string, newName: string): Promise<void> {
    const graphId = this.toGraphId(folderId, 'folder');
    await this.client.renameMailFolder(graphId, newName);
  }

  async moveFolderAsync(folderId: string, destinationParentId: string): Promise<void> {
    const graphFolderId = this.toGraphId(folderId, 'folder');
    const graphParentId = this.toGraphId(destinationParentId, 'folder');
    await this.client.moveMailFolder(graphFolderId, graphParentId);
  }

  async emptyFolderAsync(folderId: string): Promise<void> {
    const graphId = this.toGraphId(folderId, 'folder');
    await this.client.emptyMailFolder(graphId);
  }

  // ===========================================================================
  // Draft & Send Operations (Async)
  // ===========================================================================

  /**
   * Creates a new draft message.
   *
   * Converts email address strings to Recipient objects, calls the Graph client,
   * and returns a durable `em_` token plus the raw Graph id.
   */
  async createDraftAsync(params: {
    subject: string;
    body: string;
    bodyType: 'text' | 'html';
    to?: string[];
    cc?: string[];
    bcc?: string[];
  }): Promise<{ token: string; graphId: string }> {
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
    return { token: mintSelfEncoded('message', graphId), graphId };
  }

  /**
   * Updates an existing draft message.
   *
   * Resolves the draft id to its Graph id, then calls the client.
   */
  async updateDraftAsync(draftId: string, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.toGraphId(draftId, 'message');
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
   * Resolves the id to its Graph id, then calls the client.
   */
  async sendDraftAsync(draftId: string): Promise<void> {
    const graphId = this.toGraphId(draftId, 'message');
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
   * Resolves the id to its Graph id, then calls the client.
   */
  async replyMessageAsync(messageId: string, comment: string, replyAll: boolean): Promise<void> {
    const graphId = this.toGraphId(messageId, 'message');
    await this.client.replyMessage(graphId, comment, replyAll);
  }

  /**
   * Forwards a message to specified recipients.
   *
   * Resolves the id to its Graph id, converts recipient email strings to
   * Recipient objects, then calls the client.
   */
  async forwardMessageAsync(messageId: string, toRecipients: string[], comment?: string): Promise<void> {
    const graphId = this.toGraphId(messageId, 'message');
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
   * Resolves the source message id to its Graph id, creates the draft via the
   * client, and optionally updates the body.
   *
   * @returns A durable `em_` token and the raw Graph id of the new draft.
   */
  async replyAsDraftAsync(
    messageId: string,
    replyAll = false,
    comment?: string,
    bodyType: string = 'text',
  ): Promise<{ token: string; graphId: string }> {
    const graphMessageId = this.toGraphId(messageId, 'message');

    // Pass comment/body through createReply so the quoted thread is preserved
    const body = comment != null ? { contentType: bodyType, content: comment } : undefined;
    const draft = replyAll
      ? await this.client.createReplyAllDraft(graphMessageId, undefined, body)
      : await this.client.createReplyDraft(graphMessageId, undefined, body);

    const graphId = draft.id!;
    return { token: mintSelfEncoded('message', graphId), graphId };
  }

  /**
   * Creates a forward draft for a message.
   *
   * Resolves the source message id to its Graph id, creates the draft via the
   * client, and optionally updates the recipients and body.
   *
   * @returns A durable `em_` token and the raw Graph id of the new draft.
   */
  async forwardAsDraftAsync(
    messageId: string,
    toRecipients?: string[],
    comment?: string,
    bodyType: string = 'text',
  ): Promise<{ token: string; graphId: string }> {
    const graphMessageId = this.toGraphId(messageId, 'message');

    const draft = await this.client.createForwardDraft(graphMessageId);

    const graphId = draft.id!;

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

    return { token: mintSelfEncoded('message', graphId), graphId };
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
   * Resolves the id to its Graph id, calls client.listAttachments, and mints
   * an `at_` composite alias token ({ messageId, attachmentId }) per item.
   *
   * @returns Array of attachment metadata objects.
   */
  async listAttachmentsAsync(emailId: string): Promise<Array<{
    id: string;
    name: string;
    size: number;
    contentType: string;
    isInline: boolean;
  }>> {
    const graphMessageId = this.toGraphId(emailId, 'message');

    const attachments = await this.client.listAttachments(graphMessageId);

    return attachments.map((att) => {
      const attId = att.id ?? '';
      return {
        id: this.mintAliasComposite('attachment', { messageId: graphMessageId, attachmentId: attId }),
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
   * Resolves the `at_` composite token to { messageId, attachmentId }, then
   * delegates to the downloadAttachment helper which fetches the content and
   * writes it to disk.
   *
   * @returns Metadata about the downloaded file including its local path.
   */
  async downloadAttachmentAsync(
    attachmentId: string,
  ): Promise<{ filePath: string; name: string; size: number; contentType: string }> {
    const cached = this.toGraphParts(attachmentId, 'attachment', ['messageId', 'attachmentId']);

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
    calendarId?: string;
    is_online_meeting?: boolean;
    online_meeting_provider?: string;
  }): Promise<string> {
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

    if (params.is_online_meeting === true) {
      graphEvent.isOnlineMeeting = true;
      graphEvent.onlineMeetingProvider = params.online_meeting_provider ?? 'teamsForBusiness';
    }

    const graphCalendarId = params.calendarId != null
      ? this.toGraphId(params.calendarId, 'folder')
      : undefined;

    const created = await this.client.createEvent(graphEvent, graphCalendarId);
    const graphId = created.id!;
    return mintSelfEncoded('event', graphId);
  }

  /**
   * Updates an existing calendar event.
   *
   * Looks up the Graph string ID from idCache.events, then calls
   * client.updateEvent(). Throws if the event is not cached.
   */
  async updateEventAsync(eventId: string, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.toGraphId(eventId, 'event');
    await this.client.updateEvent(graphId, updates);
  }

  /**
   * Deletes a calendar event.
   *
   * Looks up the Graph string ID from idCache.events, calls
   * client.deleteEvent(), and removes the entry from idCache.
   * Throws if the event is not cached.
   */
  async deleteEventAsync(eventId: string): Promise<void> {
    const graphId = this.toGraphId(eventId, 'event');
    await this.client.deleteEvent(graphId);
  }

  /**
   * Responds to a calendar event invitation.
   *
   * Looks up the Graph string ID from idCache.events, then calls
   * client.respondToEvent(). Throws if the event is not cached.
   */
  async respondToEventAsync(
    eventId: string,
    response: 'accept' | 'decline' | 'tentative',
    sendResponse: boolean,
    comment?: string
  ): Promise<void> {
    const graphId = this.toGraphId(eventId, 'event');
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
  }): Promise<string> {
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
    return mintSelfEncoded('contact', graphId);
  }

  /**
   * Updates an existing contact.
   *
   * Looks up the Graph string ID from idCache.contacts, then calls
   * client.updateContact(). Throws if the contact is not cached.
   */
  async updateContactAsync(contactId: string, updates: Record<string, unknown>): Promise<void> {
    const graphId = this.toGraphId(contactId, 'contact');
    await this.client.updateContact(graphId, updates);
  }

  /**
   * Deletes a contact.
   *
   * Looks up the Graph string ID from idCache.contacts, calls
   * client.deleteContact(), and removes the entry from idCache.
   * Throws if the contact is not cached.
   */
  async deleteContactAsync(contactId: string): Promise<void> {
    const graphId = this.toGraphId(contactId, 'contact');
    await this.client.deleteContact(graphId);
  }

  // ===========================================================================
  // Task Write Operations (Async)
  // ===========================================================================

  /**
   * Creates a new task in a task list.
   *
   * Resolves the `tl_` task list token to a Graph ID, builds a Graph API task
   * object from the given params, calls client.createTask(), and returns a
   * durable `td_` token for the created task.
   */
  async createTaskAsync(params: {
    title: string;
    task_list_id: string;
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
    categories?: string[];
  }): Promise<string> {
    const graphListId = this.toGraphId(params.task_list_id, 'taskList');

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

    if (params.categories != null) {
      graphTask.categories = params.categories;
    }

    if (params.recurrence != null) {
      (graphTask as unknown as Record<string, unknown>).recurrence = {
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
    return this.mintAliasComposite('task', { taskListId: graphListId, taskId: created.id! });
  }

  /**
   * Updates an existing task.
   */
  async updateTaskAsync(taskId: string, updates: Record<string, unknown>): Promise<void> {
    const { taskListId, taskId: gTaskId } = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    await this.client.updateTask(taskListId, gTaskId, updates);
  }

  /**
   * Marks a task as completed.
   *
   * Convenience method that calls updateTaskAsync with status: 'completed'
   * and the current time as completedDateTime.
   */
  async completeTaskAsync(taskId: string): Promise<void> {
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
   */
  async deleteTaskAsync(taskId: string): Promise<void> {
    const { taskListId, taskId: gTaskId } = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    await this.client.deleteTask(taskListId, gTaskId);
  }

  /**
   * Creates a new task list.
   */
  async createTaskListAsync(displayName: string): Promise<string> {
    const created = await this.client.createTaskList(displayName);
    return this.mintAlias('taskList', created.id!);
  }

  /**
   * Renames a task list.
   */
  async renameTaskListAsync(listId: string, name: string): Promise<void> {
    const graphId = this.toGraphId(listId, 'taskList');
    await this.client.updateTaskList(graphId, { displayName: name });
  }

  /**
   * Deletes a task list.
   */
  async deleteTaskListAsync(listId: string): Promise<void> {
    const graphId = this.toGraphId(listId, 'taskList');
    await this.client.deleteTaskList(graphId);
  }

  // ===========================================================================
  // Checklist Items
  // ===========================================================================

  async listChecklistItemsAsync(taskId: string): Promise<Array<{
    id: string; displayName: string; isChecked: boolean; createdDateTime: string;
  }>> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const items = await this.client.listChecklistItems(taskInfo.taskListId, taskInfo.taskId);
    return items.map((item) => {
      const graphId = item.id!;
      return {
        id: this.mintAliasComposite('checklistItem', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, checklistItemId: graphId }),
        displayName: item.displayName ?? '',
        isChecked: item.isChecked ?? false,
        createdDateTime: item.createdDateTime ?? '',
      };
    });
  }

  async createChecklistItemAsync(taskId: string, displayName: string, isChecked: boolean = false): Promise<string> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const item = await this.client.createChecklistItem(taskInfo.taskListId, taskInfo.taskId, displayName, isChecked);
    return this.mintAliasComposite('checklistItem', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, checklistItemId: item.id! });
  }

  async updateChecklistItemAsync(checklistItemId: string, updates: { displayName?: string; isChecked?: boolean }): Promise<void> {
    const cached = this.toGraphParts(checklistItemId, 'checklistItem', ['taskListId', 'taskId', 'checklistItemId']);
    const graphUpdates: Record<string, unknown> = {};
    if (updates.displayName != null) graphUpdates['displayName'] = updates.displayName;
    if (updates.isChecked != null) graphUpdates['isChecked'] = updates.isChecked;
    await this.client.updateChecklistItem(cached.taskListId, cached.taskId, cached.checklistItemId, graphUpdates);
  }

  async deleteChecklistItemAsync(checklistItemId: string): Promise<void> {
    const cached = this.toGraphParts(checklistItemId, 'checklistItem', ['taskListId', 'taskId', 'checklistItemId']);
    await this.client.deleteChecklistItem(cached.taskListId, cached.taskId, cached.checklistItemId);
  }

  // ===========================================================================
  // Linked Resources
  // ===========================================================================

  async listLinkedResourcesAsync(taskId: string): Promise<Array<{
    id: string; webUrl: string; applicationName: string; displayName: string;
  }>> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const items = await this.client.listLinkedResources(taskInfo.taskListId, taskInfo.taskId);
    return items.map((item) => {
      const graphId = item.id!;
      return {
        id: this.mintAliasComposite('linkedResource', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, linkedResourceId: graphId }),
        webUrl: item.webUrl ?? '',
        applicationName: item.applicationName ?? '',
        displayName: item.displayName ?? '',
      };
    });
  }

  async createLinkedResourceAsync(taskId: string, webUrl: string, applicationName: string, displayName?: string): Promise<string> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const item = await this.client.createLinkedResource(taskInfo.taskListId, taskInfo.taskId, webUrl, applicationName, displayName);
    return this.mintAliasComposite('linkedResource', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, linkedResourceId: item.id! });
  }

  async deleteLinkedResourceAsync(linkedResourceId: string): Promise<void> {
    const cached = this.toGraphParts(linkedResourceId, 'linkedResource', ['taskListId', 'taskId', 'linkedResourceId']);
    await this.client.deleteLinkedResource(cached.taskListId, cached.taskId, cached.linkedResourceId);
  }

  // ===========================================================================
  // Task Attachments
  // ===========================================================================

  async listTaskAttachmentsAsync(taskId: string): Promise<Array<{
    id: string; name: string; size: number; contentType: string;
  }>> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const items = await this.client.listTaskAttachments(taskInfo.taskListId, taskInfo.taskId);
    return items.map((item) => {
      const graphId = item.id!;
      return {
        id: this.mintAliasComposite('taskAttachment', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, attachmentId: graphId }),
        name: (item as Record<string, unknown>)['name'] as string ?? '',
        size: item.size ?? 0,
        contentType: item.contentType ?? '',
      };
    });
  }

  async createTaskAttachmentAsync(taskId: string, name: string, contentBytes: string, contentType?: string): Promise<string> {
    const taskInfo = this.toGraphParts(taskId, 'task', ['taskListId', 'taskId']);
    const item = await this.client.createTaskAttachment(taskInfo.taskListId, taskInfo.taskId, name, contentBytes, contentType);
    return this.mintAliasComposite('taskAttachment', { taskListId: taskInfo.taskListId, taskId: taskInfo.taskId, attachmentId: item.id! });
  }

  async deleteTaskAttachmentAsync(taskAttachmentId: string): Promise<void> {
    const cached = this.toGraphParts(taskAttachmentId, 'taskAttachment', ['taskListId', 'taskId', 'attachmentId']);
    await this.client.deleteTaskAttachment(cached.taskListId, cached.taskId, cached.attachmentId);
  }

  // ===========================================================================
  // Mail Rules (Async)
  // ===========================================================================

  /**
   * Lists all inbox mail rules.
   */
  async listMailRulesAsync(): Promise<Array<{ id: string; displayName: string; sequence: number; isEnabled: boolean; conditions: unknown; actions: unknown }>> {
    const rules = await this.client.listMailRules();
    return rules.map((rule) => {
      const graphId = rule.id!;
      return {
        id: this.mintAlias('mailRule', graphId),
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
  async createMailRuleAsync(rule: Record<string, unknown>): Promise<string> {
    const created = await this.client.createMailRule(rule);
    return this.mintAlias('mailRule', created.id!);
  }

  /**
   * Deletes an inbox mail rule.
   */
  async deleteMailRuleAsync(ruleId: string): Promise<void> {
    const graphId = this.toGraphId(ruleId, 'mailRule');
    await this.client.deleteMailRule(graphId);
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
    const startDt = settings.scheduledStartDateTime as Record<string, unknown> | undefined;
    const endDt = settings.scheduledEndDateTime as Record<string, unknown> | undefined;
    return {
      status: (settings.status as string | undefined) ?? 'disabled',
      externalAudience: (settings.externalAudience as string | undefined) ?? 'none',
      internalReplyMessage: (settings.internalReplyMessage as string | undefined) ?? '',
      externalReplyMessage: (settings.externalReplyMessage as string | undefined) ?? '',
      scheduledStartDateTime: (startDt?.dateTime as string | undefined) ?? null,
      scheduledEndDateTime: (endDt?.dateTime as string | undefined) ?? null,
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
    workingHours: unknown;
  }> {
    const settings = await this.client.getMailboxSettings();
    const lang = settings.language as Record<string, unknown> | undefined;
    return {
      language: (lang?.locale as string | undefined) ?? null,
      timeZone: (settings.timeZone as string | undefined) ?? null,
      dateFormat: (settings.dateFormat as string | undefined) ?? null,
      timeFormat: (settings.timeFormat as string | undefined) ?? null,
      workingHours: settings.workingHours ?? null,
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
  async listCategoriesAsync(): Promise<Array<{ id: string; name: string; color: string }>> {
    const categories = await this.client.listMasterCategories();
    return categories.map((cat) => {
      const graphId = cat.id!;
      return {
        id: this.mintAlias('category', graphId),
        name: cat.displayName ?? '',
        color: cat.color ?? 'none',
      };
    });
  }

  /**
   * Creates a new master category.
   */
  async createCategoryAsync(name: string, color: string): Promise<string> {
    const created = await this.client.createMasterCategory(name, color);
    return this.mintAlias('category', created.id!);
  }

  /**
   * Deletes a master category.
   */
  async deleteCategoryAsync(categoryId: string): Promise<void> {
    const graphId = this.toGraphId(categoryId, 'category');
    await this.client.deleteMasterCategory(graphId);
  }

  // ===========================================================================
  // Focused Inbox Overrides
  // ===========================================================================

  /**
   * Lists all focused inbox overrides.
   */
  async listFocusedOverridesAsync(): Promise<Array<{ id: string; senderAddress: string; classifyAs: string }>> {
    const overrides = await this.client.listFocusedOverrides();
    return overrides.map((o) => {
      const graphId = o.id!;
      return {
        id: this.mintAlias('focusedOverride', graphId),
        senderAddress: o.senderEmailAddress?.address ?? '',
        classifyAs: o.classifyAs ?? '',
      };
    });
  }

  /**
   * Creates a focused inbox override.
   */
  async createFocusedOverrideAsync(senderAddress: string, classifyAs: 'focused' | 'other'): Promise<string> {
    const created = await this.client.createFocusedOverride(senderAddress, classifyAs);
    return this.mintAlias('focusedOverride', created.id!);
  }

  /**
   * Deletes a focused inbox override.
   */
  async deleteFocusedOverrideAsync(overrideId: string): Promise<void> {
    const graphId = this.toGraphId(overrideId, 'focusedOverride');
    await this.client.deleteFocusedOverride(graphId);
  }

  // ===========================================================================
  // Message Headers & MIME
  // ===========================================================================

  /**
   * Gets internet message headers for an email.
   */
  async getMessageHeadersAsync(emailId: string): Promise<Array<{ name: string; value: string }>> {
    const graphId = this.toGraphId(emailId, 'message');
    return await this.client.getMessageHeaders(graphId);
  }

  /**
   * Gets the MIME content of a message and saves it as an .eml file.
   */
  async getMessageMimeAsync(emailId: string): Promise<{ filePath: string }> {
    const graphId = this.toGraphId(emailId, 'message');
    const mime = await this.client.getMessageMime(graphId);
    const downloadDir = getDownloadDir();
    const filePath = path.join(downloadDir, `email-${createHash('sha1').update(graphId).digest('hex').slice(0, 16)}.eml`);
    fs.writeFileSync(filePath, mime, 'utf-8');
    return { filePath };
  }

  // ===========================================================================
  // Mail Tips
  // ===========================================================================

  /**
   * Gets mail tips for the specified email addresses.
   */
  async getMailTipsAsync(emailAddresses: string[]): Promise<Array<{
    emailAddress: string; automaticReplies: { message: string } | null;
    mailboxFull: boolean; deliveryRestricted: boolean;
    externalMemberCount: number; maxMessageSize: number;
  }>> {
    const tips = await this.client.getMailTips(emailAddresses);
    return tips.map((tip) => {
      const t = tip;
      const emailAddr = t.emailAddress as Record<string, unknown> | undefined;
      const autoReplies = t.automaticReplies as Record<string, unknown> | undefined;
      return {
        emailAddress: (emailAddr?.address as string | undefined) ?? '',
        automaticReplies: (autoReplies?.message != null && autoReplies.message !== '') ? { message: autoReplies.message as string } : null,
        mailboxFull: (t.mailboxFull as boolean | undefined) ?? false,
        deliveryRestricted: (t.deliveryRestricted as boolean | undefined) ?? false,
        externalMemberCount: (t.externalMemberCount as number | undefined) ?? 0,
        maxMessageSize: (t.maxMessageSize as number | undefined) ?? 0,
      };
    });
  }

  // ===========================================================================
  // Calendar Groups
  // ===========================================================================

  /**
   * Lists all calendar groups.
   *
   * Orphan entity: no Graph URL takes a calendar-group id as a path segment,
   * so this returns the raw Graph id string rather than minting a token.
   */
  async listCalendarGroupsAsync(): Promise<Array<{ id: string; name: string; classId: string }>> {
    const groups = await this.client.listCalendarGroups();
    return groups.map((group) => {
      const graphId = group.id!;
      return {
        id: graphId,
        name: group.name ?? '',
        classId: group.classId?.toString() ?? '',
      };
    });
  }

  /**
   * Creates a new calendar group.
   */
  async createCalendarGroupAsync(name: string): Promise<string> {
    const created = await this.client.createCalendarGroup(name);
    return created.id!;
  }

  // ===========================================================================
  // Calendar Permissions
  // ===========================================================================

  /**
   * Lists all permissions for a calendar.
   */
  async listCalendarPermissionsAsync(calendarId: string): Promise<Array<{ id: string; emailAddress: string; role: string; isRemovable: boolean; isInsideOrganization: boolean }>> {
    const graphCalendarId = this.toGraphId(calendarId, 'folder');

    const permissions = await this.client.listCalendarPermissions(graphCalendarId);
    return permissions.map((perm) => {
      const graphPermId = perm.id!;
      return {
        id: this.mintAliasComposite('calendarPermission', { calendarId: graphCalendarId, permissionId: graphPermId }),
        emailAddress: perm.emailAddress?.address ?? '',
        role: perm.role ?? 'none',
        isRemovable: perm.isRemovable ?? false,
        isInsideOrganization: perm.isInsideOrganization ?? false,
      };
    });
  }

  /**
   * Creates a calendar permission (shares a calendar with someone).
   */
  async createCalendarPermissionAsync(calendarId: string, email: string, role: string): Promise<string> {
    const graphCalendarId = this.toGraphId(calendarId, 'folder');

    const permission = await this.client.createCalendarPermission(graphCalendarId, {
      emailAddress: {
        address: email,
        name: email,
      },
      role,
    });

    return this.mintAliasComposite('calendarPermission', { calendarId: graphCalendarId, permissionId: permission.id! });
  }

  /**
   * Deletes a calendar permission.
   */
  async deleteCalendarPermissionAsync(permissionId: string): Promise<void> {
    const { calendarId, permissionId: permId } = this.toGraphParts(permissionId, 'calendarPermission', ['calendarId', 'permissionId']);
    await this.client.deleteCalendarPermission(calendarId, permId);
  }

  // ===========================================================================
  // Room Lists & Rooms
  // ===========================================================================

  /**
   * Lists all room lists.
   */
  async listRoomListsAsync(): Promise<Array<{ name: string; address: string }>> {
    const lists = await this.client.listRoomLists();
    return lists.map((item) => ({
      name: item.name ?? '',
      address: item.address ?? '',
    }));
  }

  /**
   * Lists rooms, optionally filtered by a room list email.
   */
  async listRoomsAsync(roomListEmail?: string): Promise<Array<{ name: string; address: string }>> {
    const rooms = await this.client.listRooms(roomListEmail);
    return rooms.map((item) => ({
      name: item.name ?? '',
      address: item.address ?? '',
    }));
  }

  // ===========================================================================
  // Teams
  // ===========================================================================

  /**
   * Lists all joined teams with cached numeric IDs.
   */
  async listTeamsAsync(): Promise<Array<{ id: string; name: string; description: string }>> {
    const teams = await this.client.listJoinedTeams();
    return teams.map((team) => {
      const graphId = team.id!;
      return { id: this.mintAlias('team', graphId), name: team.displayName ?? '', description: team.description ?? '' };
    });
  }

  /**
   * Lists all channels in a team with cached numeric IDs.
   */
  async listChannelsAsync(teamId: string): Promise<Array<{ id: string; name: string; description: string; membershipType: string }>> {
    const graphTeamId = await this.resolveTeamId(teamId);
    const channels = await this.client.listChannels(graphTeamId);
    return channels.map((ch) => {
      const graphId = ch.id!;
      const id = this.mintAliasComposite('channel', { teamId: graphTeamId, channelId: graphId });
      return { id, name: ch.displayName ?? '', description: ch.description ?? '', membershipType: ch.membershipType ?? 'standard' };
    });
  }

  /**
   * Gets a specific channel by durable cn_ token.
   */
  async getChannelAsync(channelId: string): Promise<{ id: string; name: string; description: string; membershipType: string; webUrl: string }> {
    const { teamId, channelId: graphChannelId } = this.toGraphParts(channelId, 'channel', ['teamId', 'channelId']);
    const ch = await this.client.getChannel(teamId, graphChannelId);
    return { id: String(channelId), name: ch.displayName ?? '', description: ch.description ?? '', membershipType: ch.membershipType ?? 'standard', webUrl: ch.webUrl ?? '' };
  }

  /**
   * Creates a new channel in a team.
   */
  async createChannelAsync(teamId: string, name: string, description?: string): Promise<string> {
    const graphTeamId = await this.resolveTeamId(teamId);
    const ch = await this.client.createChannel(graphTeamId, name, description);
    return this.mintAliasComposite('channel', { teamId: graphTeamId, channelId: ch.id! });
  }

  /**
   * Updates a channel's properties.
   */
  async updateChannelAsync(channelId: string, updates: { name?: string; description?: string }): Promise<void> {
    const { teamId, channelId: graphChannelId } = this.toGraphParts(channelId, 'channel', ['teamId', 'channelId']);
    const graphUpdates: Record<string, unknown> = {};
    if (updates.name != null) graphUpdates['displayName'] = updates.name;
    if (updates.description != null) graphUpdates['description'] = updates.description;
    await this.client.updateChannel(teamId, graphChannelId, graphUpdates);
  }

  /**
   * Deletes a channel.
   */
  async deleteChannelAsync(channelId: string): Promise<void> {
    const { teamId, channelId: graphChannelId } = this.toGraphParts(channelId, 'channel', ['teamId', 'channelId']);
    await this.client.deleteChannel(teamId, graphChannelId);
  }

  /**
   * Lists members of a team.
   */
  async listTeamMembersAsync(teamId: string): Promise<Array<{ id: string; displayName: string; email: string; roles: string[] }>> {
    const graphTeamId = await this.resolveTeamId(teamId);
    const members = await this.client.listTeamMembers(graphTeamId);
    return members.map((m) => ({
      id: m.id ?? '',
      displayName: m.displayName ?? '',
      email: ((m as unknown as Record<string, unknown>).email as string | undefined) ?? '',
      roles: m.roles ?? [],
    }));
  }

  // ===========================================================================
  // Channel Messages
  // ===========================================================================

  /**
   * Lists recent messages in a channel.
   */
  async listChannelMessagesAsync(channelId: string, limit: number = 25): Promise<Array<{
    id: string; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>> {
    const { teamId, channelId: graphChannelId } = this.toGraphParts(channelId, 'channel', ['teamId', 'channelId']);
    const messages = await this.client.listChannelMessages(teamId, graphChannelId, limit);
    return messages.map((msg) => {
      const graphId = msg.id!;
      return {
        id: this.mintAliasComposite('channelMessage', { teamId, channelId: graphChannelId, messageId: graphId }),
        senderName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? '',
        senderEmail: ((msg.from?.user as unknown as Record<string, unknown> | undefined)?.email as string | undefined) ?? '',
        bodyPreview: msg.body?.content?.substring(0, 200) ?? '',
        bodyContent: msg.body?.content ?? '',
        contentType: msg.body?.contentType ?? 'html',
        createdDateTime: msg.createdDateTime ?? '',
      };
    });
  }

  /**
   * Gets a specific channel message with its replies.
   */
  async getChannelMessageAsync(messageId: string): Promise<{
    id: string; senderName: string; senderEmail: string; bodyContent: string;
    contentType: string; createdDateTime: string;
    replies: Array<{ id: string; senderName: string; senderEmail: string; bodyContent: string; contentType: string; createdDateTime: string }>;
  }> {
    const { teamId, channelId, messageId: graphMessageId } = this.toGraphParts(messageId, 'channelMessage', ['teamId', 'channelId', 'messageId']);
    const [msg, repliesRaw] = await Promise.all([
      this.client.getChannelMessage(teamId, channelId, graphMessageId),
      this.client.listChannelMessageReplies(teamId, channelId, graphMessageId),
    ]);
    const replies = repliesRaw.map((r) => {
      const rGraphId = r.id!;
      // Reply ids are display-only: a reply is a child entity addressed at
      // /messages/{parentId}/replies/{replyId}, but this xm_ token (like the
      // pre-migration numeric id before it) carries only {teamId, channelId,
      // messageId}, so feeding it back to a message op resolves then 404s. No
      // tool consumes a reply id — react/reply take the parent message id.
      // Reply-level operations are tracked in issue #50.
      return {
        id: this.mintAliasComposite('channelMessage', { teamId, channelId, messageId: rGraphId }),
        senderName: r.from?.user?.displayName ?? r.from?.application?.displayName ?? '',
        senderEmail: ((r.from?.user as unknown as Record<string, unknown> | undefined)?.email as string | undefined) ?? '',
        bodyContent: r.body?.content ?? '',
        contentType: r.body?.contentType ?? 'html',
        createdDateTime: r.createdDateTime ?? '',
      };
    });
    return {
      id: String(messageId),
      senderName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? '',
      senderEmail: ((msg.from?.user as unknown as Record<string, unknown> | undefined)?.email as string | undefined) ?? '',
      bodyContent: msg.body?.content ?? '',
      contentType: msg.body?.contentType ?? 'html',
      createdDateTime: msg.createdDateTime ?? '',
      replies,
    };
  }

  /**
   * Sends a new message to a channel.
   */
  async sendChannelMessageAsync(channelId: string, body: string, contentType: string = 'html'): Promise<string> {
    const { teamId, channelId: graphChannelId } = this.toGraphParts(channelId, 'channel', ['teamId', 'channelId']);
    const msg = await this.client.sendChannelMessage(teamId, graphChannelId, body, contentType);
    const graphId = msg.id!;
    return this.mintAliasComposite('channelMessage', { teamId, channelId: graphChannelId, messageId: graphId });
  }

  /**
   * Replies to a channel message.
   */
  async replyToChannelMessageAsync(messageId: string, body: string, contentType: string = 'html'): Promise<string> {
    const { teamId, channelId, messageId: graphMessageId } = this.toGraphParts(messageId, 'channelMessage', ['teamId', 'channelId', 'messageId']);
    const reply = await this.client.replyToChannelMessage(teamId, channelId, graphMessageId, body, contentType);
    return this.mintAliasComposite('channelMessage', { teamId, channelId, messageId: reply.id! });
  }

  // ===========================================================================
  // Chats
  // ===========================================================================

  async listChatsAsync(limit: number = 25): Promise<Array<{
    id: string; topic: string; chatType: string; lastMessagePreview: string; createdDateTime: string;
  }>> {
    const chats = await this.client.listChats(limit);
    return chats.map((chat) => {
      const graphId = chat.id!;
      const chatRecord = chat as unknown as Record<string, unknown>;
      const preview = chatRecord.lastMessagePreview as Record<string, unknown> | undefined;
      const previewBody = preview?.body as Record<string, unknown> | undefined;
      const previewContent = (previewBody?.content as string | undefined)?.substring(0, 200) ?? '';
      return {
        id: this.mintAlias('chat', graphId),
        topic: chat.topic ?? '',
        chatType: chat.chatType ?? 'oneOnOne',
        lastMessagePreview: previewContent,
        createdDateTime: chat.createdDateTime ?? '',
      };
    });
  }

  async getChatAsync(chatId: string): Promise<{
    id: string; topic: string; chatType: string; createdDateTime: string; webUrl: string;
  }> {
    const graphId = await this.resolveChatId(chatId);
    const chat = await this.client.getChat(graphId);
    return {
      id: String(chatId),
      topic: chat.topic ?? '',
      chatType: chat.chatType ?? 'oneOnOne',
      createdDateTime: chat.createdDateTime ?? '',
      webUrl: ((chat as unknown as Record<string, unknown>).webUrl as string | undefined) ?? '',
    };
  }

  async listChatMessagesAsync(chatId: string, limit: number = 25): Promise<Array<{
    id: string; senderName: string; senderEmail: string; bodyPreview: string;
    bodyContent: string; contentType: string; createdDateTime: string;
  }>> {
    const graphChatId = await this.resolveChatId(chatId);
    const messages = await this.client.listChatMessages(graphChatId, limit);
    return messages.map((msg) => {
      const graphId = msg.id!;
      return {
        id: this.mintAliasComposite('chatMessage', { chatId: graphChatId, messageId: graphId }),
        senderName: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? '',
        senderEmail: ((msg.from?.user as unknown as Record<string, unknown> | undefined)?.email as string | undefined) ?? '',
        bodyPreview: msg.body?.content?.substring(0, 200) ?? '',
        bodyContent: msg.body?.content ?? '',
        contentType: msg.body?.contentType ?? 'html',
        createdDateTime: msg.createdDateTime ?? '',
      };
    });
  }

  async sendChatMessageAsync(chatId: string, body: string, contentType: string = 'html'): Promise<string> {
    const graphChatId = await this.resolveChatId(chatId);
    const msg = await this.client.sendChatMessage(graphChatId, body, contentType);
    const graphId = msg.id!;
    return this.mintAliasComposite('chatMessage', { chatId: graphChatId, messageId: graphId });
  }

  // ===========================================================================
  // Message Reactions
  // ===========================================================================

  async listMessageReactionsAsync(messageId: string, messageType: 'channel' | 'chat'): Promise<Array<{
    reactionType: string;
    user: { displayName: string };
    createdDateTime: string;
  }>> {
    const mapReactions = (msg: Record<string, unknown>): Array<{ reactionType: string; user: { displayName: string }; createdDateTime: string }> => {
      const reactions = (msg.reactions ?? []) as Array<Record<string, unknown>>;
      return reactions.map((r) => {
        const userObj = r.user as Record<string, unknown> | undefined;
        const innerUser = userObj?.user as Record<string, unknown> | undefined;
        return {
          reactionType: (r.reactionType as string | undefined) ?? '',
          user: { displayName: (innerUser?.displayName as string | undefined) ?? '' },
          createdDateTime: (r.createdDateTime as string | undefined) ?? '',
        };
      });
    };

    if (messageType === 'channel') {
      const { teamId, channelId, messageId: graphMessageId } = this.toGraphParts(messageId, 'channelMessage', ['teamId', 'channelId', 'messageId']);
      const msg = await this.client.getChannelMessage(teamId, channelId, graphMessageId);
      return mapReactions(msg as unknown as Record<string, unknown>);
    } else {
      const { chatId, messageId: graphMessageId } = this.toGraphParts(messageId, 'chatMessage', ['chatId', 'messageId']);
      const msg = await this.client.getChatMessage(chatId, graphMessageId);
      return mapReactions(msg as unknown as Record<string, unknown>);
    }
  }

  async addMessageReactionAsync(messageId: string, messageType: 'channel' | 'chat', reactionType: string): Promise<void> {
    if (messageType === 'channel') {
      const { teamId, channelId, messageId: graphMessageId } = this.toGraphParts(messageId, 'channelMessage', ['teamId', 'channelId', 'messageId']);
      await this.client.setChannelMessageReaction(teamId, channelId, graphMessageId, reactionType);
    } else {
      const { chatId, messageId: graphMessageId } = this.toGraphParts(messageId, 'chatMessage', ['chatId', 'messageId']);
      await this.client.setChatMessageReaction(chatId, graphMessageId, reactionType);
    }
  }

  async removeMessageReactionAsync(messageId: string, messageType: 'channel' | 'chat', reactionType: string): Promise<void> {
    if (messageType === 'channel') {
      const { teamId, channelId, messageId: graphMessageId } = this.toGraphParts(messageId, 'channelMessage', ['teamId', 'channelId', 'messageId']);
      await this.client.unsetChannelMessageReaction(teamId, channelId, graphMessageId, reactionType);
    } else {
      const { chatId, messageId: graphMessageId } = this.toGraphParts(messageId, 'chatMessage', ['chatId', 'messageId']);
      await this.client.unsetChatMessageReaction(chatId, graphMessageId, reactionType);
    }
  }

  async listChatMembersAsync(chatId: string): Promise<Array<{ displayName: string; email: string; roles: string[] }>> {
    const graphChatId = await this.resolveChatId(chatId);
    const members = await this.client.listChatMembers(graphChatId);
    return members.map((m) => ({
      displayName: m.displayName ?? '',
      email: ((m as unknown as Record<string, unknown>).email as string | undefined) ?? '',
      roles: m.roles ?? [],
    }));
  }

  // ===========================================================================
  // Planner Plans
  // ===========================================================================

  /**
   * Lists all plans the current user has, minting durable pl_ tokens.
   */
  async listPlansAsync(): Promise<Array<{ id: string; title: string; owner: string; createdDateTime: string }>> {
    const plans = await this.client.listPlans();
    return plans.map((plan) => {
      const graphId = plan.id!;
      return {
        id: this.mintAlias('plan', graphId),
        title: plan.title ?? '',
        owner: plan.owner ?? '',
        createdDateTime: plan.createdDateTime ?? '',
      };
    });
  }

  /**
   * Gets a specific plan.
   */
  async getPlanAsync(planId: string): Promise<{ id: string; title: string; owner: string; createdDateTime: string; etag: string }> {
    const graphPlanId = await this.resolvePlanId(planId);
    const plan = await this.client.getPlan(graphPlanId);
    return {
      id: String(planId),
      title: plan.title ?? '',
      owner: plan.owner ?? '',
      createdDateTime: plan.createdDateTime ?? '',
      etag: this.extractEtag(plan),
    };
  }

  /**
   * Creates a new plan.
   */
  async createPlanAsync(title: string, groupId: string): Promise<string> {
    const plan = await this.client.createPlan(title, groupId);
    return this.mintAlias('plan', plan.id!);
  }

  /**
   * Updates a plan (U5b-5: fetches a fresh etag immediately before the write —
   * etags are never cached across calls).
   */
  async updatePlanAsync(planId: string, updates: { title?: string }): Promise<void> {
    const graphPlanId = await this.resolvePlanId(planId);
    const graphUpdates: Record<string, unknown> = {};
    if (updates.title != null) graphUpdates['title'] = updates.title;
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getPlan(graphPlanId)),
      (etag) => this.client.updatePlan(graphPlanId, graphUpdates, etag),
    );
  }

  // ===========================================================================
  // Planner Buckets
  // ===========================================================================

  /**
   * Lists all buckets in a plan, minting durable pb_ tokens.
   */
  async listBucketsAsync(planId: string): Promise<Array<{ id: string; name: string; planId: string; orderHint: string }>> {
    const graphPlanId = await this.resolvePlanId(planId);
    const buckets = await this.client.listBuckets(graphPlanId);
    return buckets.map((bucket) => {
      const graphId = bucket.id!;
      return {
        id: this.mintAlias('plannerBucket', graphId),
        name: bucket.name ?? '',
        planId: String(planId),
        orderHint: bucket.orderHint ?? '',
      };
    });
  }

  /**
   * Creates a new bucket in a plan.
   */
  async createBucketAsync(planId: string, name: string): Promise<string> {
    const graphPlanId = await this.resolvePlanId(planId);
    const bucket = await this.client.createBucket(graphPlanId, name);
    return this.mintAlias('plannerBucket', bucket.id!);
  }

  /**
   * Updates a bucket (U5b-5: fetches a fresh etag immediately before the write).
   */
  async updateBucketAsync(bucketId: string, updates: { name?: string }): Promise<void> {
    const graphBucketId = this.toGraphId(bucketId, 'plannerBucket');
    const graphUpdates: Record<string, unknown> = {};
    if (updates.name != null) graphUpdates['name'] = updates.name;
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getBucket(graphBucketId)),
      (etag) => this.client.updateBucket(graphBucketId, graphUpdates, etag),
    );
  }

  /**
   * Deletes a bucket (U5b-5: fetches a fresh etag immediately before the write).
   */
  async deleteBucketAsync(bucketId: string): Promise<void> {
    const graphBucketId = this.toGraphId(bucketId, 'plannerBucket');
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getBucket(graphBucketId)),
      (etag) => this.client.deleteBucket(graphBucketId, etag),
    );
  }

  // ===========================================================================
  // Planner Tasks
  // ===========================================================================

  /**
   * Lists all tasks in a plan, minting durable pt_ tokens.
   */
  async listPlannerTasksAsync(planId: string): Promise<Array<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string;
  }>> {
    const graphPlanId = await this.resolvePlanId(planId);
    const tasks = await this.client.listPlannerTasks(graphPlanId);
    return tasks.map((task) => {
      const graphId = task.id!;
      return {
        id: this.mintAlias('plannerTask', graphId),
        title: task.title ?? '',
        bucketId: task.bucketId != null ? this.mintAlias('plannerBucket', task.bucketId) : null,
        assignees: task.assignments != null ? Object.keys(task.assignments) : [],
        percentComplete: task.percentComplete ?? 0,
        priority: task.priority ?? 5,
        startDateTime: task.startDateTime ?? '',
        dueDateTime: task.dueDateTime ?? '',
        createdDateTime: task.createdDateTime ?? '',
      };
    });
  }

  /**
   * Lists all Planner tasks assigned to the signed-in user, across every plan
   * (`GET /me/planner/tasks`). Unlike the plan-scoped list, each task carries its
   * own `planId`; the plan/bucket/task durable tokens are minted so follow-up
   * get/update calls resolve without a re-list.
   */
  async listMyPlannerTasksAsync(): Promise<Array<{
    id: string; title: string; planId: string; bucketId: string | null;
    assignees: string[]; percentComplete: number; priority: number;
    startDateTime: string; dueDateTime: string; createdDateTime: string;
  }>> {
    const tasks = await this.client.listMyPlannerTasks();
    return tasks.map((task) => {
      const graphId = task.id!;
      return {
        id: this.mintAlias('plannerTask', graphId),
        title: task.title ?? '',
        planId: task.planId != null ? this.mintAlias('plan', task.planId) : '',
        bucketId: task.bucketId != null ? this.mintAlias('plannerBucket', task.bucketId) : null,
        assignees: task.assignments != null ? Object.keys(task.assignments) : [],
        percentComplete: task.percentComplete ?? 0,
        priority: task.priority ?? 5,
        startDateTime: task.startDateTime ?? '',
        dueDateTime: task.dueDateTime ?? '',
        createdDateTime: task.createdDateTime ?? '',
      };
    });
  }

  /**
   * Gets a specific planner task.
   */
  async getPlannerTaskAsync(taskId: string): Promise<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string; conversationThreadId: string;
    orderHint: string; etag: string;
  }> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const task = await this.client.getPlannerTask(gTaskId);
    return {
      id: String(taskId),
      title: task.title ?? '',
      bucketId: task.bucketId != null ? this.mintAlias('plannerBucket', task.bucketId) : null,
      assignees: task.assignments != null ? Object.keys(task.assignments) : [],
      percentComplete: task.percentComplete ?? 0,
      priority: task.priority ?? 5,
      startDateTime: task.startDateTime ?? '',
      dueDateTime: task.dueDateTime ?? '',
      createdDateTime: task.createdDateTime ?? '',
      conversationThreadId: task.conversationThreadId ?? '',
      orderHint: task.orderHint ?? '',
      etag: this.extractEtag(task),
    };
  }

  /**
   * Creates a new planner task.
   */
  async createPlannerTaskAsync(
    planId: string,
    title: string,
    bucketId?: string,
    assignments?: Record<string, object>,
    priority?: number,
    startDate?: string,
    dueDate?: string,
  ): Promise<string> {
    const graphPlanId = await this.resolvePlanId(planId);
    const body: Record<string, unknown> = { planId: graphPlanId, title };
    if (bucketId != null) {
      body.bucketId = this.toGraphId(bucketId, 'plannerBucket');
    }
    if (assignments != null) body.assignments = assignments;
    if (priority != null) body.priority = priority;
    if (startDate != null) body.startDateTime = startDate;
    if (dueDate != null) body.dueDateTime = dueDate;
    const task = await this.client.createPlannerTask(body);
    return this.mintAlias('plannerTask', task.id!);
  }

  /**
   * Updates a planner task (U5b-5: fetches a fresh etag immediately before the write).
   */
  async updatePlannerTaskAsync(
    taskId: string,
    updates: {
      title?: string;
      bucketId?: string;
      percentComplete?: number;
      priority?: number;
      startDate?: string;
      dueDate?: string;
      assignments?: Record<string, object>;
    },
  ): Promise<void> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const graphUpdates: Record<string, unknown> = {};
    if (updates.title != null) graphUpdates['title'] = updates.title;
    if (updates.bucketId != null) graphUpdates['bucketId'] = this.toGraphId(updates.bucketId, 'plannerBucket');
    if (updates.percentComplete != null) graphUpdates['percentComplete'] = updates.percentComplete;
    if (updates.priority != null) graphUpdates['priority'] = updates.priority;
    if (updates.startDate != null) graphUpdates['startDateTime'] = updates.startDate;
    if (updates.dueDate != null) graphUpdates['dueDateTime'] = updates.dueDate;
    if (updates.assignments != null) graphUpdates['assignments'] = updates.assignments;
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getPlannerTask(gTaskId)),
      (etag) => this.client.updatePlannerTask(gTaskId, graphUpdates, etag),
    );
  }

  /**
   * Deletes a planner task (U5b-5: fetches a fresh etag immediately before the write).
   */
  async deletePlannerTaskAsync(taskId: string): Promise<void> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getPlannerTask(gTaskId)),
      (etag) => this.client.deletePlannerTask(gTaskId, etag),
    );
  }

  // ===========================================================================
  // Planner Task Details
  // ===========================================================================

  /**
   * Gets details for a planner task (description, checklist, references). Task
   * details piggyback the pt_ task token — they have no id of their own.
   */
  async getPlannerTaskDetailsAsync(taskId: string): Promise<{
    id: string; description: string; checklist: Record<string, unknown>;
    references: Record<string, unknown>; etag: string;
  }> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const details = await this.client.getPlannerTaskDetails(gTaskId);
    return {
      id: String(taskId),
      description: details.description ?? '',
      checklist: (details.checklist as Record<string, unknown>) ?? {},
      references: (details.references as Record<string, unknown>) ?? {},
      etag: this.extractEtag(details),
    };
  }

  /**
   * Lists all tasks in a plan with their details (description, checklist, references)
   * fetched in batched requests. This avoids N+1 queries when you need both the task
   * list and each task's details.
   *
   * Details are fetched via the Graph $batch API (up to 20 per batch).
   * Partial failures are handled gracefully: tasks whose detail fetch failed will
   * have `details` set to `undefined`.
   */
  async listPlannerTasksWithDetailsAsync(planId: string): Promise<Array<{
    id: string; title: string; bucketId: string | null; assignees: string[];
    percentComplete: number; priority: number; startDateTime: string;
    dueDateTime: string; createdDateTime: string;
    details?: {
      description: string;
      checklist: Record<string, unknown>;
      references: Record<string, unknown>;
      etag: string;
    } | undefined;
  }>> {
    // Step 1: List all tasks (mints pt_ tokens, registered in the alias store)
    const tasks = await this.listPlannerTasksAsync(planId);

    if (tasks.length === 0) return [];

    // Step 2: Build batch requests for each task's details
    const batchRequests: BatchRequest[] = tasks.map((task) => {
      const graphTaskId = this.toGraphId(task.id, 'plannerTask');
      return {
        id: task.id,
        method: 'GET',
        url: `/planner/tasks/${graphTaskId}/details`,
      };
    });

    // Step 3: Execute batch requests (automatically splits into batches of 20)
    const batchResults = await this.client.batchRequests(batchRequests);

    // Step 4: Merge details into tasks
    return tasks.map((task) => {
      const result = batchResults.get(task.id);
      if (result != null && result.status >= 200 && result.status < 300) {
        const detailBody = result.body as Record<string, unknown>;
        const etag = result.headers?.ETag ?? result.headers?.etag ?? '';
        return {
          ...task,
          details: {
            description: (detailBody.description as string) ?? '',
            checklist: (detailBody.checklist as Record<string, unknown>) ?? {},
            references: (detailBody.references as Record<string, unknown>) ?? {},
            etag,
          },
        };
      }
      // Partial failure: return task without details
      return { ...task, details: undefined };
    });
  }

  /**
   * Updates details for a planner task (U5b-5: fetches a fresh etag immediately
   * before the write — the task details etag is independent of the task's own
   * etag).
   */
  async updatePlannerTaskDetailsAsync(
    taskId: string,
    updates: {
      description?: string;
      checklist?: Record<string, object>;
      references?: Record<string, object>;
    },
  ): Promise<void> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const graphUpdates: Record<string, unknown> = {};
    if (updates.description != null) graphUpdates['description'] = updates.description;
    if (updates.checklist != null) graphUpdates['checklist'] = updates.checklist;
    if (updates.references != null) graphUpdates['references'] = updates.references;
    await this.withFreshEtag(
      async () => this.extractEtag(await this.client.getPlannerTaskDetails(gTaskId)),
      (etag) => this.client.updatePlannerTaskDetails(gTaskId, graphUpdates, etag),
    );
  }

  // ===========================================================================
  // Planner Task Chat Messages (beta)
  // ===========================================================================

  /**
   * Resolves user ids or email addresses to Entra user ids for @mentions.
   */
  private async resolveMentionUserIdsAsync(idsOrEmails: readonly string[]): Promise<string[]> {
    const resolved: string[] = [];
    for (const value of idsOrEmails) {
      if (value.includes('@')) {
        const user = await this.client.getUserProfile(value);
        if (user.id == null || user.id.length === 0) {
          throw new Error(`Could not resolve user for mention: ${value}`);
        }
        resolved.push(user.id);
      } else {
        resolved.push(value);
      }
    }
    return resolved;
  }

  private mapPlannerTaskMessage(
    gTaskId: string,
    msg: Record<string, unknown>,
  ): {
    id: string;
    content: string;
    messageType: string;
    createdDateTime: string;
    editedTime: string | null;
    deletedTime: string | null;
    createdByUserId: string;
    mentions: unknown[];
  } {
    const graphMessageId = String(msg['id'] ?? '');
    const createdBy = msg['createdBy'] as { user?: { id?: string } } | undefined;
    return {
      id: this.mintAliasComposite('plannerTaskMessage', { taskId: gTaskId, messageId: graphMessageId }),
      content: String(msg['content'] ?? ''),
      messageType: String(msg['messageType'] ?? ''),
      createdDateTime: String(msg['createdDateTime'] ?? ''),
      editedTime: (msg['editedTime'] as string | null | undefined) ?? null,
      deletedTime: (msg['deletedTime'] as string | null | undefined) ?? null,
      createdByUserId: createdBy?.user?.id ?? '',
      mentions: (msg['mentions'] as unknown[] | undefined) ?? [],
    };
  }

  /**
   * Lists chat messages (Comments tab) on a Planner task. Beta API; delegated only.
   */
  async listPlannerTaskMessagesAsync(
    taskId: string,
    skipToken?: string,
  ): Promise<{
    messages: Array<{
      id: string;
      content: string;
      messageType: string;
      createdDateTime: string;
      editedTime: string | null;
      deletedTime: string | null;
      createdByUserId: string;
      mentions: unknown[];
    }>;
    nextSkipToken?: string;
  }> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const { messages, nextSkipToken } = await this.client.listPlannerTaskMessages(gTaskId, skipToken);
    return nextSkipToken != null
      ? {
          messages: messages.map((msg) => this.mapPlannerTaskMessage(gTaskId, msg)),
          nextSkipToken,
        }
      : {
          messages: messages.map((msg) => this.mapPlannerTaskMessage(gTaskId, msg)),
        };
  }

  /**
   * Posts a comment on a Planner task. Beta API; delegated only.
   */
  async createPlannerTaskMessageAsync(
    taskId: string,
    content: string,
    mentionUserIds?: string[],
  ): Promise<string> {
    const gTaskId = this.toGraphId(taskId, 'plannerTask');
    const resolvedMentions = mentionUserIds != null && mentionUserIds.length > 0
      ? await this.resolveMentionUserIdsAsync(mentionUserIds)
      : [];
    const payload = buildPlannerTaskMessagePayload(content, resolvedMentions);
    const body: Record<string, unknown> = { content: payload.content };
    if (payload.mentions.length > 0) {
      body['mentions'] = payload.mentions;
    }
    const created = await this.client.createPlannerTaskMessage(gTaskId, body);
    const graphMessageId = String(created['id'] ?? '');
    return this.mintAliasComposite('plannerTaskMessage', { taskId: gTaskId, messageId: graphMessageId });
  }

  /**
   * Updates a Planner task comment. Beta API; delegated only.
   */
  async updatePlannerTaskMessageAsync(
    messageId: string,
    content: string,
    mentionUserIds?: string[],
  ): Promise<void> {
    const { taskId: gTaskId, messageId: gMessageId } = this.toGraphParts(
      messageId,
      'plannerTaskMessage',
      ['taskId', 'messageId'],
    );
    const resolvedMentions = mentionUserIds != null && mentionUserIds.length > 0
      ? await this.resolveMentionUserIdsAsync(mentionUserIds)
      : [];
    const payload = buildPlannerTaskMessagePayload(content, resolvedMentions);
    const body: Record<string, unknown> = { content: payload.content };
    if (payload.mentions.length > 0) {
      body['mentions'] = payload.mentions;
    }
    await this.client.updatePlannerTaskMessage(gTaskId, gMessageId, body);
  }

  /**
   * Deletes a Planner task comment. Beta API; delegated only.
   */
  async deletePlannerTaskMessageAsync(messageId: string): Promise<void> {
    const { taskId: gTaskId, messageId: gMessageId } = this.toGraphParts(
      messageId,
      'plannerTaskMessage',
      ['taskId', 'messageId'],
    );
    await this.client.deletePlannerTaskMessage(gTaskId, gMessageId);
  }

  // ===========================================================================
  // Planner Visualization Data
  // ===========================================================================

  /**
   * Assembles plan, buckets, and tasks into a unified visualization data object.
   */
  async getPlanVisualizationDataAsync(planId: string): Promise<PlanVisualizationData> {
    const plan = await this.getPlanAsync(planId);
    const buckets = await this.listBucketsAsync(planId);
    const tasks = await this.listPlannerTasksAsync(planId);

    return {
      plan: {
        id: plan.id,
        title: plan.title,
      },
      buckets: buckets.map(b => ({
        id: b.id,
        name: b.name,
        orderHint: b.orderHint,
      })),
      tasks: tasks.map(t => ({
        id: t.id,
        title: t.title,
        bucketId: t.bucketId ?? '',
        percentComplete: t.percentComplete,
        priority: t.priority,
        startDateTime: t.startDateTime || null,
        dueDateTime: t.dueDateTime || null,
        assignments: t.assignees,
      })),
    };
  }

  // ===========================================================================
  // Online Meetings
  // ===========================================================================

  async listOnlineMeetingsAsync(limit?: number): Promise<Array<{
    id: string; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
  }>> {
    const meetings = await this.client.listOnlineMeetings(limit ?? 20);
    return meetings.map((meeting) => {
      const graphId = (meeting.id as string | undefined) ?? '';
      return {
        id: this.mintAlias('onlineMeeting', graphId),
        subject: (meeting.subject as string | undefined) ?? '',
        startDateTime: (meeting.startDateTime as string | undefined) ?? '',
        endDateTime: (meeting.endDateTime as string | undefined) ?? '',
        joinUrl: (meeting.joinWebUrl as string | undefined) ?? '',
      };
    });
  }

  async getOnlineMeetingAsync(meetingId: string): Promise<{
    id: string; subject: string; startDateTime: string; endDateTime: string; joinUrl: string;
    participants: unknown;
  } | undefined> {
    const mapMeeting = (m: Record<string, unknown>): { id: string; subject: string; startDateTime: string; endDateTime: string; joinUrl: string; participants: unknown } => ({
      id: String(meetingId),
      subject: (m.subject as string | undefined) ?? '',
      startDateTime: (m.startDateTime as string | undefined) ?? '',
      endDateTime: (m.endDateTime as string | undefined) ?? '',
      joinUrl: (m.joinWebUrl as string | undefined) ?? '',
      participants: m.participants ?? null,
    });

    // om_ tokens resolve from the alias store. On a cold miss (never listed this
    // session, or a lost store) re-list — listOnlineMeetingsAsync deterministically
    // re-mints and re-stores the same token — then retry the resolve, matching
    // the resolveTeamId self-heal pattern. A final miss means "not found" (the
    // documented contract), not an error.
    let graphId: string;
    try {
      graphId = this.toGraphId(meetingId, 'onlineMeeting');
    } catch (e) {
      if (e instanceof IdUnknownError) {
        await this.listOnlineMeetingsAsync();
        try {
          graphId = this.toGraphId(meetingId, 'onlineMeeting');
        } catch (e2) {
          if (e2 instanceof IdUnknownError) return undefined;
          throw e2;
        }
      } else {
        throw e;
      }
    }
    const meeting = await this.client.getOnlineMeeting(graphId);
    return mapMeeting(meeting);
  }

  async listMeetingRecordingsAsync(meetingId: string): Promise<Array<{
    id: string; createdDateTime: string; recordingContentUrl: string;
  }>> {
    const graphMeetingId = this.toGraphId(meetingId, 'onlineMeeting');
    const recordings = await this.client.listMeetingRecordings(graphMeetingId);
    return recordings.map((recording) => {
      const graphId = (recording.id as string | undefined) ?? '';
      return {
        id: this.mintAliasComposite('recording', { meetingId: graphMeetingId, recordingId: graphId }),
        createdDateTime: (recording.createdDateTime as string | undefined) ?? '',
        recordingContentUrl: (recording.recordingContentUrl as string | undefined) ?? '',
      };
    });
  }

  async downloadMeetingRecordingAsync(recordingId: string, outputPath: string): Promise<string> {
    const { meetingId, recordingId: recId } = this.toGraphParts(recordingId, 'recording', ['meetingId', 'recordingId']);
    const content = await this.client.getMeetingRecordingContent(meetingId, recId);
    fs.writeFileSync(outputPath, Buffer.from(content));
    return outputPath;
  }

  async listMeetingTranscriptsAsync(meetingId: string): Promise<Array<{
    id: string; createdDateTime: string; contentUrl: string;
  }>> {
    const graphMeetingId = this.toGraphId(meetingId, 'onlineMeeting');
    const transcripts = await this.client.listMeetingTranscripts(graphMeetingId);
    return transcripts.map((transcript) => {
      const graphId = (transcript.id as string | undefined) ?? '';
      return {
        id: this.mintAliasComposite('transcript', { meetingId: graphMeetingId, transcriptId: graphId }),
        createdDateTime: (transcript.createdDateTime as string | undefined) ?? '',
        contentUrl: (transcript.contentUrl as string | undefined) ?? '',
      };
    });
  }

  async getMeetingTranscriptContentAsync(transcriptId: string, format?: string): Promise<string> {
    const { meetingId, transcriptId: tId } = this.toGraphParts(transcriptId, 'transcript', ['meetingId', 'transcriptId']);
    return await this.client.getMeetingTranscriptContent(meetingId, tId, format ?? 'text/vtt');
  }

  // ===========================================================================
  // Excel Online (Workbook)
  // ===========================================================================

  async listWorksheetsAsync(fileId: string): Promise<Record<string, unknown>[]> {
    const driveItemId = this.toGraphId(fileId, 'driveItem');
    return await this.client.listWorksheets(driveItemId);
  }

  async getWorksheetRangeAsync(fileId: string, worksheetName: string, range: string): Promise<Record<string, unknown>> {
    const driveItemId = this.toGraphId(fileId, 'driveItem');
    return await this.client.getWorksheetRange(driveItemId, worksheetName, range);
  }

  async getUsedRangeAsync(fileId: string, worksheetName: string): Promise<Record<string, unknown>> {
    const driveItemId = this.toGraphId(fileId, 'driveItem');
    return await this.client.getUsedRange(driveItemId, worksheetName);
  }

  async updateWorksheetRangeAsync(fileId: string, worksheetName: string, range: string, values: unknown[][]): Promise<Record<string, unknown>> {
    const driveItemId = this.toGraphId(fileId, 'driveItem');
    return await this.client.updateWorksheetRange(driveItemId, worksheetName, range, values);
  }

  async getTableDataAsync(fileId: string, tableName: string): Promise<Record<string, unknown>[]> {
    const driveItemId = this.toGraphId(fileId, 'driveItem');
    return await this.client.getTableData(driveItemId, tableName);
  }

  // OneDrive
  // ===========================================================================

  /**
   * Lists files/folders in a drive folder (or root).
   */
  async listDriveItemsAsync(folderId?: string): Promise<Array<{
    id: string; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>> {
    const graphFolderId = folderId != null
      ? this.toGraphId(folderId, 'driveItem')
      : undefined;
    const items = await this.client.listDriveItems(graphFolderId);
    // Rows carry self-encoding dr_ tokens (minted here) — no cache.
    return items.map((item) => {
      const itemId = item.id as string;
      return {
        id: itemId.length > 0 ? mintSelfEncoded('driveItem', itemId) : '',
        name: (item.name as string | undefined) ?? '',
        size: (item.size as number | undefined) ?? 0,
        lastModified: (item.lastModifiedDateTime as string | undefined) ?? '',
        isFolder: item.folder != null,
        webUrl: (item.webUrl as string | undefined) ?? '',
      };
    });
  }

  /**
   * Searches drive items by query.
   */
  async searchDriveItemsAsync(query: string, limit?: number): Promise<Array<{
    id: string; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>> {
    const items = await this.client.searchDriveItems(query, limit);
    return items.map((item) => {
      const itemId = item.id as string;
      return {
        id: itemId.length > 0 ? mintSelfEncoded('driveItem', itemId) : '',
        name: (item.name as string | undefined) ?? '',
        size: (item.size as number | undefined) ?? 0,
        lastModified: (item.lastModifiedDateTime as string | undefined) ?? '',
        isFolder: item.folder != null,
        webUrl: (item.webUrl as string | undefined) ?? '',
      };
    });
  }

  /**
   * Gets metadata for a specific drive item.
   */
  async getDriveItemAsync(itemId: string): Promise<{
    id: string; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string; mimeType: string; createdBy: string;
  }> {
    const graphId = this.toGraphId(itemId, 'driveItem');
    const item = await this.client.getDriveItem(graphId);
    const fileObj = item.file as Record<string, unknown> | undefined;
    const createdByObj = item.createdBy as Record<string, unknown> | undefined;
    const createdByUser = createdByObj?.user as Record<string, unknown> | undefined;
    return {
      id: mintSelfEncoded('driveItem', graphId),
      name: (item.name as string | undefined) ?? '',
      size: (item.size as number | undefined) ?? 0,
      lastModified: (item.lastModifiedDateTime as string | undefined) ?? '',
      isFolder: item.folder != null,
      webUrl: (item.webUrl as string | undefined) ?? '',
      mimeType: (fileObj?.mimeType as string | undefined) ?? '',
      createdBy: (createdByUser?.displayName as string | undefined) ?? '',
    };
  }

  /**
   * Downloads a drive item to a local file.
   */
  async downloadFileAsync(itemId: string, outputPath: string): Promise<{ savedPath: string; size: number }> {
    const graphId = this.toGraphId(itemId, 'driveItem');
    const content = await this.client.downloadDriveItem(graphId);
    const buffer = Buffer.from(content);
    fs.writeFileSync(outputPath, buffer);
    return { savedPath: outputPath, size: buffer.length };
  }

  /**
   * Uploads a local file to OneDrive.
   */
  async uploadFileAsync(parentPath: string, fileName: string, localFilePath: string): Promise<string> {
    const content = fs.readFileSync(localFilePath);
    const result = await this.client.uploadDriveItem(parentPath, fileName, content);
    const resultId = result.id as string;
    return mintSelfEncoded('driveItem', resultId);
  }

  /**
   * Lists recently accessed drive items.
   */
  async listRecentFilesAsync(): Promise<Array<{
    id: string; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>> {
    const items = await this.client.listRecentDriveItems();
    return items.map((item) => {
      const itemId = item.id as string;
      return {
        id: itemId.length > 0 ? mintSelfEncoded('driveItem', itemId) : '',
        name: (item.name as string | undefined) ?? '',
        size: (item.size as number | undefined) ?? 0,
        lastModified: (item.lastModifiedDateTime as string | undefined) ?? '',
        isFolder: item.folder != null,
        webUrl: (item.webUrl as string | undefined) ?? '',
      };
    });
  }

  /**
   * Lists drive items shared with the user.
   */
  async listSharedWithMeAsync(): Promise<Array<{
    id: string; name: string; size: number; lastModified: string;
    isFolder: boolean; webUrl: string;
  }>> {
    const items = await this.client.listSharedWithMe();
    return items.map((item) => {
      const itemId = item.id as string;
      return {
        id: itemId.length > 0 ? mintSelfEncoded('driveItem', itemId) : '',
        name: (item.name as string | undefined) ?? '',
        size: (item.size as number | undefined) ?? 0,
        lastModified: (item.lastModifiedDateTime as string | undefined) ?? '',
        isFolder: item.folder != null,
        webUrl: (item.webUrl as string | undefined) ?? '',
      };
    });
  }

  /**
   * Creates a sharing link for a drive item.
   */
  async createSharingLinkAsync(itemId: string, type: string, scope: string): Promise<{
    webUrl: string; type: string; scope: string;
  }> {
    const graphId = this.toGraphId(itemId, 'driveItem');
    const result = await this.client.createSharingLink(graphId, type, scope);
    const link = result.link as Record<string, unknown> | undefined;
    return {
      webUrl: (link?.webUrl as string | undefined) ?? '',
      type: (link?.type as string | undefined) ?? type,
      scope: (link?.scope as string | undefined) ?? scope,
    };
  }

  /**
   * Deletes a drive item.
   */
  async deleteDriveItemAsync(itemId: string): Promise<void> {
    const graphId = this.toGraphId(itemId, 'driveItem');
    await this.client.deleteDriveItem(graphId);
  }

  // ===========================================================================
  // SharePoint Sites & Document Libraries
  // ===========================================================================

  /**
   * Lists followed SharePoint sites, minting durable si_ tokens.
   */
  async listSitesAsync(): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>> {
    const sites = await this.client.listFollowedSites();
    const result: Array<{ id: string; name: string; webUrl: string; displayName: string }> = [];
    for (const site of sites) {
      const siteId = site.id as string | undefined;
      if (siteId != null) {
        result.push({
          id: this.mintAlias('site', siteId),
          name: (site.name as string | undefined) ?? '',
          webUrl: (site.webUrl as string | undefined) ?? '',
          displayName: (site.displayName as string | undefined) ?? '',
        });
      }
    }
    return result;
  }

  /**
   * Searches SharePoint sites by keyword, minting durable si_ tokens.
   */
  async searchSitesAsync(query: string): Promise<Array<{ id: string; name: string; webUrl: string; displayName: string }>> {
    const sites = await this.client.searchSites(query);
    const result: Array<{ id: string; name: string; webUrl: string; displayName: string }> = [];
    for (const site of sites) {
      const siteId = site.id as string | undefined;
      if (siteId != null) {
        result.push({
          id: this.mintAlias('site', siteId),
          name: (site.name as string | undefined) ?? '',
          webUrl: (site.webUrl as string | undefined) ?? '',
          displayName: (site.displayName as string | undefined) ?? '',
        });
      }
    }
    return result;
  }

  /**
   * Gets details for a specific SharePoint site.
   */
  async getSiteAsync(siteId: string): Promise<{ id: string; name: string; webUrl: string; displayName: string; description: string }> {
    const graphId = this.toGraphId(siteId, 'site');
    const site = await this.client.getSite(graphId);
    return {
      id: String(siteId),
      name: (site.name as string | undefined) ?? '',
      webUrl: (site.webUrl as string | undefined) ?? '',
      displayName: (site.displayName as string | undefined) ?? '',
      description: (site.description as string | undefined) ?? '',
    };
  }

  /**
   * Lists document libraries for a SharePoint site, minting durable dl_ tokens.
   */
  async listDocumentLibrariesAsync(siteId: string): Promise<Array<{ id: string; name: string; webUrl: string; driveType: string }>> {
    const graphSiteId = this.toGraphId(siteId, 'site');
    const drives = await this.client.listDocumentLibraries(graphSiteId);
    const result: Array<{ id: string; name: string; webUrl: string; driveType: string }> = [];
    for (const drive of drives) {
      const driveId = drive.id as string | undefined;
      if (driveId != null) {
        result.push({
          id: this.mintAliasComposite('documentLibrary', { siteId: graphSiteId, driveId }),
          name: (drive.name as string | undefined) ?? '',
          webUrl: (drive.webUrl as string | undefined) ?? '',
          driveType: (drive.driveType as string | undefined) ?? '',
        });
      }
    }
    return result;
  }

  /**
   * Lists items in a document library or folder, minting durable li_ tokens.
   */
  async listLibraryItemsAsync(libraryId: string, folderId?: string): Promise<Array<{
    id: string; name: string; size: number; webUrl: string;
    lastModifiedDateTime: string; isFolder: boolean;
  }>> {
    const libCached = this.toGraphParts(libraryId, 'documentLibrary', ['siteId', 'driveId']);
    let folderItemId: string | undefined;
    if (folderId != null) {
      folderItemId = this.toGraphParts(folderId, 'libraryDriveItem', ['driveId', 'itemId']).itemId;
    }
    const items = await this.client.listLibraryItems(libCached.driveId, folderItemId);
    const result: Array<{
      id: string; name: string; size: number; webUrl: string;
      lastModifiedDateTime: string; isFolder: boolean;
    }> = [];
    for (const item of items) {
      const itemGraphId = item.id as string | undefined;
      if (itemGraphId != null) {
        result.push({
          id: this.mintAliasComposite('libraryDriveItem', { driveId: libCached.driveId, itemId: itemGraphId }),
          name: (item.name as string | undefined) ?? '',
          size: (item.size as number | undefined) ?? 0,
          webUrl: (item.webUrl as string | undefined) ?? '',
          lastModifiedDateTime: (item.lastModifiedDateTime as string | undefined) ?? '',
          isFolder: item.folder != null,
        });
      }
    }
    return result;
  }

  /**
   * Downloads a file from a document library to the specified path.
   */
  async downloadLibraryFileAsync(itemId: string, outputPath: string): Promise<string> {
    const cached = this.toGraphParts(itemId, 'libraryDriveItem', ['driveId', 'itemId']);
    const content = await this.client.downloadLibraryFile(cached.driveId, cached.itemId);
    const resolvedPath = path.resolve(outputPath);
    const dir = path.dirname(resolvedPath);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(resolvedPath, Buffer.from(content));
    return resolvedPath;
  }

  /**
   * Creates a folder in a document library (or subfolder), minting a durable li_
   * token for the new folder.
   */
  async createLibraryFolderAsync(libraryId: string, parentFolderId: string | undefined, folderName: string, conflictBehavior: string): Promise<{
    id: string; name: string; webUrl: string; isFolder: boolean;
  }> {
    const { driveId } = this.toGraphParts(libraryId, 'documentLibrary', ['siteId', 'driveId']);
    const parentItemId = parentFolderId != null
      ? this.toGraphParts(parentFolderId, 'libraryDriveItem', ['driveId', 'itemId']).itemId
      : undefined;
    const created = await this.client.createLibraryFolder(driveId, parentItemId, folderName, conflictBehavior);
    const itemGraphId = created.id as string;
    return {
      id: this.mintAliasComposite('libraryDriveItem', { driveId, itemId: itemGraphId }),
      name: (created.name as string | undefined) ?? '',
      webUrl: (created.webUrl as string | undefined) ?? '',
      isFolder: created.folder != null,
    };
  }

  /**
   * Uploads a local file into a document library (or subfolder), minting a durable
   * li_ token for the new item.
   */
  async uploadLibraryFileAsync(libraryId: string, parentFolderId: string | undefined, fileName: string, localFilePath: string, conflictBehavior: string): Promise<{
    id: string; name: string; webUrl: string; size: number;
  }> {
    const { driveId } = this.toGraphParts(libraryId, 'documentLibrary', ['siteId', 'driveId']);
    const parentItemId = parentFolderId != null
      ? this.toGraphParts(parentFolderId, 'libraryDriveItem', ['driveId', 'itemId']).itemId
      : undefined;
    const content = fs.readFileSync(localFilePath);
    const uploaded = await this.client.uploadLibraryFile(driveId, parentItemId, fileName, content, conflictBehavior);
    const itemGraphId = uploaded.id as string;
    return {
      id: this.mintAliasComposite('libraryDriveItem', { driveId, itemId: itemGraphId }),
      name: (uploaded.name as string | undefined) ?? '',
      webUrl: (uploaded.webUrl as string | undefined) ?? '',
      size: (uploaded.size as number | undefined) ?? 0,
    };
  }

  /**
   * Resolves a folder id (durable `fd_` token or raw Graph id) to its Graph id.
   */
  getFolderGraphId(folderId: string): string {
    return this.toGraphId(folderId, 'folder');
  }

  // ===========================================================================
  // SharePoint Lists
  // ===========================================================================

  /**
   * Lists the SharePoint lists in a site, minting durable sl_ tokens. Each list
   * token carries the {siteId, listId} tuple its Graph URL needs.
   */
  async listSharePointListsAsync(siteId: string): Promise<Array<{
    id: string; name: string; displayName: string; description: string; webUrl: string;
  }>> {
    const graphSiteId = this.toGraphId(siteId, 'site');
    const lists = await this.client.listSharePointLists(graphSiteId);
    const result: Array<{ id: string; name: string; displayName: string; description: string; webUrl: string }> = [];
    for (const list of lists) {
      const listId = list.id as string | undefined;
      if (listId != null) {
        result.push({
          id: this.mintAliasComposite('sharePointList', { siteId: graphSiteId, listId }),
          name: (list.name as string | undefined) ?? '',
          displayName: (list.displayName as string | undefined) ?? '',
          description: (list.description as string | undefined) ?? '',
          webUrl: (list.webUrl as string | undefined) ?? '',
        });
      }
    }
    return result;
  }

  /**
   * Gets a specific SharePoint list, resolving the sl_ token to its tuple.
   */
  async getSharePointListAsync(listId: string): Promise<{
    id: string; name: string; displayName: string; description: string; webUrl: string;
  }> {
    const { siteId, listId: graphListId } = this.toGraphParts(listId, 'sharePointList', ['siteId', 'listId']);
    const list = await this.client.getSharePointList(siteId, graphListId);
    return {
      id: String(listId),
      name: (list.name as string | undefined) ?? '',
      displayName: (list.displayName as string | undefined) ?? '',
      description: (list.description as string | undefined) ?? '',
      webUrl: (list.webUrl as string | undefined) ?? '',
    };
  }

  /**
   * Creates a SharePoint list in a site, minting a durable sl_ token.
   */
  async createSharePointListAsync(siteId: string, displayName: string, description?: string): Promise<string> {
    const graphSiteId = this.toGraphId(siteId, 'site');
    const body: Record<string, unknown> = {
      displayName,
      list: { template: 'genericList' },
    };
    if (description != null) {
      body.description = description;
    }
    const created = await this.client.createSharePointList(graphSiteId, body);
    const newListId = created.id as string | undefined;
    if (newListId == null || newListId.length === 0) {
      // Mint guard (matches #46/#47): a composite token minted with an empty id
      // would digest to a resolvable-but-wrong token and be reported as 'created'.
      throw new Error('SharePoint list creation returned no id.');
    }
    return this.mintAliasComposite('sharePointList', { siteId: graphSiteId, listId: newListId });
  }

  /**
   * Lists the column definitions for a SharePoint list. Columns are addressed by
   * name in item field values, so they carry no durable token.
   */
  async listSharePointListColumnsAsync(listId: string): Promise<Array<{
    id: string; name: string; displayName: string; columnType: string; required: boolean; readOnly: boolean;
  }>> {
    const { siteId, listId: graphListId } = this.toGraphParts(listId, 'sharePointList', ['siteId', 'listId']);
    const columns = await this.client.listSharePointListColumns(siteId, graphListId);
    return columns.map((col) => ({
      id: (col.id as string | undefined) ?? '',
      name: (col.name as string | undefined) ?? '',
      displayName: (col.displayName as string | undefined) ?? '',
      columnType: sharePointColumnType(col),
      required: (col.required as boolean | undefined) ?? false,
      readOnly: (col.readOnly as boolean | undefined) ?? false,
    }));
  }

  /**
   * Lists the items in a SharePoint list, minting durable sn_ tokens that carry
   * the {siteId, listId, itemId} tuple.
   */
  async listSharePointListItemsAsync(listId: string, limit: number = 50): Promise<Array<{
    id: string; fields: Record<string, unknown>; webUrl: string;
    createdDateTime: string; lastModifiedDateTime: string;
  }>> {
    const { siteId, listId: graphListId } = this.toGraphParts(listId, 'sharePointList', ['siteId', 'listId']);
    const items = await this.client.listSharePointListItems(siteId, graphListId, limit);
    const result: Array<{
      id: string; fields: Record<string, unknown>; webUrl: string;
      createdDateTime: string; lastModifiedDateTime: string;
    }> = [];
    for (const item of items) {
      const itemGraphId = item.id as string | undefined;
      if (itemGraphId != null) {
        result.push({
          id: this.mintAliasComposite('sharePointListItem', { siteId, listId: graphListId, itemId: itemGraphId }),
          fields: (item.fields as Record<string, unknown> | undefined) ?? {},
          webUrl: (item.webUrl as string | undefined) ?? '',
          createdDateTime: (item.createdDateTime as string | undefined) ?? '',
          lastModifiedDateTime: (item.lastModifiedDateTime as string | undefined) ?? '',
        });
      }
    }
    return result;
  }

  /**
   * Gets a specific SharePoint list item, resolving the sn_ token to its tuple.
   */
  async getSharePointListItemAsync(itemId: string): Promise<{
    id: string; fields: Record<string, unknown>; webUrl: string;
    createdDateTime: string; lastModifiedDateTime: string;
  }> {
    const { siteId, listId, itemId: graphItemId } = this.toGraphParts(itemId, 'sharePointListItem', ['siteId', 'listId', 'itemId']);
    const item = await this.client.getSharePointListItem(siteId, listId, graphItemId);
    return {
      id: String(itemId),
      fields: (item.fields as Record<string, unknown> | undefined) ?? {},
      webUrl: (item.webUrl as string | undefined) ?? '',
      createdDateTime: (item.createdDateTime as string | undefined) ?? '',
      lastModifiedDateTime: (item.lastModifiedDateTime as string | undefined) ?? '',
    };
  }

  /**
   * Creates an item in a SharePoint list, minting a durable sn_ token.
   */
  async createSharePointListItemAsync(listId: string, fields: Record<string, unknown>): Promise<string> {
    const { siteId, listId: graphListId } = this.toGraphParts(listId, 'sharePointList', ['siteId', 'listId']);
    const created = await this.client.createSharePointListItem(siteId, graphListId, fields);
    const newItemId = created.id as string | undefined;
    if (newItemId == null || newItemId.length === 0) {
      // Mint guard (matches #46/#47): see createSharePointListAsync.
      throw new Error('SharePoint list item creation returned no id.');
    }
    return this.mintAliasComposite('sharePointListItem', { siteId, listId: graphListId, itemId: newItemId });
  }

  /**
   * Updates the field values of a SharePoint list item.
   */
  async updateSharePointListItemAsync(itemId: string, fields: Record<string, unknown>): Promise<void> {
    const { siteId, listId, itemId: graphItemId } = this.toGraphParts(itemId, 'sharePointListItem', ['siteId', 'listId', 'itemId']);
    await this.client.updateSharePointListItem(siteId, listId, graphItemId, fields);
  }

  /**
   * Deletes an item from a SharePoint list.
   */
  async deleteSharePointListItemAsync(itemId: string): Promise<void> {
    const { siteId, listId, itemId: graphItemId } = this.toGraphParts(itemId, 'sharePointListItem', ['siteId', 'listId', 'itemId']);
    await this.client.deleteSharePointListItem(siteId, listId, graphItemId);
  }
}

/**
 * Derives a SharePoint column's type from its definition facets. Graph encodes
 * the type as the presence of a facet (`text`, `number`, `boolean`, …) rather
 * than a single field, so return the first facet present or 'unknown'.
 */
function sharePointColumnType(col: Record<string, unknown>): string {
  const facets = [
    'text', 'number', 'boolean', 'dateTime', 'choice', 'currency',
    'personOrGroup', 'lookup', 'hyperlinkOrPicture', 'calculated',
  ];
  for (const facet of facets) {
    if (col[facet] != null) return facet;
  }
  return 'unknown';
}

/**
 * Creates a Microsoft Graph API repository.
 */
export function createGraphRepository(
  deviceCodeCallback?: DeviceCodeCallback,
  store?: StateStore,
  accountId?: () => string,
): GraphRepository {
  return new GraphRepository(deviceCodeCallback, store, accountId);
}
