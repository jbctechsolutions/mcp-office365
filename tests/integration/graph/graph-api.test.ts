/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 *
 * Integration tests for the Microsoft Graph API backend.
 *
 * These tests run against a REAL Microsoft 365 account using cached OAuth tokens.
 * They are READ-ONLY and do not modify any data.
 *
 * Prerequisites:
 *   - GRAPH_INTEGRATION_TEST=1 environment variable
 *   - Valid token cache at ~/.outlook-mcp/tokens.json
 *
 * To run locally:
 *   npx @jbctechsolutions/mcp-office365 auth
 *   GRAPH_INTEGRATION_TEST=1 npx vitest run tests/integration/graph
 */

import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../../../src/index.js';

const SKIP = process.env['GRAPH_INTEGRATION_TEST'] !== '1';

// Helper to call a tool and return the parsed result
async function callTool(
  client: Client,
  name: string,
  args: Record<string, unknown> = {},
): Promise<{ text: string; parsed: unknown; isError: boolean }> {
  const result = await client.callTool({ name, arguments: args });
  const content = result.content as Array<{ type: string; text: string }>;
  const text = content?.[0]?.text ?? '';
  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    // keep as string
  }
  return { text, parsed, isError: result.isError === true };
}

describe.skipIf(SKIP)('Graph API Integration', () => {
  let client: Client;
  let server: ReturnType<typeof createServer>;

  beforeAll(async () => {
    server = createServer();
    client = new Client(
      { name: 'integration-test', version: '1.0.0' },
      { capabilities: {} },
    );

    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([
      client.connect(clientTransport),
      server.connect(serverTransport),
    ]);
  });

  afterAll(async () => {
    await client?.close();
    await server?.close();
  });

  // ── Authentication ──────────────────────────────────────────────────

  describe('authentication', () => {
    it('connects with cached token (silent auth)', async () => {
      // list_folders triggers Graph API auth on first call
      const { parsed, isError } = await callTool(client, 'list_folders');
      expect(isError).toBe(false);
      expect(parsed).toBeDefined();
    });
  });

  // ── Mail ────────────────────────────────────────────────────────────

  describe('mail', () => {
    it('lists mail folders', async () => {
      const { parsed, isError } = await callTool(client, 'list_folders');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
      // Every M365 account has at least Inbox, Sent Items, Drafts
      const folders = parsed as Array<{ name: string }>;
      expect(folders.length).toBeGreaterThanOrEqual(3);
      const names = folders.map((f) => f.name.toLowerCase());
      expect(names).toContain('inbox');
    });

    it('lists emails from inbox', async () => {
      const { parsed, isError } = await callTool(client, 'list_emails', {
        limit: 5,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });

    it('gets unread count', async () => {
      const { text, isError } = await callTool(client, 'get_unread_count');
      expect(isError).toBe(false);
      // Should contain a number
      expect(text).toMatch(/\d/);
    });

    it('searches emails', async () => {
      const { parsed, isError } = await callTool(client, 'search_emails', {
        query: 'test',
        limit: 3,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });

    it('lists drafts', async () => {
      const { parsed, isError } = await callTool(client, 'list_drafts');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Calendar ────────────────────────────────────────────────────────

  describe('calendar', () => {
    it('lists calendars', async () => {
      const { parsed, isError } = await callTool(client, 'list_calendars');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
      // Every account has at least one default calendar
      const calendars = parsed as Array<{ name: string }>;
      expect(calendars.length).toBeGreaterThanOrEqual(1);
    });

    it('lists events', async () => {
      const { parsed, isError } = await callTool(client, 'list_events', {
        limit: 5,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });

    it('searches events', async () => {
      const { parsed, isError } = await callTool(client, 'search_events', {
        query: 'meeting',
        limit: 3,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Contacts ────────────────────────────────────────────────────────

  describe('contacts', () => {
    it('lists contacts', async () => {
      const { parsed, isError } = await callTool(client, 'list_contacts', {
        limit: 5,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });

    it('lists contact folders', async () => {
      const { parsed, isError } = await callTool(client, 'list_contact_folders');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Tasks ───────────────────────────────────────────────────────────

  describe('tasks', () => {
    it('lists task lists', async () => {
      const { parsed, isError } = await callTool(client, 'list_task_lists');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
      // Every account has at least a default task list
      const lists = parsed as Array<{ name: string }>;
      expect(lists.length).toBeGreaterThanOrEqual(1);
    });

    it('lists tasks', async () => {
      const { parsed, isError } = await callTool(client, 'list_tasks', {
        limit: 5,
      });
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Teams ───────────────────────────────────────────────────────────

  describe('teams', () => {
    it('lists teams', async () => {
      const { parsed, isError } = await callTool(client, 'list_teams');
      expect(isError).toBe(false);
      // May return empty array for personal accounts, that's fine
      expect(Array.isArray(parsed)).toBe(true);
    });

    it('lists chats', async () => {
      const { parsed, isError } = await callTool(client, 'list_chats');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── People ──────────────────────────────────────────────────────────

  describe('people', () => {
    it('lists relevant people', async () => {
      const { parsed, isError } = await callTool(client, 'list_relevant_people');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Planner ─────────────────────────────────────────────────────────

  describe('planner', () => {
    it('lists plans', async () => {
      const { parsed, isError } = await callTool(client, 'list_plans');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Mailbox Settings ────────────────────────────────────────────────

  describe('mailbox settings', () => {
    it('gets mailbox settings', async () => {
      const { parsed, isError } = await callTool(client, 'get_mailbox_settings');
      expect(isError).toBe(false);
      expect(parsed).toBeDefined();
    });

    it('gets automatic replies settings', async () => {
      const { parsed, isError } = await callTool(client, 'get_automatic_replies');
      expect(isError).toBe(false);
      expect(parsed).toBeDefined();
    });

    it('gets signature', async () => {
      const { isError } = await callTool(client, 'get_signature');
      expect(isError).toBe(false);
    });
  });

  // ── Categories ──────────────────────────────────────────────────────

  describe('categories', () => {
    it('lists categories', async () => {
      const { parsed, isError } = await callTool(client, 'list_categories');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });

  // ── Mail Rules ──────────────────────────────────────────────────────

  describe('mail rules', () => {
    it('lists mail rules', async () => {
      const { parsed, isError } = await callTool(client, 'list_mail_rules');
      expect(isError).toBe(false);
      expect(Array.isArray(parsed)).toBe(true);
    });
  });
});
