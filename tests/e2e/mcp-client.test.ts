/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * End-to-end tests for the MCP server.
 *
 * These tests verify the server can be started and communicate via MCP protocol.
 * Note: Full e2e tests would require an actual Outlook database which is not
 * available in CI environments. These tests focus on protocol-level verification.
 */

import { describe, it, expect } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../../src/index.js';

describe('MCP Client E2E', () => {
  describe('protocol communication', () => {
    it('can list tools via MCP protocol', async () => {
      // Create server and client with in-memory transport
      const server = createServer();
      const client = new Client(
        {
          name: 'test-client',
          version: '1.0.0',
        },
        {
          capabilities: {},
        }
      );

      // Create linked transports
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();

      // Connect both ends
      await Promise.all([
        client.connect(clientTransport),
        server.connect(serverTransport),
      ]);

      // List tools
      const result = await client.listTools();

      // Verify tools were returned (229 tools — Graph API is the only backend)
      expect(result.tools).toBeDefined();
      expect(Array.isArray(result.tools)).toBe(true);
      const count = result.tools.length;
      expect(count).toBe(229);

      // Verify core tools exist
      const toolNames = result.tools.map((t) => t.name);
      expect(toolNames).toContain('list_folders');
      expect(toolNames).toContain('list_emails');
      expect(toolNames).toContain('search_emails');
      expect(toolNames).toContain('get_email');
      expect(toolNames).toContain('get_unread_count');
      expect(toolNames).toContain('list_attachments');
      expect(toolNames).toContain('download_attachment');
      expect(toolNames).toContain('list_calendars');
      expect(toolNames).toContain('list_events');
      expect(toolNames).toContain('get_event');
      expect(toolNames).toContain('search_events');
      expect(toolNames).toContain('create_event');
      expect(toolNames).toContain('respond_to_event');
      expect(toolNames).toContain('delete_event');
      expect(toolNames).toContain('update_event');
      expect(toolNames).toContain('list_contacts');
      expect(toolNames).toContain('search_contacts');
      expect(toolNames).toContain('get_contact');
      expect(toolNames).toContain('list_tasks');
      expect(toolNames).toContain('search_tasks');
      expect(toolNames).toContain('get_task');
      expect(toolNames).toContain('list_notebooks');
      expect(toolNames).toContain('send_email');
      // Signature + scheduling tools are always present (Graph-only server).
      expect(toolNames).toContain('set_signature');
      expect(toolNames).toContain('get_signature');
      expect(toolNames).toContain('check_availability');
      expect(toolNames).toContain('find_meeting_times');

      // Tools removed with the AppleScript backend must NOT reappear. Notes are
      // now served by OneNote (list_notebooks etc.); list_accounts had no Graph
      // analog.
      for (const removed of ['list_accounts', 'list_notes', 'get_note', 'search_notes']) {
        expect(toolNames).not.toContain(removed);
      }

      // Clean up
      await client.close();
      await server.close();
    });

    it('ignores USE_APPLESCRIPT (the backend was removed) and still serves the Graph surface', async () => {
      const prev = process.env['USE_APPLESCRIPT'];
      process.env['USE_APPLESCRIPT'] = '1';
      try {
        const server = createServer();
        const client = new Client({ name: 'test-client', version: '1.0.0' }, { capabilities: {} });
        const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
        await Promise.all([client.connect(clientTransport), server.connect(serverTransport)]);

        const result = await client.listTools();
        // Still the full Graph surface, not a (nonexistent) AppleScript one.
        expect(result.tools.length).toBe(229);
        expect(result.tools.map((t) => t.name)).toContain('list_folders');

        await client.close();
        await server.close();
      } finally {
        if (prev === undefined) delete process.env['USE_APPLESCRIPT'];
        else process.env['USE_APPLESCRIPT'] = prev;
      }
    });

    it('serves registry-migrated tools exactly once with registry-sourced annotations', async () => {
      // Verifies the ListTools union (registry tools + legacy TOOLS minus the
      // migrated names) has no duplicates and that migrated tools carry the
      // metadata from their registry definition, not the legacy TOOLS entry.
      const server = createServer();
      const client = new Client({ name: 'test-client', version: '1.0.0' }, { capabilities: {} });
      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await Promise.all([client.connect(clientTransport), server.connect(serverTransport)]);

      const result = await client.listTools();
      const names = result.tools.map((t) => t.name);

      // No duplicate tool names across the registry + legacy union.
      expect(new Set(names).size).toBe(names.length);
      // Graph-only surface count (229) — the AppleScript backend and its
      // Apple-Notes-only tools (list_notes/get_note/search_notes) and the
      // AppleScript-backed list_accounts tool were removed with it.
      expect(names.length).toBe(229);

      // The 4 migrated mail-rules tools each appear exactly once.
      for (const name of ['list_mail_rules', 'create_mail_rule', 'prepare_delete_mail_rule', 'confirm_delete_mail_rule']) {
        expect(names.filter((n) => n === name)).toHaveLength(1);
      }

      // Registry-sourced annotations are present (legacy TOOLS entries had none).
      const listRule = result.tools.find((t) => t.name === 'list_mail_rules');
      expect(listRule?.annotations?.readOnlyHint).toBe(true);
      const confirmDelete = result.tools.find((t) => t.name === 'confirm_delete_mail_rule');
      expect(confirmDelete?.annotations?.destructiveHint).toBe(true);

      await client.close();
      await server.close();
    });

    it('tools have proper schemas', async () => {
      const server = createServer();
      const client = new Client(
        { name: 'test-client', version: '1.0.0' },
        { capabilities: {} }
      );

      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await Promise.all([
        client.connect(clientTransport),
        server.connect(serverTransport),
      ]);

      const result = await client.listTools();

      // Check list_emails has proper schema
      const listEmails = result.tools.find((t) => t.name === 'list_emails');
      expect(listEmails).toBeDefined();
      expect(listEmails?.inputSchema).toBeDefined();
      expect(listEmails?.inputSchema.type).toBe('object');
      expect(listEmails?.inputSchema.properties).toHaveProperty('folder_id');
      expect(listEmails?.inputSchema.properties).toHaveProperty('limit');
      expect(listEmails?.inputSchema.properties).toHaveProperty('offset');
      expect(listEmails?.inputSchema.properties).toHaveProperty('unread_only');

      // Check get_email has proper schema
      const getEmail = result.tools.find((t) => t.name === 'get_email');
      expect(getEmail).toBeDefined();
      expect(getEmail?.inputSchema.required).toContain('email_id');

      await client.close();
      await server.close();
    });

    it('returns error for unknown tool', async () => {
      const server = createServer();
      const client = new Client(
        { name: 'test-client', version: '1.0.0' },
        { capabilities: {} }
      );

      const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
      await Promise.all([
        client.connect(clientTransport),
        server.connect(serverTransport),
      ]);

      // Call an unknown tool
      const result = await client.callTool({
        name: 'unknown_tool',
        arguments: {},
      });

      expect(result.isError).toBe(true);
      expect(result.content).toBeDefined();

      await client.close();
      await server.close();
    });
  });
});
