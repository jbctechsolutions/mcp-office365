/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for Graph API authentication configuration.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  GRAPH_SCOPES,
  loadGraphConfig,
  getAuthorityUrl,
} from '../../../../src/graph/auth/config.js';

describe('graph/auth/config', () => {
  const originalEnv = { ...process.env };

  beforeEach(() => {
    // Clear relevant env vars before each test
    delete process.env['OUTLOOK_MCP_CLIENT_ID'];
    delete process.env['OUTLOOK_MCP_TENANT_ID'];
  });

  afterEach(() => {
    // Restore original env
    process.env = { ...originalEnv };
  });

  describe('GRAPH_SCOPES', () => {
    it('contains all required scopes', () => {
      expect(GRAPH_SCOPES).toContain('Mail.ReadWrite');
      expect(GRAPH_SCOPES).toContain('Calendars.ReadWrite');
      expect(GRAPH_SCOPES).toContain('Contacts.ReadWrite');
      expect(GRAPH_SCOPES).toContain('Tasks.ReadWrite');
      expect(GRAPH_SCOPES).toContain('User.Read');
      expect(GRAPH_SCOPES).toContain('offline_access');
    });

    it('contains M365 expansion scopes', () => {
      expect(GRAPH_SCOPES).toContain('ChannelMessage.Read.All');
      expect(GRAPH_SCOPES).toContain('ChannelMessage.Send');
      expect(GRAPH_SCOPES).toContain('Channel.ReadBasic.All');
      expect(GRAPH_SCOPES).toContain('Team.ReadBasic.All');
      expect(GRAPH_SCOPES).toContain('Chat.ReadWrite');
      expect(GRAPH_SCOPES).toContain('ChatMessage.Send');
      expect(GRAPH_SCOPES).toContain('People.Read');
      expect(GRAPH_SCOPES).toContain('User.ReadBasic.All');
      expect(GRAPH_SCOPES).toContain('Presence.Read.All');
      expect(GRAPH_SCOPES).toContain('Group.Read.All');
    });

    it('has exactly 18 scopes', () => {
      expect(GRAPH_SCOPES).toHaveLength(18);
    });
  });

  describe('loadGraphConfig', () => {
    it('loads embedded default client ID', () => {
      const config = loadGraphConfig();

      expect(config.clientId).toBeTruthy();
      expect(config.clientId).not.toBe('YOUR_AZURE_APP_CLIENT_ID');
      // Verify it's a valid UUID format
      expect(config.clientId).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);
    });

    it('uses environment variable for client ID', () => {
      process.env['OUTLOOK_MCP_CLIENT_ID'] = 'test-client-id-123';

      const config = loadGraphConfig();

      expect(config.clientId).toBe('test-client-id-123');
    });

    it('uses default tenant ID when not set', () => {
      process.env['OUTLOOK_MCP_CLIENT_ID'] = 'test-client-id';

      const config = loadGraphConfig();

      expect(config.tenantId).toBe('common');
    });

    it('uses environment variable for tenant ID', () => {
      process.env['OUTLOOK_MCP_CLIENT_ID'] = 'test-client-id';
      process.env['OUTLOOK_MCP_TENANT_ID'] = 'my-tenant-id';

      const config = loadGraphConfig();

      expect(config.tenantId).toBe('my-tenant-id');
    });

    it('includes all required scopes', () => {
      process.env['OUTLOOK_MCP_CLIENT_ID'] = 'test-client-id';

      const config = loadGraphConfig();

      expect(config.scopes).toEqual(expect.arrayContaining([
        'Mail.ReadWrite',
        'Calendars.ReadWrite',
        'Contacts.ReadWrite',
        'Tasks.ReadWrite',
        'User.Read',
        'offline_access',
      ]));
    });

    it('returns scopes as a new array (not the original)', () => {
      process.env['OUTLOOK_MCP_CLIENT_ID'] = 'test-client-id';

      const config = loadGraphConfig();

      expect(config.scopes).not.toBe(GRAPH_SCOPES);
      expect(config.scopes).toEqual([...GRAPH_SCOPES]);
    });
  });

  describe('getAuthorityUrl', () => {
    it('constructs URL with common tenant', () => {
      const config = {
        clientId: 'test-client-id',
        tenantId: 'common',
        scopes: ['Mail.Read'],
      };

      const url = getAuthorityUrl(config);

      expect(url).toBe('https://login.microsoftonline.com/common');
    });

    it('constructs URL with specific tenant', () => {
      const config = {
        clientId: 'test-client-id',
        tenantId: 'my-org-tenant-id',
        scopes: ['Mail.Read'],
      };

      const url = getAuthorityUrl(config);

      expect(url).toBe('https://login.microsoftonline.com/my-org-tenant-id');
    });

    it('constructs URL with organizations tenant', () => {
      const config = {
        clientId: 'test-client-id',
        tenantId: 'organizations',
        scopes: ['Mail.Read'],
      };

      const url = getAuthorityUrl(config);

      expect(url).toBe('https://login.microsoftonline.com/organizations');
    });

    it('constructs URL with consumers tenant', () => {
      const config = {
        clientId: 'test-client-id',
        tenantId: 'consumers',
        scopes: ['Mail.Read'],
      };

      const url = getAuthorityUrl(config);

      expect(url).toBe('https://login.microsoftonline.com/consumers');
    });
  });
});
