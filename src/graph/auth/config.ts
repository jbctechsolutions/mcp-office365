/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Microsoft Graph API configuration.
 *
 * Contains Azure AD app registration settings for the device code flow.
 * The client ID is embedded in the package - users don't need to configure anything.
 */

/**
 * No client ID is embedded. Device-code sign-in requires an Azure AD app that
 * belongs to the tenant you sign in against, so a single baked-in default cannot
 * serve everyone — each deployment brings its own registration via:
 * - OUTLOOK_MCP_CLIENT_ID  (required)
 * - OUTLOOK_MCP_TENANT_ID  (required for a single-tenant app; defaults to 'common')
 *
 * For setup instructions: https://github.com/jbctechsolutions/mcp-office365#custom-azure-ad-setup
 */
const CLIENT_ID_SETUP_URL =
  'https://github.com/jbctechsolutions/mcp-office365#custom-azure-ad-setup';

/**
 * Microsoft Graph API scopes required for Outlook access.
 *
 * Includes read/write permissions for mail and calendars to support
 * future implementation of email sending and event management features.
 */
export const GRAPH_SCOPES = [
  // Outlook
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite',
  'Tasks.ReadWrite',
  'User.Read',
  'offline_access',
  // Shared mailboxes / delegate access (#40) — read another user's mailbox,
  // calendar, and OneDrive via /users/{upn}/... where the signed-in user has
  // shared/delegate access.
  'Mail.Read.Shared',
  'Calendars.Read.Shared',
  'Files.Read.All',
  // Teams
  'ChannelMessage.Read.All',
  'ChannelMessage.Send',
  'Channel.ReadBasic.All',
  'Team.ReadBasic.All',
  'Chat.ReadWrite',
  'ChatMessage.Send',
  // People & Presence
  'People.Read',
  'User.ReadBasic.All',
  'Presence.Read.All',
  // Planner
  'Group.Read.All',
  // OneDrive
  'Files.ReadWrite',
  // SharePoint
  'Sites.ReadWrite.All',
  // OneNote (read/search pages + create pages)
  'Notes.ReadWrite',
] as const;

/**
 * Microsoft Graph API configuration.
 */
export interface GraphAuthConfig {
  /** Azure AD application (client) ID */
  readonly clientId: string;
  /** Azure AD tenant ID (default: 'common' for multi-tenant) */
  readonly tenantId: string;
  /** OAuth2 scopes to request */
  readonly scopes: readonly string[];
}

/**
 * Loads the Graph API configuration.
 *
 * Uses embedded defaults but allows override via environment variables
 * for users who want to use their own Azure AD app registration.
 */
export function loadGraphConfig(): GraphAuthConfig {
  const clientId = process.env['OUTLOOK_MCP_CLIENT_ID'];
  const tenantId = process.env['OUTLOOK_MCP_TENANT_ID'] ?? 'common';

  if (clientId == null || clientId === '' || clientId === 'YOUR_AZURE_APP_CLIENT_ID') {
    throw new Error(
      'OUTLOOK_MCP_CLIENT_ID is required — no Azure AD app is embedded.\n' +
        'Register an app (Microsoft Graph delegated permissions, "Allow public client flows" = Yes), then set:\n' +
        '  OUTLOOK_MCP_CLIENT_ID=<your app (client) ID>\n' +
        '  OUTLOOK_MCP_TENANT_ID=<your tenant ID>   # required for a single-tenant app; omit for multi-tenant\n' +
        `Setup: ${CLIENT_ID_SETUP_URL}`
    );
  }

  return {
    clientId,
    tenantId,
    scopes: [...GRAPH_SCOPES],
  };
}

/**
 * Gets the Azure AD authority URL for the configured tenant.
 */
export function getAuthorityUrl(config: GraphAuthConfig): string {
  return `https://login.microsoftonline.com/${config.tenantId}`;
}
