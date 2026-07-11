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
 * Default client ID for the "MCP Office 365" Azure AD app.
 *
 * This is a single-tenant public client application. Because it is single-tenant
 * (signInAudience = AzureADMyOrg), device-code sign-in must target its home tenant
 * directly — the `common`/`organizations` authorities cannot resolve a tenant for
 * it and fail with AADSTS50059. DEFAULT_TENANT_ID below is that home tenant.
 *
 * Users can override both with their own app registration by setting:
 * - OUTLOOK_MCP_CLIENT_ID environment variable
 * - OUTLOOK_MCP_TENANT_ID environment variable
 *
 * For setup instructions: https://github.com/jbctechsolutions/mcp-office365#custom-azure-ad-setup
 */
const DEFAULT_CLIENT_ID = '79313c4f-ff74-412f-9913-88de737d5891';

/**
 * Home tenant of the embedded single-tenant app. Overridable via
 * OUTLOOK_MCP_TENANT_ID (use `common`/`organizations`/`consumers` or a specific
 * tenant ID when supplying your own multi-tenant OUTLOOK_MCP_CLIENT_ID).
 */
const DEFAULT_TENANT_ID = '761e2c5f-34bd-4872-b86c-3a9f3b29d63a';

/**
 * Microsoft Graph API scopes required for Outlook access.
 *
 * Includes read/write permissions for mail and calendars to support
 * future implementation of email sending and event management features.
 */
export const GRAPH_SCOPES = [
  // Outlook
  'Mail.ReadWrite',
  'Calendars.ReadWrite',
  'Contacts.ReadWrite',
  'Tasks.ReadWrite',
  'User.Read',
  'offline_access',
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
  'Sites.Read.All',
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
  const clientId = process.env['OUTLOOK_MCP_CLIENT_ID'] ?? DEFAULT_CLIENT_ID;
  const tenantId = process.env['OUTLOOK_MCP_TENANT_ID'] ?? DEFAULT_TENANT_ID;

  if (clientId === 'YOUR_AZURE_APP_CLIENT_ID') {
    throw new Error(
      'Azure AD app not configured. Either:\n' +
        '1. Set OUTLOOK_MCP_CLIENT_ID environment variable, or\n' +
        '2. The package maintainer needs to embed the client ID in config.ts'
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
