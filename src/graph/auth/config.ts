/**
 * Microsoft Graph API configuration.
 *
 * Contains Azure AD app registration settings for the device code flow.
 * The client ID is embedded in the package - users don't need to configure anything.
 */

/**
 * Default client ID for the Outlook MCP Server Azure AD app.
 *
 * This is a public client application registered in Azure AD.
 * Users can override this with their own app registration if needed.
 *
 * TODO: Replace with actual Azure AD app client ID before publishing.
 */
const DEFAULT_CLIENT_ID = 'YOUR_AZURE_APP_CLIENT_ID';

/**
 * Microsoft Graph API scopes required for read-only Outlook access.
 */
export const GRAPH_SCOPES = [
  'Mail.Read',
  'Calendars.Read',
  'Contacts.Read',
  'Tasks.Read',
  'User.Read',
  'offline_access',
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
  const tenantId = process.env['OUTLOOK_MCP_TENANT_ID'] ?? 'common';

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
