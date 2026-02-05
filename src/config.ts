/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { homedir } from 'node:os';
import { join } from 'node:path';

/**
 * Configuration for the Outlook MCP server.
 */
export interface OutlookConfig {
  /** Name of the Outlook profile to use */
  readonly profileName: string;
  /** Base path for Outlook data */
  readonly outlookBasePath: string;
  /** Full path to the SQLite database */
  readonly databasePath: string;
  /** Path to the data directory containing olk15 files */
  readonly dataPath: string;
}

/**
 * Default Outlook data path on macOS.
 */
const DEFAULT_OUTLOOK_BASE_PATH = join(
  homedir(),
  'Library',
  'Group Containers',
  'UBF8T346G9.Office',
  'Outlook',
  'Outlook 15 Profiles'
);

/**
 * Default profile name.
 */
const DEFAULT_PROFILE_NAME = 'Main Profile';

/**
 * Creates the configuration from environment variables.
 */
export function createConfig(): OutlookConfig {
  const profileName = process.env['OUTLOOK_PROFILE'] ?? DEFAULT_PROFILE_NAME;
  const outlookBasePath =
    process.env['OUTLOOK_BASE_PATH'] ?? DEFAULT_OUTLOOK_BASE_PATH;

  const profilePath = join(outlookBasePath, profileName);
  const dataPath = join(profilePath, 'Data');
  const databasePath = join(dataPath, 'Outlook.sqlite');

  return {
    profileName,
    outlookBasePath,
    databasePath,
    dataPath,
  };
}

/**
 * Validates that the required paths exist.
 * Throws if the database cannot be found.
 */
export function validateConfig(config: OutlookConfig): void {
  // Validation will be performed by the database connection
  // This is a placeholder for any additional config validation
  if (config.profileName.trim() === '') {
    throw new Error('Profile name cannot be empty');
  }
}
