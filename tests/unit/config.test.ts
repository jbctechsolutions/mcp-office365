import { describe, it, expect, afterEach } from 'vitest';
import { createConfig, validateConfig } from '../../src/config.js';

describe('createConfig', () => {
  afterEach(() => {
    delete process.env['OUTLOOK_PROFILE'];
    delete process.env['OUTLOOK_BASE_PATH'];
  });

  it('returns default profile name when no env var set', () => {
    const config = createConfig();
    expect(config.profileName).toBe('Main Profile');
  });

  it('returns paths that include the profile name', () => {
    const config = createConfig();
    expect(config.databasePath).toContain('Main Profile');
    expect(config.databasePath).toContain('Outlook.sqlite');
    expect(config.dataPath).toContain('Main Profile');
    expect(config.dataPath).toContain('Data');
  });

  it('respects OUTLOOK_PROFILE env var', () => {
    process.env['OUTLOOK_PROFILE'] = 'Test Profile';
    const config = createConfig();
    expect(config.profileName).toBe('Test Profile');
    expect(config.databasePath).toContain('Test Profile');
  });

  it('respects OUTLOOK_BASE_PATH env var', () => {
    process.env['OUTLOOK_BASE_PATH'] = '/custom/path';
    const config = createConfig();
    expect(config.outlookBasePath).toBe('/custom/path');
    expect(config.databasePath).toContain('/custom/path');
  });

  it('returns all required properties', () => {
    const config = createConfig();
    expect(config).toHaveProperty('profileName');
    expect(config).toHaveProperty('outlookBasePath');
    expect(config).toHaveProperty('databasePath');
    expect(config).toHaveProperty('dataPath');
  });
});

describe('validateConfig', () => {
  it('throws for empty profile name', () => {
    const config = { profileName: '', outlookBasePath: '', databasePath: '', dataPath: '' };
    expect(() => validateConfig(config)).toThrow('Profile name cannot be empty');
  });

  it('throws for whitespace-only profile name', () => {
    const config = { profileName: '   ', outlookBasePath: '', databasePath: '', dataPath: '' };
    expect(() => validateConfig(config)).toThrow('Profile name cannot be empty');
  });

  it('does not throw for valid config', () => {
    const config = createConfig();
    expect(() => validateConfig(config)).not.toThrow();
  });
});
