/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Tests for the canonical per-entity id schemas + next-action hints (U6).
 */

import { describe, it, expect } from 'vitest';
import { Id, idSchema, optionalIdSchema, describeId, ENTITY_META } from '../../../src/ids/schema.js';
import { nextActionFor } from '../../../src/ids/next-action.js';
import { prefixForEntity, mintSelfEncoded, type EntityType } from '../../../src/ids/token.js';
import { allToolDefinitions } from '../../../src/registry/all-tools.js';

describe('canonical id schema', () => {
  it('trims surrounding whitespace', () => {
    expect(Id.task.parse('  td_abc  ')).toBe('td_abc');
  });

  it('rejects an empty / whitespace-only id', () => {
    expect(Id.task.safeParse('').success).toBe(false);
    expect(Id.task.safeParse('   ').success).toBe(false);
  });

  it('is string-only: a legacy numeric-type id is a type error (not silently accepted)', () => {
    // A numeric *string* is accepted at the schema boundary and left for
    // resolveId to classify (NUMERIC_ID_UNSUPPORTED); a JSON number is a type error.
    expect(Id.message.safeParse('12345').success).toBe(true);
    expect(Id.message.safeParse(12345 as unknown as string).success).toBe(false);
  });

  it('accepts a raw graph id (passthrough — classification stays in resolveId)', () => {
    expect(Id.message.parse('AAMkAGI2=').length).toBeGreaterThan(0);
  });

  it('accepts a correctly-minted self-encoding token', () => {
    const token = mintSelfEncoded('message', 'AAMkAG-real-id');
    expect(Id.message.parse(token)).toBe(token);
  });

  it('does NOT reject a wrong-entity token at the boundary (resolveId owns that)', () => {
    // B=minimal: no wrong-entity refine in the schema — an ev_ token passes the
    // task schema here and is rejected later by resolveId (ID_ENTITY_MISMATCH).
    const eventToken = mintSelfEncoded('event', 'evt-1');
    expect(Id.task.safeParse(eventToken).success).toBe(true);
  });

  it('describes the entity with its token prefix and a source tool', () => {
    const desc = describeId('task');
    expect(desc).toContain('`td_`');
    expect(desc).toContain('list_tasks');
    expect(Id.task.description).toBe(desc);
  });

  it('every entity in Id carries a prefix-named description', () => {
    for (const [entity, schema] of Object.entries(Id)) {
      const prefix = prefixForEntity(entity as EntityType);
      expect(schema.description, entity).toContain(`\`${prefix}_\``);
    }
  });

  it('optionalIdSchema is optional but still described', () => {
    const schema = optionalIdSchema('folder');
    expect(schema.safeParse(undefined).success).toBe(true);
    expect(idSchema('folder').description).toContain('`fd_`');
  });

  it('every ENTITY_META source tool is a registered tool name (no description drift)', () => {
    const registered = new Set(allToolDefinitions().map((d) => d.name));
    for (const [entity, meta] of Object.entries(ENTITY_META)) {
      for (const tool of meta.from.split('/').map((t) => t.trim())) {
        expect(registered.has(tool), `${entity} → "${tool}"`).toBe(true);
      }
    }
  });
});

describe('nextActionFor', () => {
  it('returns a prefix-named follow-up sentence for mapped entities', () => {
    const hint = nextActionFor('plan');
    expect(hint).toContain('`pl_`');
    expect(hint).toContain('get_plan');
  });

  it('names the right prefix per entity', () => {
    expect(nextActionFor('task')).toContain('`td_`');
    expect(nextActionFor('plannerTask')).toContain('`pt_`');
    expect(nextActionFor('message')).toContain('`em_`');
  });

  it('returns null for entities without a defined follow-up', () => {
    expect(nextActionFor('recording')).toBeNull();
  });
});
