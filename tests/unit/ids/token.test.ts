/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import {
  mintSelfEncoded,
  mintComposite,
  parseToken,
  canonicalKey,
  isToken,
  isKnownPrefix,
  SELF_ENCODING_PREFIXES,
  ALIAS_PREFIXES,
  type EntityType,
} from '../../../src/ids/token.js';

describe('ids/token — self-encoding', () => {
  it('round-trips every self-encoding prefix (mint → parse → same Graph ID)', () => {
    const graphId = 'AAMkAGI2=and/some+weird_chars-123';
    for (const entity of Object.values(SELF_ENCODING_PREFIXES)) {
      const token = mintSelfEncoded(entity, graphId);
      const parsed = parseToken(token);
      expect(parsed?.kind).toBe('self');
      expect(parsed?.entityType).toBe(entity);
      expect(parsed?.graphId).toBe(graphId);
    }
  });

  it('is deterministic — the same Graph ID mints a byte-identical token', () => {
    expect(mintSelfEncoded('message', 'AAA')).toBe(mintSelfEncoded('message', 'AAA'));
  });

  it('encodes with the expected prefix', () => {
    expect(mintSelfEncoded('message', 'AAA').startsWith('em_')).toBe(true);
    expect(mintSelfEncoded('driveItem', 'AAA').startsWith('dr_')).toBe(true);
  });

  it('round-trips OneNote entities (notebook/section/page) with their nb_/ns_/np_ prefixes', () => {
    const graphId = '0-ABC123!456';
    const cases: Array<[EntityType, string]> = [
      ['noteNotebook', 'nb_'],
      ['noteSection', 'ns_'],
      ['notePage', 'np_'],
    ];
    for (const [entity, prefix] of cases) {
      const token = mintSelfEncoded(entity, graphId);
      expect(token.startsWith(prefix)).toBe(true);
      const parsed = parseToken(token);
      expect(parsed?.kind).toBe('self');
      expect(parsed?.entityType).toBe(entity);
      expect(parsed?.graphId).toBe(graphId);
    }
  });

  it('preserves Graph IDs containing base64url special chars (_ and -)', () => {
    // base64url payloads themselves contain _ / -, so a first-underscore split
    // must not corrupt the payload.
    const graphId = 'a_b-c_d/e+f=g';
    const parsed = parseToken(mintSelfEncoded('event', graphId));
    expect(parsed?.graphId).toBe(graphId);
  });

  it('rejects minting a self-encoding token for an alias-backed entity', () => {
    expect(() => mintSelfEncoded('attachment', 'X')).toThrow(/not self-encoding/);
  });

  it('rejects an empty Graph ID rather than emitting a payload-less token', () => {
    expect(() => mintSelfEncoded('message', '')).toThrow(/empty Graph ID/);
  });

  it('rejects non-canonical base64url payloads — one Graph ID has exactly one token', () => {
    // 'QQ' is the canonical encoding of "A"; 'QR'/'QS' also leniently decode to
    // "A" but are non-canonical and must be rejected so a control keyed on the
    // token string cannot be evaded by a variant.
    expect(parseToken('em_QQ')?.graphId).toBe('A');
    expect(parseToken('em_QR')).toBeNull();
    expect(parseToken('em_QS')).toBeNull();
    // The canonical mint of any id parses back and equals a re-mint.
    const t = mintSelfEncoded('message', 'A');
    expect(t).toBe('em_QQ');
  });
});

describe('ids/token — composite / alias-backed', () => {
  it('round-trips every alias prefix (mint → parse → kind alias)', () => {
    for (const entity of Object.values(ALIAS_PREFIXES)) {
      const token = mintComposite(entity, canonicalKey(entity, { a: '1', b: '2' }));
      const parsed = parseToken(token);
      expect(parsed?.kind).toBe('alias');
      expect(parsed?.entityType).toBe(entity);
      expect(parsed?.graphId).toBeUndefined();
    }
  });

  it('is deterministic — the same canonical key mints a byte-identical token', () => {
    const key = canonicalKey('attachment', { messageId: 'M', attachmentId: 'A' });
    expect(mintComposite('attachment', key)).toBe(mintComposite('attachment', key));
  });

  it('canonicalKey is order-independent', () => {
    const a = canonicalKey('attachment', { messageId: 'M', attachmentId: 'A' });
    const b = canonicalKey('attachment', { attachmentId: 'A', messageId: 'M' });
    expect(a).toBe(b);
    expect(mintComposite('attachment', a)).toBe(mintComposite('attachment', b));
  });

  it('produces a fixed-length 70-bit digest (14 base32 chars) — no length extension (D1a)', () => {
    const token = mintComposite('attachment', canonicalKey('attachment', { messageId: 'M', attachmentId: 'A' }));
    const digest = token.slice(token.indexOf('_') + 1);
    expect(digest).toHaveLength(14);
    expect(digest).toMatch(/^[a-z2-7]{14}$/);
  });

  it('different keys mint different tokens', () => {
    const t1 = mintComposite('attachment', canonicalKey('attachment', { messageId: 'M1', attachmentId: 'A' }));
    const t2 = mintComposite('attachment', canonicalKey('attachment', { messageId: 'M2', attachmentId: 'A' }));
    expect(t1).not.toBe(t2);
  });

  it('is delimiter-injection safe — a value containing &/= cannot forge a boundary', () => {
    // Without percent-encoding, {messageId:'A', attachmentId:'B'} and
    // {messageId:'A&attachmentId=B'} (single field) could collide.
    const two = canonicalKey('attachment', { messageId: 'A', attachmentId: 'B' });
    const one = canonicalKey('attachment', { messageId: 'A&attachmentId=B' });
    expect(one).not.toBe(two);
    expect(mintComposite('attachment', one)).not.toBe(mintComposite('attachment', two));
  });

  it('rejects minting a composite token for a self-encoding entity', () => {
    expect(() => mintComposite('message', 'X')).toThrow(/not alias-backed/);
  });

  it('mints an xm_ token for channelMessage and classifies it as alias', () => {
    const token = mintComposite('channelMessage', 'k');
    expect(token.startsWith('xm_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('channelMessage');
  });

  it('mints a td_ token for task (composite {taskListId, taskId}) and classifies it as alias', () => {
    const token = mintComposite('task', canonicalKey('task', { taskListId: 'L', taskId: 'T' }));
    expect(token.startsWith('td_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('task');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints a tl_ token for taskList and classifies it as alias', () => {
    const token = mintComposite('taskList', 'k');
    expect(token.startsWith('tl_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('taskList');
  });

  it('mints an mr_ token for mailRule and classifies it as alias', () => {
    const token = mintComposite('mailRule', 'k');
    expect(token.startsWith('mr_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('mailRule');
  });

  it('mints a cf_ token for contactFolder and classifies it as alias', () => {
    const token = mintComposite('contactFolder', 'k');
    expect(token.startsWith('cf_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('contactFolder');
  });

  it('mints a cg_ token for category and classifies it as alias', () => {
    const token = mintComposite('category', 'k');
    expect(token.startsWith('cg_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('category');
  });

  it('mints an fo_ token for focusedOverride and classifies it as alias', () => {
    const token = mintComposite('focusedOverride', 'k');
    expect(token.startsWith('fo_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('focusedOverride');
  });

  it('mints a cp_ token for calendarPermission (composite {calendarId, permissionId}) and classifies it as alias', () => {
    const token = mintComposite('calendarPermission', canonicalKey('calendarPermission', { calendarId: 'C', permissionId: 'P' }));
    expect(token.startsWith('cp_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('calendarPermission');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints an om_ token for onlineMeeting and classifies it as alias', () => {
    const token = mintComposite('onlineMeeting', 'k');
    expect(token.startsWith('om_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('onlineMeeting');
  });

  it('mints an rc_ token for recording (composite {meetingId, recordingId}) and classifies it as alias', () => {
    const token = mintComposite('recording', canonicalKey('recording', { meetingId: 'M', recordingId: 'R' }));
    expect(token.startsWith('rc_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('recording');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints a tr_ token for transcript (composite {meetingId, transcriptId}) and classifies it as alias', () => {
    const token = mintComposite('transcript', canonicalKey('transcript', { meetingId: 'M', transcriptId: 'T' }));
    expect(token.startsWith('tr_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('transcript');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints an si_ token for site and classifies it as alias', () => {
    const token = mintComposite('site', 'k');
    expect(token.startsWith('si_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('site');
  });

  it('mints a dl_ token for documentLibrary (composite {siteId, driveId}) and classifies it as alias', () => {
    const token = mintComposite('documentLibrary', canonicalKey('documentLibrary', { siteId: 'S', driveId: 'D' }));
    expect(token.startsWith('dl_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('documentLibrary');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints a li_ token for libraryDriveItem (composite {driveId, itemId}) and classifies it as alias', () => {
    const token = mintComposite('libraryDriveItem', canonicalKey('libraryDriveItem', { driveId: 'D', itemId: 'I' }));
    expect(token.startsWith('li_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('libraryDriveItem');
    expect(parsed?.graphId).toBeUndefined();
  });

  it('mints a pb_ token for plannerBucket (single-id) and classifies it as alias', () => {
    const token = mintComposite('plannerBucket', 'k');
    expect(token.startsWith('pb_')).toBe(true);
    const parsed = parseToken(token);
    expect(parsed?.kind).toBe('alias');
    expect(parsed?.entityType).toBe('plannerBucket');
    expect(parsed?.graphId).toBeUndefined();
  });
});

describe('ids/token — parse guards', () => {
  it('returns null for non-token strings', () => {
    for (const s of ['', 'nope', 'AAMkAGI2', 'xx_', '_payload', 'zz_abc']) {
      expect(parseToken(s)).toBeNull();
    }
  });

  it('does not misclassify Object.prototype member names as known prefixes', () => {
    // Null-prototype maps mean a prefix like "constructor"/"toString"/"hasOwnProperty"
    // is not an inherited member — these must parse as non-tokens (opaque IDs).
    for (const name of ['constructor', 'toString', 'hasOwnProperty', 'valueOf', '__proto__']) {
      expect(parseToken(`${name}_abc`), name).toBeNull();
      expect(isKnownPrefix(name), name).toBe(false);
    }
  });

  it('isToken / isKnownPrefix agree with parse', () => {
    expect(isToken(mintSelfEncoded('message', 'AAA'))).toBe(true);
    expect(isToken('raw-graph-id')).toBe(false);
    expect(isKnownPrefix('em')).toBe(true);
    expect(isKnownPrefix('at')).toBe(true);
    expect(isKnownPrefix('zz')).toBe(false);
  });

  it('every prefix maps to a distinct entity and back', () => {
    const entities = new Set<EntityType>([
      ...Object.values(SELF_ENCODING_PREFIXES),
      ...Object.values(ALIAS_PREFIXES),
    ]);
    // No prefix collisions across the two maps.
    const prefixes = [...Object.keys(SELF_ENCODING_PREFIXES), ...Object.keys(ALIAS_PREFIXES)];
    expect(new Set(prefixes).size).toBe(prefixes.length);
    expect(entities.size).toBeGreaterThan(10);
  });
});
