/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { StateStore } from '../../../src/state/store.js';
import { DeltaMirror } from '../../../src/delta/mirror.js';
import type { GraphClient } from '../../../src/graph/client/index.js';
import { parseToken } from '../../../src/ids/token.js';

interface MsgResp { messages: unknown[]; deltaLink: string }
interface EvtResp { events: unknown[]; deltaLink: string }

/** A scripted GraphClient: each delta call shifts the next queued response. */
class FakeClient {
  mailQueue: MsgResp[] = [];
  calQueue: EvtResp[] = [];
  lastMailDeltaLink: string | undefined;

  getMessagesDelta(_folderId: string, deltaLink?: string): Promise<MsgResp> {
    this.lastMailDeltaLink = deltaLink;
    const next = this.mailQueue.shift();
    if (next == null) throw new Error('no scripted mail response');
    return Promise.resolve(next);
  }

  getCalendarViewDelta(_start: string, _end: string, _deltaLink?: string): Promise<EvtResp> {
    const next = this.calQueue.shift();
    if (next == null) throw new Error('no scripted calendar response');
    return Promise.resolve(next);
  }
}

function msg(id: string, extra: Record<string, unknown> = {}): unknown {
  return { id, subject: `subject-${id}`, receivedDateTime: '2026-01-01T00:00:00Z', isRead: false, ...extra };
}

function removed(id: string): unknown {
  return { id, '@removed': { reason: 'deleted' } };
}

let dir: string;
let legacyDir: string;
let store: StateStore;
let fake: FakeClient;
let mirror: DeltaMirror;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'mcp-mirror-'));
  legacyDir = mkdtempSync(join(tmpdir(), 'mcp-mirror-legacy-'));
  store = StateStore.open({ dir, legacyDir, warn: () => {} });
  fake = new FakeClient();
  mirror = new DeltaMirror(fake as unknown as GraphClient, store, () => 'acct-1', () => 1000);
});

afterEach(() => {
  try { store.close(); } catch { /* already closed */ }
  rmSync(dir, { recursive: true, force: true });
  rmSync(legacyDir, { recursive: true, force: true });
});

describe('DeltaMirror mail', () => {
  it('first sync establishes a baseline without reporting per-item creates', async () => {
    fake.mailQueue.push({ messages: [msg('a'), msg('b')], deltaLink: 'link-1' });

    const report = await mirror.sync(['mail']);
    const mail = report.resources[0]!;

    expect(fake.lastMailDeltaLink).toBeUndefined(); // no prior cursor
    expect(mail.baseline).toBe(true);
    expect(mail.created).toHaveLength(0);
    expect(mail.updated).toHaveLength(0);
    expect(mail.trackedCount).toBe(2);
    expect(store.delta.getDeltaLink('acct-1', 'mail:inbox')).toBe('link-1');
  });

  it('classifies created / updated / deleted on the second sync and mints em_ tokens', async () => {
    fake.mailQueue.push({ messages: [msg('a'), msg('b')], deltaLink: 'link-1' });
    await mirror.sync(['mail']);

    // a updated, c new, b removed.
    fake.mailQueue.push({ messages: [msg('a', { subject: 'renamed' }), msg('c'), removed('b')], deltaLink: 'link-2' });
    const report = await mirror.sync(['mail']);
    const mail = report.resources[0]!;

    expect(fake.lastMailDeltaLink).toBe('link-1'); // followed the stored cursor
    expect(mail.baseline).toBe(false);
    expect(mail.created.map((c) => c.graphId)).toEqual(['c']);
    expect(mail.updated.map((u) => u.graphId)).toEqual(['a']);
    expect(mail.deleted.map((d) => d.graphId)).toEqual(['b']);
    expect(mail.trackedCount).toBe(2); // a + c

    // Tokens are durable self-encoding em_ tokens that decode to the graph id.
    const created = mail.created[0]!;
    expect(created.token.startsWith('em_')).toBe(true);
    expect(parseToken(created.token)?.graphId).toBe('c');
    // Deleted entry carries the summary from the mirror.
    expect(mail.deleted[0]!.summary).toBe('subject-b');
  });

  it('re-baselines and flags a note when Graph returns an empty deltaLink', async () => {
    fake.mailQueue.push({ messages: [msg('a')], deltaLink: '' });
    const first = await mirror.sync(['mail']);
    expect(first.resources[0]!.note).toBeDefined();
    expect(store.delta.getDeltaLink('acct-1', 'mail:inbox')).toBeNull();

    // Next call has no cursor → baseline again.
    fake.mailQueue.push({ messages: [msg('a')], deltaLink: 'link-2' });
    const second = await mirror.sync(['mail']);
    expect(second.resources[0]!.baseline).toBe(true);
  });

  it('reset clears tracking so the next sync re-baselines', async () => {
    fake.mailQueue.push({ messages: [msg('a')], deltaLink: 'link-1' });
    await mirror.sync(['mail']);
    mirror.reset(['mail']);
    expect(store.delta.getDeltaLink('acct-1', 'mail:inbox')).toBeNull();

    fake.mailQueue.push({ messages: [msg('a')], deltaLink: 'link-2' });
    const report = await mirror.sync(['mail']);
    expect(report.resources[0]!.baseline).toBe(true);
  });
});

describe('DeltaMirror calendar', () => {
  it('baselines then reports an updated event with an ev_ token', async () => {
    fake.calQueue.push({
      events: [{ id: 'e1', subject: 'Standup', start: { dateTime: '2026-01-02T09:00:00' }, end: { dateTime: '2026-01-02T09:30:00' } }],
      deltaLink: 'cal-1',
    });
    const baseline = await mirror.sync(['calendar']);
    expect(baseline.resources[0]!.baseline).toBe(true);
    expect(baseline.resources[0]!.trackedCount).toBe(1);

    fake.calQueue.push({
      events: [{ id: 'e1', subject: 'Standup (moved)', start: { dateTime: '2026-01-02T10:00:00' }, end: { dateTime: '2026-01-02T10:30:00' } }],
      deltaLink: 'cal-2',
    });
    const report = await mirror.sync(['calendar']);
    const cal = report.resources[0]!;

    expect(cal.updated.map((u) => u.graphId)).toEqual(['e1']);
    const updated = cal.updated[0]!;
    expect(updated.token.startsWith('ev_')).toBe(true);
    expect(parseToken(updated.token)?.graphId).toBe('e1');
  });
});

describe('DeltaMirror multi-resource', () => {
  it('syncs mail and calendar together by default', async () => {
    fake.mailQueue.push({ messages: [msg('a')], deltaLink: 'm1' });
    fake.calQueue.push({ events: [{ id: 'e1', subject: 'x' }], deltaLink: 'c1' });

    const report = await mirror.sync();
    expect(report.resources.map((r) => r.resource)).toEqual(['mail', 'calendar']);
    expect(report.resources.every((r) => r.baseline)).toBe(true);
  });
});
