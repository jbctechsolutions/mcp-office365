/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import { compileEmailSearch } from '../../../src/search/compiler.js';
import { ErrorCode } from '../../../src/utils/errors.js';

describe('compileEmailSearch — mechanism selection (D9)', () => {
  it('property-only criteria compile to $filter', () => {
    const c = compileEmailSearch({ received_after: '2026-01-01' });
    expect(c.mechanism).toBe('filter');
    if (c.mechanism === 'filter') {
      expect(c.filter).toBe('receivedDateTime ge 2026-01-01T00:00:00Z');
    }
  });

  it('from/to are $search terms, not $filter (Graph $filter cannot express them reliably)', () => {
    const cf = compileEmailSearch({ from: 'alice@example.com' });
    expect(cf.mechanism).toBe('search');
    if (cf.mechanism === 'search') expect(cf.search).toBe('"from:alice@example.com"');
    const ct = compileEmailSearch({ to: 'bob@example.com' });
    expect(ct.mechanism).toBe('search');
    if (ct.mechanism === 'search') expect(ct.search).toBe('"to:bob@example.com"');
  });

  it('free-text-only criteria compile to a quoted $search', () => {
    const c = compileEmailSearch({ text: 'quarterly report' });
    expect(c.mechanism).toBe('search');
    if (c.mechanism === 'search') {
      expect(c.search).toBe('"quarterly report"');
    }
  });

  it('mixed property + free-text compile to /search/query KQL', () => {
    const c = compileEmailSearch({ received_after: '2026-01-01', body_contains: 'invoice' });
    expect(c.mechanism).toBe('searchQuery');
    if (c.mechanism === 'searchQuery') {
      expect(c.kql).toBe('received>=2026-01-01 AND body:"invoice"');
    }
  });

  it('from + a property filter is mixed → /search/query KQL', () => {
    const c = compileEmailSearch({ from: 'alice@example.com', received_after: '2026-01-01' });
    expect(c.mechanism).toBe('searchQuery');
    if (c.mechanism === 'searchQuery') {
      expect(c.kql).toBe('from:"alice@example.com" AND received>=2026-01-01');
    }
  });
});

describe('compileEmailSearch — $filter compilation', () => {
  it('combines two properties with a single "and"-joined $filter (stable order)', () => {
    const c = compileEmailSearch({ received_after: '2026-01-01', has_attachments: true });
    expect(c).toEqual({
      mechanism: 'filter',
      filter: 'receivedDateTime ge 2026-01-01T00:00:00Z and hasAttachments eq true',
    });
  });

  it('normalizes a bare date to midnight UTC and passes a full datetime through', () => {
    expect(compileEmailSearch({ received_before: '2026-03-15' })).toEqual({
      mechanism: 'filter',
      filter: 'receivedDateTime le 2026-03-15T00:00:00Z',
    });
    expect(compileEmailSearch({ received_after: '2026-03-15T09:30:00Z' })).toEqual({
      mechanism: 'filter',
      filter: 'receivedDateTime ge 2026-03-15T09:30:00Z',
    });
  });

  it('maps flags correctly (has_attachments, is_unread, importance)', () => {
    expect((compileEmailSearch({ has_attachments: true }) as { filter: string }).filter).toBe(
      'hasAttachments eq true',
    );
    // is_unread:true means isRead eq false.
    expect((compileEmailSearch({ is_unread: true }) as { filter: string }).filter).toBe('isRead eq false');
    expect((compileEmailSearch({ is_unread: false }) as { filter: string }).filter).toBe('isRead eq true');
    expect((compileEmailSearch({ importance: 'high' }) as { filter: string }).filter).toBe(
      "importance eq 'high'",
    );
  });

});

describe('compileEmailSearch — $search compilation', () => {
  it('quotes each term and AND-joins subject/body/free-text', () => {
    const c = compileEmailSearch({ subject_contains: 'budget', body_contains: 'Q3' }) as { search: string };
    expect(c.search).toBe('"subject:budget" AND "body:Q3"');
  });

  it('strips embedded double quotes from a free-text term (no in-phrase escape)', () => {
    const c = compileEmailSearch({ text: 'say "hi"' }) as { search: string };
    expect(c.search).toBe('"say hi"');
  });

  it('a quote-injection attempt cannot break out of the $search phrase', () => {
    // Without stripping, `a" AND "b` would inject an AND. Stripped → literal.
    const c = compileEmailSearch({ text: 'a" AND "b' }) as { search: string };
    expect(c.search).toBe('"a AND b"');
    expect(c.search).not.toContain('" AND "');
  });
});

describe('compileEmailSearch — mixed KQL compilation', () => {
  it('emits KQL dates as YYYY-MM-DD and quotes terms', () => {
    const c = compileEmailSearch({
      received_after: '2026-01-01',
      subject_contains: 'quarterly report',
    }) as { kql: string };
    expect(c.kql).toBe('received>=2026-01-01 AND subject:"quarterly report"');
  });

  it('combines many mixed criteria deterministically (all free values quoted)', () => {
    const c = compileEmailSearch({
      from: 'alice@example.com',
      received_after: '2026-01-01',
      has_attachments: true,
      text: 'invoice',
    }) as { kql: string };
    expect(c.kql).toBe(
      'from:"alice@example.com" AND received>=2026-01-01 AND hasattachment:true AND "invoice"',
    );
  });

  it('a from value with KQL operators is mixed only via a real filter; injection stays quoted', () => {
    // from + a property filter → /search/query KQL; the from value stays inside
    // one quoted phrase, so it cannot inject an extra clause.
    const c = compileEmailSearch({
      from: 'alice@x.com AND body:secret',
      received_after: '2026-01-01',
    }) as { kql: string };
    expect(c.kql).toBe('from:"alice@x.com AND body:secret" AND received>=2026-01-01');
  });
});

describe('compileEmailSearch — validation', () => {
  it('rejects empty params with a VALIDATION error', () => {
    try {
      compileEmailSearch({});
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.VALIDATION_ERROR);
      expect((e as Error).message).toMatch(/at least one/i);
    }
  });

  it('rejects a non-ISO date (e.g. "yesterday") with a VALIDATION error naming the field', () => {
    try {
      compileEmailSearch({ received_after: 'yesterday' });
      expect.unreachable('should throw');
    } catch (e) {
      expect((e as { code?: string }).code).toBe(ErrorCode.VALIDATION_ERROR);
      expect((e as Error).message).toMatch(/received_after/);
    }
  });

  it('rejects an impossible date that matches the shape but is invalid', () => {
    expect(() => compileEmailSearch({ received_after: '2026-13-45' })).toThrow(/received_after/);
  });

  it('rejects rollover-invalid calendar dates (Feb 30, non-leap Feb 29, Apr 31)', () => {
    // Date.parse rolls these to a valid instant, so the strict re-serialize
    // check must reject them rather than silently shifting the window.
    for (const bad of ['2026-02-30', '2026-02-29', '2026-04-31', '2026-06-31']) {
      expect(() => compileEmailSearch({ received_after: bad }), bad).toThrow(/valid ISO date/);
    }
    // A real leap day is accepted.
    expect(() => compileEmailSearch({ received_after: '2024-02-29' })).not.toThrow();
  });

  it('treats whitespace-only free-text as absent (empty → validation error)', () => {
    expect(() => compileEmailSearch({ text: '   ' })).toThrow(/at least one/i);
  });
});
