# Email Signatures & Calendar Scheduling Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add 4 new MCP tools — `set_signature`, `get_signature`, `check_availability`, `find_meeting_times` — with auto-append signature behavior on email creation/send tools.

**Architecture:** Two independent features sharing the same layered architecture (GraphClient → GraphRepository → Tools → index.ts dispatcher). Signatures are stored as HTML files at `~/.outlook-mcp/signature.html` and auto-appended to email bodies. Scheduling uses existing `Calendars.ReadWrite` Graph API permissions.

**Tech Stack:** TypeScript, Zod, Microsoft Graph API, Vitest, Node.js `fs` module

---

### Task 1: Signature Storage Module

**Files:**
- Create: `src/signature.ts`
- Create: `tests/unit/signature.test.ts`

**Step 1: Write the failing tests**

Create `tests/unit/signature.test.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { join } from 'node:path';

// Mock fs and os before imports
const mockReadFileSync = vi.fn();
const mockWriteFileSync = vi.fn();
const mockExistsSync = vi.fn();
const mockMkdirSync = vi.fn();
const mockHomedir = vi.fn().mockReturnValue('/mock/home');

vi.mock('node:fs', () => ({
  readFileSync: mockReadFileSync,
  writeFileSync: mockWriteFileSync,
  existsSync: mockExistsSync,
  mkdirSync: mockMkdirSync,
}));

vi.mock('node:os', () => ({
  homedir: mockHomedir,
}));

import { readSignature, writeSignature, appendSignature } from '../../src/signature.js';

describe('signature', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('readSignature', () => {
    it('returns signature content when file exists', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>-- Joel</p>');

      const result = readSignature();

      expect(result).toBe('<p>-- Joel</p>');
      expect(mockExistsSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html')
      );
      expect(mockReadFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        'utf-8'
      );
    });

    it('returns null when file does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      const result = readSignature();

      expect(result).toBeNull();
      expect(mockReadFileSync).not.toHaveBeenCalled();
    });
  });

  describe('writeSignature', () => {
    it('writes HTML content to signature file', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('<p>Best regards,<br>Joel</p>');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<p>Best regards,<br>Joel</p>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });

    it('creates directory if it does not exist', () => {
      mockExistsSync.mockReturnValue(false);

      writeSignature('<p>Sig</p>');

      expect(mockMkdirSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp'),
        { recursive: true, mode: 0o700 }
      );
    });

    it('wraps plain text in <pre> tag when content_type is text', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('-- Joel\nSenior Dev', 'text');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<pre>-- Joel\nSenior Dev</pre>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });

    it('stores HTML content directly when content_type is html', () => {
      mockExistsSync.mockReturnValue(true);

      writeSignature('<b>Joel</b>', 'html');

      expect(mockWriteFileSync).toHaveBeenCalledWith(
        join('/mock/home', '.outlook-mcp', 'signature.html'),
        '<b>Joel</b>',
        { encoding: 'utf-8', mode: 0o600 }
      );
    });
  });

  describe('appendSignature', () => {
    it('appends signature to HTML body with <br><br> separator', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>-- Joel</p>');

      const result = appendSignature('<p>Hello</p>', 'html', true);

      expect(result).toBe('<p>Hello</p><br><br><p>-- Joel</p>');
    });

    it('appends signature to text body with \\n\\n--\\n separator and strips HTML', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<p>Best regards,<br>Joel</p>');

      const result = appendSignature('Hello World', 'text', true);

      expect(result).toBe('Hello World\n\n--\nBest regards,\nJoel');
    });

    it('returns body unchanged when includeSignature is false', () => {
      const result = appendSignature('Hello', 'text', false);

      expect(result).toBe('Hello');
      expect(mockExistsSync).not.toHaveBeenCalled();
    });

    it('returns body unchanged when no signature file exists', () => {
      mockExistsSync.mockReturnValue(false);

      const result = appendSignature('Hello', 'html', true);

      expect(result).toBe('Hello');
    });

    it('handles signature with nested HTML tags for text stripping', () => {
      mockExistsSync.mockReturnValue(true);
      mockReadFileSync.mockReturnValue('<div><b>Joel</b> | <a href="https://example.com">Site</a></div>');

      const result = appendSignature('Hi', 'text', true);

      expect(result).toBe('Hi\n\n--\nJoel | Site');
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/signature.test.ts`
Expected: FAIL — `Cannot find module '../../src/signature.js'`

**Step 3: Implement the signature module**

Create `src/signature.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Email signature storage and auto-append logic.
 *
 * Signatures are stored as HTML at ~/.outlook-mcp/signature.html
 * and auto-appended to email bodies when creating/sending emails.
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const SIGNATURE_DIR = join(homedir(), '.outlook-mcp');
const SIGNATURE_FILE = join(SIGNATURE_DIR, 'signature.html');

/**
 * Ensures the signature directory exists.
 */
function ensureDir(): void {
  if (!existsSync(SIGNATURE_DIR)) {
    mkdirSync(SIGNATURE_DIR, { recursive: true, mode: 0o700 });
  }
}

/**
 * Strips HTML tags from a string and converts <br> to newlines.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '');
}

/**
 * Reads the stored signature. Returns null if no signature file exists.
 */
export function readSignature(): string | null {
  if (!existsSync(SIGNATURE_FILE)) return null;
  return readFileSync(SIGNATURE_FILE, 'utf-8');
}

/**
 * Writes a signature to disk.
 * If content_type is 'text', wraps in <pre> tag for HTML storage.
 */
export function writeSignature(content: string, contentType: 'html' | 'text' = 'html'): void {
  ensureDir();
  const html = contentType === 'text' ? `<pre>${content}</pre>` : content;
  writeFileSync(SIGNATURE_FILE, html, { encoding: 'utf-8', mode: 0o600 });
}

/**
 * Appends the stored signature to an email body.
 *
 * For HTML bodies: appends with <br><br> separator.
 * For text bodies: appends with \n\n--\n separator and strips HTML from signature.
 * Returns the body unchanged if includeSignature is false or no signature exists.
 */
export function appendSignature(
  body: string,
  bodyType: 'html' | 'text',
  includeSignature: boolean
): string {
  if (!includeSignature) return body;

  const signature = readSignature();
  if (signature == null) return body;

  if (bodyType === 'html') {
    return `${body}<br><br>${signature}`;
  }

  return `${body}\n\n--\n${stripHtml(signature)}`;
}
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/signature.test.ts`
Expected: All 10 tests PASS

**Step 5: Commit**

```bash
git add src/signature.ts tests/unit/signature.test.ts
git commit -m "feat: add signature storage module with read, write, and append"
```

---

### Task 2: Signature Tools (set_signature, get_signature)

**Files:**
- Modify: `src/tools/mail-send.ts` — add `set_signature`/`get_signature` handlers to MailSendTools
- Modify: `tests/unit/tools/mail-send.test.ts` — add tests for new tools

**Step 1: Write the failing tests**

Add to `tests/unit/tools/mail-send.test.ts`, inside the top-level `describe('mail-send tools', ...)`:

```typescript
// Add these imports at the top of the file:
// import { readSignature, writeSignature } from '../../../src/signature.js';

// Add mock at top level (before imports):
// vi.mock('../../../src/signature.js', () => ({
//   readSignature: vi.fn(),
//   writeSignature: vi.fn(),
//   appendSignature: vi.fn((body: string) => body),
// }));

describe('setSignature', () => {
  it('writes HTML signature and returns success', async () => {
    const result = await tools.setSignature({ content: '<p>Joel</p>', content_type: 'html' });

    expect(result).toEqual({ success: true, message: 'Signature saved successfully.' });
    expect(writeSignature).toHaveBeenCalledWith('<p>Joel</p>', 'html');
  });

  it('writes text signature and returns success', async () => {
    const result = await tools.setSignature({ content: '-- Joel', content_type: 'text' });

    expect(result).toEqual({ success: true, message: 'Signature saved successfully.' });
    expect(writeSignature).toHaveBeenCalledWith('-- Joel', 'text');
  });
});

describe('getSignature', () => {
  it('returns signature content when set', async () => {
    vi.mocked(readSignature).mockReturnValue('<p>-- Joel</p>');

    const result = await tools.getSignature();

    expect(result).toEqual({ has_signature: true, content: '<p>-- Joel</p>' });
  });

  it('returns no-signature message when not set', async () => {
    vi.mocked(readSignature).mockReturnValue(null);

    const result = await tools.getSignature();

    expect(result).toEqual({ has_signature: false, message: 'No signature is set. Use set_signature to create one.' });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: FAIL — `tools.setSignature is not a function`

**Step 3: Implement the signature tools**

Modify `src/tools/mail-send.ts`:

1. Add import at top (after existing imports, around line 32):
```typescript
import { readSignature, writeSignature, appendSignature } from '../signature.js';
```

2. Add Zod schemas (after `ForwardAsDraftInput`, around line 182):
```typescript
export const SetSignatureInput = z.strictObject({
  content: z.string().describe('Signature content (HTML or plain text)'),
  content_type: z.enum(['html', 'text']).default('html').describe('Content type of the signature'),
});

export const GetSignatureInput = z.strictObject({});
```

3. Add methods to `MailSendTools` class (after `forwardAsDraft` method, before the private helpers):
```typescript
  // ---------------------------------------------------------------------------
  // Signature Management
  // ---------------------------------------------------------------------------

  async setSignature(params: z.infer<typeof SetSignatureInput>): Promise<{ success: boolean; message: string }> {
    writeSignature(params.content, params.content_type);
    return { success: true, message: 'Signature saved successfully.' };
  }

  async getSignature(): Promise<{ has_signature: boolean; content?: string; message?: string }> {
    const signature = readSignature();
    if (signature == null) {
      return { has_signature: false, message: 'No signature is set. Use set_signature to create one.' };
    }
    return { has_signature: true, content: signature };
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: All tests PASS (existing + 4 new)

**Step 5: Commit**

```bash
git add src/tools/mail-send.ts tests/unit/tools/mail-send.test.ts
git commit -m "feat: add set_signature and get_signature tools"
```

---

### Task 3: Signature Auto-Append on Email Creation/Send

**Files:**
- Modify: `src/tools/mail-send.ts` — add `include_signature` param to affected Zod schemas, call `appendSignature` in affected methods
- Modify: `tests/unit/tools/mail-send.test.ts` — add tests for auto-append behavior

**Step 1: Write the failing tests**

Add to `tests/unit/tools/mail-send.test.ts`:

```typescript
describe('signature auto-append', () => {
  it('appends signature to body in createDraft when include_signature is true', async () => {
    vi.mocked(appendSignature).mockReturnValue('Hello<br><br><p>-- Joel</p>');
    (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 42, graphId: 'AAA' });

    await tools.createDraft({
      subject: 'Test', body: 'Hello', body_type: 'html',
      include_signature: true,
    });

    expect(appendSignature).toHaveBeenCalledWith('Hello', 'html', true);
    expect(repo.createDraftAsync).toHaveBeenCalledWith(
      expect.objectContaining({ body: 'Hello<br><br><p>-- Joel</p>' })
    );
  });

  it('does not append signature when include_signature is false', async () => {
    vi.mocked(appendSignature).mockReturnValue('Hello');
    (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 42, graphId: 'AAA' });

    await tools.createDraft({
      subject: 'Test', body: 'Hello', body_type: 'text',
      include_signature: false,
    });

    expect(appendSignature).toHaveBeenCalledWith('Hello', 'text', false);
  });

  it('defaults include_signature to true in createDraft', async () => {
    vi.mocked(appendSignature).mockReturnValue('Hello');
    (repo.createDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 42, graphId: 'AAA' });

    await tools.createDraft({ subject: 'Test', body: 'Hello', body_type: 'text' });

    expect(appendSignature).toHaveBeenCalledWith('Hello', 'text', true);
  });

  it('appends signature in prepareSendEmail', () => {
    vi.mocked(appendSignature).mockReturnValue('Hi<br><br><sig>');

    tools.prepareSendEmail({
      to: ['bob@example.com'], subject: 'Test', body: 'Hi', body_type: 'html',
      include_signature: true,
    });

    expect(appendSignature).toHaveBeenCalledWith('Hi', 'html', true);
  });

  it('appends signature in replyAsDraft', async () => {
    vi.mocked(appendSignature).mockReturnValue('reply text with sig');
    (repo.replyAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 10, graphId: 'BBB' });

    await tools.replyAsDraft({ message_id: 1, comment: 'reply text', include_signature: true });

    expect(appendSignature).toHaveBeenCalledWith('reply text', 'text', true);
  });

  it('appends signature in forwardAsDraft', async () => {
    vi.mocked(appendSignature).mockReturnValue('fwd comment with sig');
    (repo.forwardAsDraftAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ numericId: 11, graphId: 'CCC' });

    await tools.forwardAsDraft({ message_id: 1, comment: 'fwd comment', include_signature: true });

    expect(appendSignature).toHaveBeenCalledWith('fwd comment', 'text', true);
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: FAIL — Zod rejects `include_signature` as unknown field (strict schemas)

**Step 3: Add `include_signature` param and call `appendSignature`**

Modify `src/tools/mail-send.ts`:

1. Add `include_signature` to affected Zod schemas:

In `CreateDraftInput` (around line 90), add after `attachments`:
```typescript
  include_signature: z.boolean().default(true).describe('Include email signature (default: true)'),
```

In `PrepareSendEmailInput` (around line 130), add after `attachments`:
```typescript
  include_signature: z.boolean().default(true).describe('Include email signature (default: true)'),
```

In `ReplyAsDraftInput` (around line 172), add after `reply_all`:
```typescript
  include_signature: z.boolean().default(true).describe('Include email signature (default: true)'),
```

In `ForwardAsDraftInput` (around line 178), add after `comment`:
```typescript
  include_signature: z.boolean().default(true).describe('Include email signature (default: true)'),
```

2. Call `appendSignature` in affected methods:

In `createDraft` (around line 283), before calling `this.repository.createDraftAsync`:
```typescript
    const body = appendSignature(params.body, params.body_type, params.include_signature);
```
Then use `body` instead of `params.body` in the call.

In `prepareSendEmail` (around line 355), before `hashDirectSendForApproval`:
```typescript
    const body = appendSignature(params.body, params.body_type, params.include_signature);
```
Then use `body` instead of `params.body` in the hash, token metadata, and preview.

In `replyAsDraft` (around line 553), before calling `this.repository.replyAsDraftAsync`:
```typescript
    const comment = params.comment != null
      ? appendSignature(params.comment, 'text', params.include_signature)
      : params.comment;
```
Then pass `comment` instead of `params.comment`.

In `forwardAsDraft` (around line 570), before calling `this.repository.forwardAsDraftAsync`:
```typescript
    const comment = params.comment != null
      ? appendSignature(params.comment, 'text', params.include_signature)
      : params.comment;
```
Then pass `comment` instead of `params.comment`.

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/tools/mail-send.test.ts`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add src/tools/mail-send.ts tests/unit/tools/mail-send.test.ts
git commit -m "feat: auto-append signature to email creation and send tools"
```

---

### Task 4: GraphClient Scheduling Methods

**Files:**
- Modify: `src/graph/client/graph-client.ts` — add `getSchedule` and `findMeetingTimes` methods
- Modify: `tests/unit/graph/client/graph-client.test.ts` — add tests

**Step 1: Write the failing tests**

Add to `tests/unit/graph/client/graph-client.test.ts`:

```typescript
describe('getSchedule', () => {
  it('calls POST /me/calendar/getSchedule with correct body', async () => {
    const mockResponse = {
      value: [
        {
          scheduleId: 'bob@example.com',
          availabilityView: '0120',
          scheduleItems: [
            { status: 'busy', start: { dateTime: '2026-02-24T10:00:00' }, end: { dateTime: '2026-02-24T11:00:00' } },
          ],
        },
      ],
    };
    const postBuilder = { post: vi.fn().mockResolvedValue(mockResponse) };
    mockApi.mockReturnValue(postBuilder);

    const result = await graphClient.getSchedule({
      schedules: ['bob@example.com'],
      startTime: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
      endTime: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });

    expect(mockApi).toHaveBeenCalledWith('/me/calendar/getSchedule');
    expect(postBuilder.post).toHaveBeenCalledWith({
      schedules: ['bob@example.com'],
      startTime: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
      endTime: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });
    expect(result).toEqual(mockResponse.value);
  });
});

describe('findMeetingTimes', () => {
  it('calls POST /me/findMeetingTimes with correct body', async () => {
    const mockResponse = {
      meetingTimeSuggestions: [
        {
          confidence: 100,
          meetingTimeSlot: {
            start: { dateTime: '2026-02-24T14:00:00', timeZone: 'UTC' },
            end: { dateTime: '2026-02-24T15:00:00', timeZone: 'UTC' },
          },
          attendeeAvailability: [
            { attendee: { emailAddress: { address: 'bob@example.com' } }, availability: 'free' },
          ],
        },
      ],
      emptySuggestionsReason: '',
    };
    const postBuilder = { post: vi.fn().mockResolvedValue(mockResponse) };
    mockApi.mockReturnValue(postBuilder);

    const result = await graphClient.findMeetingTimes({
      attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
      meetingDuration: 'PT1H',
      timeConstraint: {
        timeslots: [{
          start: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
          end: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
        }],
      },
      maxCandidates: 5,
    });

    expect(mockApi).toHaveBeenCalledWith('/me/findMeetingTimes');
    expect(postBuilder.post).toHaveBeenCalledWith({
      attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
      meetingDuration: 'PT1H',
      timeConstraint: {
        timeslots: [{
          start: { dateTime: '2026-02-24T08:00:00', timeZone: 'UTC' },
          end: { dateTime: '2026-02-24T18:00:00', timeZone: 'UTC' },
        }],
      },
      maxCandidates: 5,
    });
    expect(result).toEqual(mockResponse);
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/client/graph-client.test.ts`
Expected: FAIL — `graphClient.getSchedule is not a function`

**Step 3: Implement the methods**

Add to `src/graph/client/graph-client.ts`, after the `createForwardDraft` method (around line 875):

```typescript
  // ---------------------------------------------------------------------------
  // Calendar Scheduling
  // ---------------------------------------------------------------------------

  /**
   * Gets the free/busy schedule for one or more people.
   * POST /me/calendar/getSchedule
   */
  async getSchedule(params: {
    schedules: string[];
    startTime: { dateTime: string; timeZone: string };
    endTime: { dateTime: string; timeZone: string };
    availabilityViewInterval?: number;
  }): Promise<unknown[]> {
    const client = await this.getClient();
    const response = await client.api('/me/calendar/getSchedule').post(params) as { value: unknown[] };
    return response.value;
  }

  /**
   * Suggests meeting times for a set of attendees.
   * POST /me/findMeetingTimes
   */
  async findMeetingTimes(params: {
    attendees: Array<{ emailAddress: { address: string }; type: string }>;
    meetingDuration: string;
    timeConstraint?: {
      timeslots: Array<{
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      }>;
    };
    maxCandidates?: number;
  }): Promise<unknown> {
    const client = await this.getClient();
    return await client.api('/me/findMeetingTimes').post(params);
  }
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/client/graph-client.test.ts`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add src/graph/client/graph-client.ts tests/unit/graph/client/graph-client.test.ts
git commit -m "feat: add getSchedule and findMeetingTimes to GraphClient"
```

---

### Task 5: GraphRepository Scheduling Methods

**Files:**
- Modify: `src/graph/repository.ts` — add `getScheduleAsync` and `findMeetingTimesAsync`
- Modify: `tests/unit/graph/repository.test.ts` — add tests

**Step 1: Write the failing tests**

Add to `tests/unit/graph/repository.test.ts`:

```typescript
describe('getScheduleAsync', () => {
  it('calls client.getSchedule with formatted params and returns result', async () => {
    const mockSchedules = [
      { scheduleId: 'bob@example.com', availabilityView: '0120', scheduleItems: [] },
    ];
    mockClient.getSchedule.mockResolvedValue(mockSchedules);

    const result = await repository.getScheduleAsync({
      emailAddresses: ['bob@example.com'],
      startTime: '2026-02-24T08:00:00Z',
      endTime: '2026-02-24T18:00:00Z',
      availabilityViewInterval: 30,
    });

    expect(mockClient.getSchedule).toHaveBeenCalledWith({
      schedules: ['bob@example.com'],
      startTime: { dateTime: '2026-02-24T08:00:00Z', timeZone: 'UTC' },
      endTime: { dateTime: '2026-02-24T18:00:00Z', timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });
    expect(result).toEqual(mockSchedules);
  });

  it('uses default interval of 30 when not specified', async () => {
    mockClient.getSchedule.mockResolvedValue([]);

    await repository.getScheduleAsync({
      emailAddresses: ['bob@example.com'],
      startTime: '2026-02-24T08:00:00Z',
      endTime: '2026-02-24T18:00:00Z',
    });

    expect(mockClient.getSchedule).toHaveBeenCalledWith(
      expect.objectContaining({ availabilityViewInterval: 30 })
    );
  });
});

describe('findMeetingTimesAsync', () => {
  it('calls client.findMeetingTimes with formatted attendees and duration', async () => {
    const mockResult = {
      meetingTimeSuggestions: [{ confidence: 100 }],
      emptySuggestionsReason: '',
    };
    mockClient.findMeetingTimes.mockResolvedValue(mockResult);

    const result = await repository.findMeetingTimesAsync({
      attendees: ['bob@example.com', 'alice@example.com'],
      durationMinutes: 60,
      startTime: '2026-02-24T08:00:00Z',
      endTime: '2026-02-24T18:00:00Z',
      maxCandidates: 5,
    });

    expect(mockClient.findMeetingTimes).toHaveBeenCalledWith({
      attendees: [
        { emailAddress: { address: 'bob@example.com' }, type: 'required' },
        { emailAddress: { address: 'alice@example.com' }, type: 'required' },
      ],
      meetingDuration: 'PT1H0M',
      timeConstraint: {
        timeslots: [{
          start: { dateTime: '2026-02-24T08:00:00Z', timeZone: 'UTC' },
          end: { dateTime: '2026-02-24T18:00:00Z', timeZone: 'UTC' },
        }],
      },
      maxCandidates: 5,
    });
    expect(result).toEqual(mockResult);
  });

  it('omits timeConstraint when startTime/endTime not provided', async () => {
    mockClient.findMeetingTimes.mockResolvedValue({ meetingTimeSuggestions: [] });

    await repository.findMeetingTimesAsync({
      attendees: ['bob@example.com'],
      durationMinutes: 30,
    });

    expect(mockClient.findMeetingTimes).toHaveBeenCalledWith({
      attendees: [{ emailAddress: { address: 'bob@example.com' }, type: 'required' }],
      meetingDuration: 'PT0H30M',
      maxCandidates: 5,
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/graph/repository.test.ts`
Expected: FAIL — `repository.getScheduleAsync is not a function`

**Step 3: Implement the repository methods**

Add to `src/graph/repository.ts`, after `forwardAsDraftAsync` (around line 940):

```typescript
  // ---------------------------------------------------------------------------
  // Calendar Scheduling
  // ---------------------------------------------------------------------------

  async getScheduleAsync(params: {
    emailAddresses: string[];
    startTime: string;
    endTime: string;
    availabilityViewInterval?: number;
  }): Promise<unknown[]> {
    return await this.client.getSchedule({
      schedules: params.emailAddresses,
      startTime: { dateTime: params.startTime, timeZone: 'UTC' },
      endTime: { dateTime: params.endTime, timeZone: 'UTC' },
      availabilityViewInterval: params.availabilityViewInterval ?? 30,
    });
  }

  async findMeetingTimesAsync(params: {
    attendees: string[];
    durationMinutes: number;
    startTime?: string;
    endTime?: string;
    maxCandidates?: number;
  }): Promise<unknown> {
    const hours = Math.floor(params.durationMinutes / 60);
    const minutes = params.durationMinutes % 60;
    const meetingDuration = `PT${hours}H${minutes}M`;

    const attendees = params.attendees.map(addr => ({
      emailAddress: { address: addr },
      type: 'required' as const,
    }));

    const request: Record<string, unknown> = {
      attendees,
      meetingDuration,
      maxCandidates: params.maxCandidates ?? 5,
    };

    if (params.startTime != null && params.endTime != null) {
      request['timeConstraint'] = {
        timeslots: [{
          start: { dateTime: params.startTime, timeZone: 'UTC' },
          end: { dateTime: params.endTime, timeZone: 'UTC' },
        }],
      };
    }

    return await this.client.findMeetingTimes(request as Parameters<typeof this.client.findMeetingTimes>[0]);
  }
```

Also add `getSchedule` and `findMeetingTimes` to the mock client in the test setup (inside the `vi.mock` for `GraphClient`).

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/graph/repository.test.ts`
Expected: All tests PASS

**Step 5: Commit**

```bash
git add src/graph/repository.ts tests/unit/graph/repository.test.ts
git commit -m "feat: add getScheduleAsync and findMeetingTimesAsync to GraphRepository"
```

---

### Task 6: Scheduling Tools Module

**Files:**
- Create: `src/tools/scheduling.ts`
- Create: `tests/unit/tools/scheduling.test.ts`

**Step 1: Write the failing tests**

Create `tests/unit/tools/scheduling.test.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { SchedulingTools, createSchedulingTools, type ISchedulingRepository } from '../../../src/tools/scheduling.js';

function createMockRepository(): ISchedulingRepository {
  return {
    getScheduleAsync: vi.fn(),
    findMeetingTimesAsync: vi.fn(),
  };
}

describe('scheduling tools', () => {
  let repo: ISchedulingRepository;
  let tools: SchedulingTools;

  beforeEach(() => {
    repo = createMockRepository();
    tools = new SchedulingTools(repo);
  });

  describe('checkAvailability', () => {
    it('returns schedule data for requested attendees', async () => {
      const mockSchedules = [
        {
          scheduleId: 'bob@example.com',
          availabilityView: '020120',
          scheduleItems: [
            { status: 'busy', start: { dateTime: '2026-02-24T10:00:00' }, end: { dateTime: '2026-02-24T11:00:00' } },
          ],
        },
      ];
      (repo.getScheduleAsync as ReturnType<typeof vi.fn>).mockResolvedValue(mockSchedules);

      const result = await tools.checkAvailability({
        email_addresses: ['bob@example.com'],
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
        availability_view_interval: 30,
      });

      expect(result).toEqual({ schedules: mockSchedules });
      expect(repo.getScheduleAsync).toHaveBeenCalledWith({
        emailAddresses: ['bob@example.com'],
        startTime: '2026-02-24T08:00:00Z',
        endTime: '2026-02-24T18:00:00Z',
        availabilityViewInterval: 30,
      });
    });

    it('uses default interval when not specified', async () => {
      (repo.getScheduleAsync as ReturnType<typeof vi.fn>).mockResolvedValue([]);

      await tools.checkAvailability({
        email_addresses: ['bob@example.com'],
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
      });

      expect(repo.getScheduleAsync).toHaveBeenCalledWith(
        expect.objectContaining({ availabilityViewInterval: 30 })
      );
    });
  });

  describe('findMeetingTimes', () => {
    it('returns meeting time suggestions', async () => {
      const mockResult = {
        meetingTimeSuggestions: [
          { confidence: 100, meetingTimeSlot: { start: {}, end: {} } },
        ],
        emptySuggestionsReason: '',
      };
      (repo.findMeetingTimesAsync as ReturnType<typeof vi.fn>).mockResolvedValue(mockResult);

      const result = await tools.findMeetingTimes({
        attendees: ['bob@example.com', 'alice@example.com'],
        duration_minutes: 60,
        start_time: '2026-02-24T08:00:00Z',
        end_time: '2026-02-24T18:00:00Z',
        max_candidates: 3,
      });

      expect(result).toEqual(mockResult);
      expect(repo.findMeetingTimesAsync).toHaveBeenCalledWith({
        attendees: ['bob@example.com', 'alice@example.com'],
        durationMinutes: 60,
        startTime: '2026-02-24T08:00:00Z',
        endTime: '2026-02-24T18:00:00Z',
        maxCandidates: 3,
      });
    });

    it('uses defaults for optional params', async () => {
      (repo.findMeetingTimesAsync as ReturnType<typeof vi.fn>).mockResolvedValue({ meetingTimeSuggestions: [] });

      await tools.findMeetingTimes({
        attendees: ['bob@example.com'],
        duration_minutes: 30,
      });

      expect(repo.findMeetingTimesAsync).toHaveBeenCalledWith({
        attendees: ['bob@example.com'],
        durationMinutes: 30,
        startTime: undefined,
        endTime: undefined,
        maxCandidates: 5,
      });
    });
  });

  describe('createSchedulingTools', () => {
    it('returns a SchedulingTools instance', () => {
      const result = createSchedulingTools(repo);
      expect(result).toBeInstanceOf(SchedulingTools);
    });
  });
});
```

**Step 2: Run tests to verify they fail**

Run: `npx vitest run tests/unit/tools/scheduling.test.ts`
Expected: FAIL — `Cannot find module '../../../src/tools/scheduling.js'`

**Step 3: Implement the scheduling tools module**

Create `src/tools/scheduling.ts`:

```typescript
/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Calendar scheduling MCP tools.
 *
 * Provides tools for checking free/busy availability and
 * finding optimal meeting times for groups of attendees.
 */

import { z } from 'zod';

// =============================================================================
// Repository Interface
// =============================================================================

export interface ISchedulingRepository {
  getScheduleAsync(params: {
    emailAddresses: string[];
    startTime: string;
    endTime: string;
    availabilityViewInterval?: number;
  }): Promise<unknown[]>;

  findMeetingTimesAsync(params: {
    attendees: string[];
    durationMinutes: number;
    startTime?: string;
    endTime?: string;
    maxCandidates?: number;
  }): Promise<unknown>;
}

// =============================================================================
// Zod Schemas
// =============================================================================

export const CheckAvailabilityInput = z.strictObject({
  email_addresses: z.array(z.string().email()).min(1).describe('Email addresses to check availability for'),
  start_time: z.string().describe('Start of time window (ISO 8601)'),
  end_time: z.string().describe('End of time window (ISO 8601)'),
  availability_view_interval: z.number().int().min(5).max(1440).default(30).describe('Time slot interval in minutes (default: 30)'),
});

export const FindMeetingTimesInput = z.strictObject({
  attendees: z.array(z.string().email()).min(1).describe('Attendee email addresses'),
  duration_minutes: z.number().int().min(1).describe('Meeting duration in minutes'),
  start_time: z.string().optional().describe('Start of search window (ISO 8601)'),
  end_time: z.string().optional().describe('End of search window (ISO 8601)'),
  max_candidates: z.number().int().min(1).max(25).default(5).describe('Max time suggestions to return (default: 5)'),
});

// =============================================================================
// Scheduling Tools Class
// =============================================================================

export class SchedulingTools {
  constructor(private readonly repository: ISchedulingRepository) {}

  async checkAvailability(params: z.infer<typeof CheckAvailabilityInput>): Promise<{ schedules: unknown[] }> {
    const schedules = await this.repository.getScheduleAsync({
      emailAddresses: params.email_addresses,
      startTime: params.start_time,
      endTime: params.end_time,
      availabilityViewInterval: params.availability_view_interval,
    });
    return { schedules };
  }

  async findMeetingTimes(params: z.infer<typeof FindMeetingTimesInput>): Promise<unknown> {
    return await this.repository.findMeetingTimesAsync({
      attendees: params.attendees,
      durationMinutes: params.duration_minutes,
      startTime: params.start_time,
      endTime: params.end_time,
      maxCandidates: params.max_candidates,
    });
  }
}

// =============================================================================
// Factory
// =============================================================================

export function createSchedulingTools(repository: ISchedulingRepository): SchedulingTools {
  return new SchedulingTools(repository);
}
```

**Step 4: Run tests to verify they pass**

Run: `npx vitest run tests/unit/tools/scheduling.test.ts`
Expected: All 6 tests PASS

**Step 5: Commit**

```bash
git add src/tools/scheduling.ts tests/unit/tools/scheduling.test.ts
git commit -m "feat: add scheduling tools module with check_availability and find_meeting_times"
```

---

### Task 7: Wire All 4 New Tools into index.ts

**Files:**
- Modify: `src/index.ts` — add 4 tool definitions, add `include_signature` to 4 existing tool schemas, wire handlers
- Modify: `tests/e2e/mcp-client.test.ts` — update tool count 74 → 78

**Step 1: Add tool definitions to TOOLS array**

In `src/index.ts`, add `include_signature` property to these existing tool inputSchemas:
- `create_draft` (line ~1276): add `include_signature: { type: 'boolean', default: true, description: 'Include email signature (default: true)' }` to properties
- `prepare_send_email` (line ~1416): add same property
- `reply_as_draft` (line ~1539): add same property
- `forward_as_draft` (line ~1552): add same property

Then, before the closing `];` of the TOOLS array (around line 1568), add 4 new tool definitions:

```typescript
  // Signature tools
  {
    name: 'set_signature',
    description: 'Save an email signature that will be auto-appended to outgoing emails',
    inputSchema: {
      type: 'object' as const,
      properties: {
        content: { type: 'string', description: 'Signature content (HTML or plain text)' },
        content_type: {
          type: 'string',
          enum: ['html', 'text'],
          default: 'html',
          description: 'Content type of the signature (default: html)',
        },
      },
      required: ['content'],
    },
  },
  {
    name: 'get_signature',
    description: 'Get the currently stored email signature',
    inputSchema: {
      type: 'object' as const,
      properties: {},
      required: [],
    },
  },
  // Scheduling tools
  {
    name: 'check_availability',
    description: 'Check free/busy availability for one or more people in a time window',
    inputSchema: {
      type: 'object' as const,
      properties: {
        email_addresses: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Email addresses to check availability for',
        },
        start_time: { type: 'string', description: 'Start of time window (ISO 8601)' },
        end_time: { type: 'string', description: 'End of time window (ISO 8601)' },
        availability_view_interval: {
          type: 'number',
          default: 30,
          description: 'Time slot interval in minutes (default: 30)',
        },
      },
      required: ['email_addresses', 'start_time', 'end_time'],
    },
  },
  {
    name: 'find_meeting_times',
    description: 'Find available meeting time slots for a group of attendees',
    inputSchema: {
      type: 'object' as const,
      properties: {
        attendees: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Attendee email addresses',
        },
        duration_minutes: { type: 'number', description: 'Meeting duration in minutes' },
        start_time: { type: 'string', description: 'Start of search window (ISO 8601, optional)' },
        end_time: { type: 'string', description: 'End of search window (ISO 8601, optional)' },
        max_candidates: {
          type: 'number',
          default: 5,
          description: 'Maximum number of time suggestions (default: 5)',
        },
      },
      required: ['attendees', 'duration_minutes'],
    },
  },
```

**Step 2: Add imports and state variable**

At the top of `src/index.ts`, add imports:
```typescript
import { createSchedulingTools, CheckAvailabilityInput, FindMeetingTimesInput, SetSignatureInput, GetSignatureInput } from './tools/scheduling.js';
```

Wait — the signature schemas are in `mail-send.ts`. Make sure those are imported:
```typescript
import { ..., SetSignatureInput, GetSignatureInput } from './tools/mail-send.js';
```

And import scheduling:
```typescript
import { createSchedulingTools, CheckAvailabilityInput, FindMeetingTimesInput } from './tools/scheduling.js';
```

Add state variable (around line 1605, alongside `sendTools`):
```typescript
let schedulingTools: ReturnType<typeof createSchedulingTools> | null = null;
```

Initialize in `initializeGraphBackend` (around line 1655):
```typescript
schedulingTools = createSchedulingTools(graphRepository);
```

**Step 3: Add handler function and wire into dispatcher**

Add after `handleSendToolCall` (around line 1989):

```typescript
async function handleSchedulingToolCall(
  name: string,
  args: unknown,
  schedulingTools: ReturnType<typeof createSchedulingTools>
): Promise<ToolResult | null> {
  switch (name) {
    case 'check_availability': {
      const params = CheckAvailabilityInput.parse(args);
      const result = await schedulingTools.checkAvailability(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'find_meeting_times': {
      const params = FindMeetingTimesInput.parse(args);
      const result = await schedulingTools.findMeetingTimes(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    default:
      return null;
  }
}
```

Add `set_signature` and `get_signature` cases to `handleSendToolCall` (before the `default: return null`):
```typescript
    case 'set_signature': {
      const params = SetSignatureInput.parse(args);
      const result = await sendTools.setSignature(params);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
    case 'get_signature': {
      const result = await sendTools.getSignature();
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
```

In `handleGraphToolCall`, add scheduling dispatch after sendResult (around line 2603):
```typescript
  const schedulingResult = await handleSchedulingToolCall(name, args, schedulingTools!);
  if (schedulingResult != null) return schedulingResult;
```

Pass `schedulingTools` through `handleGraphToolCall` signature:
```typescript
async function handleGraphToolCall(
  name: string,
  args: unknown,
  repository: GraphRepository,
  contentReaders: GraphContentReaders,
  orgTools: ReturnType<typeof createMailboxOrganizationTools>,
  sendTools: ReturnType<typeof createMailSendTools>,
  schedulingTools: ReturnType<typeof createSchedulingTools>,
  tokenManager: ApprovalTokenManager
): Promise<ToolResult> {
```

Update the call site in the `CallToolRequestSchema` handler to pass `schedulingTools!`.

**Step 4: Update E2E tool count**

In `tests/e2e/mcp-client.test.ts` line 49, change:
```typescript
expect(result.tools.length).toBe(78);
```

Add tool name checks:
```typescript
expect(toolNames).toContain('set_signature');
expect(toolNames).toContain('get_signature');
expect(toolNames).toContain('check_availability');
expect(toolNames).toContain('find_meeting_times');
```

**Step 5: Run the full test suite**

Run: `npx vitest run`
Expected: All tests PASS, 78 tools

**Step 6: Commit**

```bash
git add src/index.ts src/tools/scheduling.ts tests/e2e/mcp-client.test.ts
git commit -m "feat: wire 4 new tools (signatures + scheduling) into MCP server"
```

---

### Task 8: Export Barrel Updates

**Files:**
- Modify: `src/tools/index.ts` — re-export scheduling module

**Step 1: Add export**

In `src/tools/index.ts`, add:
```typescript
export * from './scheduling.js';
```

**Step 2: Run tests**

Run: `npx vitest run`
Expected: All tests PASS

**Step 3: Commit**

```bash
git add src/tools/index.ts
git commit -m "chore: re-export scheduling tools from barrel"
```

---

### Summary

| Task | Description | New Tests | Files |
|------|-------------|-----------|-------|
| 1 | Signature storage module | ~10 | `src/signature.ts`, `tests/unit/signature.test.ts` |
| 2 | set_signature + get_signature tools | ~4 | `src/tools/mail-send.ts`, `tests/unit/tools/mail-send.test.ts` |
| 3 | Signature auto-append on create/send | ~6 | `src/tools/mail-send.ts`, `tests/unit/tools/mail-send.test.ts` |
| 4 | GraphClient scheduling methods | ~2 | `src/graph/client/graph-client.ts`, `tests/unit/graph/client/graph-client.test.ts` |
| 5 | GraphRepository scheduling methods | ~4 | `src/graph/repository.ts`, `tests/unit/graph/repository.test.ts` |
| 6 | Scheduling tools module | ~6 | `src/tools/scheduling.ts`, `tests/unit/tools/scheduling.test.ts` |
| 7 | Wire all 4 tools into index.ts | ~1 (E2E) | `src/index.ts`, `tests/e2e/mcp-client.test.ts` |
| 8 | Export barrel update | 0 | `src/tools/index.ts` |
| **Total** | | **~33** | |
