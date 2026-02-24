# Design: Reply/Forward as Draft

**Date:** 2026-02-24
**Status:** Approved

## Problem

The current `prepare_reply_email` and `prepare_forward_email` tools send immediately (via two-phase approval). There is no way to create a reply or forward as an editable draft that can be modified before sending.

## Solution

Add two new non-destructive tools that create drafts using the Microsoft Graph API's native `createReply`, `createReplyAll`, and `createForward` endpoints:

- `reply_as_draft` — creates a reply (or reply-all) draft
- `forward_as_draft` — creates a forward draft

The returned `draft_id` plugs into the existing `update_draft` / `send_draft` / `prepare_send_email` flow.

## New Tools

### `reply_as_draft`

No two-phase approval (drafts are non-destructive).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `message_id` | `number` | yes | Message to reply to |
| `comment` | `string` | no | Initial reply body |
| `reply_all` | `boolean` | no (default: false) | Reply to all recipients |

Returns: `{ draft_id: number, subject: string, to: string[] }`

### `forward_as_draft`

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `message_id` | `number` | yes | Message to forward |
| `to_recipients` | `string[]` | no | Forward recipients (can add later via `update_draft`) |
| `comment` | `string` | no | Initial forward body |

Returns: `{ draft_id: number, subject: string, to: string[] }`

## Graph API Endpoints

| Operation | Endpoint | Response |
|-----------|----------|----------|
| Reply draft | `POST /me/messages/{id}/createReply` | `Message` with `isDraft: true` |
| Reply-all draft | `POST /me/messages/{id}/createReplyAll` | `Message` with `isDraft: true` |
| Forward draft | `POST /me/messages/{id}/createForward` | `Message` with `isDraft: true` |

All endpoints accept an empty body (no request body required). The returned draft includes the original message's quoted content, proper subject prefix (Re:/Fwd:), and correct recipients.

## Layer Changes

### GraphClient (3 new methods)

- `createReplyDraft(messageId: string): Promise<Message>` — `POST /me/messages/{id}/createReply`
- `createReplyAllDraft(messageId: string): Promise<Message>` — `POST /me/messages/{id}/createReplyAll`
- `createForwardDraft(messageId: string): Promise<Message>` — `POST /me/messages/{id}/createForward`

### GraphRepository (2 new methods)

- `replyAsDraftAsync(messageId: number, comment?: string, replyAll?: boolean): Promise<{ numericId: number; graphId: string }>`
  - Looks up graph ID from idCache.messages
  - Calls `createReplyDraft` or `createReplyAllDraft`
  - If comment provided, calls `updateDraft` to set body
  - Caches returned draft in idCache.messages
  - Returns numeric + graph IDs

- `forwardAsDraftAsync(messageId: number, toRecipients?: string[], comment?: string): Promise<{ numericId: number; graphId: string }>`
  - Looks up graph ID from idCache.messages
  - Calls `createForwardDraft`
  - If toRecipients or comment provided, calls `updateDraft` to set them
  - Caches returned draft in idCache.messages
  - Returns numeric + graph IDs

### MailSendTools (2 new handler methods)

- `replyAsDraft(params)` — orchestrates repository call, returns draft info
- `forwardAsDraft(params)` — orchestrates repository call, returns draft info

### index.ts

- 2 tool definitions in TOOLS array
- 2 Zod schemas (`ReplyAsDraftInput`, `ForwardAsDraftInput`)
- 2 handler cases in `handleGraphToolCall`

## User Flow

```
reply_as_draft(message_id=42, reply_all=true)
  → draft_id: 99, subject: "Re: Project Update", to: ["alice@co.com", "bob@co.com"]

update_draft(draft_id=99, body="Thanks for the update!\n\nI'll review by Friday.")
  → success

prepare_send_email(draft_id=99)
  → token_id, preview...

confirm_send_email(token_id=...)
  → sent
```

Or the user can simply leave the draft in Outlook for manual editing/sending.

## No Two-Phase Approval

Creating a draft is non-destructive — it just adds an editable message to the Drafts folder. The existing `prepare_send_email` / `confirm_send_email` flow handles the destructive send action.

## Test Plan

- GraphClient: endpoint + HTTP method assertions for all 3 new methods
- Repository: idCache management, comment/recipient handling, error cases
- MailSendTools: orchestration tests
- E2E: tool count update (72 → 74)
