---
title: Graph createReply/createForward — message.body silently discards the quoted thread
date: 2026-08-10
category: integration-issues
module: graph/mail
problem_type: integration_issue
component: email_processing
symptoms:
  - Reply and forward drafts open with no conversation history below the cursor
  - "Draft body collapses from ~65 KB to ~131 bytes; no `divRplyFwdMsg`, no `<hr>`, no `From:` header"
  - Editing the draft afterwards with update_draft never brings the thread back
root_cause: wrong_api
resolution_type: code_fix
severity: high
related_components:
  - testing_framework
tags:
  - microsoft-graph
  - createreply
  - createforward
  - quoted-thread
  - reply-as-draft
  - mocked-tests
  - live-probe
---

# Graph createReply/createForward — message.body silently discards the quoted thread

## Problem

`reply_as_draft` and `forward_as_draft` produced drafts containing only the new comment — the quoted original message was gone before the user ever opened the draft. Shipped broken twice: v2.5.4 (`763ea53`) introduced it while attempting to *fix* the same symptom, and it survived until PR #105.

## Symptoms

- Reply/forward drafts have no conversation history
- Draft body is ~131 bytes instead of ~65 KB; no `<div id="divRplyFwdMsg">`, no `<hr>`, no `From:`/`Sent:`/`To:` header block
- A later `update_draft` can't recover it — its quote-preservation logic finds no marker to anchor on, so it plain-replaces the body

## What Didn't Work

- **Passing the body as `message.body` on `createReply`** (the v2.5.4 approach). This was adopted *because* the previous approach — creating a bare draft and then `PATCH`ing `body` — was known to wipe the quote. Both fail for the same underlying reason; `message` on a create-action is applied as a property overwrite on the draft Graph just generated, so it is equivalent to the PATCH it replaced.
- **Reading the Microsoft Learn page for `createReply`.** It documents "Specify either a comment or the **body** property of the `message` parameter" but never states that `message.body` *replaces* the generated quoted body. The doc reads as if the two are interchangeable. They are not.

## Solution

Route the body through Graph's `comment` parameter. Graph inserts it into the generated HTML body immediately after `<body>`, above the `<hr>` and the quote block.

```ts
// src/graph/client/graph-client.ts — before
async createReplyDraft(messageId: string, comment?: string, body?: { contentType: string; content: string }) {
  const postBody: Record<string, unknown> = {};
  if (comment != null) postBody.comment = comment;
  if (body != null) postBody.message = { body };   // <-- wipes the quoted thread
  ...
}

// after — `body` param removed entirely so it cannot be re-armed
async createReplyDraft(messageId: string, comment?: string) {
  const result = await client
    .api(`/me/messages/${messageId}/createReply`)
    .post(comment != null ? { comment } : null) as MicrosoftGraph.Message;
  ...
}
```

For forwards, the recipients still need a `PATCH` — but only the recipients:

```ts
// src/graph/repository.ts
const draft = await this.client.createForwardDraft(graphMessageId, toDraftComment(comment, bodyType));
if (toRecipients != null && toRecipients.length > 0) {
  await this.client.updateDraft(graphId, {
    toRecipients: toRecipients.map(addr => ({ emailAddress: { address: addr } })),
  });   // never `body` here
}
```

`comment` lands in an HTML body, so a plain-text comment is wrapped to keep its line breaks:

```ts
function toDraftComment(comment: string | undefined, bodyType: string): string | undefined {
  if (comment == null) return undefined;
  return bodyType === 'html' ? comment : `<pre>${comment}</pre>`;
}
```

## Why This Works

Graph's create-actions (`createReply`, `createReplyAll`, `createForward`) generate a complete draft body containing the quoted original. The `comment` parameter is *composed into* that generated body. The `message` parameter is a set of writable properties **applied on top of** the generated draft — so `message.body` overwrites the whole thing. A post-creation `PATCH /messages/{id}` with `body` does exactly the same. Only `comment` composes; everything else replaces.

Two behaviors measured against live Graph that aren't in the docs:

- **`comment` does not HTML-escape.** `<b>x</b>` arrives raw and renders. This is what makes it safe for the HTML-signature path — the escaping concern that would otherwise push you back toward `message.body` is unfounded.
- **`appendonsend` never appears in Graph-created drafts.** That id is an OWA-composer artifact. The real marker on API-created drafts is `divRplyFwdMsg`, with an `<hr>` immediately preceding it.

An aggravating factor worth noting for any similar tool: `include_signature` defaults to `true`, which made `comment` non-null on *every* call (`params.comment ?? ''` then signature-appended). So a bug that looked conditional actually fired 100% of the time. When auditing a "sometimes" bug, check whether a default makes the guarded branch unconditional.

## Prevention

**The real lesson: mocked unit tests cannot catch a wrong belief about a third-party API.** The existing tests asserted the call shape against a mocked client:

```ts
expect(mockClient.createReplyDraft).toHaveBeenCalledWith('msg-comment', undefined, {
  contentType: 'text', content: 'Thanks for sharing!',
});
```

That assertion passes whether or not `message.body` destroys the quote — it only proves the code calls the client the way the test author *believed* was correct. The belief was the bug, so the test ratified it. This is why v2.5.4 shipped inverted with a green suite.

Concrete guardrails:

1. **Probe the real API before encoding a belief about it.** A throwaway script against a real mailbox settles in minutes what docs leave ambiguous. Create scratch drafts, inspect the returned body, delete them — nothing is sent, and it is fully reversible:

   ```js
   const bare  = await c.createReplyDraft(src.id);
   const withBody = await c.createReplyDraft(src.id, undefined, { contentType: 'html', content: '<p>x</p>' });
   console.log(bare.body.content.length, withBody.body.content.length);  // 65313 vs 131
   ```

2. **Assert the negative on API call shape**, so a future refactor can't quietly reintroduce it:

   ```ts
   expect(apiCalls[0].body).toEqual({ comment: '<p>Reply text</p>' });
   expect(apiCalls[0].body).not.toHaveProperty('message');
   ```

3. **Delete the dangerous parameter** rather than documenting around it. `createReplyDraft`'s `body` argument existed only to be misused; removing it makes the mistake unrepresentable.

4. **Encode the *why* at the call site**, not just in a changelog. The client methods now carry a comment explaining the overwrite semantics, so the next person tempted to "just set the body" is stopped where they'd make the change.

## Related Issues

- PR jbctechsolutions/mcp-office365#105 — the fix
- `763ea53` (v2.5.4) — the inverted fix this replaces
- Microsoft Learn: [message: createReply](https://learn.microsoft.com/en-us/graph/api/message-createreply?view=graph-rest-1.0) — documents the either/or constraint, not the overwrite behavior
