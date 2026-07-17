# Using the Microsoft 365 connector in Claude (JP staff guide)

This connector lets Claude (on claude.ai) work with your Joshua Project Microsoft
365 account — your email, calendar, contacts, Teams, tasks, files, and more — on
your behalf. You ask Claude in plain language ("what meetings do I have Thursday?",
"draft a reply to Priya"), and it uses your M365 account to do it.

You don't need to know anything about "MCP" or servers. This guide is everything
you need.

---

## Before you start

- You need a **Joshua Project** Microsoft 365 account (your `@…` JP work account).
  Personal or guest accounts won't work — the connector only serves assigned JP
  staff.
- Someone (Joel / IT) must have **added you to the connector** first. If sign-in
  fails with "not authorized", that assignment hasn't happened yet — ask to be
  added.
- Use it in a browser signed into claude.ai.

---

## One-time setup — add the connector

1. In claude.ai, open **Settings → Connectors** (or "Custom connectors").
2. Choose **Add custom connector**.
3. Enter the connector URL you were given (it looks like
   `https://mcp-o365.jp.example.org/mcp`). Add it.
4. Claude will send you to a **Microsoft sign-in** page. Sign in with your **JP
   work account** and approve the requested access.
5. When it returns to claude.ai and shows the connector as connected, you're done.

You only do this once. After that, Claude can use your M365 account whenever you
ask.

---

## Using it

Just ask Claude naturally. Examples:

- "Summarize my unread email from today."
- "What's on my calendar tomorrow afternoon?"
- "Draft a reply to the last message from Daniel saying I'll have the numbers by Friday."
- "Create a task to follow up with the Nepal team next Tuesday."

Claude will tell you what it's doing and show you results.

### Keep the approval prompts ON (important)

For anything that **changes** something — sending an email, deleting a message,
moving items — Claude will ask you to **approve** before it acts. **Leave these
prompts on.** They're your chance to catch a mistake before an email actually
goes out from your account. Read the "send" / "delete" confirmations before you
approve them. Every write action is also recorded in an audit log JP admins can
review.

---

## What the errors mean

| What you see | What's happening | What to do |
|--------------|------------------|------------|
| **Sign-in succeeds, but the connector then shows "failed" / "not authorized"** | Your account reached the server but isn't an assigned JP **member** (e.g. it's a guest account, or you haven't been assigned yet). | Ask Joel / IT to assign your JP account to the connector. Make sure you signed in with your JP work account, not a personal or guest one. |
| **Claude says it needs you to reconnect / sign in again** | Your sign-in session expired (this is normal after a while). | Reconnect the connector (repeat "add / sign in"). Nothing is lost. |
| **Sign-in loops or is blocked before you can approve** | A Conditional Access / security policy is blocking the connector. | Tell IT — the connector may need a Conditional Access exemption. Not something you can fix yourself. |
| **"Service busy" / try-again errors** | The M365 service is rate-limiting (all pilot users share one connection). | Wait a moment and retry. If it's persistent, flag it — it's a capacity signal for the pilot. |
| **A large file upload/download times out** | Very large transfers can exceed Claude's tool time limit. | Break it up or use OneDrive/Outlook directly for very large files. |

---

## Your privacy & what's recorded

- Claude acts **as you** — it can only see and do what your own M365 account can.
- **Write actions** (send, delete, move, upload) are recorded in an audit log
  (who, which action, when) that JP admins can review. The **content** of your
  messages is **not** stored in that log — only that an action happened.
- If you leave JP or lose access, an admin can revoke the connector server-side;
  removing it from your own claude.ai settings does **not** by itself clear
  server access, so tell an admin.

Questions or something not working? Contact Joel / JP IT.
