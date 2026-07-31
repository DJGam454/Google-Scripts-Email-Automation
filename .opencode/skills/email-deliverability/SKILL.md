---
name: email-deliverability
description: Cold-email deliverability, sending limits, and sender-reputation safety for this outreach automation. Use when working on DAILY_LIMIT enforcement, random send delays, Gemini rate limiting, TEST_MODE / FAST_TEST_MODE, bounce detection, threaded follow-ups (In-Reply-To/References), working-day and sending-window scheduling, spam-filter avoidance, or anything that could affect Gmail deliverability.
---

# Email Deliverability & Sending Safety

## Mission

This system sends unsolicited B2B cold email at scale. Every line of code that
touches sending must protect the sender domain's reputation and inbox
placement. When in doubt, choose the slower, safer path.

## Existing safety mechanisms (never weaken these)

### 1. Daily sending limit
- `canSendEmail()` (Validation.js) enforces `DAILY_LIMIT` from the Config
  sheet by counting `SUCCESS` email actions in ActivityLog for today
  (`getEmailsSentToday()`).
- Called **before** every send (`processEmails`, `sendFollowUp`). When the
  limit is reached, `processEmails` returns immediately — do not change that
  to "try anyway".
- Both `EMAIL_1` and follow-ups count toward the same daily budget.

### 2. Human-like pacing
- `_getRandomDelay(3, 4)` inserts a **random 3–4 second** sleep between sends.
  Jitter (not fixed delays) is deliberate — it mimics human sending patterns.
- Follow-up sends add an additional random delay before dispatch.
- Do not remove, shorten to a constant, or replace with `Utilities.sleep(0)`
  unless `FAST_TEST_MODE` explicitly handles it (see below).

### 3. Rate limiting external AI calls
- `_waitForAIRateLimit()` (AI.js) throttles Gemini calls using `AI_RPM_LIMIT`
  (default 10 rpm) with jitter. Bypassing it risks 429s and API bans.
- `callGemini` and `callGeminiWithUrlContext` implement exponential backoff
  with jitter on 429/5xx (max 3 retries) and deliberately do **not** retry
  4xx client errors or JSON parse failures.

### 4. Test modes (your safety valve)
- `TEST_MODE` + `TEST_EMAIL`: every send redirects to the test address via
  `getActualRecipient()` (Modes.js). Thread search uses the actual recipient
  so test runs still record correct thread IDs.
- `FAST_TEST_MODE`: bypasses working-day/window checks and uses minute-based
  follow-up delays (`FOLLOWUP_n_MINUTES`) and `NEXT_SERVICE_MINUTES`.
- `PREVIEW_MODE`: preview-only flows (`isPreviewMode`, `generatePreview`).
- Always use a test mode before any real run; never send production mail
  without verifying a test-mode pass first.

### 5. Schedule discipline
- `isAutomationScheduleAllowed()` (Scheduling.js): working days
  (`WORKING_DAYS`, default Mon–Fri) and a sending window
  (`SEND_START_HOUR`–`SEND_END_HOUR`) in the configured `TIMEZONE`
  (default `Europe/London`; this project's manifest timezone is
  `Asia/Kolkata`).
- Production runs outside the window are blocked; `FAST_TEST_MODE` bypasses
  these checks. Do not add a "force send" flag.

## Bounce handling

- `checkBounces()` (Bounces.js) runs **first** in `runAutomation()`: it
  searches Gmail for delivery-failure threads from the last 2 days
  (`from:mailer-daemon` / "Delivery Status Notification" / "Delivery
  incomplete"), extracts the failing address with `extractBouncedEmail()`,
  and marks matching leads `INVALID` in the Leads sheet with an
  `EMAIL_BOUNCED` activity entry.
- `extractBouncedEmail()` matches standard DSN patterns (`Final-Recipient:`,
  `Original-Recipient:`, `Recipient:`, "Your message wasn't delivered to").
- Leads already `INVALID` are skipped (idempotent).
- Never send to a lead flagged `INVALID` — it poisons sender reputation.
- Bounce checks search the **last 2 days**; if bounce volume is high,
  hard-bounce rates above ~5% indicate a list-quality problem that code
  cannot fix — surface it in logs rather than masking it.

## Threaded follow-ups (engagement signal)

- Follow-ups reuse the original Gmail conversation via
  `Gmail.Users.Threads.get` + `Gmail.Users.Messages.send` with
  `In-Reply-To` and `References` headers taken from the **latest** message
  (`sendThreadedFollowUp`, EmailEngine.js).
- Threading matters for deliverability: replies inside a real conversation
  thread are treated as expected conversation traffic rather than new cold
  mail.
- Every `EMAIL_1` send stores its thread ID in Leads col P; follow-ups
  hard-fail (`FOLLOWUP_THREAD_MISSING`) if the thread ID is absent rather
  than sending a broken follow-up.
- `hasDuplicateCampaign()` prevents sending `EMAIL_1` for the same
  email+service from a different Lead ID, but intentionally permits
  Website → SEO → Google Ads progression for the same lead.

## Content factors (spam filtering)

- Every email is personalised (`{{PersonalisedIntro}}`); identical mass-send
  copy is a top spam signal. Never bulk-send the same body verbatim. Emails
  are HTML with a plain-text fallback (multipart/alternative on the threaded
  path); the HTML is sparse and text-like by design (see html-email skill).
- Follow-up subjects: the selected variant's subject is used when non-empty;
  otherwise the thread subject is reused — threading is preserved in both
  cases, and subjects are never randomized independently of their template.
- No attachments, no tracking pixels, no read receipts, no embedded images
  beyond the optional hero/portfolio images on the template.
- Sender identity is explicit: `SENDER_NAME` display name and `REPLY_TO_EMAIL`
  are set from Config on every `EMAIL_1` send. Never send with an empty or
  misleading sender.
- Greetings come from the template (`Hi {{FirstName}}`); the AI intro must
  not repeat greetings, names, or pitch content (enforced in `AI.js`).

## Infrastructure-level factors (document, don't code)

The codebase cannot control these, but any deliverability discussion must
acknowledge them:

- **SPF, DKIM, DMARC** must be configured for the sending domain (e.g.
  `Google Workspace` or the domain of the Gmail account). `GmailApp.sendEmail`
  sends from the authenticated account — the account's domain records apply.
- **Sending account warming**: new domains/accounts should start far below
  the daily limit and ramp up over weeks. `DAILY_LIMIT` is the knob; when
  changing it, prefer conservative values (e.g. 20–50/day for a cold domain).
- **Reply monitoring**: `REPLIED` status stops automation. Inbox placement
  should be monitored; a spike in spam-folder placement means pacing/content
  must change.
- **List hygiene**: invalid/mistyped addresses (marked `INVALID` via
  `isValidEmail`) are dropped before sending. Prefer verified lists.

## What not to do

- Do not parallelise sends (no `UrlFetchApp.fetchAll` batching for Gmail,
  no multi-threaded sends) — GAS is single-threaded by design, and
  concurrency would defeat pacing.
- Do not reduce or remove the 3–4 s jitter delays in production paths.
- Do not bypass `canSendEmail()`, bounce checks, thread-ID checks, or
  stop-statuses (`REPLIED`/`INVALID`/`DO_NOT_CONTACT`).
- Do not add retries that double-send: sends are logged once, and failure
  paths log `FAILED` without auto-retry (manual review is safer than
  accidental duplicates).
- Do not send to `TEST_EMAIL` in production configs — verify `TEST_MODE` is
  off before real runs.
