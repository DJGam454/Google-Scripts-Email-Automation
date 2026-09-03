# Google Scripts Email Automation

An outreach automation system built with **Google Apps Script + Google Sheets + Gmail**.

It manages:
- new cold outreach sends
- timed follow-ups
- bounced email handling
- campaign progression across multiple services
- activity logging and dashboard reporting

---

## What this application does

This project automates a multi-step outbound email workflow using your spreadsheet as the control center.

For each lead in the `Leads` sheet, the system can:

1. Validate the recipient email
2. Generate personalization text (AI module)
3. Send **Email 1** from a service-specific template
4. Track Gmail thread IDs
5. Send follow-ups (`FOLLOWUP_1` to `FOLLOWUP_3`) in the same thread
6. Stop sends for leads marked replied / invalid / do-not-contact
7. Detect bounced emails and mark leads invalid
8. Move no-reply leads to the next service campaign (if configured)
9. Write all actions into `ActivityLog`
10. Surface metrics in `Dashboard`

Main orchestration entrypoint:
- `runAutomation()` in `EmailEngine.js`

Execution flow inside `runAutomation()`:
1. `checkBounces()` — classify DSNs and suppress hard bounces
2. `checkReplies()` — detect human replies (skippable via `SKIP_REPLIES_IN_MAIN_RUN`)
3. `generateMissingPersonalizations()` — AI intros for NEW leads
4. `processFollowUps()` — threaded follow-ups
5. `processEmails()` — email 1 for NEW leads

Every stage boundary checks the 299-second execution self-cap so the script
lock is always released cleanly.

---

## Core architecture (code map)

- `EmailEngine.js`  
  Sending engine, follow-up logic, automation runner, delay handling,
  thread-aware follow-ups, per-run send budget (`MAX_SENDS_PER_RUN`),
  new-lead floor, execution self-cap, suppression gates.

- `HtmlEmailEngine.js`  
  Renders HTML + plain-text emails, component-based layout, theme support,
  MIME construction for Gmail API sends (From/Reply-To and RFC 8058
  `List-Unsubscribe` headers).

- `Templates.js`  
  Template lookup, active variant selection, placeholder expansion
  (including the signed `{{UnsubscribeLink}}`).

- `Leads.js`  
  Canonical lead column map (A–V), lead object builder, statuses and
  follow-up step mapping, plus the suppression layer (`SuppressionList`
  sheet, `isSuppressed`, `addToSuppressionList`, `ensureLeadsBounceColumns`).

- `Replies.js`  
  Automatic human-reply detection: thread scan via the Gmail API plus a
  mailbox fallback for replies that start a new thread; filters
  auto-replies/bounces so only human replies set `REPLIED`.

- `Unsubscribe.js`  
  One-click unsubscribe web app (RFC 8058) with HMAC-signed per-lead links;
  marks leads `DO_NOT_CONTACT` and adds them to the SuppressionList.

- `Bounces.js`  
  DSN classification into `HARD_BOUNCE` / `SOFT_BOUNCE` /
  `POLICY_REJECTION` / `UNKNOWN`, suppression writes, soft-retry cap,
  message-ID dedup.

- `SuppressionReconcile.js`  
  One-time 30-day mailbox bounce backfill and historical suppression audit;
  config-driven domain blocking (`SUPPRESSED_DOMAINS`).

- `VerificationCleanup.js`  
  Applies a bulk email-verification export (undeliverable → HARD_BOUNCE,
  risky → UNKNOWN). Lists are empty placeholders — no lead data committed.

- `TestModeFix.js`  
  One-time helpers: test-date reconciliation against Gmail, ActivityLog
  burst purge, local-time display column.

- `Campaigns.js`  
  Service ordering, next-service movement, campaign completion transitions.

- `Validation.js`  
  Email format validation, daily sent count, daily-limit gate, hourly and
  per-domain caps, Gmail quota circuit breaker.

- `Modes.js`  
  Mode parsing (`TEST_MODE`, `FAST_TEST_MODE`, `PREVIEW_MODE`) and recipient redirection.

- `Scheduling.js`  
  Working day + sending window checks with timezone support (minute
  precision via `SEND_WINDOWS`).

- `Dashboard.js`  
  Dashboard setup, KPI formulas, preview generation, activity logging.

- `Config.js`  
  Reads configuration from `Config` and `AIConfig` sheets (cached); one-click
  hardening migration (`ensureHardeningMigration`).

- `appsscript.json`  
  Apps Script manifest with Gmail advanced service enabled (`gmail` v1).

---

## Prerequisites

### Required accounts

1. **Google account** with Gmail access
2. Access to **Google Sheets**
3. Access to **Google Apps Script**

### Recommended account setup

For best reliability and sender reputation, use:
- a dedicated sender account (or mailbox alias) for outreach
- a business domain mailbox (preferred over a brand-new personal mailbox)
- warmed-up mailbox behavior before high volume sending

### Required Google permissions

When you first run the script, authorize scopes for:
- Spreadsheet read/write
- Gmail send/read/search access
- Script services (cache, lock, utilities)

### Script/Repo tooling (optional but recommended)

- Node.js + npm (if syncing locally)
- `@google/clasp` for push/pull Apps Script files (`.clasp.json` already exists in repo)

---

## Spreadsheet requirements

The automation expects these sheets (exact names):

- `Leads`
- `Templates`
- `Config`
- `AIConfig`
- `ActivityLog`
- `Services`
- `Dashboard`
- `Preview`
- `Themes` (optional, but supported)
- `SuppressionList` (created automatically by `ensureHardeningMigration()`)

If required sheets are missing, functions throw errors (for example: “Leads sheet not found.”).

---

## Leads sheet contract

The lead schema is defined in `Leads.js` (`LEADS_COL`):

- A `Lead ID`
- B `Company`
- C `Name`
- D `Email`
- E `Website`
- F `Industry`
- G `Personalised Intro`
- H `Service Assigned`
- I `Current Campaign`
- J `Email 1 Sent`
- K `Follow-up 1 Sent`
- L `Follow-up 2 Sent`
- M `Follow-up 3 Sent`
- N `Last Email Date`
- O `Status`
- P `Gmail Thread ID`
- Q `Notes`
- R `Last Updated`
- S `Bounce Category` (`HARD_BOUNCE` / `SOFT_BOUNCE` / `POLICY_REJECTION` / `UNKNOWN` / `UNSUBSCRIBED`)
- T `Bounce Diagnostic`
- U `Suppressed At`
- V `Retry Count`

Columns S–V and the `SuppressionList` sheet are created by running
`ensureHardeningMigration()` once after deploying (append-only; existing
values are never overwritten).

Important statuses include:
- `NEW`
- `EMAIL_1_SENT`
- `FOLLOWUP_1_SENT`
- `FOLLOWUP_2_SENT`
- `FOLLOWUP_3_SENT`
- `REPLIED`
- `INVALID`
- `DO_NOT_CONTACT`
- `COMPLETED`

Automation hard-stop statuses:
- `REPLIED`
- `INVALID`
- `DO_NOT_CONTACT`

---

## Templates sheet contract

Expected columns in `Templates`:

- A Service
- B Campaign Step (`EMAIL_1`, `FOLLOWUP_1`, `FOLLOWUP_2`, `FOLLOWUP_3`)
- C Subject
- D Plain Text Template
- E Variant ID
- F HTML Template
- G CTA Button Text
- H CTA URL
- I Theme
- J Hero Image
- K Portfolio Image
- L Status (`ACTIVE` / `INACTIVE`)
- M Notes

Behavior notes:
- Only `ACTIVE` variants are eligible.
- A random variant is selected per send.
- Subject is tied to the selected variant.

---

## Config sheet requirements

`Config` stores runtime switches and delivery settings. Key fields used by the code include:

- `AUTOMATION_ENABLED`
- `DAILY_LIMIT`
- `HOURLY_LIMIT` (+ per-group overrides `HOURLY_LIMIT_YAHOO`, `HOURLY_LIMIT_OUTLOOK`, `HOURLY_LIMIT_ICLOUD`)
- `MAX_SENDS_PER_RUN`
- `NEW_LEADS_DAILY_FLOOR` (default 8% of `DAILY_LIMIT`, min 2)
- `SUPPRESSED_DOMAINS` (comma-separated domain blocklist)
- `SKIP_REPLIES_IN_MAIN_RUN`
- `UNSUBSCRIBE_URL` (your deployed Unsubscribe.js web app URL)
- `TEST_MODE`
- `TEST_EMAIL`
- `PREVIEW_MODE`
- `FAST_TEST_MODE`
- `TIMEZONE`
- `WORKING_DAYS`
- `SEND_START_HOUR`
- `SEND_END_HOUR`
- `SEND_WINDOWS` (minute-precision ranges, e.g. `9:30-12:25,14:00-17:00` — overrides start/end hours)

Follow-up timing:
- production: `FOLLOWUP_1_DAYS`, `FOLLOWUP_2_DAYS`, `FOLLOWUP_3_DAYS`, `NEXT_SERVICE_DAYS`
- fast test: `FOLLOWUP_1_MINUTES`, `FOLLOWUP_2_MINUTES`, `FOLLOWUP_3_MINUTES`, `NEXT_SERVICE_MINUTES`

Brand / signature placeholders (examples):
- `SENDER_NAME`, `REPLY_TO_EMAIL`, `SENDER_ROLE`
- `COMPANY_NAME`, `COMPANY_WEBSITE`, `COMPANY_EMAIL`, `COMPANY_PHONE`, `COMPANY_ADDRESS`
- `SOCIAL_LINKEDIN`, `SOCIAL_TWITTER`, `SOCIAL_GITHUB`
- `PORTFOLIO_URL`, `UNSUBSCRIBE_MESSAGE`

### AIConfig

`AIConfig` is read separately for AI personalization settings (e.g., `AI_ENABLED`, model/tone settings).

---

## Gmail + Apps Script setup

1. Open the Apps Script project linked to your sheet.
2. Confirm `appsscript.json` includes Gmail advanced service (`gmail` v1).
3. In Apps Script editor, ensure **Gmail API advanced service** is enabled.
4. Run a safe function first (e.g., preview/test function) and grant permissions.
5. Verify emails can be sent from the intended mailbox.

---

## Sending limits and preferred operating limits

This project enforces a configurable `DAILY_LIMIT` through `canSendEmail()`.

### How limit enforcement works
- Sent count is derived from `ActivityLog` rows where:
  - action is one of `EMAIL_1`, `FOLLOWUP_1`, `FOLLOWUP_2`, `FOLLOWUP_3`
  - result is `SUCCESS`
  - timestamp is today
- If sent count reaches `DAILY_LIMIT`, new sends are blocked.

### Additional layers (hardening)
- `HOURLY_LIMIT` — global cap per rolling hour (defaults to `DAILY_LIMIT / 4`).
- `HOURLY_LIMIT_YAHOO` / `HOURLY_LIMIT_OUTLOOK` / `HOURLY_LIMIT_ICLOUD` —
  per-provider caps for domains that defer under bursts.
- `MAX_SENDS_PER_RUN` — per-execution budget so a run always finishes
  before the 6-minute trigger limit.
- Gmail quota circuit breaker — a `quota exceeded` error blocks all sends
  until midnight instead of hammering the API.
- `NEW_LEADS_DAILY_FLOOR` — guarantees new-lead slots (default 8% of the
  daily limit) so fresh leads are not starved by the follow-up queue;
  overdue follow-ups keep priority.

### Preferred limit strategy (recommended)
Use conservative limits to protect sender reputation and avoid account restrictions.

Suggested approach:
- Start low (example: 20–50/day) for new or cold mailboxes
- Increase gradually based on bounce/reply health
- Keep follow-up cadence human and spread through working window
- Avoid sudden volume jumps
- Monitor bounce rates and complaint signals

**Important:** Provider limits vary by account type and trust history. Keep `DAILY_LIMIT` comfortably below your mailbox’s safe threshold.

---

## Test mode, preview mode, and production

### `TEST_MODE = TRUE`
- Actual recipient is redirected to `TEST_EMAIL`
- Lets you validate formatting, templates, threading logic safely

### `FAST_TEST_MODE = TRUE`
- Bypasses working-day and sending-window schedule checks
- Uses minute-based delays instead of day-based delays
- Best for QA and functional testing

### `PREVIEW_MODE`
- Use preview generation to inspect rendered content before live sends

### Production best practice
1. Keep `TEST_MODE=TRUE` during setup
2. Send test rows to your own inbox
3. Review `Preview`, `ActivityLog`, and thread behavior
4. Switch to production only after validation

---

## Scheduling behavior

Production sends are gated by:
- Working day check (`WORKING_DAYS`)
- Sending window check (`SEND_START_HOUR` → `SEND_END_HOUR` in configured `TIMEZONE`)

When `FAST_TEST_MODE` is enabled, schedule checks are bypassed.

---

## Campaign progression model

- Active services are read from `Services` sheet and ordered by sequence.
- A lead starts with assigned service and campaign (`Service N`).
- After `FOLLOWUP_3_SENT` and the configured next-service delay:
  - if next service exists: lead is rotated to next service and campaign reset
  - if no next service: lead status becomes `COMPLETED`

---

## Bounce handling

`checkBounces()` scans recent delivery-failure messages (7-day window) and
classifies each DSN into one of four categories:

- `HARD_BOUNCE` — permanent (mailbox not found, disabled) → status `INVALID`
  + added to `SuppressionList`
- `SOFT_BOUNCE` — temporary (greylist, mailbox full) → retry counter
  incremented; after 3 retries it is promoted to `UNKNOWN` for review
- `POLICY_REJECTION` — reputation/blocked (5.7.x) → logged, not retried
- `UNKNOWN` — unparseable → logged for manual review

Hard-bounced addresses land in the `SuppressionList` sheet and are blocked
before every send even if re-imported later. Duplicate Gmail DSN messages
are skipped via a 7-day message-ID dedup cache.

## Reply handling

`checkReplies()` (Replies.js) scans every active lead's Gmail thread for a
human reply:

- messages from the lead's own address, newer than the last send, and not
  auto-replies/bounces (filters `Auto-Submitted`, `Precedence`, etc.)
- a mailbox fallback catches replies sent via the mailto "reply to this
  email" link, which start a brand-new thread

On match the lead is set to `REPLIED` (terminal) and logged as
`REPLY_DETECTED`. Add a separate 5-minute trigger for `checkReplies()` and
set `SKIP_REPLIES_IN_MAIN_RUN=TRUE` to free the main run's send budget.

## Unsubscribe (one-click, RFC 8058)

`Unsubscribe.js` is a web app that lets recipients opt out with one click:

1. Deploy as a web app (Execute as: me, Access: Anyone).
2. Paste the URL into Config as `UNSUBSCRIBE_URL`.
3. Emails carry a signed `List-Unsubscribe` header and an
   `{{UnsubscribeLink}}` footer link (HMAC token, secret in Script
   Properties as `UNSUB_TOKEN_SECRET`).

On success the lead becomes `DO_NOT_CONTACT` and the address is added to
the `SuppressionList` as `UNSUBSCRIBED`.

---

## How to use (quick start)

1. **Create/prepare spreadsheet tabs** listed above.
2. **Populate `Config` and `AIConfig`** with required keys.
3. **Add services** in `Services` and mark active rows.
4. **Load templates** in `Templates` (EMAIL_1 + follow-ups per service).
5. **Add leads** in `Leads` with valid email + assigned service.
6. Set `TEST_MODE=TRUE` and set `TEST_EMAIL`.
7. Run `generatePreview()` and inspect output.
8. Run `runAutomation()` manually for first test cycle.
9. Verify `ActivityLog`, status transitions, and Gmail threads.
10. Set production config (disable test mode), then run on trigger schedule.

---

## Suggested trigger strategy

Use a time-driven trigger for `runAutomation()` (for example, every 15–60 minutes during your send window).

Guidance:
- choose an interval that matches your lead volume and daily limit
- avoid excessive trigger frequency if volume is low
- monitor logs and dashboard weekly

---

## Operational checklist before going live

- [ ] Run `ensureHardeningMigration()` once (creates Leads S–V + SuppressionList + config keys)
- [ ] All required sheets exist and headers are correct
- [ ] `DAILY_LIMIT` is set conservatively
- [ ] `WORKING_DAYS`, `SEND_START_HOUR`, `SEND_END_HOUR`, `TIMEZONE` are correct
- [ ] `TEST_MODE` validated with your own inbox
- [ ] Template placeholders resolve correctly
- [ ] Bounce handling confirmed (`testBounceParser`, `dryRunHistoricalSuppression`)
- [ ] Reply detection confirmed (`testReplyDetection` dry run)
- [ ] Unsubscribe deployed and `UNSUBSCRIBE_URL` set (`testUnsubscribeLink`)
- [ ] Sender identity (`SENDER_NAME`, `REPLY_TO_EMAIL`) configured

---

## Troubleshooting

### ��Sheet not found” errors
Create missing sheet tab using exact expected name.

### Emails not sending
Check:
- `AUTOMATION_ENABLED`
- `DAILY_LIMIT` not exhausted
- send window/day rules
- valid recipient emails
- Apps Script permissions and Gmail advanced service

### Follow-ups not sending
Check:
- lead has `threadId`
- lead status is one of follow-up source statuses
- `lastEmailDate` is valid
- required delay has elapsed

### Test mode confusion
When `TEST_MODE=TRUE`, all outgoing messages route to `TEST_EMAIL` (not the lead’s original address).

---

## Security and compliance notes

- Respect consent and local email regulations in your target region.
- Keep unsubscribe language in templates/config where appropriate.
- Do not send to purchased/unverified lists.
- Regularly clean invalid/bounced leads.
- Setup DKIM, SPF and DMARC for your domain for verified sending and increase inbox rates.


