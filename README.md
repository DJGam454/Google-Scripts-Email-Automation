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
1. `checkBounces()`
2. `generateMissingPersonalizations()`
3. `processFollowUps()`
4. `processEmails()`

---

## Core architecture (code map)

- `EmailEngine.js`  
  Sending engine, follow-up logic, automation runner, delay handling, thread-aware follow-ups.

- `HtmlEmailEngine.js`  
  Renders HTML + plain-text emails, component-based layout, theme support, MIME construction for Gmail API thread replies.

- `Templates.js`  
  Template lookup, active variant selection, placeholder expansion.

- `Leads.js`  
  Canonical lead column map, lead object builder, statuses and follow-up step mapping.

- `Campaigns.js`  
  Service ordering, next-service movement, campaign completion transitions.

- `Validation.js`  
  Email format validation, daily sent count, daily-limit gate.

- `Bounces.js`  
  Bounce scanning via Gmail search and lead invalidation.

- `Modes.js`  
  Mode parsing (`TEST_MODE`, `FAST_TEST_MODE`, `PREVIEW_MODE`) and recipient redirection.

- `Scheduling.js`  
  Working day + sending window checks with timezone support.

- `Dashboard.js`  
  Dashboard setup, KPI formulas, preview generation, activity logging.

- `Config.js`  
  Reads configuration from `Config` and `AIConfig` sheets (cached).

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
- `TEST_MODE`
- `TEST_EMAIL`
- `PREVIEW_MODE`
- `FAST_TEST_MODE`
- `TIMEZONE`
- `WORKING_DAYS`
- `SEND_START_HOUR`
- `SEND_END_HOUR`

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

`checkBounces()` scans recent delivery-failure messages and tries to extract bounced recipient addresses.

When a match is found in `Leads`:
- status is set to `INVALID`
- activity is logged as `EMAIL_BOUNCED`

This prevents continued sends to bad addresses.

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

- [ ] All required sheets exist and headers are correct
- [ ] `DAILY_LIMIT` is set conservatively
- [ ] `WORKING_DAYS`, `SEND_START_HOUR`, `SEND_END_HOUR`, `TIMEZONE` are correct
- [ ] `TEST_MODE` validated with your own inbox
- [ ] Template placeholders resolve correctly
- [ ] Bounce handling confirmed
- [ ] Reply-handling/status update process defined
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


