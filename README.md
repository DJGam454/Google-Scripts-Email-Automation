# Google Apps Script Email Automation

Automated B2B cold-email outreach powered entirely by a bound Google
Spreadsheet. Google Apps Script (V8) orchestrates the pipeline: Gemini writes
personalised intros, Gmail sends initial emails and threaded follow-ups,
bounces and human replies are detected automatically, and leads rotate
through services (Website Development → SEO → Google Ads) until every
campaign completes.

## How it works

A time-based trigger calls `runAutomation()` in `EmailEngine.js`. Each run:

1. **Checks bounces** — DSNs are classified (`HARD_BOUNCE`, `SOFT_BOUNCE`,
   `POLICY_REJECTION`, `UNKNOWN`); hard bounces mark the lead `INVALID` and
   land in the SuppressionList.
2. **Checks replies** — the Gmail API is scanned for human replies; any
   inbound mail newer than the last send marks the lead `REPLIED`
   (terminal stop-status) so no further emails go out.
3. **Generates AI personalisation** — Gemini writes a personalised intro for
   each `NEW` lead that lacks one (rate-limited by `AI_RPM_LIMIT`).
4. **Processes follow-ups** — leads in `EMAIL_1_SENT` / `FOLLOWUP_1_SENT` /
   `FOLLOWUP_2_SENT` advance through the remaining steps, threaded via
   `In-Reply-To` / `References`.
5. **Sends email 1** — `NEW` leads receive the first email, then rotate to
   the next service after the final follow-up until `COMPLETED`.

Safety is enforced at every layer: `DAILY_LIMIT`, hourly and per-domain caps,
a per-run send budget (`MAX_SENDS_PER_RUN`), a 299-second execution
self-cap, random 3–4 s send delays, a Gmail quota circuit breaker, a script
lock, and stop-statuses (`REPLIED` / `INVALID` / `DO_NOT_CONTACT`) that
permanently halt automation for a lead.

## Spreadsheet contract

The bound spreadsheet drives everything. Sheets: `Leads` (A–V), `Config`,
`AIConfig`, `Templates`, `Themes`, `Services`, `ActivityLog`, `AIResearch`,
`SuppressionList`. See `.opencode/skills/excel-sheet/SKILL.md` for the full
column map, statuses, and config keys.

## File map

| File | Owns |
| --- | --- |
| `EmailEngine.js` | Pipeline orchestration, send budgets, execution deadline, MIME sends |
| `Leads.js` | Leads sheet contract, lead builder, suppression layer |
| `Templates.js` | Template lookup, variants, `{{Placeholder}}` personalisation |
| `HtmlEmailEngine.js` | HTML rendering, themes, multipart MIME builder |
| `AI.js` | Gemini calls, rate limiter, website research, intros |
| `Replies.js` | Automatic human-reply detection (thread + mailbox scan) |
| `Bounces.js` | DSN classification, suppression writes, dedup |
| `Unsubscribe.js` | One-click unsubscribe web app (RFC 8058, HMAC-signed links) |
| `SuppressionReconcile.js` | 30-day mailbox backfill + config-driven domain blocking |
| `VerificationCleanup.js` | Bulk verification export application (lists empty by default) |
| `TestModeFix.js` | One-time test-date reconcile + ActivityLog purge helpers |
| `Campaigns.js` | Service rotation, campaign lifecycle, duplicate detection |
| `Config.js` | Config/AIConfig reads, hardening migration helper |
| `Dashboard.js` | Dashboard rebuild, previews, activity logging |
| `Scheduling.js` | Working-day + sending-window gates |
| `Validation.js` | Daily/hourly/domain limits, quota breaker, email validation |
| `Modes.js` | Test/preview/fast-test modes, recipient resolution |
| `Tests.js` | Cross-cutting test helpers |

## Setup

1. Open the bound spreadsheet → Extensions → Apps Script (or push with
   [clasp](https://developers.google.com/apps-script/guides/clasp)).
2. Script Properties: add `GEMINI_API_KEY` (never commit it).
3. Run `ensureHardeningMigration()` once — it creates the Leads S–V
   columns, the SuppressionList sheet, and the throttling config keys.
4. Optional: deploy `Unsubscribe.js` as a web app and set `UNSUBSCRIBE_URL`
   in Config to enable one-click unsubscribes.
5. Configure the Config/AIConfig sheets (see the excel-sheet skill).
6. Create a time-based trigger for `runAutomation()`; optionally add a
   separate 5-minute trigger for `checkReplies()` and set
   `SKIP_REPLIES_IN_MAIN_RUN` to `TRUE`.

## Testing

Apps Script has no compiler. Every decision path logs to the execution
transcript (`console.log` / `console.error`). Verify changes with the
`test*()` helpers (Apps Script editor → Run), then `clasp push`. Dry-run
helpers exist for every destructive operation:
`reconcileMailboxBounces(true)`, `applyVerificationCleanup(true)`,
`purgeTestActivityLog(true)`, `reconcileTestDates(true)`.

## Deliverability rules (non-negotiable)

- Never weaken `DAILY_LIMIT` / `canSendEmail()` / hourly or domain caps.
- Never bypass the quota breaker, the run budget, or the 3–4 s delays.
- Test-mode sends and previews must be verified before production runs.
- Secrets (API keys, token secrets) live in Script Properties only.