---
name: excel-sheet
description: The Google Sheets data layer behind this automation (driven by Leads.xlsx). Use when reading or writing the Leads, Config, AIConfig, Templates, Services, ActivityLog, or AIResearch sheets; dealing with the Leads column map (A–R), lead statuses, boolean cell parsing, batch getValues/setValues operations, or spreadsheet-driven configuration.
---

# Spreadsheet / Excel Data Layer

## Overview

The entire application is configured and driven by a Google Spreadsheet
(originally `Leads.xlsx` in this repo, bound to the Apps Script project).
`SpreadsheetApp.getActiveSpreadsheet()` returns the bound spreadsheet; there
are **ten** sheets:

| Sheet | Purpose |
| --- | --- |
| `Leads` | Lead records and campaign state (the operational core) |
| `Config` | Global automation configuration (key/value, rows 2+) |
| `AIConfig` | Gemini/AI configuration (key/value, rows 2+) |
| `Templates` | Email templates with variants: Service | Step | Subject | Plain Text | Variant ID | HTML | CTA | Theme | Images | Status |
| `Themes` | Optional per-theme colour sets (falls back to built-in defaults) |
| `Services` | Active services in outreach order: Order | Name | Active |
| `ActivityLog` | Append-only audit log of every action |
| `AIResearch` | Cached website research per lead (append-only) |
| `SuppressionList` | Permanently suppressed addresses (Email | Category | Diagnostic | First Seen | Last Seen | Count | Source) |

## Leads sheet column map (critical — do not break)

Headers live in row 1; data starts at row 2. Indexes used throughout the code:

| Col | Letter | Header (as used in code) |
| --- | --- | --- |
| 1 | A | Lead ID |
| 2 | B | Company |
| 3 | C | Contact Name |
| 4 | D | Email |
| 5 | E | Website |
| 6 | F | Industry |
| 7 | G | Personalised Intro (AI-generated) |
| 8 | H | Service Assigned |
| 9 | I | Current Campaign (e.g. `Service 1`) |
| 10 | J | Email 1 Sent (timestamp) |
| 11 | K | Follow-up 1 Sent |
| 12 | L | Follow-up 2 Sent |
| 13 | M | Follow-up 3 Sent |
| 14 | N | Last Email Date |
| 15 | O | Status |
| 16 | P | Gmail Thread ID |
| 17 | Q | Notes |
| 18 | R | Last Updated |
| 19 | S | Bounce Category (`HARD_BOUNCE`/`SOFT_BOUNCE`/`POLICY_REJECTION`/`UNKNOWN`/`UNSUBSCRIBED`) |
| 20 | T | Bounce Diagnostic (DSN snippet) |
| 21 | U | Suppressed At (timestamp) |
| 22 | V | Retry Count (soft-bounce counter) |

S–V are created by `ensureLeadsBounceColumns()` (Leads.js) — run
`ensureHardeningMigration()` once after deploying to append them plus the
SuppressionList sheet. Every consumer reads the bounce fields through
`buildLeadFromRow()` (`bounceCategory`, `bounceDiagnostic`, `suppressedAt`,
`retryCount`) with `row.length` guards for sheets that predate the columns.

## Lead statuses (col O)

| Status | Meaning |
| --- | --- |
| `NEW` | Eligible for initial email (only statuses processed by `processEmails`) |
| `EMAIL_1_SENT` | Initial email sent, follow-up 1 due |
| `FOLLOWUP_1_SENT` / `FOLLOWUP_2_SENT` | Follow-up sent, next due |
| `FOLLOWUP_3_SENT` | Campaign complete; waiting for next-service delay |
| `REPLIED` | Human reply — automation must stop (terminal, stop-status) |
| `INVALID` | Invalid email or bounced (terminal, stop-status) |
| `DO_NOT_CONTACT` | Manual opt-out (terminal, stop-status) |
| `COMPLETED` | All campaigns finished |
| `NOT_INTERESTED_SERVICE_n` | Logged by `finishCurrentCampaign` via activity log |

Rules:

- Only `NEW` leads receive `EMAIL_1`.
- Only statuses in `ACTIVE_CAMPAIGN_STATUSES` (`[EMAIL_1_SENT, FOLLOWUP_1_SENT, FOLLOWUP_2_SENT, FOLLOWUP_3_SENT]`) are considered for follow-ups (`processFollowUps`) or reply detection (`checkReplies`).
- Stop-statuses (`REPLIED`, `INVALID`, `DO_NOT_CONTACT`) short-circuit all further automation.
- When a campaign finishes (`finishCurrentCampaign`), cols G (intro), J–N (timestamps), and P (thread ID) are cleared/reset and status returns to `NEW` with the next service in `H`. Suppression columns S–V are untouched by rotation (suppression is permanent).

## Templates sheet column map

Headers in row 1, data from row 2. Columns A–D are legacy-compatible
(Service | Step | Subject | Body); E–M power the variant/HTML system:

| Col | Letter | Header | Code field | Notes |
| --- | --- | --- | --- | --- |
| 1 | A | Service | `service` | e.g. `Website Development` |
| 2 | B | Campaign Step | `step` | `EMAIL_1`, `FOLLOWUP_1`..`FOLLOWUP_3` |
| 3 | C | Subject | `subject` | may contain `{{Placeholders}}`; travels with the variant |
| 4 | D | Plain Text Template | `body` | legacy "Body" column |
| 5 | E | Variant ID | `variantId` | 1, 2, ... one row per variant |
| 6 | F | HTML Template | `html` | full skeleton; blank → plain-text-only email |
| 7 | G | CTA Button Text | `ctaText` | blank → `DEFAULT_CTA_TEXT` |
| 8 | H | CTA URL | `ctaUrl` | blank → `DEFAULT_CTA_URL` → `PORTFOLIO_URL` |
| 9 | I | Theme | `theme` | blank = service default (website/seo/google-ads) |
| 10 | J | Hero Image | `heroImage` | public image URL, optional |
| 11 | K | Portfolio Image | `portfolioImage` | public image URL, optional |
| 12 | L | Status | `status` | `ACTIVE` / `INACTIVE`; blank counts as active (legacy rows) |
| 13 | M | Notes | `notes` | free text |

Read through `Templates.js` — `getTemplate` (legacy first-match),
`getTemplateVariants` (ACTIVE only), `getRandomTemplate` (random ACTIVE
variant; subject always comes with its variant). The send path and
`buildEmailPreview` never fetch subjects independently of templates.

## Themes sheet columns

Optional per-theme override of the built-in colour sets
(`HtmlEmailEngine.DEFAULT_THEMES`). Columns: Theme (A) | Primary (B) |
Background (C) | Text (D) | Muted (E) | Accent (F) | Border (G) | Font Stack (H)
| Status (I) | Notes (J). Status `Active` rows win over defaults; blank or
missing Themes sheet = built-in defaults (website blue, SEO green, Google Ads
orange).

## Boolean cells: text vs boolean

Sheets may return booleans as real booleans or as the strings `"TRUE"`/`"FALSE"`.
Every config boolean (e.g. `AUTOMATION_ENABLED`, `TEST_MODE`,
`FAST_TEST_MODE`, `PREVIEW_MODE`, `AI_ENABLED`, service `Active`) must be read
with the dual check pattern:

```js
const enabled = config.X === true || String(config.X).trim().toUpperCase() === "TRUE";
```

See `Modes.js` (`isFastTestMode`, `isPreviewMode`, `isTestMode`) and
`getServices()` in `Campaigns.js` for canonical implementations. Reuse those
helpers — never inline new boolean parsing.

## Config sheets (Config, AIConfig)

- Read via `getConfig()` / `getAIConfig()` in `Config.js`, which cache results
  in `CacheService` for 300 s.
- Key in column A, value in column B, data from row 2 down. Blank keys are
  skipped.
- When adding a new config key, add it to the sheet **and** read it through
  these functions — never read the sheet directly elsewhere.
- Keys in use include: `AUTOMATION_ENABLED`, `DAILY_LIMIT`, `HOURLY_LIMIT`,
  `HOURLY_LIMIT_YAHOO`, `HOURLY_LIMIT_OUTLOOK`, `HOURLY_LIMIT_ICLOUD`,
  `MAX_SENDS_PER_RUN`, `NEW_LEADS_DAILY_FLOOR`, `SUPPRESSED_DOMAINS`
  (comma-separated, e.g. `example.com,example.org`), `SKIP_REPLIES_IN_MAIN_RUN`,
  `SENDER_NAME`, `REPLY_TO_EMAIL`, `EMAIL_SIGNATURE` (now placeholder-based, e.g.
  `"Regards,\n{{SenderName}}\n{{SenderRole}}, {{CompanyName}}"`), `SENDER_ROLE`,
  `TEST_MODE`, `TEST_EMAIL`, `FAST_TEST_MODE`, `PREVIEW_MODE`, `TIMEZONE`,
  `WORKING_DAYS`, `SEND_START_HOUR`, `SEND_END_HOUR`, `SEND_WINDOWS`
  (minute-precision ranges, e.g. `9:30-12:25,14:00-17:00` — overrides the
  legacy start/end hours),
  `FOLLOWUP_1_MINUTES` … `FOLLOWUP_3_MINUTES` (test mode),
  `FOLLOWUP_1_DAYS` … `FOLLOWUP_3_DAYS` (production),
  `NEXT_SERVICE_MINUTES`, `NEXT_SERVICE_DAYS`, `UNSUBSCRIBE_URL` (web-app
  deployment of Unsubscribe.js); brand/HTML keys:
  `COMPANY_NAME`, `COMPANY_WEBSITE`, `COMPANY_EMAIL`, `COMPANY_PHONE`,
  `COMPANY_ADDRESS`, `SOCIAL_LINKEDIN`, `SOCIAL_TWITTER`, `SOCIAL_GITHUB`,
  `PORTFOLIO_URL`, `DEFAULT_CTA_TEXT`, `DEFAULT_CTA_URL`,
  `UNSUBSCRIBE_MESSAGE`, `HERO_TITLE`, `HERO_SUBTITLE`,
  `SERVICE_BENEFIT_WEBSITE_DEVELOPMENT`/`_SEO`/`_GOOGLE_ADS`,
  `SERVICE_FEATURES_WEBSITE_DEVELOPMENT`/`_SEO`/`_GOOGLE_ADS` (JSON string
  arrays rendered as `• item` bullets); AIConfig: `AI_ENABLED`, `AI_MODEL`,
  `TEMPERATURE`, `MAX_INTRO_WORDS`, `AI_RPM_LIMIT`.
- `getConfig()` applies fail-safe defaults: blank `DAILY_LIMIT` → 0 (blocks
  all sends), blank `HOURLY_LIMIT` → `ceil(DAILY_LIMIT / 4)` (or 15).
- New config keys are appended by `ensureHardeningMigration()` /
  `refreshHardeningConfig()` / `refreshSendWindowsConfig()` (Config.js) —
  never overwriting existing values.
- Timing keys are read as numbers and validated (`Number()` + `isNaN` checks)
  in `getFollowUpDelay` / `getNextServiceDelay` (EmailEngine.js).

## ActivityLog sheet columns

Append-only log written by `logActivity()` (Dashboard.js) and read by
`getEmailsSentToday()` (Validation.js) and `hasDuplicateCampaign()`
(Campaigns.js). Column indexes used by code:

- A(0) Timestamp, B(1) Lead ID, C(2) Email, D(3) Action (`EMAIL_1`,
  `FOLLOWUP_1..3`, `EMAIL_BOUNCED`, `CAMPAIGN_*`, ...), F(5) Service, J(9) Result
  (`SUCCESS`/`FAILED`).
- Comment in `hasDuplicateCampaign()` explicitly warns: adjust indexes only if
  columns change. Never reorder ActivityLog columns without auditing these two
  consumers.

## Access patterns (performance)

- **Shared data layer**: `Leads.js` owns the Leads sheet contract —
  `getLeadsSheet()`, `getLeadsData()` (per-execution cached `getValues()`),
  `buildLeadFromRow(data, index)` (row array → lead object with named fields).
  All consumers (`EmailEngine.js`, `AI.js`, `Bounces.js`, `Replies.js`,
  `Dashboard.js`, `Unsubscribe.js`, `SuppressionReconcile.js`) must go through
  these; never re-parse Leads rows in place.
- **Suppression layer**: `isSuppressed(email)`, `getSuppressionCategory(email)`,
  `addToSuppressionList(...)` (idempotent), `ensureLeadsBounceColumns()`,
  `writeLeadBounceState(row, category, diagnostic, incrementRetry)` all live in
  `Leads.js`. The SuppressionList sheet is the permanent blocklist — check it
  before every send, never resurrect a suppressed address.
- **Read** the whole sheet once with `sheet.getDataRange().getValues()` and
  iterate the 2D array; never call `getValue()` per row in a loop over the
  full dataset. This is the established pattern in `processEmails`,
  `processFollowUps`, `checkBounces`, `generateMissingPersonalizations`.
- **Write** in batches: `sheet.getRange(row, 10, 1, 9).setValues([[ ... ]])`
  for the J–R block (EmailEngine.js), `getRange(row, 10, 1, 5).clearContent()`
  for J–N resets (Campaigns.js). Prefer one batched write over many
  `setValue` calls.
- Timestamps are `new Date()` objects written directly to cells.
- Guard every `getSheetByName(...)` with an existence check that throws a
  descriptive error (`throw new Error("Leads sheet not found.")`), matching
  the existing codebase convention.

## Testing

Existing helpers you can extend rather than inventing new infrastructure:

- `testReadSheet()` — dumps the Leads sheet.
- `testConfig()` / `testAIConfig()` — dump config maps.
- `testServices()` — dumps active services in order.
- `testDailyCount()` — reads ActivityLog counts.

When changing sheet structure, update the header row and this column map, and
verify with the relevant `test*()` function before deploying with `clasp push`.
