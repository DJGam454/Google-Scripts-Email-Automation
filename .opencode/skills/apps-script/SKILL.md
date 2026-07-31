---
name: apps-script
description: Google Apps Script (V8) development for this GAS email-automation project. Use when editing .js files deployed via clasp, working with SpreadsheetApp, GmailApp, the Gmail advanced service, CacheService, LockService, Script Properties, time-based triggers, or appsscript.json / .clasp.json configuration.
---

# Google Apps Script Development

## About this project

This repository is a Google Apps Script (V8 runtime) application that automates
B2B cold-email outreach. All source files are plain `.js` files deployed to a
single Apps Script project via `clasp` (see `.clasp.json` for the target
`scriptId`). There is no build step, no module system, and no bundler.

## Key facts you must respect

- **V8 runtime only.** Modern ES2015+ syntax is allowed: `let`/`const`, arrow
  functions, template literals, `String.prototype.replaceAll`, spread. Do not
  use ES modules (`import`/`export`) — Apps Script does not support them in
  this project layout.
- **Global namespace.** All top-level functions across all `.js` files share a
  single global scope. There are no file imports. A function in one file can
  (and does) call a function defined in another file — this is by design
  (e.g. `EmailEngine.js` calls `getConfig()` from `Config.js`,
  `isValidEmail()` from `Validation.js`, `logActivity()` from `Dashboard.js`).
- **No bundlers / npm packages.** External APIs are called with
  `UrlFetchApp.fetch()` only (see `callGemini` in `AI.js` for the pattern).
- **Entry points are plain functions.** Any top-level function can be run
  manually from the Apps Script editor and can be attached to triggers.
  Production entry point is `runAutomation()` in `EmailEngine.js`.
- **Test functions.** Helper `test*()` functions (in `Tests.js` and per-file
  test helpers such as `testConfig`, `testGeminiKey`, `testBounceParser`) are
  runnable from the Apps Script editor. Keep them isolated and never call them
  from production code paths.

## Files in this project

| File | Responsibility |
| --- | --- |
| `EmailEngine.js` | Orchestration: `runAutomation()`, `processEmails()`, `processFollowUps()`, `sendFollowUp()`, `sendThreadedFollowUp()`, delay helpers. Sends HTML via `renderHtmlEmail()` (`htmlBody` on GmailApp path, multipart/alternative MIME on Gmail API path) |
| `Leads.js` | Shared data layer: `LEADS_COL` map, `LEAD_STATUS`, `STOP_STATUSES`, `FOLLOWUP_STEPS`, `getLeadsSheet()`, `getLeadsData()`, `buildLeadFromRow()` (per-execution cached reads) |
| `Templates.js` | `getTemplate()`, `getTemplateVariants()`, `getRandomTemplate()`, `personaliseTemplate()`, `buildPlaceholderValues()`, `applyPlaceholderValues()`, `refreshTemplatesCache()` |
| `HtmlEmailEngine.js` | `renderHtmlEmail()` (HTML skeleton + `{{Component}}` tokens + theme), `getTheme()`, `buildMultipartAlternative()`, `htmlToPlainText()`, `DEFAULT_THEMES` |
| `AI.js` | Gemini calls, rate limiter, website research, personalisation generation |
| `Bounces.js` | `checkBounces()`, `extractBouncedEmail()` |
| `Campaigns.js` | Service rotation: `getServices()`, `getNextService()`, `finishCurrentCampaign()`, `hasDuplicateCampaign()` |
| `Config.js` | `getConfig()`, `getAIConfig()` — reads Config/AIConfig sheets with `CacheService` |
| `Dashboard.js` | `setupDashboard()`, `buildEmailPreview()`, `generatePreview()`, `logActivity()` |
| `Modes.js` | `isTestMode()`, `isPreviewMode()`, `isFastTestMode()`, `getActualRecipient()` |
| `Scheduling.js` | Working-day and sending-window checks |
| `Tests.js` | Cross-cutting `test*()` helpers |
| `Validation.js` | `canSendEmail()`, `getEmailsSentToday()`, `isValidEmail()` |
| `appsscript.json` | Project manifest (timezone `Asia/Kolkata`, Gmail advanced service, V8) |

## Services and APIs used

- **`SpreadsheetApp`** — the primary data store. All sheets are accessed via
  `SpreadsheetApp.getActiveSpreadsheet().getSheetByName(...)`.
- **`GmailApp`** — simple sends (`GmailApp.sendEmail`), searching
  (`GmailApp.search`), and thread lookup (`GmailApp.getThreadById`).
- **`Gmail` (advanced service)** — `Gmail.Users.Messages.send` for threaded
  follow-ups and `Gmail.Users.Threads.get` to read raw headers. Declared in
  `appsscript.json` as `enabledAdvancedServices`. It is enabled by default and
  must remain enabled for follow-ups to work.
- **`UrlFetchApp`** — Gemini REST API calls. Never call external APIs without
  `muteHttpExceptions: true` and explicit response-code handling.
- **`CacheService.getScriptCache()`** — config caching (see `Config.js`;
  TTL 300 s). Only cache JSON-serialisable data, never PII or secrets.
- **`LockService.getScriptLock()`** — `runAutomation()` takes a lock so
  concurrent trigger executions cannot run simultaneously.
- **`PropertiesService.getScriptProperties()`** — stores secrets such as
  `GEMINI_API_KEY`. Secrets must never be hard-coded in source files.
- **`Utilities.sleep()` / `Utilities.formatDate()` / `Utilities.base64EncodeWebSafe()`** —
  delays, timezone-aware formatting, and MIME encoding respectively.

## Conventions

- File header: section banner comments like
  `// ============================================================`.
- Inline section comments:
  `// =================================` used to delimit steps inside functions.
- Private helpers start with an underscore: `_getRandomDelay`, `_waitForAIRateLimit`,
  `_normaliseWebsiteUrl`. Treat any function with a leading `_` as internal —
  do not call it from other files unless it already is.
- Logging: `console.log(...)` for informational output, `console.error(...)`
  for failures. Every send/failure path logs the lead email for traceability.
- Never `throw` into the void inside per-lead loops — per-lead errors are
  caught and logged so one bad lead cannot kill the run
  (see `generateMissingPersonalizations()` for the pattern).
- Column indices in sheet reads are documented in comments; when you change a
  column contract, update every consumer.

## Quotas and limits (critical)

Apps Script imposes hard quotas. Never design code that assumes unlimited
executions, API calls, or email sends:

- **Gmail send quota** — self-imposed `DAILY_LIMIT` from the Config sheet is
  enforced by `canSendEmail()`; always respect it.
- **External API quota** — Gemini calls are throttled by `_waitForAIRateLimit()`
  using `AI_RPM_LIMIT` (default 10). Never bypass the rate limiter.
- **Execution time** — a single Apps Script execution caps at 6 minutes (with
  triggers, 30 minutes for some trigger types). Do not add unbounded sleeps or
  huge loops without considering the limit; the existing 3–4 s random delays
  between sends are deliberate.
- **`Utilities.sleep()`** maxes at 5 minutes per call; keep waits below it.

## Deployment

- Deploy with `clasp push` (the project is already configured via
  `.clasp.json`). After pushing, changes are live for the bound spreadsheet.
- `appsscript.json` changes (runtime, services, timezone) also require a push.
- If the `Gmail` advanced service is ever disabled in the manifest, threaded
  follow-ups and website research will break — verify it stays enabled.

## Golden rules when editing code

1. Preserve the existing safety checks: status gating (`NEW` only), email
   validation, duplicate-campaign blocking, thread-ID presence, daily limit.
   Never remove a guard to "make it work faster".
2. Prefer `getValues()`/`setValues()` batch reads/writes over cell-by-cell
   access inside loops.
3. Keep functions small and single-purpose. If a function exceeds ~200 lines,
   split it or move logic to the appropriate file.
4. Any change that alters sheet columns, statuses, or config keys must be
   mirrored across every consumer file — run `testReadSheet()` /
   `testConfig()` / `testAIConfig()` (or write a new `test*()` helper) to
   verify.
5. Do not add npm dependencies, HTML files, or other files not deployable by
   clasp unless the user explicitly asks.
