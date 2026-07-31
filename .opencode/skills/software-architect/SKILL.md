---
name: software-architect
description: Architecture and engineering conventions for this Apps Script email-automation codebase. Use when planning new features, deciding which file a function belongs in, refactoring or reorganising code, reviewing changes for separation of concerns, naming, logging, or safety, and when designing extensions to the automation pipeline.
---

# Software Architecture & Engineering Conventions

## System overview

A Google Apps Script (V8) application orchestrating B2B cold-email campaigns
entirely from a bound Google Spreadsheet:

```
Trigger (time-based) ──► runAutomation() [EmailEngine.js]
                           │ LockService guard
                           │ Config gate (AUTOMATION_ENABLED)
                           │ Schedule gate (Scheduling.js)
                           ▼
                   1. checkBounces()          [Bounces.js]
                   2. generateMissingPersonalizations() [AI.js]
                   3. processFollowUps()      [EmailEngine.js]
                   4. processEmails()         [EmailEngine.js]
```

The pipeline order is deliberate and must be preserved:

1. **Bounces first** — purge bad addresses before any new sends.
2. **AI personalisation next** — NEW leads need intros before they can send.
3. **Follow-ups before new emails** — existing conversations are prioritised
   over new cold outreach.
4. **Email 1 last** — new leads are the tail end of the run.

## File ownership (separation of concerns)

Each `.js` file owns one concern, and cross-file calls are allowed because
Apps Script shares one global scope. Keep it that way:

| File | Owns |
| --- | --- |
| `EmailEngine.js` | Pipeline orchestration and send mechanics (delays, threading, status transitions, HTML wiring into both send paths) |
| `Leads.js` | Leads sheet contract: column map, statuses, follow-up step map, cached sheet reads, `buildLeadFromRow()` |
| `Templates.js` | Template lookup (first-match + variants + random selection) and placeholder personalisation |
| `HtmlEmailEngine.js` | HTML email rendering: themes, `{{Component}}` tokens, plain-text fallback, multipart MIME builder |
| `AI.js` | All Gemini interaction: rate limiter, website research, personalisation generation and validation |
| `Bounces.js` | Gmail bounce detection and DSN parsing |
| `Campaigns.js` | Service rotation, campaign lifecycle, duplicate detection |
| `Config.js` | Reading Config/AIConfig sheets through cache |
| `Dashboard.js` | UI/setup, previews (`buildEmailPreview`, `generatePreview`), `logActivity()` |
| `Modes.js` | Mode flags and recipient resolution (`getActualRecipient`), `isFlagTrue()` |
| `Scheduling.js` | Working-day and sending-window logic |
| `Validation.js` | Email validation and daily-limit counting |
| `Tests.js` | Cross-cutting `test*()` helpers |

Placement rules:

- New send logic → `EmailEngine.js`. New Gemini logic → `AI.js`. New sheet
  read/write of config → `Config.js`. New campaign lifecycle logic →
  `Campaigns.js`. New lead-data access → `Leads.js`. New HTML/MIME content
  work → `HtmlEmailEngine.js`; new template/placeholder logic → `Templates.js`.
- Do not create new top-level files unless the concern is genuinely new
  (a future example: a `Replies.js` for reply detection).
- Test helpers live next to the code they test (`testConfig` in Config.js) or
  in `Tests.js` for cross-cutting checks — never inline in production paths.

## Naming conventions

- Top-level functions: `camelCase`, descriptive verb-first
  (`processEmails`, `getActualRecipient`, `finishCurrentCampaign`).
- Private helpers: `_underscorePrefix` (`_getRandomDelay`, `_waitForAIRateLimit`,
  `_normaliseWebsiteUrl`). Leading underscore signals "internal to this file's
  domain" — do not call from other files.
- Test functions: `test` + PascalCase (`testGeminiKey`, `testBounceParser`).
- Constants in ALL_CAPS only when they are config keys (e.g.
  `FOLLOWUP_1_MINUTES`) or status strings (`EMAIL_1_SENT`).
- Variables: lowercase `camelCase`; avoid single-letter names except loop
  counters (`i`, `m`, `t`), matching existing style.

## Code style (match the codebase)

- Banner headers between files/sections:
  `// ============================================================`.
- Inline step dividers inside functions:
  `// =================================` followed by an UPPERCASE comment label
  (e.g. `// SEND EMAIL`), then a blank line.
- One statement per line; blank lines between logical steps.
- `const` for everything not reassigned; `let` only when needed; `var` only
  in legacy/private helpers (do not introduce new `var`).
- Use `String(...)`, `Number(...)`, `.trim()`, `.toUpperCase()` defensively on
  sheet data — cells can hold any type.
- No comments that restate the code; comments explain *why* (safety intent,
  sheet column contracts, Gmail quirks).

## Safety-first design patterns (non-negotiable)

1. **Per-lead error isolation.** Wrap per-lead work in `try/catch`, log with
   `console.error`, and `continue`. One bad lead must never abort the run
   (`generateMissingPersonalizations`, `processEmails`, `processFollowUps`).
2. **Validate before act.** Every send path validates: status gating (`NEW`),
   email format (`isValidEmail`), actual recipient validity, template
   existence, thread-ID presence, daily limit — in that spirit of order.
3. **Terminal states are absolute.** `REPLIED`, `INVALID`, `DO_NOT_CONTACT`
   halt automation. Never add paths that resurrect these leads.
4. **Idempotency where cheap.** Bounce processing skips already-`INVALID`
   leads; duplicate-campaign detection blocks re-sends from duplicate Lead IDs.
5. **Concurrency guard.** `runAutomation()` takes `LockService.getScriptLock()`
   (`tryLock(10000)`) and always releases in `finally`. Keep the lock window
   tight — never hold it across long sleeps if avoidable.
6. **Fail-safe config gates.** `AUTOMATION_ENABLED` and schedule checks
   default to *off/blocked* when misconfigured; a config typo should stop the
   pipeline, not accelerate it.

## Logging & observability

- `console.log(...)` for every decision point, including the *reason* a lead
  was skipped (e.g. `"Skipping X: AI personalisation is missing."`).
- `console.error(...)` for failures, always including the lead email and
  `error.message`.
- `logActivity(...)` (Dashboard.js) is the persistence layer for audit:
  call it on every SUCCESS and FAILED send, bounce, and campaign transition.
- Use `=== MARKERS ===` (`"=== AUTOMATION STARTED ==="`) for run boundaries.
- No secrets in logs: never log `GEMINI_API_KEY` or full config blobs.

## Extending the pipeline

When adding a new step or stage, follow this checklist:

1. Place the logic in the owning file (see table above).
2. Hook it into `runAutomation()` in the correct order with a marker log.
3. Add a stop-status / new status to the Leads contract (col O) and update
   the excel-sheet skill's status table if it changes the contract.
4. Add a `test*()` helper and verify via the Apps Script editor before
   `clasp push`.
5. If it consumes config, add the key through `getConfig()`/`getAIConfig()`
   with a defensive `Number()`/`String()` read and a descriptive error on
   invalid values.
6. Never bypass existing safety guards to accommodate the new step.

## Refactoring rules

- Renames must be global: Apps Script has no compiler. Grep every `.js` file
  for the symbol before renaming (the duplicate `testResearchPersonalisation`
  in AI.js was deduplicated during the HTML-email refactor; one copy remains).
- Column-index changes ripple: update the excel-sheet skill map, the inline
  comments, and every `data[i][n]` consumer in one change set.
- Preserve function names that may be referenced by triggers, the Dashboard
  UI, or clasp-pushed versions — if a rename is required, confirm with the
  user first.
- Batch reads/writes before micro-optimising; GAS quota math beats loop
  micro-optimisations every time.
