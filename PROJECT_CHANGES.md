# PROJECT_CHANGES.md

## 0. Git-Style Implementation Plan (committed before any change)

The work is split into ordered, independently verifiable commits (one
architectural area at a time). Each step preserves behaviour and keeps the
system deployable.

### Step 1 — Data-layer cleanup (new file `Leads.js`)
- Introduce a single lead model: column constants, status constants,
  `buildLeadFromRow()`, cached `getLeadsSheet()`/`getLeadsData()`.
- Remove the five duplicated lead-parsing blocks
  (EmailEngine.js `processEmails`/`processFollowUps`, AI.js
  `generateMissingPersonalizations`, Bounces.js `checkBounces`,
  Dashboard.js `generatePreview`).
- Add `FOLLOWUP_STEPS` map to collapse the three near-identical
  follow-up branches in `processFollowUps()`.
- **Verify:** `node --check`, grep for stale `data[i][N]` consumers,
  per-file `test*()` helpers still run.

### Step 2 — Shared boolean parsing (Modes.js)
- Add `isFlagTrue(value)` and refactor `isTestMode`/`isPreviewMode`/
  `isFastTestMode`, EmailEngine `AUTOMATION_ENABLED` + `AI_ENABLED`,
  AI.js `AI_ENABLED`, Campaigns.js service `Active`.
- **Verify:** grep all `toUpperCase() === "TRUE"` sites resolved.

### Step 3 — Template system redesign (Templates.js)
- Template record: service, step, variantId, subject, plain-text body,
  HTML template, CTA text/URL, theme, hero image, portfolio image,
  status, notes. Columns A–D of the Templates sheet keep their old
  positions (backwards compatible); new columns are appended E–M.
- `getTemplate(service, step)` kept (first-match, legacy callers).
- New `getTemplateVariants()`, `getRandomTemplate()` (random ACTIVE
  variant — subject travels with the variant, never randomized
  independently). In-memory per-execution cache.
- Generic placeholder system: unlimited placeholders resolved from
  template row → lead → config → computed values. New aliases
  (`{{Business}}`, `{{Name}}`) and new tokens (`{{City}}`,
  `{{RecentObservation}}`, `{{Competitor}}`, `{{ServiceBenefit}}`,
  `{{Offer}}`, `{{PortfolioUrl}}`, `{{CurrentYear}}`, company/social
  keys, etc.). Unresolved placeholders remain visible for QA.
- **Verify:** Node smoke harness personalises a sample lead; no
  leftover tokens in subject/body.

### Step 4 — HTML Email Engine (new file `HtmlEmailEngine.js`)
- Component system: Header, Hero, FeatureList, PortfolioShowcase,
  CTA Button, Divider, SignatureBlock, SocialLinks, Footer,
  UnsubscribeBlock. Reusable across services.
- Theme system: per-service default themes (Website=blue, SEO=green,
  Google Ads=orange) with built-in fallbacks; optional Themes sheet
  override; per-template Theme column override.
- `renderHtmlEmail(template, lead)` → `{ subject, htmlBody,
  plainTextBody }`. HTML shell: 600 px, table layout, inline CSS,
  Google-safe fonts, Outlook `bgcolor` fallbacks, rounded cards,
  alt text on images.
- MIME builder `buildMultipartAlternative()` for threaded Gmail API
  follow-ups (multipart/alternative with plain-text fallback).
- **Verify:** Node smoke harness renders HTML for every
  service/step/variant; MIME builder output structurally valid.

### Step 5 — Send-flow integration (EmailEngine.js)
- `processEmails()`: random variant selection, HTML + plain send via
  `GmailApp.sendEmail(..., { htmlBody })`.
- `sendFollowUp()`: variant selection + HTML engine; threaded send via
  MIME multipart/alternative. Follow-up subject: variant subject when
  present, otherwise the thread's existing subject (threading is
  preserved via In-Reply-To/References in either case).
- `processFollowUps()` refactored onto `FOLLOWUP_STEPS` map
  (identical behaviour, one branch instead of three).
- All safety guards untouched: `canSendEmail`, 3–4 s jitter, thread-ID
  requirement, stop-statuses, schedule gates.
- **Verify:** `node --check`; grep every send path logs; smoke harness
  on personalise + render paths.

### Step 6 — Preview system (Dashboard.js)
- `buildEmailPreview()` renders the full HTML engine output and keeps
  legacy `.body` (plain text).
- `generatePreview()` writes new Preview columns
  (Subject, Variant ID, Plain Text, HTML) — headers updated in the
  workbook.
- **Verify:** harness renders preview payloads for sample leads.

### Step 7 — Remaining cleanup (AI.js, Bounces.js, Campaigns.js)
- Deduplicate the double `testResearchPersonalisation` in AI.js
  (known debt, behaviour-preserving).
- Switch lead parsing to `buildLeadFromRow()`; use `isFlagTrue()`.
- **Verify:** `node --check`; grep for renamed/removed symbols.

### Step 8 — Spreadsheet redesign (Leads.xlsx via openpyxl script)
- Templates sheet: new headers (Service, Campaign Step, Subject,
  Plain Text Template, Variant ID, HTML Template, CTA Button Text,
  CTA URL, Theme, Hero Image, Portfolio Image, Status, Notes) and a
  full content set: 3 services × 4 steps × 2 variants with distinct
  follow-up angles (intro → reminder → case study → final).
- New Themes sheet (website/seo/google-ads colour sets).
- Config sheet: appenditive branding/social/portfolio/CTA/service
  benefit keys; EMAIL_SIGNATURE upgraded to placeholders.
- Preview sheet headers updated.
- Leads, Services, ActivityLog, AIResearch, AIConfig, Dashboard
  untouched (existing data preserved).
- **Verify:** reopen workbook; check headers, row counts, no broken
  cells.

### Step 9 — Verification & docs
- `node --check` every file; Node smoke harness covers template
  selection, placeholders, HTML render, MIME build.
- Update skill docs (excel-sheet, html-email, apps-script,
  software-architect, email-deliverability) to the new contracts.
- Complete this document (sections 1–10 below).

---

## 1. Architecture Overview

Google Apps Script (V8) application that automates B2B cold-email campaigns
from a bound Google Spreadsheet (deployed with clasp). The runtime flow,
unchanged in shape:

```
Trigger (time-based) ──► runAutomation() [EmailEngine.js]
                           │ LockService guard
                           │ Config gate (AUTOMATION_ENABLED, isFlagTrue)
                           │ Schedule gate (Scheduling.js)
                           ▼
                   1. checkBounces()                    [Bounces.js]
                   2. generateMissingPersonalizations() [AI.js]
                   3. processFollowUps()                [EmailEngine.js]
                   4. processEmails()                   [EmailEngine.js]
```

Pipeline order is deliberate and preserved: bounces first, AI intros next,
follow-ups before new emails, EMAIL_1 last.

What changed:

- **Shared data layer (`Leads.js`)** — one canonical lead model
  (`LEADS_COL` 0-based column map, `LEAD_STATUS`, `STOP_STATUSES`,
  `FOLLOWUP_STEPS`) replacing five duplicated row-parsing blocks across
  EmailEngine.js, AI.js, Bounces.js, Dashboard.js.
- **Template system (`Templates.js`)** — variants + random ACTIVE selection
  with the subject bound to its variant; generic placeholder engine
  (context → lead → config → computed, nested expansion).
- **HTML email engine (`HtmlEmailEngine.js`)** — component-based HTML
  rendering with per-service themes, plain-text fallback, and a
  multipart/alternative MIME builder for the threaded send path.
- **Send flow (`EmailEngine.js`)** — EMAIL_1 sends HTML via
  `GmailApp.sendEmail` `htmlBody`; follow-ups send multipart MIME via the
  Gmail advanced service; `processFollowUps` collapsed onto the
  `FOLLOWUP_STEPS` map (behaviour-preserving).
- **Preview (`Dashboard.js`)** — `buildEmailPreview` returns subject,
  plain text, and HTML; `generatePreview` writes 10 preview columns.

Every safety guard is intact: `DAILY_LIMIT` (`canSendEmail`), 3–4 s jitter
(`_getRandomDelay`), Gemini rate limiter, bounce checks before sends,
thread-ID requirements, stop-statuses, `AUTOMATION_ENABLED`, schedule
windows, per-lead try/catch.

---

## 2. Files Modified

| File | Change |
| --- | --- |
| `Leads.js` | **New.** Column map, statuses, `FOLLOWUP_STEPS`, cached sheet reads, `buildLeadFromRow()`. |
| `Templates.js` | **Rewritten.** Variant lookup, random selection, generic placeholder system, per-execution cache, tests. |
| `HtmlEmailEngine.js` | **New.** Themes, components, shell, `renderHtmlEmail()`, `buildMultipartAlternative()`, tests. |
| `EmailEngine.js` | `processEmails`/`processFollowUps` on shared layer + `FOLLOWUP_STEPS`; HTML sends both paths; `AUTOMATION_ENABLED`/`getDelay` via `isFlagTrue`/`isFastTestMode`; `_getRandomDelay` now `const`. |
| `Dashboard.js` | `buildEmailPreview` returns `{subject, body, htmlBody, plainTextBody, variantId, theme, template}`; `generatePreview` writes 10 columns. |
| `AI.js` | `buildLeadFromRow`/`isFlagTrue`; `_cleanGeminiJson`; duplicate `testResearchPersonalisation` removed (one copy kept). |
| `Bounces.js` | `buildLeadFromRow`; Status→`INVALID` write restored alongside Last Updated. |
| `Campaigns.js` | `isFlagTrue`; indentation fix in `finishCurrentCampaign`. |
| `Modes.js` | New `isFlagTrue()`; mode helpers refactored onto it. |
| `Leads.xlsx` | Templates/Themes/Config/Preview redesigned (see §3). |
| `PROJECT_CHANGES.md` | Plan + this documentation. |

Unchanged: `Config.js`, `Scheduling.js`, `Validation.js`, `Tests.js`,
`appsscript.json`, `.clasp.json`.

---

## 3. Spreadsheet Changes

**Templates sheet** (headers row 1, data row 2+). Columns A–D keep legacy
positions; E–M are new:

| Col | Header | Meaning |
| --- | --- | --- |
| A | Service | `Website Development` / `SEO` / `Google Ads` |
| B | Campaign Step | `EMAIL_1`, `FOLLOWUP_1`..`FOLLOWUP_3` |
| C | Subject | template with `{{Placeholders}}`; travels with its variant |
| D | Plain Text Template | legacy Body column |
| E | Variant ID | 1, 2, ... one row per variant |
| F | HTML Template | full skeleton with `{{Component}}` tokens (blank → plain-text-only) |
| G | CTA Button Text | falls back to `DEFAULT_CTA_TEXT` |
| H | CTA URL | falls back to `DEFAULT_CTA_URL` → `PORTFOLIO_URL` |
| I | Theme | blank = service default (website/seo/google-ads) |
| J | Hero Image | public URL, optional |
| K | Portfolio Image | public URL, optional |
| L | Status | `ACTIVE`/`INACTIVE` (blank counts as active) |
| M | Notes | free text |

Content: 3 services × 4 steps × 2 variants (24 rows) with distinct follow-up
angles per service (intro → reminder → case study → final ask).

**New Themes sheet** — per-theme overrides: Theme | Primary | Background |
Text | Muted | Accent | Border | Font Stack | Status | Notes. Rows with
Status `Active` override the built-in defaults in `HtmlEmailEngine.js`.

**Config sheet** — appenditive keys only:
`COMPANY_NAME`, `COMPANY_WEBSITE`, `COMPANY_EMAIL`, `COMPANY_PHONE`,
`COMPANY_ADDRESS`, `SOCIAL_LINKEDIN`, `SOCIAL_TWITTER`, `SOCIAL_GITHUB`,
`PORTFOLIO_URL`, `DEFAULT_CTA_TEXT`, `DEFAULT_CTA_URL`,
`UNSUBSCRIBE_MESSAGE`, `SENDER_ROLE`, `HERO_TITLE`, `HERO_SUBTITLE`,
`SERVICE_BENEFIT_*` (×3), `SERVICE_FEATURES_*` (×3, JSON string arrays).
`EMAIL_SIGNATURE` upgraded to placeholders:
`"Regards,\n{{SenderName}}\n{{SenderRole}}, {{CompanyName}}"`.

**Preview sheet headers** updated for the 10 output columns of
`generatePreview`.

Untouched (data preserved): Leads (A1:R1000), Services, ActivityLog (212
rows), AIResearch, AIConfig, Dashboard.

---

## 4. HTML Template System

- **Skeleton** (variant column F): 600 px table layout, inline CSS only,
  web-safe fonts, `bgcolor` fallbacks for Outlook, alt text on images.
- **Component tokens** replaced at render time: `{{Header}}`, `{{Hero}}`,
  `{{FeatureList}}`, `{{PortfolioShowcase}}`, `{{CtaButton}}`,
  `{{Divider}}`, `{{SignatureBlock}}`, `{{SocialLinks}}`, `{{Footer}}`,
  `{{UnsubscribeBlock}}`. Missing data strips the token to a blank line.
- **Themes**: built-in defaults per service — website `#2563EB` (blue), SEO
  `#059669` (green), Google Ads `#EA580C` (orange) — overridden by the
  variant's Theme column, then by `Active` Themes-sheet rows.
- **Placeholders**: any `{{Token}}` resolves via `buildPlaceholderValues`
  (precedence: context → lead → config → computed), with aliases
  (`{{Business}}`, `{{FirstName}}`, `{{ServiceName}}`) and optional fields
  (`{{City}}`, `{{RecentObservation}}`, `{{Competitor}}`,
  `{{WebsiteIssue}}`, `{{Offer}}`). Values containing tokens get up to 3
  nested-expansion passes; unresolved tokens stay visible for QA.
- **Plain text**: `htmlToPlainText()` derives a text version; email is
  always sent with both (GmailApp `htmlBody` / MIME multipart/alternative).
- Variant subjects are never randomized independently — the subject rides
  with the selected variant; threaded follow-ups fall back to the thread
  subject when the variant subject is blank.

---

## 5. Campaign Flow

Unchanged end-to-end pipeline, now on the shared data layer:

1. `NEW` lead → AI intro (`AI.js`) → EMAIL_1 (random variant, HTML + plain,
   thread ID stored in P) → status `EMAIL_1_SENT`.
2. After follow-up delays, `processFollowUps` advances through
   `FOLLOWUP_STEPS`: `EMAIL_1_SENT → FOLLOWUP_1 (K, 1)`,
   `FOLLOWUP_1_SENT → FOLLOWUP_2 (L, 2)`, `FOLLOWUP_2_SENT → FOLLOWUP_3
   (M, 3)`; each follow-up is threaded onto the stored thread via
   `sendThreadedFollowUp`.
3. `FOLLOWUP_3_SENT` → `finishCurrentCampaign` rotates service
   (Website Development → SEO → Google Ads) after `NEXT_SERVICE_*` delay;
   after the last service the lead is `COMPLETED`.
4. `REPLIED` / `INVALID` / `DO_NOT_CONTACT` are terminal stop-statuses —
   automation never touches them again.

---

## 6. Folder Structure

```
.
├── EmailEngine.js        # Orchestration + send mechanics (HTML wiring)
├── Leads.js              # Lead model / sheet contract (NEW)
├── Templates.js          # Template lookup, variants, placeholders
├── HtmlEmailEngine.js    # HTML render, themes, MIME builder (NEW)
├── AI.js                 # Gemini: research, intros, rate limiter
├── Bounces.js            # Bounce detection
├── Campaigns.js          # Service rotation, lifecycle
├── Config.js             # Config/AIConfig readers (cache)
├── Dashboard.js          # Setup, previews, activity log
├── Modes.js              # Test/preview/fast modes, isFlagTrue
├── Scheduling.js         # Working days, windows
├── Validation.js         # canSendEmail, daily count
├── Tests.js              # Cross-cutting tests
├── Leads.xlsx            # Bound spreadsheet (source of truth)
├── appsscript.json       # Manifest (V8, Gmail advanced service)
├── .clasp.json           # clasp deployment config
├── AGENTS.md             # Engineering rules
├── PROJECT_CHANGES.md    # This document
└── .opencode/skills/     # 5 skills docs (updated to new contracts)
```

---

## 7. Configuration

Everything below comes from the Config sheet via `getConfig()`; templates
reference them with `{{PascalCase}}` tokens. New/updated keys:

- **Brand**: `COMPANY_NAME`, `COMPANY_WEBSITE`, `COMPANY_EMAIL`,
  `COMPANY_PHONE`, `COMPANY_ADDRESS`
- **Social**: `SOCIAL_LINKEDIN`, `SOCIAL_TWITTER`, `SOCIAL_GITHUB`
- **Portfolio/CTA**: `PORTFOLIO_URL`, `DEFAULT_CTA_TEXT`, `DEFAULT_CTA_URL`
- **Footer**: `UNSUBSCRIBE_MESSAGE`
- **Hero**: `HERO_TITLE`, `HERO_SUBTITLE` (may contain other tokens)
- **Service-aware**: `SERVICE_BENEFIT_WEBSITE_DEVELOPMENT`,
  `SERVICE_BENEFIT_SEO`, `SERVICE_BENEFIT_GOOGLE_ADS` and
  `SERVICE_FEATURES_*` JSON arrays (`[ "...", ... ]` → `• item` bullets)
- **Sender**: `SENDER_ROLE`; `EMAIL_SIGNATURE` now uses `{{SenderName}}`,
  `{{SenderRole}}`, `{{CompanyName}}`

All pre-existing keys are unchanged and still honoured (limits, delays,
modes, timezone, AIConfig).

---

## 8. Testing Guide

Apps Script has no compiler — every change was verified before deploy:

1. **Syntax**: `node --check` on all 13 `.js` files (all pass).
2. **Smoke harness** (Node with mocks for SpreadsheetApp/GmailApp/Gmail/
   CacheService/Utilities, driven by a live export of the workbook at
   `/var/folders/rl/4bzbbmys42s74sccm2wl3vf00000gn/T/opencode/smoke_test.js`
   + `sheets.json`): **102 checks pass**, including:
   - `getTemplateVariants` returns 2 ACTIVE variants for every
     service/step pair; `getRandomTemplate` hits both variants over 50 draws.
   - Zero leftover placeholders across all 24 variants (subject, plain,
     HTML), nested expansion, CTA/hero defaults.
   - Theme resolution (default + Themes-sheet override), HTML render
     structure (table shell, inline CSS, alt text).
   - MIME structure of `buildMultipartAlternative` (multipart/alternative,
     plain + HTML parts, In-Reply-To/References).
   - End-to-end: `processEmails` passes `htmlBody` to
     `GmailApp.sendEmail`; `processFollowUps` sends threaded MIME via the
     Gmail API; blank variant subject falls back to thread subject.
3. **In-editor tests**: `testTemplateVariants`,
   `testRandomTemplateSelection`, `testPlaceholderExpansion`,
   `testHtmlEmailRender`, `testBuildMultipartAlternative` plus all
   pre-existing `test*()` helpers run from the Apps Script editor.
4. **Workbook**: `Leads.backup.xlsx` is the pre-change backup; the migration
   script (`redesign_workbook.py`, temp dir) can regenerate the layout.

Run order after any change: `node --check` → smoke harness → in-editor
tests → preview via `generatePreview` in TEST_MODE before touching
production.

---

## 9. Deployment Guide

1. `clasp push` (deploys all `.js` + `appsscript.json` to the bound script;
   scriptId in `.clasp.json`).
2. Verify in the Apps Script editor: run `testTemplateVariants` and
   `testHtmlEmailRender`; confirm Templates/Themes/Config sheets are as
   documented in §3.
3. Sanity-send in `PREVIEW_MODE`/`TEST_MODE` first (preview sheet, then a
   real test send to `TEST_EMAIL`).
4. Re-enable/confirm the time-based trigger for `runAutomation()`.
5. If the spreadsheet was re-uploaded from `Leads.xlsx`, re-verify sheet
   names and the Config/AIConfig keys — the Apps Script project is bound by
   spreadsheet ID, and content changes are already in the workbook.

---

## 10. Future Improvements

- **A/B testing analytics**: track variantId → reply/status outcomes in
  ActivityLog to steer variant weights.
- **Dashboard UI** upgrade to surface variant distribution and per-theme
  previews.
- **Replies.js** — automatic reply detection (`Gmail.Users.Threads.list`)
  to set `REPLIED` without manual review.
- **More variants per step** — content set currently ships 2 per
  service/step; the selection system supports any number.
- **Custom Themes sheet rows** — the mechanism exists; default sheet rows
  for bespoke client brands can be added anytime without code changes.
- **Optional fields on Leads sheet** — `City`, `RecentObservation`,
  `Competitor`, `WebsiteIssue`, `Offer` placeholders are supported; adding
  the columns enables richer personalisation.
- **Track opens** only if the user explicitly opts in — current policy is
  no tracking pixels (deliverability-first).

---

## 11. Hardening & Frontier Port (2026-09-03)

Ported the reusable engineering features from the frontier project
(Doctor-Email-Automation) into this public repo. Nothing private was
carried over: no real lead data, addresses, phone numbers, brand names,
logos, script IDs or project-specific window dates. Everything doctor- or
clinic-specific was either generalised (config-driven) or left out.

### New files
- `Replies.js` — automatic human-reply detection. Thread scan via
  `Gmail.Users.Threads.get` (metadata format) plus a mailbox fallback
  (recent-INBOX list + `from:` search) for replies that start a new
  thread. Filters auto-replies/bounces via `Auto-Submitted`/`Precedence`/
  `X-Auto-Response-Suppress` and never marks a message newer than the
  lead's last send as a reply. `REPLIED` is terminal, so each lead flips
  exactly once.
- `Unsubscribe.js` — one-click unsubscribe web app (RFC 8058). HMAC-signed
  per-lead links (`UNSUB_TOKEN_SECRET` in Script Properties), GET
  confirmation page, POST one-click endpoint, writes
  `DO_NOT_CONTACT` + SuppressionList `UNSUBSCRIBED` + ActivityLog
  `UNSUBSCRIBE`. Confirmation page is generic (company name from
  `COMPANY_NAME` config), no external image assets.
- `SuppressionReconcile.js` — 30-day Gmail DSN backfill
  (`reconcileMailboxBounces`, dry-run default), historical suppression
  audit (`dryRunHistoricalSuppression`), and config-driven domain blocking
  (`SUPPRESSED_DOMAINS` in Config) with `dryRunDomainSuppression` /
  `liveDomainSuppression`. No hardcoded provider domains.
- `TestModeFix.js` — `reconcileTestDates` (FAST_TEST minute-scale dates vs
  Gmail `internalDate`), `purgeTestActivityLog` (minute-burst detection,
  archive-then-delete), `ensureLocalTimeDisplayHelper`. Windows are
  parameters, not hardcoded dates.
- `VerificationCleanup.js` — `applyVerificationCleanup` applies a bulk
  verification export (undeliverable → HARD_BOUNCE, risky → UNKNOWN).
  Lists are empty placeholders; no lead data committed.

### Leads contract (A–V)
- Appended columns: S Bounce Category, T Bounce Diagnostic, U Suppressed
  At, V Retry Count (`ensureLeadsBounceColumns`).
- New `SuppressionList` sheet (Email | Category | Diagnostic | First Seen |
  Last Seen | Count | Source) with `isSuppressed` /
  `getSuppressionCategory` / `addToSuppressionList` (idempotent).
- New `ACTIVE_CAMPAIGN_STATUSES`; `buildLeadFromRow` now exposes
  `bounceCategory`, `bounceDiagnostic`, `suppressedAt`, `retryCount`.

### Send-path hardening (EmailEngine.js / Validation.js)
- `MAX_SENDS_PER_RUN` per-execution send budget.
- `NEW_LEADS_DAILY_FLOOR` (default 8% of `DAILY_LIMIT`, min 2) so NEW
  leads are never starved by the follow-up queue; overdue follow-ups
  (>24 h past due) keep priority.
- 299 s execution self-cap (`_runDeadline`) — lock is always released.
- Gmail quota circuit breaker (`tripQuotaBreaker`, blocks until midnight,
  `QUOTA_BLOCK_UNTIL` Script Property).
- Hourly cap (`HOURLY_LIMIT`) + per-group caps (`HOURLY_LIMIT_YAHOO`,
  `HOURLY_LIMIT_OUTLOOK`, `HOURLY_LIMIT_ICLOUD`).
- Suppression + blocked-domain gates before every send path.
- EMAIL_1 now sends raw multipart/alternative via the Gmail API with an
  authoritative `threadId` (no post-send search), matching follow-ups.
- `List-Unsubscribe` / `List-Unsubscribe-Post` headers (RFC 8058) and
  `{{UnsubscribeLink}}` in the HTML footer + plain-text fallback.
- `runAutomation` pipeline: bounces → replies → AI personalisation →
  follow-ups → email 1, with deadline checks between stages; optional
  `SKIP_REPLIES_IN_MAIN_RUN` for a dedicated reply trigger.
- Bounce classifier: 4 categories (HARD/SOFT/POLICY/UNKNOWN), soft-retry
  cap (3 → UNKNOWN), message-ID dedup cache, 7-day DSN window.

### Cleanups applied
- Removed dead `getDelay()` (superseded by `getFollowUpDelay`).
- Removed legacy `body` field from `buildEmailPreview` return.
- Renamed `Tests.js: testBounceParser` → `testExtractBouncedEmail`
  (classifier test lives next to the classifier in Bounces.js — no
  duplicate function names).
- Dashboard mode cells now parse text booleans via `isFlagTrue`; Failed
  Actions KPI excludes quota-flood rows.
- Added `.claspignore` so docs/workbooks never reach the Apps Script
  project.

### Config keys added by `ensureHardeningMigration`
`HOURLY_LIMIT`, `HOURLY_LIMIT_YAHOO`, `HOURLY_LIMIT_OUTLOOK`,
`HOURLY_LIMIT_ICLOUD`, `MAX_SENDS_PER_RUN`, `NEW_LEADS_DAILY_FLOOR`,
`SUPPRESSED_DOMAINS`, `SKIP_REPLIES_IN_MAIN_RUN`, `SEND_WINDOWS`,
`UNSUBSCRIBE_URL` (set after deploying the web app).

### Post-deploy checklist
1. Run `ensureHardeningMigration()` once.
2. Deploy Unsubscribe.js as a web app; paste `UNSUBSCRIBE_URL` in Config.
3. Optional: add a 5-minute `checkReplies()` trigger and set
   `SKIP_REPLIES_IN_MAIN_RUN=TRUE`.
4. Dry-run the one-time helpers before any live run
   (`reconcileMailboxBounces(true)`, `applyVerificationCleanup(true)`).
