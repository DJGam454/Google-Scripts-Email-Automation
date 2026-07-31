# AGENTS.md — Engineering Rules

Global rules for working in this repository. Applies to every task unless the
user explicitly overrides it.

## Project in one paragraph

Google Apps Script (V8) application that automates B2B cold-email campaigns
from a bound Google Spreadsheet: Gemini generates personalised intros, Gmail
sends initial emails and threaded follow-ups, bounces are detected and leads
marked INVALID, and leads rotate through services (Website Development → SEO →
Google Ads) until all campaigns complete. Deployed with clasp (`scriptId` in
`.clasp.json`).

## Domain-specific skills (load before relevant work)

This repository ships OpenCode skills under `.opencode/skills/`. Consult the
matching skill before doing related work:

- `apps-script` — editing/deploying the GAS code, services, quotas, conventions
- `html-email` — email template content, MIME, previewing, personalisation
- `excel-sheet` — Leads/Config/AIConfig/Templates/Services/ActivityLog/AIResearch sheet contracts, column map A–R, statuses
- `email-deliverability` — daily limits, pacing, test modes, bounces, threading, spam safety
- `software-architect` — file ownership, naming, pipeline order, refactoring rules

## Global rules

1. **Do not modify application code unless asked.** This repo's source files
   (`*.js`, `appsscript.json`, `.clasp.json`) are live automation code — any
   change can send real email. When the task is only documentation,
   configuration, or skills, leave application code untouched.
2. **Never weaken safety guards.** The following are non-negotiable: `DAILY_LIMIT`
   enforcement (`canSendEmail`), random 3–4 s send delays, Gemini rate limiting
   (`AI_RPM_LIMIT`), bounce checks before sends, thread-ID requirements for
   follow-ups, stop-statuses (`REPLIED` / `INVALID` / `DO_NOT_CONTACT`),
   `AUTOMATION_ENABLED` gate, and schedule windows.
3. **Test mode before production.** Anything that sends mail should be
   verified under `TEST_MODE`/`FAST_TEST_MODE` first. Never bypass test-mode
   logic.
4. **Secrets never enter code or logs.** `GEMINI_API_KEY` lives in Script
   Properties only. No API keys, tokens, or full config dumps in source files,
   logs, or commits.
5. **No new dependencies.** No npm packages, no bundlers, no ES modules.
   Everything runs as plain global-scope `.js` in Apps Script V8. External
   calls use `UrlFetchApp` only.
6. **Match the existing style.** Banner comments, section dividers, blank
   lines between steps, `const`-first, defensive `String()`/`Number()` reads
   on sheet data, `console.log`/`console.error` on every decision and failure
   path.
7. **Own your file.** Follow the file-ownership table in the
   `software-architect` skill: send logic → `EmailEngine.js`, Gemini → `AI.js`,
   campaigns → `Campaigns.js`, config → `Config.js`, etc. Keep `test*()`
   helpers next to their code or in `Tests.js`.
8. **Preserve sheet contracts.** The Leads column map (A–R), statuses, and
   ActivityLog indexes in the `excel-sheet` skill are load-bearing. Any change
   must update every consumer plus the skill's contract documentation.
9. **One lead's failure never kills the run.** Per-lead `try/catch` + `continue`
   is the house pattern.
10. **Verify before finishing.** Apps Script has no compiler: grep the whole
    repo for renamed symbols, run the relevant `test*()` function in the Apps
    Script editor where possible, and only then suggest `clasp push`.
11. **No duplicate functions.** Check for existing definitions before adding a
    function (note: `testResearchPersonalisation` exists twice in `AI.js` —
    treat as known debt; deduplicating it is welcome only if it does not change
    behavior).
12. **Do not commit unless asked.** Git commits are only made on explicit user
    request, and never with secrets, `.clasp.json` tokens, or `Leads.xlsx`
    content changes beyond what the user asked for.
