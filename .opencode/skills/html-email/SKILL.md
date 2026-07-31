---
name: html-email
description: HTML email authoring and MIME construction for this cold-email automation. Use when writing or modifying email templates, htmlBody content, inline-CSS email markup, MIME messages, plain-text fallbacks, previewing emails via buildEmailPreview/generatePreview, or personalising templates with {{Placeholder}} tokens.
---

# HTML Email Authoring

## How emails are sent in this project

Both sending paths now carry **HTML with a plain-text fallback**:

1. `GmailApp.sendEmail(recipient, subject, body, { htmlBody: ... })` — initial
   `EMAIL_1` sends in `EmailEngine.js`; `renderHtmlEmail()` produces
   `{ subject, htmlBody, plainTextBody }`, so Gmail renders the HTML with the
   plain text as fallback.
2. `Gmail.Users.Messages.send()` via `sendThreadedFollowUp(lead, email,
   template)` — follow-ups built by `buildMultipartAlternative(email)` as raw
   multipart/alternative MIME (`multipart/alternative; boundary=...` with
   `text/plain` then `text/html` parts) plus `In-Reply-To`/`References`
   headers for threading. The variant's subject is used when non-empty;
   otherwise the thread subject is reused (threading is preserved either way).

Clients that block HTML or read plain text must never receive an empty
message — both paths always include plain text.

## Template system

Templates live in the **Templates** sheet (see excel-sheet skill for the full
A–M column map): Service | Step | Subject | Plain Text | Variant ID | HTML |
CTA Text | CTA URL | Theme | Hero Image | Portfolio Image | Status | Notes.

- `getTemplate(service, step)` — legacy first-match lookup.
- `getTemplateVariants(service, step)` — all ACTIVE variants.
- `getRandomTemplate(service, step)` — random ACTIVE variant; **the chosen
  variant carries its own subject** — subjects are never randomized
  independently of their template (see `EmailEngine.js`).
- `personaliseTemplate(text, lead, context)` — placeholder substitution.
- `renderHtmlEmail(template, lead)` in `HtmlEmailEngine.js` — renders the
  variant's HTML skeleton (with `{{Component}}` tokens), personalises
  subject/body/CTA, and returns `{ subject, htmlBody, plainTextBody, values,
  theme }`. If the variant has no HTML, the plain-text body is used for both.

## HTML skeleton & components

The variant's `html` column holds a full email skeleton (600 px table, inline
CSS) that may use component tokens, each replaced by `HtmlEmailEngine.js`:

| Token | Rendered by |
| --- | --- |
| `{{Header}}` | logo/brand bar |
| `{{Hero}}` | hero image (J) or title/subtitle block |
| `{{FeatureList}}` | `SERVICE_FEATURES_*` bullets (rendered from `{{Features}}`) |
| `{{PortfolioShowcase}}` | portfolio image (K) with `PORTFOLIO_URL` link |
| `{{CtaButton}}` | CTA button (G/H, falls back to `DEFAULT_CTA_*`) |
| `{{Divider}}` | horizontal rule |
| `{{SignatureBlock}}` | `{{EmailSignature}}` |
| `{{SocialLinks}}` | LinkedIn/Twitter/GitHub icons from Config |
| `{{Footer}}` | `{{UnsubscribeMessage}}`, company, `{{CurrentYear}}` |
| `{{UnsubscribeBlock}}` | plain unsubscribe line for the plain-text part |

If a component token appears in a skeleton and its component has nothing to
render (e.g. no social links configured), the token is stripped to a blank
line instead of erroring.

## Themes

Each template renders with a theme: built-in defaults per service (website
blue `#2563EB`, SEO green `#059669`, Google Ads orange `#EA580C`), overridden
by the template's Theme column (I), which itself can be overridden by an
`Active` row in the Themes sheet. Theme fields: primary, background, text,
muted, accent, border, font stack (see `DEFAULT_THEMES` in
`HtmlEmailEngine.js`).

## Supported placeholders

Any `{{Token}}` is resolved from a values map built in
`buildPlaceholderValues(lead, context)` (precedence: context → lead →
config → computed). Beyond the legacy set:

| Placeholder | Source |
| --- | --- |
| `{{FirstName}}` | Leads col C — contact name |
| `{{Company}}` | Leads col B (alias `{{Business}}`) |
| `{{Website}}` | Leads col E |
| `{{Industry}}` | Leads col F |
| `{{PersonalisedIntro}}` | Leads col G — AI-generated intro (never duplicate this in the template) |
| `{{Service}}` / `{{ServiceName}}` | Leads col H |
| `{{CurrentYear}}` | computed |
| `{{ServiceBenefit}}` | `SERVICE_BENEFIT_*` for the lead's service |
| `{{Features}}` | `SERVICE_FEATURES_*` JSON array → `• item` bullet list |
| `{{CtaText}}` / `{{CtaUrl}}` | variant G/H, then `DEFAULT_CTA_*` |
| `{{HeroTitle}}` / `{{HeroSubtitle}}` | Config (may contain other tokens) |
| `{{SenderName}}` | Config `SENDER_NAME` |
| `{{SenderRole}}` | Config `SENDER_ROLE` |
| `{{ReplyToEmail}}` | Config `REPLY_TO_EMAIL` |
| `{{EmailSignature}}` | Config `EMAIL_SIGNATURE` (may itself contain `{{SenderName}}`, `{{SenderRole}}`, `{{CompanyName}}` …) |
| `{{CompanyName}}`, `{{CompanyWebsite}}`, `{{CompanyEmail}}`, `{{CompanyPhone}}`, `{{CompanyAddress}}` | Config brand keys |
| `{{SocialLinkedIn}}`, `{{SocialTwitter}}`, `{{SocialGithub}}` | Config |
| `{{PortfolioUrl}}` | Config |
| `{{UnsubscribeMessage}}` | Config |
| `{{HeroImage}}`, `{{PortfolioImage}}` | variant J/K |

Any camelCase lead property maps to a PascalCase token (e.g. `lead.website`
→ `{{Website}}`), so optional fields like `city`, `recentObservation`,
`competitor`, `websiteIssue`, `offer` become placeholders when present.
Values may themselves contain `{{Tokens}}` — a nested-expansion pass resolves
them (up to 3 passes). Unresolved placeholders are left as-is so QA can spot
them.

Rules:

- Placeholders are substituted with `replaceAll` (via a replacer function so
  replacement values containing `$` are inserted literally); tokens must
  match exactly, case-sensitively, including the double braces.
- Never hard-code sender name, reply-to, signature, or company details inside
  a template — always use placeholders.
- The AI-generated intro must not contain a greeting; the template supplies
  `Hi {{FirstName}}`. Keep that separation intact.
- Subjects should stay under ~60 characters for deliverability and mobile
  rendering.

## HTML email best practices

If an HTML body is added to a template, follow these rules:

- **Inline CSS only.** Gmail strips `<style>` blocks and `<head>` content.
  Every style must be a `style="..."` attribute on the element.
- **Table-based layout.** Use `<table>`/`<td>` for structure, not `div`
  grids. Target a max width of 600 px.
- **No JavaScript, no forms, no external stylesheets, no background images
  on body.**
- **Web-safe fonts only** (Arial, Helvetica, Georgia, Verdana, Courier New)
  with fallbacks: `font-family: Arial, Helvetica, sans-serif`.
- **Gmail-specific caveats:** Gmail does not support `<video>`, CSS
  animations in some clients, `position: fixed`, or `@media` queries in all
  renderers. Design for degradation.
- Use `bgcolor` attributes on `<td>` alongside CSS for Outlook.
- Alt text on every image; images should never be the only content — if the
  recipient blocks images, the message must still make sense.
- Include a plain-text fallback via multipart/alternative when using raw MIME.
- Set explicit `width`/`height` on images and use `display:block` to avoid
  whitespace gaps.

## Cold-email constraints

This is a cold-outreach system. Deliverability trumps design:

- **Keep it text-like.** Sparse HTML with a personal tone performs better
  than marketing-grade design. A single `{{PersonalisedIntro}}`, a short body,
  and a signature.
- **Never embed tracking pixels or read-receipt scripts.** They hurt
  spam-scoring and are not used anywhere in this project — do not introduce
  them.
- **No attachments on cold emails** (also avoids spam filters and quota
  issues). Attachments are only acceptable if the user explicitly requests
  them.
- Avoid spam-trigger words, excessive punctuation (`!!!`), red-flag fonts
  (large red text), and excessive use of `font-weight: bold`.
- One primary link maximum in the body; use `{{Website}}` naturally rather
  than bare `http://` URLs with anchor text like "click here".

## Previewing before sending

Use the existing preview machinery instead of sending real mail:

- `buildEmailPreview(lead, step)` in `Dashboard.js` renders a random variant
  (subject + plain text + HTML + variant id + theme) for a lead without
  sending.
- `generatePreview()` writes 10 columns: Generated At, Lead ID, Email,
  Company, Service, Step, Subject, Variant ID, Plain Text Preview, HTML
  Preview.
- When validating a new template, call `personaliseTemplate()` on a sample
  lead and inspect the output for un-substituted placeholders before relying
  on it.

## Verification checklist

Before any email content change ships:

1. `getRandomTemplate(service, step)` returns an ACTIVE variant for every
   service/step pair the pipeline uses.
2. Every placeholder is substituted (no leftover `{{...}}` in subject, plain
   text, or HTML).
3. Plain-text path reads well; HTML path has inline CSS and a plain-text
   fallback (multipart/alternative on the threaded path).
4. Length checks: subject ≤ ~60 chars, body scannable in 5 seconds.
5. Personalisation intro ≤ `MAX_INTRO_WORDS` from AIConfig.
6. Previewed through `generatePreview()` and checked in an actual Gmail
   client render.
