// ============================================================
// EMAIL TEMPLATES
// ============================================================
// The Templates sheet is the single source of email content:
//
//   A  Service            (e.g. "Website Development")
//   B  Campaign Step      (EMAIL_1, FOLLOWUP_1..3)
//   C  Subject            (may contain {{Placeholders}})
//   D  Plain Text Template
//   E  Variant ID         (1, 2, 3... — multiple variants per step)
//   F  HTML Template      (full skeleton, may use {{Component}} tokens)
//   G  CTA Button Text
//   H  CTA URL
//   I  Theme              (blank = service default theme)
//   J  Hero Image         (public image URL, optional)
//   K  Portfolio Image    (public image URL, optional)
//   L  Status             (ACTIVE / INACTIVE)
//   M  Notes
//
// Columns A–D keep their original positions so legacy readers of
// Service | Step | Subject | Body keep working.

// ============================================================
// TEMPLATE LOOKUP (per-execution in-memory cache)
// ============================================================

var _templatesCache = null;

function _loadTemplates() {

  if (_templatesCache) {
    return _templatesCache;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Templates");

  if (!sheet) {
    throw new Error("Templates sheet not found.");
  }

  _templatesCache = sheet.getDataRange().getValues();

  return _templatesCache;
}

// Clears the in-memory template cache (useful after editing
// the Templates sheet inside the same execution, e.g. tests).
function refreshTemplatesCache() {
  _templatesCache = null;
}

// ============================================================
// TEMPLATE RECORD
// ============================================================

function _buildTemplateRecord(row) {

  const rawStatus = String(row[11] || "")
    .trim()
    .toUpperCase();

  return {
    service: row[0],
    step: row[1],
    subject: row[2],
    body: row[3],            // Plain text body (legacy "Body" column)
    variantId: row[4],
    html: row[5],            // HTML template (may be blank → plain text only)
    ctaText: row[6],
    ctaUrl: row[7],
    theme: row[8],
    heroImage: row[9],
    portfolioImage: row[10],
    // Blank status on legacy sheets counts as active.
    status: rawStatus || "ACTIVE",
    notes: row[12]
  };
}

// ============================================================
// GET TEMPLATE (legacy, first match)
// ============================================================
// Returns the first matching row for (service, step) or null.
// Kept for backwards compatibility; new code should prefer
// getRandomTemplate().

function getTemplate(service, step) {

  const data = _loadTemplates();

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === service &&
      String(data[i][1]) === step
    ) {
      return _buildTemplateRecord(data[i]);
    }
  }

  return null;
}

// ============================================================
// TEMPLATE VARIANTS
// ============================================================
// All ACTIVE variants for a (service, step) pair. Multiple
// variants per step enable future A/B testing; the send path
// picks one at random.

function getTemplateVariants(service, step) {

  const data = _loadTemplates();

  const variants = [];

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][0]) === service &&
      String(data[i][1]) === step
    ) {

      const record = _buildTemplateRecord(data[i]);

      if (record.status !== "ACTIVE") {
        continue;
      }

      variants.push(record);
    }
  }

  return variants;
}

// ============================================================
// RANDOM TEMPLATE SELECTION
// ============================================================
// Picks ONE random ACTIVE variant for (service, step). The
// chosen variant carries its own subject — subjects are never
// randomized independently of their template.
//
// Fallback: if the sheet has no ACTIVE variants (legacy layout
// without a Status column), first-match behaviour is preserved.

function getRandomTemplate(service, step) {

  const variants =
    getTemplateVariants(service, step);

  if (variants.length === 0) {
    return getTemplate(service, step);
  }

  const index =
    Math.floor(
      Math.random() * variants.length
    );

  return variants[index];
}

// ============================================================
// PLACEHOLDER PERSONALISATION
// ============================================================
// Generic system: ANY {{Token}} is resolved from a values map
// built from (in precedence order):
//
//   1. context   — template-row values (CTA, images) passed by caller
//   2. lead      — every lead field becomes a placeholder
//                  (Company, Name, Website, Industry,
//                   PersonalisedIntro, Service, Campaign, ...)
//                  plus aliases (Business, FirstName) and
//                  optional fields (City, RecentObservation,
//                  Competitor, WebsiteIssue, Offer, ...)
//   3. config    — SenderName, ReplyToEmail, CompanyName,
//                  CompanyWebsite, CompanyEmail, CompanyPhone,
//                  CompanyAddress, SocialLinkedIn, SocialTwitter,
//                  SocialGithub, PortfolioUrl, UnsubscribeMessage,
//                  SenderRole, ...
//   4. computed  — CurrentYear, ServiceBenefit, Features,
//                  HeroTitle, HeroSubtitle, EmailSignature
//                  (these may themselves contain {{Tokens}},
//                   expanded by the nested-expansion pass)
//
// Unresolved placeholders are left as-is so QA can spot them
// (see html-email skill: inspect for un-substituted tokens).

function personaliseTemplate(text, lead, context) {

  if (!text) {
    return "";
  }

  const values =
    buildPlaceholderValues(lead, context);

  return _applyPlaceholders(text, values);
}

// ============================================================
// VALUES MAP BUILDER
// ============================================================

function buildPlaceholderValues(lead, context) {

  const config = getConfig();

  const values = {};

  // --------------------------------
  // 1. COMPUTED DEFAULTS
  // --------------------------------

  values["CurrentYear"] =
    String(new Date().getFullYear());

  // --------------------------------
  // 2. LEAD FIELDS (generic pass-through)
  // --------------------------------
  // Any property on the lead object becomes a placeholder.
  // camelCase keys map to PascalCase tokens:
  // lead.recentObservation -> {{RecentObservation}}

  if (lead) {

    Object.keys(lead).forEach(function(key) {

      const token = _camelToPlaceholder(key);

      if (token) {
        values[token] = lead[key];
      }
    });

    // Friendly aliases
    values["FirstName"] = lead.name;
    values["Business"] = lead.company;
    values["ServiceName"] = lead.service;

    // Optional personalisation fields (empty when the Leads
    // sheet has no such columns — system stays generic).
    const optionalFields = [
      "city",
      "recentObservation",
      "competitor",
      "websiteIssue",
      "offer"
    ];

    optionalFields.forEach(function(field) {

      if (
        lead[field] !== undefined &&
        lead[field] !== null
      ) {
        values[_camelToPlaceholder(field)] = lead[field];
      }
    });
  }

  // --------------------------------
  // 3. CONFIG VALUES
  // --------------------------------
  // Token name -> Config sheet key. Tokens are PascalCase for
  // templates; Config sheet keys are ALL_CAPS.

  const configMapping = [
    ["SenderName", "SENDER_NAME"],
    ["ReplyToEmail", "REPLY_TO_EMAIL"],
    ["SenderRole", "SENDER_ROLE"],
    ["CompanyName", "COMPANY_NAME"],
    ["CompanyWebsite", "COMPANY_WEBSITE"],
    ["CompanyEmail", "COMPANY_EMAIL"],
    ["CompanyPhone", "COMPANY_PHONE"],
    ["CompanyAddress", "COMPANY_ADDRESS"],
    ["SocialLinkedIn", "SOCIAL_LINKEDIN"],
    ["SocialTwitter", "SOCIAL_TWITTER"],
    ["SocialGithub", "SOCIAL_GITHUB"],
    ["PortfolioUrl", "PORTFOLIO_URL"],
    ["UnsubscribeMessage", "UNSUBSCRIBE_MESSAGE"],
    ["EmailSignature", "EMAIL_SIGNATURE"]
  ];

  configMapping.forEach(function(pair) {

    if (config[pair[1]]) {
      values[pair[0]] = config[pair[1]];
    }
  });

  // --------------------------------
  // 4. CONTEXT (template-row) VALUES
  // --------------------------------
  // Highest precedence. Only non-empty values are applied so
  // config defaults survive blank template cells.

  if (context) {

    const contextKeys = [
      ["ctaText", "CtaText"],
      ["ctaUrl", "CtaUrl"],
      ["heroImage", "HeroImage"],
      ["portfolioImage", "PortfolioImage"]
    ];

    contextKeys.forEach(function(pair) {

      const value = context[pair[0]];

      if (value) {
        values[pair[1]] = value;
      }
    });
  }

  // --------------------------------
  // 5. COMPUTED (service-aware)
  // --------------------------------

  const serviceKey = _serviceConfigKey(
    lead && lead.service
  );

  const serviceBenefit =
    config["SERVICE_BENEFIT_" + serviceKey];

  if (serviceBenefit) {
    values["ServiceBenefit"] = serviceBenefit;
  }

  const featuresRaw =
    config["SERVICE_FEATURES_" + serviceKey];

  if (featuresRaw) {

    let features = [];

    try {
      features = JSON.parse(featuresRaw);
    } catch (error) {
      features = [];
    }

    if (!Array.isArray(features)) {
      features = [];
    }

    values["Features"] = features;
  }

  // Defaults for CTA and hero text
  if (!values["CtaText"]) {
    values["CtaText"] =
      config.DEFAULT_CTA_TEXT ||
      "Reply to this email";
  }

  if (!values["CtaUrl"]) {
    values["CtaUrl"] =
      config.DEFAULT_CTA_URL ||
      config.PORTFOLIO_URL ||
      config.COMPANY_WEBSITE ||
      "";
  }

  values["HeroTitle"] =
    config.HERO_TITLE ||
    "Helping {{Company}} grow";

  values["HeroSubtitle"] =
    config.HERO_SUBTITLE ||
    "A short note from {{SenderName}}";

  // --------------------------------
  // UNSUBSCRIBE DEEP LINK
  // --------------------------------
  // Built only when a web-app URL is configured in the Config
  // sheet (UNSUBSCRIBE_URL). The recipient's address and Lead ID
  // are URL-encoded so doGet (Unsubscribe.js) can mark the lead
  // DO_NOT_CONTACT with one click.

  const unsubscribeUrl = String(
    config.UNSUBSCRIBE_URL || ""
  ).trim();

  if (unsubscribeUrl && lead) {

    // Prefer the HMAC-signed link so the List-Unsubscribe header
    // and the visible footer link are identical and One-Click
    // verifiable. Fall back to the plain link if signing is not
    // available (e.g. during early bootstrap before PropertiesService
    // is accessible).
    let signed = "";

    try {
      if (typeof buildSignedUnsubscribeLink === "function") {
        signed = buildSignedUnsubscribeLink(lead, config);
      }
    } catch (e) {
      signed = "";
    }

    if (signed) {
      values["UnsubscribeLink"] = signed;
    } else {
      values["UnsubscribeLink"] =
        unsubscribeUrl +
        "?email=" +
        encodeURIComponent(String(lead.email || "")) +
        "&id=" +
        encodeURIComponent(String(lead.leadId || ""));
    }
  }

  // --------------------------------
  // 6. NESTED EXPANSION
  // --------------------------------
  // Values such as {{EmailSignature}} or {{HeroTitle}} may
  // themselves contain {{Tokens}} — expand up to 3 passes.

  for (let pass = 0; pass < 3; pass++) {

    let changed = false;

    Object.keys(values).forEach(function(key) {

      const value = values[key];

      if (
        typeof value !== "string" ||
        value.indexOf("{{") === -1
      ) {
        return;
      }

      const expanded =
        _applyPlaceholders(value, values);

      if (expanded !== value) {
        values[key] = expanded;
        changed = true;
      }
    });

    if (!changed) {
      break;
    }
  }

  return values;
}

// Substitute every {{Token}} in text using a pre-built values
// map (see buildPlaceholderValues). Public wrapper so the HTML
// email engine can reuse the same substitution logic.
function applyPlaceholderValues(text, values) {

  if (!text) {
    return "";
  }

  return _applyPlaceholders(text, values);
}

// ============================================================
// INTERNAL HELPERS
// ============================================================

// camelCase -> PascalCase token ("recentObservation" ->
// "RecentObservation", "leadId" -> "LeadId").
function _camelToPlaceholder(key) {

  if (!key) {
    return "";
  }

  return String(key)
    .charAt(0)
    .toUpperCase() +
    String(key).slice(1);
}

// Service name -> config key suffix
// ("Website Development" -> "WEBSITE_DEVELOPMENT").
function _serviceConfigKey(service) {

  if (!service) {
    return "";
  }

  return String(service)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "_");
}

// Substitute every {{Token}} in text using the values map.
// Uses a replacer function so replacement values containing
// "$" are inserted literally.
function _applyPlaceholders(text, values) {

  let output = String(text);

  Object.keys(values).forEach(function(key) {

    const value = values[key];

    if (value === undefined || value === null) {
      return;
    }

    const replacement = Array.isArray(value)
      ? value.map(function(item) {
          return "\u2022 " + item;
        }).join("\n")
      : String(value);

    const token = "{{" + key + "}}";

    if (output.indexOf(token) === -1) {
      return;
    }

    output = output.replaceAll(token, function() {
      return replacement;
    });
  });

  return output;
}

// ============================================================
// TESTS
// ============================================================

function testTemplateVariants() {

  const services = [
    "Website Development",
    "SEO",
    "Google Ads"
  ];

  const steps = [
    "EMAIL_1",
    "FOLLOWUP_1",
    "FOLLOWUP_2",
    "FOLLOWUP_3"
  ];

  services.forEach(function(service) {

    steps.forEach(function(step) {

      const variants =
        getTemplateVariants(service, step);

      console.log(
        service +
        " / " +
        step +
        " -> " +
        variants.length +
        " variant(s)"
      );

      variants.forEach(function(variant) {

        console.log(
          "  Variant " +
          variant.variantId +
          ": subject='" +
          variant.subject +
          "' | theme=" +
          (variant.theme || "default") +
          " | CTA='" +
          variant.ctaText +
          "' | html=" +
          (variant.html ? "yes" : "no")
        );
      });
    });
  });
}

function testRandomTemplateSelection(service, step, rounds) {

  const count = rounds || 10;

  console.log(
    "Random selection: " +
    service +
    " / " +
    step +
    " x" +
    count
  );

  const seen = {};

  for (let i = 0; i < count; i++) {

    const template =
      getRandomTemplate(service, step);

    const key =
      String(template.variantId) +
      ":" +
      template.subject;

    seen[key] = (seen[key] || 0) + 1;

    console.log(
      "  Pick " + (i + 1) +
      ": variant " +
      template.variantId +
      " | " +
      template.subject
    );
  }

  console.log("Distribution:", JSON.stringify(seen));
}

function testPlaceholderExpansion() {

  const lead = {
    name: "Divyan",
    company: "Test Co",
    website: "https://test.com",
    industry: "Technology",
    service: "SEO",
    personalisedIntro:
      "Test introduction.",
    recentObservation:
      "A generic observation."
  };

  const text =
    "Hi {{FirstName}},\n" +
    "{{PersonalisedIntro}}\n" +
    "Business: {{Business}} | Industry: {{Industry}}\n" +
    "Observation: {{RecentObservation}}\n" +
    "Benefit: {{ServiceBenefit}}\n" +
    "CTA: {{CtaText}} -> {{CtaUrl}}\n" +
    "Year: {{CurrentYear}} | Sender: {{SenderName}}\n" +
    "Signature: {{EmailSignature}}\n" +
    "Leftover check: {{Name}}";

  const output =
    personaliseTemplate(text, lead, {
      ctaText: "Let's talk",
      ctaUrl: "https://example.com/cta"
    });

  console.log("=== PLACEHOLDER EXPANSION ===");
  console.log(output);

  const leftover = output.match(/\{\{[A-Za-z]+\}\}/g);

  if (leftover) {
    console.log(
      "WARNING: Unresolved placeholders: " +
      leftover.join(", ")
    );
  } else {
    console.log("All placeholders resolved.");
  }
}
