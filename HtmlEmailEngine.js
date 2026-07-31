// ============================================================
// HTML EMAIL ENGINE
// ============================================================
// Renders professional HTML emails from spreadsheet templates
// using a reusable component + theme system.
//
// How it works:
//   1. A template record (see Templates.js) may contain an HTML
//      skeleton that composes {{Component}} tokens, e.g.
//      {{Header}} {{Hero}} {{CtaButton}} {{Footer}}.
//   2. Each component is a reusable block renderer that takes
//      (values, theme, config) and returns inline-CSS HTML.
//   3. The engine wraps the composed content in a 600px
//      table-based shell and runs placeholder substitution.
//   4. The same template also drives a plain-text fallback
//      (the Plain Text Template column; otherwise HTML-stripped).
//
// Deliverability constraints honoured (see email-deliverability
// and html-email skills): inline CSS only, table layout, Google
// safe fonts, Outlook bgcolor fallbacks, no scripts/tracking,
// alt text on images, one primary CTA.

// ============================================================
// THEME SYSTEM
// ============================================================
// Built-in fallback themes (used when the Themes sheet has no
// row for a theme). Services map to a default theme; a template
// row's Theme column can override. All colours are configurable
// via the optional Themes sheet.

var DEFAULT_THEMES = {
  "website": {
    name: "website",
    accent: "#2563EB",
    accentDark: "#1D4ED8",
    heroBg: "#2563EB",
    heroText: "#FFFFFF",
    heroSubtext: "#DBEAFE",
    buttonBg: "#2563EB",
    buttonText: "#FFFFFF",
    cardBorder: "#E2E8F0",
    cardBg: "#FFFFFF"
  },
  "seo": {
    name: "seo",
    accent: "#059669",
    accentDark: "#047857",
    heroBg: "#059669",
    heroText: "#FFFFFF",
    heroSubtext: "#D1FAE5",
    buttonBg: "#059669",
    buttonText: "#FFFFFF",
    cardBorder: "#E2E8F0",
    cardBg: "#FFFFFF"
  },
  "google-ads": {
    name: "google-ads",
    accent: "#EA580C",
    accentDark: "#C2410C",
    heroBg: "#EA580C",
    heroText: "#FFFFFF",
    heroSubtext: "#FFEDD5",
    buttonBg: "#EA580C",
    buttonText: "#FFFFFF",
    cardBorder: "#E2E8F0",
    cardBg: "#FFFFFF"
  }
};

// Service -> default theme name.
var SERVICE_DEFAULT_THEME = {
  "Website Development": "website",
  "SEO": "seo",
  "Google Ads": "google-ads"
};

// Themes sheet (optional) columns:
// A Theme | B Accent | C Accent Dark | D Hero Background |
// E Hero Text | F Hero Subtext | G Button Background |
// H Button Text | I Card Border | J Status

var _themesSheetCache = null;

function _loadThemesSheet() {

  if (_themesSheetCache !== null) {
    return _themesSheetCache;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Themes");

  if (!sheet) {
    _themesSheetCache = null;
    return null;
  }

  _themesSheetCache =
    sheet.getDataRange().getValues();

  return _themesSheetCache;
}

// Resolve the theme for an email. themeName comes from the
// template's Theme column (blank -> service default).
function getTheme(themeName, service) {

  const fallbackName =
    String(themeName || "")
      .trim()
      .toLowerCase() ||
    SERVICE_DEFAULT_THEME[service] ||
    "website";

  const fallback =
    DEFAULT_THEMES[fallbackName] ||
    DEFAULT_THEMES["website"];

  const sheet = _loadThemesSheet();

  if (sheet) {

    for (let i = 1; i < sheet.length; i++) {

      const rowName = String(sheet[i][0] || "")
        .trim()
        .toLowerCase();

      if (rowName !== fallbackName) {
        continue;
      }

      const status = String(sheet[i][9] || "")
        .trim()
        .toUpperCase();

      if (status && status !== "ACTIVE") {
        break;
      }

      return {
        name: fallbackName,
        accent: sheet[i][1] || fallback.accent,
        accentDark: sheet[i][2] || fallback.accentDark,
        heroBg: sheet[i][3] || fallback.heroBg,
        heroText: sheet[i][4] || fallback.heroText,
        heroSubtext: sheet[i][5] || fallback.heroSubtext,
        buttonBg: sheet[i][6] || fallback.buttonBg,
        buttonText: sheet[i][7] || fallback.buttonText,
        cardBorder: sheet[i][8] || fallback.cardBorder,
        cardBg: fallback.cardBg
      };
    }
  }

  return fallback;
}

// ============================================================
// COMPONENT RENDERERS
// ============================================================
// Each renderer returns a <tr> fragment for the email shell.
// All styles are inline for Gmail compatibility.

function renderHeader(values, theme, config) {

  const companyName =
    values["CompanyName"] || "Our Company";

  return (
    '<tr>' +
    '<td bgcolor="#0F172A" style="background-color:#0F172A;' +
    'padding:20px 32px;border-radius:12px 12px 0 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr>' +
    '<td style="font-size:18px;font-weight:bold;color:#FFFFFF;' +
    'font-family:Arial,Helvetica,sans-serif;">' +
    companyName +
    '</td>' +
    '<td align="right" style="font-size:10px;color:#94A3B8;' +
    'font-family:Arial,Helvetica,sans-serif;letter-spacing:1px;">' +
    'B2B&nbsp;OUTREACH' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '</td>' +
    '</tr>'
  );
}

function renderHero(values, theme, config) {

  const title = values["HeroTitle"] || "";
  const subtitle = values["HeroSubtitle"] || "";

  let inner = "";

  if (title) {

    inner +=
      '<div style="font-size:20px;font-weight:bold;color:' +
      theme.heroText +
      ';line-height:1.4;font-family:Arial,Helvetica,sans-serif;">' +
      title +
      '</div>';
  }

  if (subtitle) {

    inner +=
      '<div style="font-size:12px;color:' +
      theme.heroSubtext +
      ';margin-top:8px;line-height:1.5;font-family:Arial,Helvetica,sans-serif;">' +
      subtitle +
      '</div>';
  }

  if (!inner) {
    return "";
  }

  return (
    '<tr>' +
    '<td bgcolor="' + theme.heroBg + '" style="background-color:' +
    theme.heroBg +
    ';padding:28px 32px;">' +
    inner +
    '</td>' +
    '</tr>'
  );
}

function renderFeatureList(values, theme, config) {

  const features = values["Features"];

  if (!features || features.length === 0) {
    return "";
  }

  const items = features.map(function(feature) {

    return (
      '<div style="padding:5px 0;font-size:14px;color:#334155;' +
      'line-height:1.5;font-family:Arial,Helvetica,sans-serif;">' +
      '<span style="color:' + theme.accent + ';font-weight:bold;">' +
      '&#8226;&nbsp;</span>' +
      feature +
      '</div>'
    );
  }).join("");

  return (
    '<tr>' +
    '<td style="padding:16px 32px 4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
    'style="background-color:' + theme.cardBg + ';border:1px solid ' +
    theme.cardBorder + ';border-radius:10px;">' +
    '<tr>' +
    '<td style="padding:18px 22px;">' +
    '<div style="font-size:11px;font-weight:bold;color:#64748B;' +
    'text-transform:uppercase;letter-spacing:1px;font-family:Arial,Helvetica,sans-serif;">' +
    'Why this matters' +
    '</div>' +
    '<div style="margin-top:8px;">' +
    items +
    '</div>' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '</td>' +
    '</tr>'
  );
}

function renderPortfolioShowcase(values, theme, config) {

  const heroImage = values["HeroImage"] || "";
  const portfolioImage = values["PortfolioImage"] || "";
  const portfolioUrl = values["PortfolioUrl"] || "";
  const companyName = values["CompanyName"] || "our";

  // No images configured -> simple portfolio link card.
  if (!heroImage && !portfolioImage) {

    if (!portfolioUrl) {
      return "";
    }

    return (
      '<tr>' +
      '<td style="padding:16px 32px 4px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
      'style="background-color:' + theme.cardBg + ';border:1px solid ' +
      theme.cardBorder + ';border-radius:10px;">' +
      '<tr>' +
      '<td align="center" style="padding:20px 22px;">' +
      '<div style="font-size:14px;font-weight:bold;color:#0F172A;' +
      'font-family:Arial,Helvetica,sans-serif;">Recent work</div>' +
      '<div style="margin-top:6px;font-size:13px;color:#64748B;line-height:1.5;' +
      'font-family:Arial,Helvetica,sans-serif;">' +
      'A few examples of what we build and the results we get.' +
      '</div>' +
      '<div style="margin-top:12px;">' +
      '<a href="' + portfolioUrl + '" style="color:' + theme.accent +
      ';font-weight:bold;font-size:13px;text-decoration:none;' +
      'font-family:Arial,Helvetica,sans-serif;">' +
      'View our portfolio &#8594;' +
      '</a>' +
      '</div>' +
      '</td>' +
      '</tr>' +
      '</table>' +
      '</td>' +
      '</tr>'
    );
  }

  // Images configured -> two rounded image cards.
  const images = [
    [heroImage, "Recent work by " + companyName],
    [portfolioImage, "More recent work by " + companyName]
  ].filter(function(pair) {
    return pair[0];
  });

  const cards = images.map(function(pair) {

    return (
      '<td width="48%" style="padding:4px;">' +
      '<a href="' + portfolioUrl + '">' +
      '<img src="' + pair[0] + '" alt="' + pair[1] +
      '" width="260" style="display:block;width:100%;height:auto;' +
      'border-radius:10px;border:1px solid ' + theme.cardBorder + ';" />' +
      '</a>' +
      '</td>'
    );
  }).join("");

  return (
    '<tr>' +
    '<td style="padding:16px 32px 4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
    'style="background-color:' + theme.cardBg + ';border:1px solid ' +
    theme.cardBorder + ';border-radius:10px;">' +
    '<tr>' +
    '<td style="padding:18px 22px;">' +
    '<div style="font-size:11px;font-weight:bold;color:#64748B;' +
    'text-transform:uppercase;letter-spacing:1px;font-family:Arial,Helvetica,sans-serif;">' +
    'Recent work' +
    '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ' +
    'style="margin-top:10px;">' +
    '<tr>' +
    cards +
    '</tr>' +
    '</table>' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '</td>' +
    '</tr>'
  );
}

function renderCtaButton(values, theme, config) {

  const text = values["CtaText"] || "";
  const url = values["CtaUrl"] || "";

  if (!text) {
    return "";
  }

  const href = url
    ? 'href="' + url + '"'
    : 'href="mailto:' + (values["ReplyToEmail"] || "") + '"';

  return (
    '<tr>' +
    '<td align="center" style="padding:24px 32px 8px;">' +
    '<a ' + href + ' style="display:inline-block;background-color:' +
    theme.buttonBg + ';color:' + theme.buttonText +
    ';font-weight:bold;font-size:15px;text-decoration:none;' +
    'border-radius:8px;padding:12px 30px;' +
    'font-family:Arial,Helvetica,sans-serif;">' +
    text +
    '</a>' +
    '</td>' +
    '</tr>'
  );
}

function renderDivider(values, theme, config) {

  return (
    '<tr>' +
    '<td style="padding:20px 32px 4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr>' +
    '<td style="border-top:1px solid ' + theme.cardBorder +
    ';font-size:0;line-height:1px;height:1px;">&nbsp;</td>' +
    '</tr>' +
    '</table>' +
    '</td>' +
    '</tr>'
  );
}

function renderSignatureBlock(values, theme, config) {

  const name = values["SenderName"] || "";
  const role = values["SenderRole"] || "";
  const company = values["CompanyName"] || "";
  const email = values["ReplyToEmail"] || "";
  const website = values["CompanyWebsite"] || "";

  let lines = "";

  if (name) {

    lines +=
      '<div style="font-size:15px;font-weight:bold;color:#0F172A;' +
      'font-family:Arial,Helvetica,sans-serif;">' +
      name +
      '</div>';
  }

  const roleLine = [role, company]
    .filter(Boolean)
    .join(" &#183; ");

  if (roleLine) {

    lines +=
      '<div style="font-size:13px;color:' + theme.accent +
      ';font-family:Arial,Helvetica,sans-serif;">' +
      roleLine +
      '</div>';
  }

  const contactLine = [email, website]
    .filter(Boolean)
    .join(" &#183; ");

  if (contactLine) {

    lines +=
      '<div style="font-size:12px;color:#64748B;' +
      'font-family:Arial,Helvetica,sans-serif;">' +
      contactLine +
      '</div>';
  }

  if (!lines) {
    return "";
  }

  return (
    '<tr>' +
    '<td style="padding:8px 32px 4px;">' +
    lines +
    '</td>' +
    '</tr>'
  );
}

function renderSocialLinks(values, theme, config) {

  const pairs = [
    ["SocialLinkedIn", "LinkedIn"],
    ["SocialTwitter", "Twitter"],
    ["SocialGithub", "GitHub"]
  ];

  const links = pairs
    .filter(function(pair) {
      return values[pair[0]];
    })
    .map(function(pair) {

      return (
        '<a href="' + values[pair[0]] +
        '" style="color:#64748B;text-decoration:none;font-size:12px;' +
        'font-family:Arial,Helvetica,sans-serif;">' +
        pair[1] +
        '</a>'
      );
    });

  if (links.length === 0) {
    return "";
  }

  return (
    '<tr>' +
    '<td align="center" style="padding:12px 32px 4px;">' +
    '<div style="font-size:12px;color:#64748B;font-family:Arial,Helvetica,sans-serif;">' +
    links.join("&nbsp;&nbsp;&#183;&nbsp;&nbsp;") +
    '</div>' +
    '</td>' +
    '</tr>'
  );
}

function renderFooter(values, theme, config) {

  const company = values["CompanyName"] || "";
  const address = values["CompanyAddress"] || "";
  const phone = values["CompanyPhone"] || "";
  const email =
    values["CompanyEmail"] || values["ReplyToEmail"] || "";
  const website = values["CompanyWebsite"] || "";

  const details = [address, phone, email, website]
    .filter(Boolean)
    .join("&nbsp;&nbsp;|&nbsp;&nbsp;");

  return (
    '<tr>' +
    '<td bgcolor="#F8FAFC" style="background-color:#F8FAFC;' +
    'border-top:1px solid ' + theme.cardBorder +
    ';border-radius:0 0 12px 12px;padding:16px 32px;' +
    'color:#94A3B8;font-size:11px;line-height:1.6;' +
    'font-family:Arial,Helvetica,sans-serif;">' +
    company +
    (company && details ? "<br/>" : "") +
    details +
    '</td>' +
    '</tr>'
  );
}

function renderUnsubscribeBlock(values, theme, config) {

  const message = values["UnsubscribeMessage"] || "";

  if (!message) {
    return "";
  }

  return (
    '<tr>' +
    '<td align="center" style="padding:12px 32px 20px;color:#94A3B8;' +
    'font-size:11px;line-height:1.6;font-family:Arial,Helvetica,sans-serif;">' +
    message +
    '</td>' +
    '</tr>'
  );
}

// ============================================================
// COMPONENT REGISTRY
// ============================================================
// Token name -> renderer. Templates compose these tokens inside
// their HTML skeleton.

var EMAIL_COMPONENTS = {
  Header: renderHeader,
  Hero: renderHero,
  FeatureList: renderFeatureList,
  PortfolioShowcase: renderPortfolioShowcase,
  CtaButton: renderCtaButton,
  Divider: renderDivider,
  SignatureBlock: renderSignatureBlock,
  SocialLinks: renderSocialLinks,
  Footer: renderFooter,
  UnsubscribeBlock: renderUnsubscribeBlock
};

// ============================================================
// EMAIL SHELL
// ============================================================
// 600px centred table, page background, white content card.

function _buildShell(contentHtml) {

  return (
    '<div style="background-color:#F4F6F8;padding:24px 12px;' +
    'font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
    '<tr>' +
    '<td align="center">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" ' +
    'style="max-width:600px;width:100%;background-color:#FFFFFF;' +
    'border:1px solid #E2E8F0;border-radius:12px;">' +
    contentHtml +
    '</table>' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '</div>'
  );
}

// ============================================================
// RENDERER
// ============================================================
// Renders a template record (Templates.js) into a complete
// email payload: subject, HTML body, plain-text fallback.

function renderHtmlEmail(template, lead) {

  const config = getConfig();

  // Context = template record so CTA text/URL and images
  // become {{CtaText}} / {{CtaUrl}} / {{HeroImage}} /
  // {{PortfolioImage}} placeholders.
  const values =
    buildPlaceholderValues(lead, template);

  const theme =
    getTheme(template && template.theme, lead && lead.service);

  const subject =
    personaliseTemplate(template.subject, lead, template);

  // --------------------------------
  // HTML BODY
  // --------------------------------

  let htmlBody = "";

  if (template.html) {

    let content = String(template.html);

    // 1. Swap component tokens for unique sentinels so rendered
    //    component HTML can never be re-parsed as components.
    const sentinels = {};

    Object.keys(EMAIL_COMPONENTS).forEach(function(name, index) {

      const token = "{{" + name + "}}";

      if (content.indexOf(token) !== -1) {

        const sentinel = "@@EMAIL_COMPONENT_" + index + "@@";

        sentinels[sentinel] =
          EMAIL_COMPONENTS[name](values, theme, config);

        content = content.split(token).join(sentinel);
      }
    });

    // 2. Replace sentinels with rendered component HTML.
    Object.keys(sentinels).forEach(function(sentinel) {

      content = content.split(sentinel).join(sentinels[sentinel]);
    });

    // 3. Wrap in the shell, then substitute data placeholders.
    content = _buildShell(content);

    htmlBody =
      applyPlaceholderValues(content, values);
  }

  // --------------------------------
  // PLAIN-TEXT FALLBACK
  // --------------------------------
  // Preferred: the template's own plain-text copy. Fallback:
  // strip the HTML rendering.

  let plainTextBody = "";

  if (template.body) {

    plainTextBody =
      personaliseTemplate(template.body, lead, template);

  } else if (htmlBody) {

    plainTextBody = htmlToPlainText(htmlBody);
  }

  return {
    subject: subject,
    htmlBody: htmlBody,
    plainTextBody: plainTextBody,
    values: values,
    theme: theme
  };
}

// ============================================================
// HTML -> PLAIN TEXT (fallback only)
// ============================================================

function htmlToPlainText(html) {

  if (!html) {
    return "";
  }

  let text = String(html);

  // Convert block-level tags to line breaks BEFORE stripping.
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<\/tr>/gi, "\n");
  text = text.replace(/<\/table>/gi, "\n\n");
  text = text.replace(/<[^>]+>/g, "");

  // Decode common entities.
  text = text.replace(/&nbsp;/gi, " ");
  text = text.replace(/&amp;/gi, "&");
  text = text.replace(/&lt;/gi, "<");
  text = text.replace(/&gt;/gi, ">");
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&#39;/gi, "'");
  text = text.replace(/&#8226;/gi, "-");
  text = text.replace(/&#8594;/gi, "->");
  text = text.replace(/&#183;/gi, "|");

  // Collapse excessive blank lines.
  text = text.replace(/\n{3,}/g, "\n\n");

  return text.trim();
}

// ============================================================
// MIME BUILDER (threaded follow-ups via Gmail API)
// ============================================================
// Builds a multipart/alternative raw message with both the
// plain-text and HTML versions plus In-Reply-To / References
// for correct threading.

function buildMultipartAlternative(options) {

  const boundary =
    "----=_boundary_" +
    Utilities.getUuid().replace(/-/g, "");

  const headers = [
    "To: " + options.to,
    "Subject: " + (options.subject || "")
  ];

  if (options.inReplyTo) {
    headers.push("In-Reply-To: " + options.inReplyTo);
  }

  if (options.references) {
    headers.push("References: " + options.references);
  }

  headers.push("MIME-Version: 1.0");
  headers.push(
    'Content-Type: multipart/alternative; boundary="' +
    boundary +
    '"'
  );

  const parts = [
    headers.join("\r\n"),
    "",
    "--" + boundary,
    'Content-Type: text/plain; charset="UTF-8"',
    "",
    options.plainTextBody || "",
    "",
    "--" + boundary,
    'Content-Type: text/html; charset="UTF-8"',
    "",
    options.htmlBody || "",
    "",
    "--" + boundary + "--"
  ];

  return parts.join("\r\n");
}

// ============================================================
// TESTS
// ============================================================

function testThemeResolver() {

  const services = [
    "Website Development",
    "SEO",
    "Google Ads",
    "Unknown Service"
  ];

  services.forEach(function(service) {

    const theme =
      getTheme("", service);

    console.log(
      service +
      " -> theme '" +
      theme.name +
      "' | accent: " +
      theme.accent
    );
  });

  console.log(
    "Explicit theme 'seo': " +
    getTheme("seo", "Website Development").accent
  );
}

function testHtmlEmailRender(service, step) {

  const template =
    getRandomTemplate(service, step);

  if (!template) {
    console.log("No template for " + service + " / " + step);
    return;
  }

  const lead = {
    name: "Divyan",
    company: "Test Co",
    website: "https://test.com",
    industry: "Technology",
    service: service,
    personalisedIntro:
      "Your company stood out while I was researching " +
      "the industry."
  };

  const email =
    renderHtmlEmail(template, lead);

  console.log("=== HTML RENDER: " + service + " / " + step + " ===");
  console.log("Variant: " + template.variantId);
  console.log("Subject: " + email.subject);
  console.log(
    "HTML length: " +
    email.htmlBody.length +
    " | Plain length: " +
    email.plainTextBody.length
  );

  const leftover =
    email.htmlBody.match(/\{\{[A-Za-z]+\}\}/g);

  if (leftover) {
    console.log(
      "WARNING: Unresolved placeholders in HTML: " +
      leftover.join(", ")
    );
  } else {
    console.log("HTML: all placeholders resolved.");
  }

  console.log("Plain text preview:");
  console.log(email.plainTextBody.slice(0, 300));
}

function testAllVariantsRender() {

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

  let rendered = 0;
  let failed = 0;

  services.forEach(function(service) {

    steps.forEach(function(step) {

      const variants =
        getTemplateVariants(service, step);

      variants.forEach(function(template) {

        try {

          const lead = {
            name: "Divyan",
            company: "Test Co",
            website: "https://test.com",
            industry: "Technology",
            service: service,
            personalisedIntro:
              "Your company stood out while I was researching."
          };

          const email =
            renderHtmlEmail(template, lead);

          const leftover = (
            (email.subject || "") +
            (email.htmlBody || "") +
            (email.plainTextBody || "")
          ).match(/\{\{[A-Za-z]+\}\}/g);

          if (leftover) {

            failed++;

            console.log(
              "FAIL " +
              service + " / " + step +
              " / variant " + template.variantId +
              ": leftover " + leftover.join(", ")
            );
          } else {

            rendered++;

            console.log(
              "OK   " +
              service + " / " + step +
              " / variant " + template.variantId +
              " | html=" + email.htmlBody.length +
              " plain=" + email.plainTextBody.length
            );
          }

        } catch (error) {

          failed++;

          console.log(
            "ERROR " +
            service + " / " + step +
            " / variant " + template.variantId +
            ": " + error.message
          );
        }
      });
    });
  });

  console.log(
    "Render summary | rendered: " +
    rendered +
    " | failed: " +
    failed
  );
}

function testMimeBuilder() {

  const mime = buildMultipartAlternative({
    to: "recipient@example.com",
    subject: "A test subject",
    inReplyTo: "<abc123@mail.gmail.com>",
    references: "<abc123@mail.gmail.com>",
    plainTextBody: "Hello,\n\nThis is the plain text version.",
    htmlBody: "<div style='font-family:Arial'>Hello, <b>HTML</b></div>"
  });

  const lines = mime.split("\r\n");

  console.log("=== MIME BUILDER ===");
  console.log("Total length: " + mime.length);

  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    console.log(lines[i].slice(0, 120));
  }

  console.log("Boundaries: " + (mime.match(/boundary/g) || []).length);
}
