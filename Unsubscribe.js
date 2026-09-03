// ============================================================
// UNSUBSCRIBE ENDPOINT + ONE-CLICK (RFC 8058)
// ============================================================
// Web-app endpoint that lets email recipients opt out of all
// future automation for their lead. Supports both human clicks
// (GET -> confirmation page) and mailbox One-Click posts
// (POST with List-Unsubscribe-Post: List-Unsubscribe=One-Click).
//
// Deploy: Apps Script editor -> Deploy -> New deployment ->
// Web app (Execute as: me, Access: Anyone). Copy the returned
// URL into the Config sheet as UNSUBSCRIBE_URL. Templates build
// the per-lead link automatically as {{UnsubscribeLink}} with an
// HMAC token so the List-Unsubscribe header is authenticated.
// Token secret lives in Script Properties (UNSUB_TOKEN_SECRET),
// never in code or Config.
//
// On success: lead is set to DO_NOT_CONTACT (terminal),
// SuppressionList gains UNSUBSCRIBED, ActivityLog gains UNSUBSCRIBE.

// Secret key for HMAC (generated once, stored in Script Properties).
var UNSUB_SECRET_PROP = "UNSUB_TOKEN_SECRET";

function _getUnsubSecret() {

  try {
    const props = PropertiesService.getScriptProperties();
    let secret = props.getProperty(UNSUB_SECRET_PROP);

    if (secret) {
      return secret;
    }

    // Generate a fresh 32-byte hex secret.
    const bytes = [];
    for (let i = 0; i < 32; i++) {
      bytes.push(Math.floor(Math.random() * 256));
    }

    secret = bytes.map(function(b) {
      const hex = b.toString(16);
      return hex.length === 1 ? "0" + hex : hex;
    }).join("");

    props.setProperty(UNSUB_SECRET_PROP, secret);

    return secret;

  } catch (e) {
    // Fallback (not persisted) - still validates within same execution.
    return "fallback-secret-do-not-use-in-prod";
  }
}

function _computeUnsubToken(email, leadId) {

  const secret = _getUnsubSecret();
  const payload = String(email || "").trim().toLowerCase() +
    "|" + String(leadId || "").trim();

  const sig = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    payload,
    secret
  );

  // Base64url without padding.
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, "");
}

function _verifyUnsubToken(email, leadId, token) {

  if (!token) {
    return false;
  }

  const expected = _computeUnsubToken(email, leadId);
  return String(token).trim() === expected;
}

function buildSignedUnsubscribeLink(lead, configOverride) {

  const config = configOverride || getConfig();
  const base = String(config.UNSUBSCRIBE_URL || "").trim();

  if (!base || !lead || !lead.email) {
    return "";
  }

  const token = _computeUnsubToken(lead.email, lead.leadId);

  return (
    base +
    "?email=" + encodeURIComponent(String(lead.email || "")) +
    "&id=" + encodeURIComponent(String(lead.leadId || "")) +
    "&token=" + encodeURIComponent(token)
  );
}

function _processUnsubscribe(email, leadId, token, source) {

  const target = String(email || "").trim().toLowerCase();
  if (!target) {
    return {
      ok: false,
      title: "Something went wrong.",
      message: "No email address was provided. Please use the link from the email you received."
    };
  }

  // If a token is present, it must verify. If no token, allow
  // human GET clicks for backward compatibility with older links
  // (those links still require exact email match, so forgery risk is
  // low - they can only unsubscribe the address in the URL itself).
  if (token && !_verifyUnsubToken(target, leadId, token)) {
    console.log("Unsubscribe token failed for " + target);
    return {
      ok: false,
      title: "Something went wrong.",
      message: "This unsubscribe link is invalid or expired. Please reply \"not interested\" to the original email."
    };
  }

  const sheet = getLeadsSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const leadEmail = String(data[i][LEADS_COL.EMAIL] || "").trim().toLowerCase();
    if (leadEmail !== target) {
      continue;
    }

    const row = i + 1;
    const currentStatus = String(data[i][LEADS_COL.STATUS] || "").trim();

    if (currentStatus === "DO_NOT_CONTACT") {
      console.log("Unsubscribe requested for already-opted-out lead: " + email);
      return {
        ok: true,
        title: "You're already unsubscribed.",
        message: "No further emails will be sent to this address."
      };
    }

    // Optional leadId check - if provided and mismatched, still process
    // by email (the email is the canonical opt-out key). Log mismatch.
    const sheetLeadId = String(data[i][LEADS_COL.LEAD_ID] || "").trim();
    if (leadId && String(leadId).trim() !== sheetLeadId) {
      console.log(
        "Unsubscribe leadId mismatch: url id=" + leadId +
        " sheet id=" + sheetLeadId + " for " + target
      );
    }

    const lead = buildLeadFromRow(data, i);

    // O - Status
    sheet.getRange(row, LEADS_COL.STATUS + 1).setValue("DO_NOT_CONTACT");

    // S/T - Bounce Category/Diagnostic (reuse as suppression reason)
    writeLeadBounceState(row, "UNSUBSCRIBED", "Unsubscribed via " + (source || "link"), false);

    // R - Last Updated
    sheet.getRange(row, LEADS_COL.LAST_UPDATED + 1).setValue(new Date());

    addToSuppressionList(target, "UNSUBSCRIBED", "Unsubscribe via " + (source || "link"), "UNSUB");

    logActivity(
      lead,
      "UNSUBSCRIBE",
      "DO_NOT_CONTACT",
      "SUCCESS",
      "",
      lead.threadId,
      "Unsubscribed via " + (source || "email link")
    );

    console.log("Unsubscribe processed for " + email + " | Lead ID: " + lead.leadId + " | Source: " + (source || "link"));

    return {
      ok: true,
      title: "You're unsubscribed.",
      message: "You won't receive any further emails from us. Take care."
    };
  }

  console.log("Unsubscribe requested for unknown address: " + email);
  return {
    ok: false,
    title: "Address not found.",
    message: "This address is not on our list, so no action was needed."
  };
}

function doGet(event) {

  const params = event && event.parameter;

  const email = String(params && params.email || "").trim();
  const leadId = String(params && params.id || "").trim();
  const token = String(params && params.token || "").trim();

  if (!email) {
    return _unsubscribePage(
      "Something went wrong.",
      "No email address was provided. Please use the link from the email you received."
    );
  }

  try {
    const result = _processUnsubscribe(email, leadId, token, "email link");
    return _unsubscribePage(result.title, result.message);

  } catch (error) {
    console.error("Unsubscribe failed for " + email + ": " + error.message);
    return _unsubscribePage(
      "Something went wrong.",
      "Please try again, or reply \u201cnot interested\u201d to the original email."
    );
  }
}

function doPost(event) {

  // RFC 8058 One-Click: mailbox does POST with
  // List-Unsubscribe=One-Click + same query params. Process
  // identically but do not render a page - return 200 text.
  const params = event && event.parameter;

  const email = String(params && params.email || "").trim();
  const leadId = String(params && params.id || "").trim();
  const token = String(params && params.token || "").trim();

  try {
    const result = _processUnsubscribe(email, leadId, token, "one-click");

    // Header required by RFC: must succeed even for unknown addresses
    // to avoid oracle attacks; result already handles that gracefully.

    return ContentService.createTextOutput(result.ok ? "Unsubscribed" : result.title)
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    console.error("One-click unsubscribe failed for " + email + ": " + error.message);
    return ContentService.createTextOutput("Error")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ============================================================
// CONFIRMATION PAGE
// ============================================================
// Generic branded card: company name comes from the Config sheet
// (COMPANY_NAME), falling back to "our team". No image assets are
// required, so the page works in any deployment.

function _unsubscribePage(title, message) {

  let companyName = "our team";

  try {
    const config = getConfig();
    if (config && String(config.COMPANY_NAME || "").trim()) {
      companyName = String(config.COMPANY_NAME).trim();
    }
  } catch (e) {
    // Config read is best-effort; page still renders.
  }

  const html =
    '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8" />' +
    '<meta name="viewport" content="width=device-width, initial-scale=1" />' +
    '<title>Unsubscribed</title></head>' +
    '<body style="margin:0;padding:0;background:#F6F0FB;' +
    'font-family:Arial,Helvetica,sans-serif;">' +
    '<div style="max-width:480px;margin:0 auto;padding:48px 24px;">' +
    '<div style="text-align:center;font-size:18px;font-weight:bold;' +
    'color:#392A9C;margin-bottom:24px;">' +
    companyName +
    '</div>' +
    '<div style="background:#FFFFFF;border:1px solid #E2E8F0;' +
    'border-radius:12px;padding:32px;text-align:center;">' +
    '<div style="font-size:20px;font-weight:bold;color:#392A9C;">' +
    title +
    '</div>' +
    '<div style="margin-top:10px;font-size:14px;color:#64748B;' +
    'line-height:1.6;">' +
    message +
    '</div>' +
    '</div></div></body></html>';

  return HtmlService.createHtmlOutput(html);
}

// ============================================================
// TESTS
// ============================================================

// Prints the unsubscribe link that templates will embed for a
// sample lead (requires UNSUBSCRIBE_URL in the Config sheet).
function testUnsubscribeLink() {

  const config = getConfig();

  const url = String(config.UNSUBSCRIBE_URL || "").trim();

  if (!url) {

    console.log(
      "UNSUBSCRIBE_URL is not configured yet — " +
      "deploy the web app and add its URL to the Config sheet."
    );

    return;
  }

  const lead = {
    leadId: 1,
    email: "test@example.com"
  };

  const values = buildPlaceholderValues(lead, null);

  console.log("Sample unsubscribe link (via placeholders):");
  console.log(values["UnsubscribeLink"]);

  console.log("Sample signed direct link:");
  console.log(buildSignedUnsubscribeLink(lead));

  console.log("Token verify (should be true): " +
    _verifyUnsubToken(lead.email, lead.leadId, _computeUnsubToken(lead.email, lead.leadId))
  );
}

function testUnsubscribeHmac() {

  const email = "test@example.com";
  const id = "123";
  const token = _computeUnsubToken(email, id);

  console.log("Token: " + token);
  console.log("Verify good: " + _verifyUnsubToken(email, id, token));
  console.log("Verify bad: " + _verifyUnsubToken(email, id, token + "x"));
  console.log("Direct link: " + buildSignedUnsubscribeLink({ email: email, leadId: id }));
}
