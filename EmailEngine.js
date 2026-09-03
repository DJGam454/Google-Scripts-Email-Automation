// ============================================================
// RANDOM DELAY HELPER
// ============================================================

function _getRandomDelay(minSeconds, maxSeconds) {
  const min = minSeconds * 1000;
  const max = maxSeconds * 1000;
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function testSendEmail() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  const name = sheet.getRange(2, 2).getValue();
  const email = sheet.getRange(2, 3).getValue();
  const company = sheet.getRange(2, 4).getValue();

  const subject = "Cold Email Automation Test";

  const body =
    "Hi " + name + ",\n\n" +
    "This is my first automated email from Google Apps Script.\n\n" +
    "Company: " + company + "\n\n" +
    "Regards,\n" +
    "Divyan";

  GmailApp.sendEmail(email, subject, body);
}

// ============================================================
// RUN SEND BUDGET
// ============================================================
// MAX_SENDS_PER_RUN caps how many emails (EMAIL_1 + follow-ups
// combined) a single execution may send, so a run finishes well
// before the 6-minute trigger limit instead of being killed
// mid-batch. The daily limit (canSendEmail) still applies on top.

var _runSendCount = 0;
var _runSendCap = null;

function _initRunSendCap(config) {

  const cap = Number(config.MAX_SENDS_PER_RUN);

  _runSendCap =
    (!isNaN(cap) && cap > 0)
      ? cap
      : null;

  _runSendCount = 0;
}

function _runSendBudgetExhausted() {

  return (
    _runSendCap !== null &&
    _runSendCount >= _runSendCap
  );
}

function _recordRunSend() {
  _runSendCount++;
}

// Combined budget check: daily + hourly + domain-hourly + per-run.
// Hourly protects Gmail quota bursts; per-domain protects Yahoo/Outlook.
function _sendBudgetOk(emailForDomainCheck) {

  if (isQuotaBlocked()) {
    console.log("Send blocked: Gmail quota breaker tripped until midnight.");
    return false;
  }

  if (!canSendEmail()) {
    return false;
  }

  if (!canSendEmailHourly()) {
    return false;
  }

  if (emailForDomainCheck && !canSendToDomain(emailForDomainCheck)) {
    console.log(
      "Per-domain hourly cap blocks send to " +
      emailForDomainCheck
    );
    return false;
  }

  return !_runSendBudgetExhausted();
}

// ============================================================
// NEW-LEAD FLOOR (6-8% = 2-3 new/day)
// ============================================================
// Guarantees 6-8% of DAILY_LIMIT for NEW leads so fresh leads
// are not starved by a large follow-up queue. Follow-ups still
// have priority when overdue (>24h) - only non-overdue follow-ups
// yield their slot. Keeps 2/5/7 day timing intact.

function _getNewLeadFloor(config) {

  const explicit = Number(config.NEW_LEADS_DAILY_FLOOR);
  if (!isNaN(explicit) && explicit > 0) {
    return Math.floor(explicit);
  }

  const daily = Number(config.DAILY_LIMIT);
  if (!isNaN(daily) && daily > 0) {
    return Math.max(2, Math.ceil(daily * 0.08));
  }

  return 3;
}

function _hasOverdueFollowUps(config, data) {

  const graceMs = 24 * 60 * 60 * 1000;
  const now = Date.now();

  for (let i = 1; i < data.length; i++) {
    const lead = buildLeadFromRow(data, i);
    if (!ACTIVE_CAMPAIGN_STATUSES.includes(lead.status)) continue;
    if (!lead.threadId || !lead.lastEmailDate) continue;
    const followUp = FOLLOWUP_STEPS[lead.status];
    if (!followUp) continue;
    const delay = getFollowUpDelay(config, followUp.number);
    const last = new Date(lead.lastEmailDate).getTime();
    if (isNaN(last)) continue;
    // Due and overdue by grace.
    if (now - last >= delay.milliseconds + graceMs) {
      return true;
    }
  }

  return false;
}

function _countRemainingNewLeads(data) {

  let c = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][LEADS_COL.STATUS] || "").trim() === "NEW" && String(data[i][LEADS_COL.EMAIL] || "").trim() !== "") {
      c++;
    }
  }
  return c;
}

// ============================================================
// UNSUBSCRIBE HEADER HELPERS
// ============================================================
// Shared by both send paths: build the signed List-Unsubscribe
// URL (empty when UNSUBSCRIBE_URL is not configured or signing
// is unavailable) and append it to the plain-text fallback when
// the HTML footer would show it - some clients only show text.

function _buildListUnsubscribeForLead(lead, config) {

  try {
    if (typeof buildSignedUnsubscribeLink === "function") {
      return buildSignedUnsubscribeLink(lead, config) || "";
    }
  } catch (e) {
    // Signing is best-effort - fall through to no header.
  }

  return "";
}

function _appendUnsubscribeLine(plainText, unsubscribeUrl) {

  let output = plainText || "";

  if (
    unsubscribeUrl &&
    output.indexOf("Unsubscribe") === -1
  ) {
    output += "\n\nUnsubscribe: " + unsubscribeUrl;
  }

  return output;
}

// ============================================================
// RUN EXECUTION BUDGET
// ============================================================
// runAutomation can be killed by the Apps Script time limit
// (5-6 min depending on the account). A killed run skips the
// finally block, so the script lock lingers and the next run
// starts with "Another automation is already running". This
// self-cap ends the run cleanly (~299s) so the lock is always
// released and leftover leads defer to the next scheduled run.

var _runDeadline = null;
var _RUN_BUDGET_MILLISECONDS = 299000; // 299s hard self-cap

function _initRunDeadline() {

  _runDeadline =
    Date.now() + _RUN_BUDGET_MILLISECONDS;
}

function _pastRunDeadline() {

  return (
    _runDeadline !== null &&
    Date.now() >= _runDeadline
  );
}

// Test helper: verifies the execution-budget flag flips once the
// deadline passes. Safer than logging the live 299s window.
function testRunDeadline() {

  const previousDeadline = _runDeadline;

  _runDeadline = Date.now() + 1000;

  Utilities.sleep(1100);

  const past = _pastRunDeadline();

  _runDeadline = previousDeadline;

  console.log(
    "Run deadline test | After 1.1s sleep the 1s budget " +
    "is spent: " +
    past
  );

  return past;
}

function processEmails() {

  const aiConfig = getAIConfig();

  const aiEnabled =
    isFlagTrue(aiConfig.AI_ENABLED);


  // Load normal Config for TEST_MODE,
  // sender name and reply-to configuration.
  const config = getConfig();

  // Standalone runs (manual invocations outside runAutomation)
  // get the same execution-time self-cap.
  if (_runDeadline === null) {
    _initRunDeadline();
  }


  const sheet = getLeadsSheet();

  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    const row = i + 1;


    // =================================
    // RUN EXECUTION BUDGET
    // =================================

    if (_pastRunDeadline()) {

      console.log(
        "Run budget reached (299s). " +
        "Stopping EMAIL_1 processing."
      );

      return;
    }


    // =================================
    // BUILD LEAD OBJECT
    // =================================

    const lead =
      buildLeadFromRow(data, i);


    // =================================
    // ONLY PROCESS NEW LEADS
    // =================================

    if (lead.status !== "NEW") {
      continue;
    }


    // =================================
    // SUPPRESSION + DOMAIN BLOCK (S + SuppressionList + SUPPRESSED_DOMAINS)
    // =================================

    const suppressed =
      isSuppressed(lead.email);

    const domainBlocked =
      typeof isBlockedDomainEmail === "function" &&
      isBlockedDomainEmail(lead.email);

    if (suppressed || domainBlocked) {
      const cat = suppressed
        ? getSuppressionCategory(lead.email)
        : "POLICY_REJECTION (SUPPRESSED_DOMAINS block)";

      console.log(
        "Suppressed skipped: " +
        lead.email +
        " | Category: " + cat
      );

      continue;
    }

    // Defensive: Leads S already marked HARD but sheet not yet INVALID.
    if (
      String(lead.bounceCategory || "").trim() === "HARD_BOUNCE" &&
      String(lead.status || "").trim() !== "INVALID"
    ) {
      console.log(
        "Bounce Category HARD_BOUNCE without INVALID status - " +
        "skipping " + lead.email
      );
      continue;
    }


    // =================================
    // AI PERSONALISATION SAFETY
    // =================================

    if (
      aiEnabled &&
      (
        !lead.personalisedIntro ||
        String(lead.personalisedIntro).trim() === ""
      )
    ) {

      console.log(
        "Skipping " +
        lead.email +
        ": AI personalisation is missing."
      );

      continue;
    }


    // =================================
    // BASIC SAFETY
    // =================================

    if (!lead.email) {

      console.log(
        "Skipping Lead " +
        lead.leadId +
        ": Email is blank."
      );

      continue;
    }


    // =================================
    // VALIDATE ORIGINAL LEAD EMAIL
    // =================================

    if (!isValidEmail(lead.email)) {

      console.log(
        "Invalid email format: " +
        lead.email
      );


      // O - Status
      sheet
        .getRange(row, 15)
        .setValue("INVALID");


      // R - Last Updated
      sheet
        .getRange(row, 18)
        .setValue(new Date());


      logActivity(
        lead,
        "EMAIL_VALIDATION_FAILED",
        "INVALID",
        "FAILED",
        "",
        "",
        "Invalid email format"
      );


      continue;
    }


    // =================================
    // SERVICE SAFETY
    // =================================

    if (!lead.service) {

      console.log(
        "Skipping " +
        lead.email +
        ": No service assigned."
      );

      continue;
    }


    // =================================
    // DETERMINE ACTUAL RECIPIENT
    // =================================

    let recipient;

    try {

      recipient =
        getActualRecipient(
          lead,
          config
        );

    } catch (error) {

      console.error(
        "Recipient resolution failed for " +
        lead.email +
        ": " +
        error.message
      );

      continue;
    }


    // Validate actual destination too.
    // This matters especially in TEST_MODE.
    if (!isValidEmail(recipient)) {

      console.error(
        "Actual recipient is invalid: " +
        recipient
      );

      continue;
    }


    // =================================
    // DUPLICATE CAMPAIGN SAFETY
    // =================================
    //
    // Blocks:
    // Same email + same service +
    // different Lead ID where EMAIL_1
    // was already successfully sent.
    //
    // Does NOT block:
    // Website -> SEO -> Google Ads
    // progression for the same lead/email.
    // =================================

    if (
      hasDuplicateCampaign(
        lead.email,
        lead.service,
        lead.leadId
      )
    ) {

      console.log(
        "Duplicate campaign skipped: " +
        lead.email +
        " | Service: " +
        lead.service +
        " | Lead ID: " +
        lead.leadId
      );

      continue;
    }


    // =================================
    // GET EMAIL 1 TEMPLATE (random variant)
    // =================================
    // One variant is chosen at random; the variant carries its
    // own subject — subjects are never randomized separately.

    const template =
      getRandomTemplate(
        lead.service,
        "EMAIL_1"
      );


    if (!template) {

      console.log(
        "ERROR: EMAIL_1 template not found for " +
        lead.service
      );

      continue;
    }


    // =================================
    // PERSONALISE TEMPLATE
    // =================================

    const email =
      renderHtmlEmail(
        template,
        lead
      );

    const subject = email.subject;

    const body = email.plainTextBody;

    const htmlBody = email.htmlBody;


    // =================================
    // SEND BUDGET (daily + hourly + per-domain + per-run)
    // =================================

    if (!_sendBudgetOk(lead.email)) {

      console.log(
        "Send budget reached (daily/hourly/domain/per-run). " +
        "Email 1 not sent to " +
        lead.email
      );

      // Per-domain hourly blocks should skip this lead but let
      // later leads with different domains proceed.
      if (
        !canSendEmail() ||
        isQuotaBlocked() ||
        _runSendBudgetExhausted()
      ) {
        return;
      }

      continue;
    }

    if (_pastRunDeadline()) {

      console.log(
        "Run budget reached (299s). " +
        "Stopping EMAIL_1 processing."
      );

      return;
    }


    // =================================
    // LOG WHAT IS ABOUT TO HAPPEN
    // =================================

    console.log(
      "Sending EMAIL_1 for lead " +
      lead.email +
      " | Service: " +
      lead.service
    );


    if (isTestMode(config)) {

      console.log(
        "TEST MODE ACTIVE | Lead: " +
        lead.email +
        " | Actual Gmail recipient: " +
        recipient
      );
    }


    // =================================
    // SEND EMAIL
    // =================================

    let threadId = "";
    let messageId = "";

    try {

      // -------------------------------------------------
      // BUILD RAW MIME (multipart/alternative)
      // -------------------------------------------------
      // Same MIME structure as threaded follow-ups, so
      // EMAIL_1 and follow-ups share one format: From
      // display name (SENDER_NAME), Reply-To, HTML body
      // with plain-text fallback.

      // List-Unsubscribe header (signed). Empty when UNSUBSCRIBE_URL
      // is not configured - header is omitted entirely.
      const listUnsub =
        _buildListUnsubscribeForLead(lead, config);

      // Ensure plain-text fallback always contains the unsubscribe
      // URL when the HTML footer would - some clients only show text.
      const plainForSend =
        _appendUnsubscribeLine(body, listUnsub);

      const mimeMessage =
        buildMultipartAlternative({
          to: recipient,
          fromName:
            config.SENDER_NAME || "",
          fromEmail:
            Session.getEffectiveUser().getEmail() || "",
          replyTo:
            config.REPLY_TO_EMAIL || "",
          subject: subject,
          plainTextBody: plainForSend,
          htmlBody: htmlBody,
          listUnsubscribeUrl: listUnsub || ""
        });


      // -------------------------------------------------
      // SEND THROUGH GMAIL API
      // -------------------------------------------------

      const sentMessage =
        Gmail.Users.Messages.send(
          {
            raw:
              Utilities.base64EncodeWebSafe(
                Utilities
                  .newBlob(mimeMessage)
                  .getBytes()
              )
          },
          "me"
        );


      if (!sentMessage || !sentMessage.id) {

        throw new Error(
          "Gmail API returned no message id."
        );
      }


      // Authoritative thread ID from the send response —
      // no post-send search needed.
      threadId =
        sentMessage.threadId || "";

      messageId =
        sentMessage.id;


      console.log(
        "Gmail API sent message " +
        messageId +
        " | Thread: " +
        threadId +
        " | Lead: " +
        lead.email +
        " | Actual recipient: " +
        recipient
      );

    } catch (error) {

      console.error(
        "Failed to send EMAIL_1 for " +
        lead.email +
        ": " +
        error.message
      );

      // Quota errors trip the day-long breaker.
      if (_isQuotaError(error.message)) {
        tripQuotaBreaker();

        logActivity(
          lead,
          "EMAIL_1",
          lead.status,
          "FAILED",
          "",
          "",
          "QUOTA_EXCEEDED: " + error.message
        );

        // Stop the run - hammering further only burns quota.
        return;
      }

      logActivity(
        lead,
        "EMAIL_1",
        lead.status,
        "FAILED",
        "",
        "",
        error.message
      );

      // Failure pacing - mirrors success path so a quota/burst
      // does not hammer Gmail at 80/min.
      if (!_pastRunDeadline() && i < data.length - 1) {
        Utilities.sleep(_getRandomDelay(3, 4));
      }

      continue;
    }


    const now = new Date();


    // =================================
    // BATCH UPDATE LEAD
    // =================================
    // Columns: J(10) K(11) L(12) M(13)
    //          N(14) O(15) P(16) Q(17) R(18)
    // P holds the authoritative thread ID returned by the
    // Gmail API send response.

    sheet
      .getRange(row, 10, 1, 9)
      .setValues([[
        now,                // J - Email 1 Sent
        "",                 // K
        "",                 // L
        "",                 // M
        now,                // N - Last Email Date
        "EMAIL_1_SENT",     // O - Status
        threadId,           // P - Gmail Thread ID
        "",                 // Q
        now                 // R - Last Updated
      ]]);


    // =================================
    // ACTIVITY LOG
    // =================================

    logActivity(
      lead,
      "EMAIL_1",
      "EMAIL_1_SENT",
      "SUCCESS",
      messageId,
      threadId,
      ""
    );

    // Keep the in-memory daily count and run budget accurate.
    _incrementDailyCountCache();
    _recordRunSend();


    console.log(
      "EMAIL_1 completed for lead " +
      lead.email
    );


    // =================================
    // RANDOM DELAY BEFORE NEXT SEND
    // =================================

    if (i < data.length - 1) {

      // Never sleep past the execution budget.
      if (_pastRunDeadline()) {
        break;
      }

      const wait = _getRandomDelay(3, 4);

      console.log(
        "Waiting " + Math.round(wait / 1000) +
        "s before processing next lead..."
      );

      Utilities.sleep(wait);
    }
  }
}

function processFollowUps() {

  const config = getConfig();

  // Standalone runs (manual invocations outside runAutomation)
  // get the same execution-time self-cap.
  if (_runDeadline === null) {
    _initRunDeadline();
  }

  const sheet = getLeadsSheet();

  const data =
    sheet.getDataRange().getValues();

  // Aggregated counts for leads skipped by terminal status —
  // per-lead lines for these would flood the execution log
  // (thousands of leads can be DO_NOT_CONTACT/REPLIED).
  const skipCounts = {};


  for (let i = 1; i < data.length; i++) {

    // =================================
    // RUN EXECUTION BUDGET
    // =================================

    if (_pastRunDeadline()) {

      console.log(
        "Run budget reached (299s). " +
        "Stopping follow-up processing."
      );

      break;
    }


    const lead =
      buildLeadFromRow(data, i);


    // =================================
    // BASIC EMAIL CHECK
    // =================================

    if (!lead.email) {

      console.log(
        "Skipping Lead " +
        lead.leadId +
        ": No email"
      );

      continue;
    }

    // =================================
    // SUPPRESSION + DOMAIN BLOCK (pre-follow-up)
    // =================================

    if (isSuppressed(lead.email) || (typeof isBlockedDomainEmail === "function" && isBlockedDomainEmail(lead.email))) {
      skipCounts["SUPPRESSED"] =
        (skipCounts["SUPPRESSED"] || 0) + 1;
      continue;
    }


    // =================================
    // MANUAL STOP
    // =================================

    if (STOP_STATUSES.includes(lead.status)) {

      skipCounts[lead.status] =
        (skipCounts[lead.status] || 0) + 1;

      continue;
    }


    // =================================
    // CAMPAIGN FINISHED
    // =================================
    // FOLLOWUP_3_SENT leads wait for the next-service delay
    // before rotating into their next campaign.

    if (lead.status === "FOLLOWUP_3_SENT") {

      const lastEmail =
        new Date(
          lead.lastEmailDate
        );


      if (
        isNaN(
          lastEmail.getTime()
        )
      ) {

        console.error(
          "Invalid Last Email Date for " +
          lead.email
        );

        continue;
      }


      const elapsedMilliseconds =
        new Date().getTime() -
        lastEmail.getTime();


      const delay =
        getNextServiceDelay(
          config
        );


      const due =
        elapsedMilliseconds >=
        delay.milliseconds;


      console.log(
        "Campaign completion timing | " +
        lead.email +
        " | Required: " +
        delay.value +
        " " +
        delay.mode +
        " | Due: " +
        due
      );


      if (!due) {
        continue;
      }


      console.log(
        "Finishing current campaign for " +
        lead.email
      );


      finishCurrentCampaign(
        sheet,
        i + 1,
        lead
      );


      continue;
    }


    // =================================
    // FOLLOW-UP STAGE RESOLUTION
    // =================================
    // FOLLOWUP_STEPS (Leads.js) maps the current status to the
    // template step, next status, timestamp column and
    // follow-up number.

    const followUp =
      FOLLOWUP_STEPS[lead.status];


    if (!followUp) {
      continue;
    }


    // =================================
    // THREAD ID CHECK
    // =================================

    if (!lead.threadId) {

      console.error(
        "Cannot process follow-up for " +
        lead.email +
        ": Gmail Thread ID missing."
      );


      logActivity(
        lead,
        "FOLLOWUP_THREAD_MISSING",
        lead.status,
        "FAILED",
        "",
        "",
        "Gmail Thread ID missing"
      );


      continue;
    }


    // =================================
    // LAST EMAIL DATE CHECK
    // =================================

    if (!lead.lastEmailDate) {

      console.log(
        "Skipping " +
        lead.email +
        ": No Last Email Date"
      );

      continue;
    }


    // =================================
    // FOLLOW-UP DUE CHECK
    // =================================

    const due =
      isFollowUpDue(
        lead.lastEmailDate,
        followUp.number,
        config
      );


    if (!due) {
      continue;
    }

    // =================================
    // SEND BUDGET CHECK
    // =================================
    // Once the daily limit, hourly, per-domain, or MAX_SENDS_PER_RUN is
    // reached for a global reason, stop scanning. Per-domain hourly
    // blocks are handled per-lead inside sendFollowUp, so here we only
    // break on global exhaustion.

    if (!_sendBudgetOk()) {

      console.log(
        "Send budget reached (daily/hourly/per-run/quota). " +
        "Stopping follow-up processing."
      );

      break;
    }

    // =================================
    // NEW-LEAD FLOOR (2-3/day, follow-up priority)
    // =================================
    // If no follow-up is overdue, reserve floor slots for NEW.
    // When overdue exists, follow-ups take all remaining budget.

    const floor = _getNewLeadFloor(config);
    const remainingNew = _countRemainingNewLeads(data);
    if (remainingNew > 0 && floor > 0) {
      const dailyLimit = Number(config.DAILY_LIMIT) || 0;
      const sentToday = _dailyCountCache !== null ? _dailyCountCache : getEmailsSentToday();
      const remaining = Math.max(0, dailyLimit - sentToday - _runSendCount);
      if (remaining <= floor && remaining > 0) {
        if (!_hasOverdueFollowUps(config, data)) {
          console.log(
            "New-lead floor: " + remaining + " slots left <= floor " + floor +
            ", no overdue follow-ups — yielding to NEW leads."
          );
          break;
        }
      }
    }

    if (_pastRunDeadline()) {

      console.log(
        "Run budget reached (299s). " +
        "Stopping follow-up processing."
      );

      break;
    }


    const sent =
      sendFollowUp(
        sheet,
        i + 1,
        lead,
        followUp.step,
        followUp.nextStatus,
        followUp.column
      );


    if (!sent) {

      console.log(
        followUp.step +
        " deferred for " +
        lead.email
      );
    }
  }


  const skipKeys = Object.keys(skipCounts);

  if (skipKeys.length > 0) {

    console.log(
      "Follow-up skip summary: " +
      JSON.stringify(skipCounts)
    );
  }
}

function sendFollowUp(
  sheet,
  row,
  lead,
  templateStep,
  newStatus,
  timestampColumn
) {

  const config = getConfig();


  console.log(
    "Preparing " +
    templateStep +
    " for " +
    lead.email
  );


  // =================================
  // TEST MODE LOGGING
  // =================================

  if (isTestMode(config)) {

    const recipient =
      getActualRecipient(
        lead,
        config
      );


    console.log(
      "TEST MODE ACTIVE | Lead: " +
      lead.email +
      " | Test recipient: " +
      recipient
    );
  }


  // =================================
  // GET TEMPLATE (random variant)
  // =================================

  const template =
    getRandomTemplate(
      lead.service,
      templateStep
    );


  if (!template) {

    console.error(
      "Template not found for " +
      lead.service +
      " / " +
      templateStep
    );

    return false;
  }


  // =================================
  // PERSONALISE BODY
  // =================================

  const email =
    renderHtmlEmail(
      template,
      lead
    );


  // =================================
  // GET STORED GMAIL THREAD
  // =================================

  const thread =
    GmailApp.getThreadById(
      lead.threadId
    );


  if (!thread) {

    console.error(
      "Gmail thread not found for " +
      lead.email
    );

    return false;
  }


  // =================================
  // SEND BUDGET (daily + hourly + per-domain + per-run)
  // =================================

  if (!_sendBudgetOk(lead.email)) {

    // Per-domain hourly block defers this lead only.
    if (
      !canSendToDomain(lead.email) &&
      canSendEmail() &&
      canSendEmailHourly() &&
      !isQuotaBlocked() &&
      !_runSendBudgetExhausted()
    ) {
      console.log(
        "Per-domain hourly cap defers " +
        templateStep +
        " for " +
        lead.email
      );
      return false;
    }

    console.log(
      "Send budget reached (global). " +
      templateStep +
      " not sent to " +
      lead.email
    );

    return false;
  }

  // =================================
  // SUPPRESSION + DOMAIN BLOCK (follow-up path)
  // =================================

  if (isSuppressed(lead.email) || (typeof isBlockedDomainEmail === "function" && isBlockedDomainEmail(lead.email))) {
    console.log(
      "Suppressed - skipping " +
      templateStep +
      " for " +
      lead.email
    );
    return false;
  }


  // =================================
  // RUN EXECUTION BUDGET
  // =================================
  // Defer the follow-up instead of sending past the cap —
  // the lead keeps its current status so the next run sends it.

  if (_pastRunDeadline()) {

    console.log(
      "Run budget reached (299s). " +
      templateStep +
      " deferred for " +
      lead.email
    );

    return false;
  }


  // =================================
  // RANDOM DELAY BEFORE FOLLOW-UP
  // =================================

  const wait = _getRandomDelay(3, 4);

  console.log(
    "Waiting " + Math.round(wait / 1000) +
    "s before sending " + templateStep + "..."
  );

  Utilities.sleep(wait);


  // =================================
  // SEND THREADED FOLLOW-UP
  // =================================

  try {

    sendThreadedFollowUp(
      lead,
      email,
      template
    );

  } catch (error) {

    console.error(
      "Failed to send " +
      templateStep +
      " to " +
      lead.email +
      ": " +
      error.message
    );

    if (_isQuotaError(error.message)) {
      tripQuotaBreaker();

      logActivity(
        lead,
        templateStep,
        lead.status,
        "FAILED",
        "",
        lead.threadId,
        "QUOTA_EXCEEDED: " + error.message
      );

      return false;
    }

    logActivity(
      lead,
      templateStep,
      lead.status,
      "FAILED",
      "",
      lead.threadId,
      error.message
    );

    return false;
  }


  const now = new Date();


  // =================================
  // UPDATE SHEET
  // =================================

  sheet
    .getRange(row, timestampColumn)
    .setValue(now);


  // N - Last Email Date
  sheet
    .getRange(row, 14)
    .setValue(now);


  // O - Status
  sheet
    .getRange(row, 15)
    .setValue(newStatus);


  // R - Last Updated
  sheet
    .getRange(row, 18)
    .setValue(now);


  // =================================
  // ACTIVITY LOG
  // =================================

  logActivity(
    lead,
    templateStep,
    newStatus,
    "SUCCESS",
    "",
    lead.threadId,
    ""
  );

  // Keep the in-memory daily count and run budget accurate.
  _incrementDailyCountCache();
  _recordRunSend();


  console.log(
    templateStep +
    " sent successfully for lead " +
    lead.email
  );


  return true;
}

function sendThreadedFollowUp(lead, email, template) {

  // --------------------------------
  // CONFIG + RECIPIENT
  // --------------------------------

  const config = getConfig();

  const recipient =
    getActualRecipient(
      lead,
      config
    );


  if (!isValidEmail(recipient)) {
    throw new Error(
      "Invalid follow-up recipient: " +
      recipient
    );
  }

  // Named From so follow-ups display the same sender as Email 1
  // (GmailApp gets `name` from Config; MIME sends must set the
  // From header themselves).
  const fromName =
    config.SENDER_NAME || "";

  const fromEmail =
    Session.getEffectiveUser().getEmail() ||
    config.REPLY_TO_EMAIL ||
    "";


  // --------------------------------
  // GET THREAD THROUGH GMAIL API
  // --------------------------------

  const thread = Gmail.Users.Threads.get(
    "me",
    lead.threadId
  );


  if (
    !thread ||
    !thread.messages ||
    thread.messages.length === 0
  ) {

    throw new Error(
      "Gmail API could not find thread: " +
      lead.threadId
    );
  }


  // --------------------------------
  // GET LATEST MESSAGE
  // --------------------------------
  // Using the latest message is better for
  // In-Reply-To / References than always
  // referencing the original message.

  const latestMessage =
    thread.messages[
      thread.messages.length - 1
    ];


  const headers =
    latestMessage.payload.headers;


  function getHeader(name) {

    const header =
      headers.find(function(h) {

        return (
          h.name.toLowerCase() ===
          name.toLowerCase()
        );

      });


    return header
      ? header.value
      : "";
  }


  const threadSubject =
    getHeader("Subject");

  const latestMessageId =
    getHeader("Message-ID");


  if (!threadSubject) {

    throw new Error(
      "Thread message has no Subject header."
    );
  }


  if (!latestMessageId) {

    throw new Error(
      "Thread message has no Message-ID header."
    );
  }


  // --------------------------------
  // SUBJECT RESOLUTION
  // --------------------------------
  // A variant may define its own subject. When it does, it is
  // used (threading is preserved by In-Reply-To/References);
  // when blank, the thread's existing subject is reused.

  const subject =
    email.subject
      ? email.subject
      : threadSubject;


  // --------------------------------
  // BUILD MIME EMAIL (multipart/alternative)
  // --------------------------------

  const listUnsubFollow =
    _buildListUnsubscribeForLead(lead, config);

  const plainForFollow =
    _appendUnsubscribeLine(
      email.plainTextBody,
      listUnsubFollow
    );

  const mimeMessage =
    buildMultipartAlternative({
      to: recipient,
      fromName: fromName,
      fromEmail: fromEmail,
      replyTo: config.REPLY_TO_EMAIL || "",
      subject: subject,
      inReplyTo: latestMessageId,
      references: latestMessageId,
      plainTextBody: plainForFollow,
      htmlBody: email.htmlBody,
      listUnsubscribeUrl: listUnsubFollow || ""
    });


  // --------------------------------
  // BASE64URL ENCODE
  // --------------------------------

  const encodedMessage =
    Utilities.base64EncodeWebSafe(
      Utilities
        .newBlob(mimeMessage)
        .getBytes()
    );


  // --------------------------------
  // SEND THROUGH GMAIL API
  // --------------------------------

  const message = {

    raw: encodedMessage,

    threadId: lead.threadId

  };


  const sentMessage =
    Gmail.Users.Messages.send(
      message,
      "me"
    );


  // --------------------------------
  // LOGGING
  // --------------------------------

  if (isTestMode(config)) {

    console.log(
      "TEST MODE | Threaded follow-up " +
      "for lead " +
      lead.email +
      " redirected to " +
      recipient
    );

  }


  console.log(
    "Gmail API sent threaded message " +
    sentMessage.id +
    " | Lead: " +
    lead.email +
    " | Actual recipient: " +
    recipient
  );


  return sentMessage;
}

function runAutomation() {

  const lock = LockService.getScriptLock();

  const locked = lock.tryLock(10000);

  if (!locked) {

    console.log(
      "Another automation is already running. Skipping."
    );

    return;
  }

  try {

    const config = getConfig();

    // Reset run-scoped budgets, caches and quota view.
    _initRunSendCap(config);
    _initRunDeadline();
    _resetDailyCountCache();
    _resetDuplicateCache();
    _resetHourlyDomainCache();
    _resetSuppressionCache();

    // Master ON/OFF switch
    if (!isFlagTrue(config.AUTOMATION_ENABLED)) {

      console.log(
        "Automation is currently disabled."
      );

      return;
    }

// --------------------------------
// SCHEDULE CONTROL
// --------------------------------

if (
  !isAutomationScheduleAllowed(
    config
  )
) {

  console.log(
    "Automation skipped due to schedule."
  );

  return;
}

console.log("=== AUTOMATION STARTED ===");

// 1. Detect bounced emails first
checkBounces();

if (_pastRunDeadline()) {

  console.log(
    "Run budget reached (299s). " +
    "Stopping after bounce check."
  );

  return;
}

// 2. Detect human replies before anything else sends.
//    When a dedicated checkReplies trigger runs every few
//    minutes, this scan can be skipped here to free the whole
//    send budget for this run.
if (!isFlagTrue(config.SKIP_REPLIES_IN_MAIN_RUN)) {
  checkReplies();
}

if (_pastRunDeadline()) {

  console.log(
    "Run budget reached (299s). " +
    "Stopping after reply check."
  );

  return;
}

// 3. Generate AI personalisation for NEW leads
generateMissingPersonalizations();

if (_pastRunDeadline()) {

  console.log(
    "Run budget reached (299s). " +
    "Stopping after AI personalisation."
  );

  return;
}

// 4. Existing campaigns/follow-ups get priority
processFollowUps();

if (_pastRunDeadline()) {

  console.log(
    "Run budget reached (299s). " +
    "Stopping after follow-ups."
  );

  return;
}

// 5. Send Email 1 to NEW leads
processEmails();

console.log("=== AUTOMATION COMPLETED ===");

  } catch (error) {

    console.error(
      "Automation failed: " + error.message
    );

  } finally {

    lock.releaseLock();

    console.log("=== LOCK RELEASED ===");
  }
}

function getFollowUpDelay(config, followUpNumber) {

  const fastTestMode =
    isFastTestMode(config);


  // ==============================
  // FAST TEST MODE
  // ==============================

  if (fastTestMode) {

    const key =
      "FOLLOWUP_" +
      followUpNumber +
      "_MINUTES";


    const minutes =
      Number(config[key]);


    if (
      isNaN(minutes) ||
      minutes < 0
    ) {

      throw new Error(
        "Invalid test follow-up timing: " +
        key
      );
    }


    return {
      mode: "MINUTES",
      value: minutes,
      milliseconds:
        minutes * 60 * 1000
    };
  }


  // ==============================
  // PRODUCTION MODE
  // ==============================

  const key =
    "FOLLOWUP_" +
    followUpNumber +
    "_DAYS";


  const days =
    Number(config[key]);


  if (
    isNaN(days) ||
    days < 0
  ) {

    throw new Error(
      "Invalid production follow-up timing: " +
      key
    );
  }


  return {
    mode: "DAYS",
    value: days,
    milliseconds:
      days * 24 * 60 * 60 * 1000
  };
}

function getNextServiceDelay(config) {

  if (isFastTestMode(config)) {

    const minutes =
      Number(config.NEXT_SERVICE_MINUTES);

    if (
      isNaN(minutes) ||
      minutes < 0
    ) {
      throw new Error(
        "Invalid NEXT_SERVICE_MINUTES"
      );
    }

    return {
      mode: "MINUTES",
      value: minutes,
      milliseconds:
        minutes * 60 * 1000
    };
  }


  const days =
    Number(config.NEXT_SERVICE_DAYS);

  if (
    isNaN(days) ||
    days < 0
  ) {
    throw new Error(
      "Invalid NEXT_SERVICE_DAYS"
    );
  }

  return {
    mode: "DAYS",
    value: days,
    milliseconds:
      days * 24 * 60 * 60 * 1000
  };
}

function isFollowUpDue(
  lastEmailDate,
  followUpNumber,
  config
) {

  if (!lastEmailDate) {
    return false;
  }


  const lastSent =
    new Date(lastEmailDate);


  if (isNaN(lastSent.getTime())) {
    return false;
  }


  const now =
    new Date();


  const elapsedMilliseconds =
    now.getTime() -
    lastSent.getTime();


  const delay =
    getFollowUpDelay(
      config,
      followUpNumber
    );


  const due =
    elapsedMilliseconds >=
    delay.milliseconds;


  const elapsedMinutes =
    elapsedMilliseconds /
    (1000 * 60);


  console.log(
    "Follow-up " +
    followUpNumber +
    " timing | Required: " +
    delay.value +
    " " +
    delay.mode +
    " | Elapsed minutes: " +
    elapsedMinutes.toFixed(2) +
    " | Due: " +
    due
  );


  return due;
}


