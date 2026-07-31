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

function processEmails() {

  const aiConfig = getAIConfig();

  const aiEnabled =
    isFlagTrue(aiConfig.AI_ENABLED);


  // Load normal Config for TEST_MODE,
  // sender name and reply-to configuration.
  const config = getConfig();


  const sheet = getLeadsSheet();

  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    const row = i + 1;


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
    // DAILY LIMIT
    // =================================

    if (!canSendEmail()) {

      console.log(
        "Daily sending limit reached. " +
        "Email 1 not sent to " +
        lead.email
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

    try {

      const sendOptions = {};


      // Configurable Gmail display name
      if (config.SENDER_NAME) {

        sendOptions.name =
          config.SENDER_NAME;
      }


      // Configurable Reply-To
      if (config.REPLY_TO_EMAIL) {

        sendOptions.replyTo =
          config.REPLY_TO_EMAIL;
      }


      // HTML body with plain-text fallback
      if (htmlBody) {

        sendOptions.htmlBody =
          htmlBody;
      }


      GmailApp.sendEmail(
        recipient,
        subject,
        body,
        sendOptions
      );

    } catch (error) {

      console.error(
        "Failed to send EMAIL_1 for " +
        lead.email +
        ": " +
        error.message
      );


      logActivity(
        lead,
        "EMAIL_1",
        lead.status,
        "FAILED",
        "",
        "",
        error.message
      );


      continue;
    }


    const now = new Date();


    // =================================
    // FIND GMAIL THREAD
    // =================================

    // Give Gmail a moment to expose the
    // newly-created thread to search.
    Utilities.sleep(2000);


    // Search using recipient rather than
    // lead.email because TEST_MODE may
    // redirect the actual destination.
    const threads =
      GmailApp.search(
        'in:sent to:"' +
        recipient +
        '" subject:"' +
        subject +
        '"',
        0,
        5
      );


    let threadId = "";


    if (threads.length > 0) {

      threadId =
        threads[0].getId();

      console.log(
        "Thread ID saved: " +
        threadId
      );

    } else {

      console.log(
        "WARNING: Email sent but Gmail thread " +
        "could not be found for lead " +
        lead.email +
        " | Actual recipient: " +
        recipient
      );
    }


    // =================================
    // BATCH UPDATE LEAD
    // =================================
    // Columns: J(10) K(11) L(12) M(13)
    //          N(14) O(15) P(16) Q(17) R(18)

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
      "",
      threadId,
      ""
    );


    console.log(
      "EMAIL_1 completed for lead " +
      lead.email
    );


    // =================================
    // RANDOM DELAY BEFORE NEXT SEND
    // =================================

    if (i < data.length - 1) {

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

  const sheet = getLeadsSheet();

  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

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
    // MANUAL STOP
    // =================================

    if (STOP_STATUSES.includes(lead.status)) {

      console.log(
        "Skipping " +
        lead.email +
        " because status is " +
        lead.status
      );

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


    console.log(
      "Preparing " +
      followUp.step +
      " for " +
      lead.email
    );


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
  // DAILY LIMIT
  // =================================

  if (!canSendEmail()) {

    console.log(
      "Daily sending limit reached. " +
      templateStep +
      " not sent to " +
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

  const mimeMessage =
    buildMultipartAlternative({
      to: recipient,
      subject: subject,
      inReplyTo: latestMessageId,
      references: latestMessageId,
      plainTextBody: email.plainTextBody,
      htmlBody: email.htmlBody
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

// 2. Generate AI personalisation for NEW leads
generateMissingPersonalizations();

// 3. Existing campaigns/follow-ups get priority
processFollowUps();

// 4. Send Email 1 to NEW leads
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

function getDelay(config, settingName) {

  if (isFastTestMode(config)) {
    return 0;
  }

  return Number(config[settingName]);
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


