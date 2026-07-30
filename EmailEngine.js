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
    aiConfig.AI_ENABLED === true ||
    String(aiConfig.AI_ENABLED).toUpperCase() === "TRUE";


  // Load normal Config for TEST_MODE,
  // sender name and reply-to configuration.
  const config = getConfig();


  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  if (!sheet) {
    throw new Error("Leads sheet not found.");
  }


  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    const row = i + 1;


    // =================================
    // BUILD LEAD OBJECT
    // =================================

    const lead = {
      leadId: data[i][0],              // A - Lead ID
      company: data[i][1],             // B - Company Name
      name: data[i][2],                // C - Contact Name
      email: data[i][3],               // D - Email
      website: data[i][4],             // E - Website
      industry: data[i][5],            // F - Industry
      personalisedIntro: data[i][6],   // G - Personalised Intro
      service: data[i][7],             // H - Service Assigned
      campaign: data[i][8],            // I - Current Campaign
      status: data[i][14]              // O - Status
    };


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
    // GET EMAIL 1 TEMPLATE
    // =================================

    const template =
      getTemplate(
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

    const subject =
      personaliseTemplate(
        template.subject,
        lead
      );


    const body =
      personaliseTemplate(
        template.body,
        lead
      );


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
    // UPDATE LEAD
    // =================================

    // J - Email 1 Sent
    sheet
      .getRange(row, 10)
      .setValue(now);


    // N - Last Email Date
    sheet
      .getRange(row, 14)
      .setValue(now);


    // O - Status
    sheet
      .getRange(row, 15)
      .setValue("EMAIL_1_SENT");


    // R - Last Updated
    sheet
      .getRange(row, 18)
      .setValue(now);


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


      // P - Gmail Thread ID
      sheet
        .getRange(row, 16)
        .setValue(threadId);


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
  }
}

function processFollowUps() {

  const config = getConfig();

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  if (!sheet) {
    throw new Error("Leads sheet not found.");
  }

  const data =
    sheet.getDataRange().getValues();


  for (let i = 1; i < data.length; i++) {

    const lead = {
      leadId: data[i][0],
      company: data[i][1],
      name: data[i][2],
      email: data[i][3],
      website: data[i][4],
      industry: data[i][5],
      personalisedIntro: data[i][6],
      service: data[i][7],
      campaign: data[i][8],

      email1Sent: data[i][9],
      followup1Sent: data[i][10],
      followup2Sent: data[i][11],
      followup3Sent: data[i][12],

      lastEmailDate: data[i][13],
      status: data[i][14],
      threadId: data[i][15]
    };


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

    const stopStatuses = [
      "REPLIED",
      "INVALID",
      "DO_NOT_CONTACT"
    ];


    if (
      stopStatuses.includes(
        lead.status
      )
    ) {

      console.log(
        "Skipping " +
        lead.email +
        " because status is " +
        lead.status
      );

      continue;
    }


    // =================================
    // FOLLOW-UP STATUSES
    // =================================

    const followUpStatuses = [
      "EMAIL_1_SENT",
      "FOLLOWUP_1_SENT",
      "FOLLOWUP_2_SENT",
      "FOLLOWUP_3_SENT"
    ];


    if (
      !followUpStatuses.includes(
        lead.status
      )
    ) {
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
    // FOLLOW-UP 1
    // =================================

    if (
      lead.status === "EMAIL_1_SENT"
    ) {

      const due =
        isFollowUpDue(
          lead.lastEmailDate,
          1,
          config
        );


      if (!due) {
        continue;
      }


      console.log(
        "Preparing FOLLOWUP_1 for " +
        lead.email
      );


      const sent =
        sendFollowUp(
          sheet,
          i + 1,
          lead,
          "FOLLOWUP_1",
          "FOLLOWUP_1_SENT",
          11
        );


      if (!sent) {

        console.log(
          "FOLLOWUP_1 deferred for " +
          lead.email
        );

      }


      continue;
    }


    // =================================
    // FOLLOW-UP 2
    // =================================

    if (
      lead.status ===
      "FOLLOWUP_1_SENT"
    ) {

      const due =
        isFollowUpDue(
          lead.lastEmailDate,
          2,
          config
        );


      if (!due) {
        continue;
      }


      console.log(
        "Preparing FOLLOWUP_2 for " +
        lead.email
      );


      const sent =
        sendFollowUp(
          sheet,
          i + 1,
          lead,
          "FOLLOWUP_2",
          "FOLLOWUP_2_SENT",
          12
        );


      if (!sent) {

        console.log(
          "FOLLOWUP_2 deferred for " +
          lead.email
        );

      }


      continue;
    }


    // =================================
    // FOLLOW-UP 3
    // =================================

    if (
      lead.status ===
      "FOLLOWUP_2_SENT"
    ) {

      const due =
        isFollowUpDue(
          lead.lastEmailDate,
          3,
          config
        );


      if (!due) {
        continue;
      }


      console.log(
        "Preparing FOLLOWUP_3 for " +
        lead.email
      );


      const sent =
        sendFollowUp(
          sheet,
          i + 1,
          lead,
          "FOLLOWUP_3",
          "FOLLOWUP_3_SENT",
          13
        );


      if (!sent) {

        console.log(
          "FOLLOWUP_3 deferred for " +
          lead.email
        );

      }


      continue;
    }


    // =================================
    // CAMPAIGN FINISHED
    // =================================

    if (
      lead.status ===
      "FOLLOWUP_3_SENT"
    ) {

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
  // GET TEMPLATE
  // =================================

  const template =
    getTemplate(
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

  const body =
    personaliseTemplate(
      template.body,
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
  // SEND THREADED FOLLOW-UP
  // =================================

  try {

    // Keep using your existing
    // working threading implementation.
    sendThreadedFollowUp(
      lead,
      body
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

function sendThreadedFollowUp(lead, body) {

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


  const subject =
    getHeader("Subject");

  const latestMessageId =
    getHeader("Message-ID");


  if (!subject) {

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
  // BUILD MIME EMAIL
  // --------------------------------

  const mimeMessage = [

    "To: " + recipient,

    "Subject: " + subject,

    "In-Reply-To: " +
      latestMessageId,

    "References: " +
      latestMessageId,

    "MIME-Version: 1.0",

    'Content-Type: text/plain; charset="UTF-8"',

    "",

    body

  ].join("\r\n");


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
    if (
      config.AUTOMATION_ENABLED !== true &&
      String(config.AUTOMATION_ENABLED).toUpperCase() !== "TRUE"
    ) {

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

  const fastTestMode =
    config.FAST_TEST_MODE === true ||
    String(config.FAST_TEST_MODE).toUpperCase() === "TRUE";

  if (fastTestMode) {
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


