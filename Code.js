function testReadSheet() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  const data = sheet.getDataRange().getValues();

  console.log(data);
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
function testFollowUp() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  // We're testing row 2 only
  const row = 2;

  // P = Gmail Thread ID
  const threadId = sheet
    .getRange(row, 16)
    .getValue();

  if (!threadId) {
    console.log("No Gmail Thread ID found.");
    return;
  }

  console.log("Using Thread ID: " + threadId);

  // Find existing Gmail conversation
  const thread = GmailApp.getThreadById(threadId);

  if (!thread) {
    console.log("Could not find Gmail thread.");
    return;
  }

  // Reply inside that conversation
  thread.reply(
    "Hi Divyan,\n\n" +
    "Just following up on my previous email.\n\n" +
    "This follow-up was sent automatically using the stored Gmail Thread ID.\n\n" +
    "Regards,\n" +
    "Divyan"
  );

  console.log("Follow-up sent inside existing Gmail thread.");
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




function logActivity(
  lead,
  action,
  status,
  result,
  messageId,
  threadId,
  errorMessage
) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("ActivityLog");

  if (!sheet) {
    console.error("ActivityLog sheet not found.");
    return;
  }

  sheet.appendRow([
    new Date(),                    // Timestamp
    lead.leadId || "",             // Lead ID
    lead.email || "",              // Email
    action || "",                  // Action
    lead.campaign || "",           // Campaign
    lead.service || "",            // Service
    status || "",                  // Status
    messageId || "",               // Gmail Message ID
    threadId || "",                // Gmail Thread ID
    result || "",                  // SUCCESS / FAILED
    errorMessage || ""             // Error
  ]);
}
function getServices() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Services");

  if (!sheet) {
    throw new Error("Services sheet not found.");
  }

  const data = sheet.getDataRange().getValues();

  const services = [];

  for (let i = 1; i < data.length; i++) {

    const order = Number(data[i][0]);
    const name = String(data[i][1]).trim();
    const active = data[i][2];

    // Skip incomplete rows
    if (!order || !name) {
      continue;
    }

    // Handle TRUE whether Sheets returns boolean or text
    const isActive =
      active === true ||
      String(active).toUpperCase() === "TRUE";

    if (!isActive) {
      continue;
    }

    services.push({
      order: order,
      name: name
    });
  }

  // Don't depend on rows being physically sorted
  services.sort(function(a, b) {
    return a.order - b.order;
  });

  return services;
}
function testServices() {

  const services = getServices();

  console.log("Active services:");

  services.forEach(function(service) {

    console.log(
      "Service " +
      service.order +
      ": " +
      service.name
    );

  });
}
function getNextService(currentService) {

  const services = getServices();

  for (let i = 0; i < services.length; i++) {

    if (services[i].name === currentService) {

      // Is there another service after this one?
      if (i + 1 < services.length) {

        return services[i + 1];

      }

      // Current service is the last service
      return null;
    }
  }

  // Current service wasn't found
  return null;
}
function testNextService() {

  let next;


  // Website Development
  next = getNextService("Website Development");

  if (next) {
    console.log(
      "After Website Development: " +
      next.name
    );
  } else {
    console.log(
      "Website Development is the last service."
    );
  }


  // SEO
  next = getNextService("SEO");

  if (next) {
    console.log(
      "After SEO: " +
      next.name
    );
  } else {
    console.log(
      "SEO is the last service."
    );
  }


  // Google Ads
  next = getNextService("Google Ads");

  if (next) {
    console.log(
      "After Google Ads: " +
      next.name
    );
  } else {
    console.log(
      "Google Ads is the last service."
    );
  }
}
function finishCurrentCampaign(sheet, row, lead) {

  // --------------------------------
  // CURRENT CAMPAIGN NUMBER
  // --------------------------------

  const campaignNumber =
    getCampaignNumber(lead.campaign);

  const notInterestedStatus =
    "NOT_INTERESTED_SERVICE_" +
    campaignNumber;


  // --------------------------------
  // RECORD NO REPLY
  // --------------------------------

  logActivity(
    lead,
    "CAMPAIGN_NO_REPLY",
    notInterestedStatus,
    "SUCCESS",
    "",
    lead.threadId,
    ""
  );

  console.log(
    lead.email +
    " finished " +
    lead.service +
    " as Service " +
    campaignNumber +
    " without reply."
  );


  // --------------------------------
  // FIND NEXT ACTIVE SERVICE
  // --------------------------------

  const nextService =
    getNextService(lead.service);


  // --------------------------------
  // NO NEXT SERVICE
  // --------------------------------

  if (!nextService) {

    const now = new Date();

    // O - Status
    sheet
      .getRange(row, 15)
      .setValue("COMPLETED");

    // R - Last Updated
    sheet
      .getRange(row, 18)
      .setValue(now);


    logActivity(
      lead,
      "ALL_CAMPAIGNS_COMPLETED",
      "COMPLETED",
      "SUCCESS",
      "",
      lead.threadId,
      ""
    );


    console.log(
      "All campaigns completed for " +
      lead.email
    );

    return;
  }


  // --------------------------------
  // NEXT CAMPAIGN NUMBER
  // --------------------------------

  const nextCampaignNumber =
    campaignNumber + 1;

  const nextCampaign =
    "Service " + nextCampaignNumber;

  const now = new Date();


  // --------------------------------
  // UPDATE SERVICE
  // --------------------------------
// G - Personalised Intro
sheet
  .getRange(row, 7)
  .clearContent();

  // H - Service Assigned
  sheet
    .getRange(row, 8)
    .setValue(nextService.name);


  // I - Current Campaign
  sheet
    .getRange(row, 9)
    .setValue(nextCampaign);


  // --------------------------------
  // RESET CURRENT CAMPAIGN DATA
  // --------------------------------

  // J:N
  //
  // Email 1 Sent
  // Follow-up 1 Sent
  // Follow-up 2 Sent
  // Follow-up 3 Sent
  // Last Email Date

  sheet
    .getRange(row, 10, 1, 5)
    .clearContent();


  // O - Status
  sheet
    .getRange(row, 15)
    .setValue("NEW");


  // P - Gmail Thread ID
  //
  // New service should start
  // a new Gmail conversation.

  sheet
    .getRange(row, 16)
    .clearContent();


  // R - Last Updated
  sheet
    .getRange(row, 18)
    .setValue(now);


  // --------------------------------
  // LOG NEW CAMPAIGN
  // --------------------------------

  const nextLead = {

    leadId: lead.leadId,

    email: lead.email,

    campaign: nextCampaign,

    service: nextService.name
  };


  logActivity(
    nextLead,
    "CAMPAIGN_STARTED",
    "NEW",
    "SUCCESS",
    "",
    "",
    ""
  );


  console.log(
    "Moved " +
    lead.email +
    " from " +
    lead.service +
    " to " +
    nextService.name +
    " as " +
    nextCampaign
  );
}
function getCampaignNumber(currentCampaign) {

  if (!currentCampaign) {
    return 1;
  }

  const match = String(currentCampaign).match(/\d+/);

  if (!match) {
    return 1;
  }

  return Number(match[0]);
}
function getEmailsSentToday() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("ActivityLog");

  if (!sheet) {
    throw new Error("ActivityLog sheet not found.");
  }

  const data = sheet.getDataRange().getValues();

  const today = new Date();

  let count = 0;

  for (let i = 1; i < data.length; i++) {

    const timestamp = data[i][0];
    const action = data[i][3];
    const result = data[i][9];

    if (!timestamp) {
      continue;
    }

    const logDate = new Date(timestamp);

    const isToday =
      logDate.getFullYear() === today.getFullYear() &&
      logDate.getMonth() === today.getMonth() &&
      logDate.getDate() === today.getDate();


    const isEmailAction = [
      "EMAIL_1",
      "FOLLOWUP_1",
      "FOLLOWUP_2",
      "FOLLOWUP_3"
    ].includes(action);


    if (
      isToday &&
      isEmailAction &&
      result === "SUCCESS"
    ) {
      count++;
    }
  }

  return count;
}
function testDailyCount() {

  const count = getEmailsSentToday();

  console.log(
    "Emails sent today: " + count
  );
}
function canSendEmail() {

  const config = getConfig();

  const dailyLimit =
    Number(config.DAILY_LIMIT);

  const sentToday =
    getEmailsSentToday();

  const remaining =
    dailyLimit - sentToday;


  console.log(
    "Daily limit: " + dailyLimit +
    " | Sent today: " + sentToday +
    " | Remaining: " + Math.max(remaining, 0)
  );


  return sentToday < dailyLimit;
}
function isValidEmail(email) {

  if (!email) {
    return false;
  }

  email = String(email).trim();

  // Simple practical validation for MVP
  const emailPattern =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return emailPattern.test(email);
}
function testEmailValidation() {

  const emails = [
    "test@gmail.com",
    "hello@company.co.uk",
    "johncompany.com",
    "abc@",
    "",
    "hello @gmail.com"
  ];

  emails.forEach(function(email) {

    console.log(
      email + " → " + isValidEmail(email)
    );

  });
}

function checkBounces() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName("Leads");

  if (!leadsSheet) {
    throw new Error("Leads sheet not found.");
  }

  console.log("=== CHECKING FOR BOUNCES ===");

  // Look for common Gmail delivery failure messages
  // received recently.
  const threads = GmailApp.search(
    'newer_than:2d (from:mailer-daemon OR subject:"Delivery Status Notification" OR subject:"Delivery incomplete")'
  );

  console.log(
    "Potential bounce threads found: " +
    threads.length
  );

  if (threads.length === 0) {
    console.log("=== BOUNCE CHECK COMPLETE ===");
    return;
  }

  const data =
    leadsSheet.getDataRange().getValues();

  // Summary counters instead of logging every
  // unrecognized bounce message individually.
  let messagesScanned = 0;
  let identifiedBounces = 0;
  let unidentifiedMessages = 0;
  let leadsMarkedInvalid = 0;
  let alreadyInvalid = 0;


  for (let t = 0; t < threads.length; t++) {

    const messages =
      threads[t].getMessages();


    for (let m = 0; m < messages.length; m++) {

      messagesScanned++;

      const message =
        messages[m];

      const body =
        message.getPlainBody();

      const bouncedEmail =
        extractBouncedEmail(body);


      if (!bouncedEmail) {

        unidentifiedMessages++;

        continue;
      }


      identifiedBounces++;

      console.log(
        "Bounce detected for: " +
        bouncedEmail
      );


      // Find bounced email in Leads
      for (let i = 1; i < data.length; i++) {

        const leadEmail =
          String(data[i][3] || "")
            .trim()
            .toLowerCase();


        if (
          leadEmail !==
          bouncedEmail.toLowerCase()
        ) {
          continue;
        }


        const row = i + 1;

        const currentStatus =
          data[i][14];


        // Don't repeatedly process an
        // already-invalid lead.
        if (currentStatus === "INVALID") {

          alreadyInvalid++;

          break;
        }


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
          status: currentStatus,
          threadId: data[i][15]

        };


        // O = Status
        leadsSheet
          .getRange(row, 15)
          .setValue("INVALID");


        // R = Last Updated
        leadsSheet
          .getRange(row, 18)
          .setValue(new Date());


        logActivity(
          lead,
          "EMAIL_BOUNCED",
          "INVALID",
          "FAILED",
          "",
          lead.threadId,
          "Delivery failure / bounced email"
        );


        leadsMarkedInvalid++;


        console.log(
          bouncedEmail +
          " marked INVALID due to bounce."
        );


        break;
      }
    }
  }


  // =================================
  // BOUNCE SCAN SUMMARY
  // =================================

  console.log(
    "Bounce scan summary | " +
    "Threads: " +
    threads.length +
    " | Messages: " +
    messagesScanned +
    " | Bounces identified: " +
    identifiedBounces +
    " | Unrecognized messages: " +
    unidentifiedMessages +
    " | Leads marked INVALID: " +
    leadsMarkedInvalid +
    " | Already INVALID: " +
    alreadyInvalid
  );


  console.log(
    "=== BOUNCE CHECK COMPLETE ==="
  );
}

function extractBouncedEmail(body) {

  if (!body) {
    return null;
  }

  /*
   * Common bounce messages often contain:
   *
   * Recipient:
   * Final-Recipient:
   * Original-Recipient:
   * or an email address in the failure text.
   */

  const patterns = [

    /Final-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i,

    /Original-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i,

    /Recipient:\s*([^\s<>]+@[^\s<>]+)/i,

    /Your message wasn't delivered to\s+([^\s<>]+@[^\s<>]+)/i

  ];


  for (let i = 0; i < patterns.length; i++) {

    const match = body.match(patterns[i]);

    if (match && match[1]) {

      return match[1]
        .replace(/[>,.;]+$/, "")
        .trim();
    }
  }

  return null;
}

function testBounceParser() {

  const testBodies = [

    "Final-Recipient: rfc822; fakeperson@example.com",

    "Original-Recipient: rfc822; customer@testcompany.com",

    "Recipient: anotherperson@example.org",

    "Your message wasn't delivered to broken@example.com because the address couldn't be found."

  ];


  testBodies.forEach(function(body) {

    const email =
      extractBouncedEmail(body);

    console.log(
      body + " → " + email
    );

  });
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
function testGeminiKey() {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    console.log("ERROR: GEMINI_API_KEY not found.");
    return;
  }

  console.log("Gemini API key found successfully.");
}



function callGemini(prompt) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");


  if (!apiKey) {
    throw new Error(
      "GEMINI_API_KEY not found in Script Properties."
    );
  }


  const config = getAIConfig();


  // Check AI toggle
  const aiEnabled =
    config.AI_ENABLED === true ||
    String(config.AI_ENABLED).toUpperCase() === "TRUE";


  if (!aiEnabled) {
    throw new Error(
      "AI generation is currently disabled."
    );
  }


  const model =
    String(config.AI_MODEL).trim();


  if (!model) {
    throw new Error(
      "AI_MODEL is missing from AIConfig."
    );
  }


  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent";


  const payload = {

    contents: [
      {
        role: "user",

        parts: [
          {
            text: prompt
          }
        ]
      }
    ],

    generationConfig: {

      temperature:
        Number(config.TEMPERATURE) || 0.4

    }
  };


  const options = {

    method: "post",

    contentType: "application/json",

    headers: {
      "x-goog-api-key": apiKey
    },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true

  };


  const response =
    UrlFetchApp.fetch(
      url,
      options
    );


  const statusCode =
    response.getResponseCode();


  const responseText =
    response.getContentText();


  // -------------------------------
  // API ERROR
  // -------------------------------

  if (
    statusCode < 200 ||
    statusCode >= 300
  ) {

    throw new Error(
      "Gemini API error " +
      statusCode +
      ": " +
      responseText
    );
  }


  const result =
    JSON.parse(responseText);


  // -------------------------------
  // VALIDATE RESPONSE
  // -------------------------------

  if (
    !result.candidates ||
    result.candidates.length === 0
  ) {

    throw new Error(
      "Gemini returned no candidates."
    );
  }


  const candidate =
    result.candidates[0];


  if (
    !candidate.content ||
    !candidate.content.parts ||
    candidate.content.parts.length === 0
  ) {

    throw new Error(
      "Gemini returned no text content."
    );
  }


  // Combine text parts
  const text =
    candidate.content.parts
      .map(function(part) {
        return part.text || "";
      })
      .join("")
      .trim();


  if (!text) {

    throw new Error(
      "Gemini returned an empty response."
    );
  }


  return text;
}
function testGeminiConnection() {

  try {

    console.log(
      "Testing Gemini connection..."
    );


    const response =
      callGemini(
        "Return exactly this text and nothing else: " +
        "Gemini connection successful"
      );


    console.log(
      "Gemini response: " +
      response
    );


  } catch (error) {

    console.error(
      "Gemini test failed: " +
      error.message
    );

  }
}

function generatePersonalisation(lead, research = null) {

  const config = getAIConfig();

  const maxWords =
    Number(config.MAX_INTRO_WORDS) || 35;

    const researchContext =
  research
    ? `
VERIFIED WEBSITE RESEARCH:

Company Summary:
${research.companySummary}

Services / Products Found:
${research.servicesFound.join(", ") || "Unknown"}

Target Audience:
${research.targetAudience}

Value Proposition:
${research.valueProposition}

Useful Facts:
${research.usefulFacts.join("\n") || "None"}

Research Confidence:
${research.researchConfidence}

The information above was obtained from the company's
supplied public website.

You may use these facts for personalisation, but do not
invent anything beyond them.
`
    : `
NO VERIFIED WEBSITE RESEARCH IS AVAILABLE.

Use only the basic lead information.
Do not claim to have inspected the website.
`;


  const prompt = `
You are writing a highly concise opening line for a B2B cold email.

Your job is NOT to sell the service.
Your job is to create a natural transition into a later sales message.

SERVICE BEING OFFERED:
${lead.service}

LEAD INFORMATION:
Company: ${lead.company || "Unknown"}
Website: ${lead.website || "Not provided"}
Industry: ${lead.industry || "Not provided"}

${researchContext}

TASK:
Write one short personalised opening sentence relevant to the company
and the service being offered.

IMPORTANT GROUNDING RULES:
- Use ONLY the information supplied above.
- Never invent company facts, statistics, rankings, traffic, customers,
  website problems, advertising activity, revenue, growth, or performance.
- The website URL is provided only as an identifier.
- Do NOT claim you visited, reviewed, analysed, or inspected the website.
- Do NOT claim the company has a specific problem unless that information
  was explicitly supplied.
- If information is limited, keep the observation broad but relevant.

WRITING STYLE:
- Write ONLY the opening observation.
- NEVER include the contact's name.
- NEVER include "Hi", "Hello", "Hey", or any greeting.
- The email template will add the greeting separately.
- Maximum ${maxWords} words.
- Target 12-25 words.
- Professional but conversational.
- Write like a human salesperson, not marketing copy.
- Avoid generic statements that could apply to every company.
- Do not use phrases such as:
  "for a company like"
  "in today's"
  "fast-paced"
  "competitive landscape"
  "robust digital presence"
  "high-intent buyers"
  "unlock"
  "leverage"
  "essential"
  "key to"
  "game-changing"
- Do not exaggerate.
- Do not pitch the service.
- Do not include a CTA.
- Do not introduce the sender.
- Do not repeat the company name unnecessarily.
- Use the company or industry context naturally.

SERVICE GUIDANCE:

Website Development:
Focus on the importance of a company's website or digital presence
without claiming its current website is poor.

SEO:
Focus on discoverability or organic search relevance without claiming
the company currently has poor rankings or traffic.

Google Ads:
Focus on reaching relevant potential customers through paid search
without claiming the company currently does or does not run ads.

Return ONLY valid JSON in exactly this structure:

{
  "intro": "one personalised sentence",
  "confidence": "high, medium, or low"
}
CONFIDENCE RULES:

"high":
Only when specific verified company information has been supplied.

"medium":
When useful company-specific context beyond basic name,
industry and URL has been supplied.

"low":
When only basic fields such as company name, industry,
website URL and service are available.

For the information provided in this prompt, do not use
"high" unless specific verified company information exists.
`;


  const response = callGemini(prompt);

  let result;


  // -------------------------------
  // PARSE JSON
  // -------------------------------

  try {

    let cleanedResponse =
      response.trim();


    // Defensive cleanup in case Gemini
    // wraps JSON in markdown fences.
    cleanedResponse = cleanedResponse
      .replace(/^```json\s*/i, "")
      .replace(/^```\s*/, "")
      .replace(/```$/, "")
      .trim();


    result =
      JSON.parse(cleanedResponse);

  } catch (error) {

    throw new Error(
      "Gemini returned invalid JSON: " +
      response
    );
  }


  // -------------------------------
  // VALIDATE INTRO
  // -------------------------------

  if (
    !result.intro ||
    typeof result.intro !== "string"
  ) {

    throw new Error(
      "Gemini response does not contain a valid intro."
    );
  }


  const intro =
    result.intro.trim();

// Reject greetings because the email template
// already adds Hi {{FirstName}}
const greetingPattern =
  /^(hi|hello|hey|dear)\b/i;

if (greetingPattern.test(intro)) {

  throw new Error(
    "AI personalisation incorrectly included a greeting: " +
    intro
  );
}

  const wordCount =
    countWords(intro);


  if (wordCount === 0) {

    throw new Error(
      "Gemini generated an empty intro."
    );
  }


  if (wordCount > maxWords) {

    throw new Error(
      "AI personalisation exceeded word limit. " +
      "Generated " +
      wordCount +
      " words. Maximum: " +
      maxWords
    );
  }


  // -------------------------------
  // VALIDATE CONFIDENCE
  // -------------------------------

  const allowedConfidence = [
    "high",
    "medium",
    "low"
  ];


  const confidence =
    String(result.confidence || "low")
      .toLowerCase()
      .trim();


  if (!allowedConfidence.includes(confidence)) {

    throw new Error(
      "Invalid AI confidence value: " +
      confidence
    );
  }


  console.log(
    "AI confidence: " +
    confidence +
    " | Words: " +
    wordCount
  );


  return intro;
}

function testPersonalisationQuality() {

  const services = [
    "Website Development",
    "SEO",
    "Google Ads"
  ];


  services.forEach(function(service) {

    const testLead = {

      name: "Divyan",

      company: "Test Co",

      website: "https://test.com",

      industry: "Technology",

      service: service

    };


    try {

      console.log(
        "============================"
      );

      console.log(
        "Testing service: " +
        service
      );


      const intro =
        generatePersonalisation(
          testLead
        );


      console.log(
        "RESULT: " +
        intro
      );


    } catch (error) {

      console.error(
        service +
        " failed: " +
        error.message
      );

    }

  });

}
function countWords(text) {

  if (!text) {
    return 0;
  }

  return String(text)
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .length;
}

function generateMissingPersonalizations() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Leads");

  if (!sheet) {
    throw new Error("Leads sheet not found.");
  }


  const config = getAIConfig();

  const aiEnabled =
    config.AI_ENABLED === true ||
    String(config.AI_ENABLED).toUpperCase() === "TRUE";


  if (!aiEnabled) {

    console.log(
      "AI personalisation is disabled."
    );

    return;
  }


  const data =
    sheet.getDataRange().getValues();


  console.log(
    "=== AI PERSONALISATION STARTED ==="
  );


  for (let i = 1; i < data.length; i++) {

    const row = i + 1;


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
      status: data[i][14]

    };


    // --------------------------------
    // ONLY NEW LEADS
    // --------------------------------

    if (lead.status !== "NEW") {
      continue;
    }


    // --------------------------------
    // ALREADY PERSONALISED
    // --------------------------------

    if (
      lead.personalisedIntro &&
      String(lead.personalisedIntro).trim() !== ""
    ) {

      console.log(
        "Skipping " +
        lead.email +
        ": Personalisation already exists."
      );

      continue;
    }


    // --------------------------------
    // SERVICE REQUIRED
    // --------------------------------

    if (
      !lead.service ||
      String(lead.service).trim() === ""
    ) {

      console.log(
        "Skipping " +
        lead.email +
        ": No service assigned."
      );

      continue;
    }


    try {

      console.log(
        "Generating personalisation for " +
        lead.email +
        " | Service: " +
        lead.service
      );


      let research = null;


      // =================================
      // WEBSITE RESEARCH
      // =================================

      if (
        lead.website &&
        String(lead.website).trim() !== ""
      ) {

        try {

          console.log(
            "Researching website: " +
            lead.website
          );


          research =
            getWebsiteResearch(lead);


          console.log(
            "Website research completed."
          );


          console.log(
            "Research confidence: " +
            research.researchConfidence
          );


          console.log(
            "Company summary: " +
            research.companySummary
          );


        } catch (researchError) {

          // Website research failure should NOT
          // stop the entire lead.

          console.error(
            "Website research failed for " +
            lead.email +
            ": " +
            researchError.message
          );


          console.log(
            "Falling back to basic lead personalisation."
          );


          research = null;
        }

      } else {

        console.log(
          "No website available for " +
          lead.email +
          ". Using basic personalisation."
        );

      }


      // =================================
      // GENERATE INTRO
      // =================================

      const intro =
        generatePersonalisation(
          lead,
          research
        );


      // =================================
      // SAVE RESULT
      // =================================

      // G - Personalised Intro
      sheet
        .getRange(row, 7)
        .setValue(intro);


      // R - Last Updated
      sheet
        .getRange(row, 18)
        .setValue(new Date());


      console.log(
        "Personalisation generated for " +
        lead.email +
        ": " +
        intro
      );


    } catch (error) {

      console.error(
        "AI personalisation failed for " +
        lead.email +
        ": " +
        error.message
      );


      // Don't allow one lead to kill
      // processing for every other lead.
      continue;
    }
  }


  console.log(
    "=== AI PERSONALISATION COMPLETED ==="
  );
}
function callGeminiWithUrlContext(prompt) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("GEMINI_API_KEY not found.");
  }

  const config = getAIConfig();

  const model = String(config.AI_MODEL).trim();

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent";

  const payload = {

    contents: [
      {
        role: "user",
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],

    // This enables Gemini URL Context
    tools: [
      {
        url_context: {}
      }
    ]

  };

  const options = {

    method: "post",

    contentType: "application/json",

    headers: {
      "x-goog-api-key": apiKey
    },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true
  };


  const response = UrlFetchApp.fetch(
    url,
    options
  );

  const statusCode =
    response.getResponseCode();

  const responseText =
    response.getContentText();


  if (
    statusCode < 200 ||
    statusCode >= 300
  ) {

    throw new Error(
      "Gemini URL Context API error " +
      statusCode +
      ": " +
      responseText
    );
  }


  const result =
    JSON.parse(responseText);


  if (
    !result.candidates ||
    result.candidates.length === 0
  ) {

    throw new Error(
      "Gemini returned no candidates."
    );
  }


  const parts =
    result.candidates[0]
      .content
      .parts;


  const text = parts
    .map(function(part) {
      return part.text || "";
    })
    .join("")
    .trim();


  if (!text) {
    throw new Error(
      "Gemini returned no text."
    );
  }


  return {
    text: text,
    raw: result
  };
}
function testWebsiteContext() {

  const website =
    "https://www.dynamisers.com/";


  const prompt = `
Analyse the following public company website:

${website}

Use information retrieved from the supplied URL.

Return ONLY valid JSON using this structure:

{
  "companySummary": "brief description",
  "servicesFound": ["service 1", "service 2"],
  "targetAudience": "brief description",
  "evidence": [
    "specific fact found on the website",
    "another specific fact found on the website"
  ]
}

Rules:
- Use only information supported by the supplied website.
- Do not invent facts.
- Keep the company summary under 50 words.
- Keep evidence concise.
- If something cannot be determined, use "unknown".
- Do not wrap the JSON in markdown.
`;


  try {

    console.log(
      "=== WEBSITE CONTEXT TEST ==="
    );

    console.log(
      "Website: " + website
    );


    const result =
      callGeminiWithUrlContext(
        prompt
      );


    console.log(
      "Gemini result:"
    );

    console.log(
      result.text
    );


  } catch (error) {

    console.error(
      "Website context test failed: " +
      error.message
    );

  }
}
function researchWebsite(lead) {

  if (
    !lead.website ||
    String(lead.website).trim() === ""
  ) {
    throw new Error(
      "Website URL missing for " + lead.email
    );
  }


  const website =
    String(lead.website).trim();


  const prompt = `
Research this company's public website:

${website}

COMPANY PROVIDED BY LEAD DATA:
${lead.company || "Unknown"}

Your job is to extract factual business context that can later
be used to personalise a B2B cold email.

Use information retrieved from the supplied website only.

Return ONLY valid JSON in exactly this structure:

{
  "companySummary": "short factual summary",
  "servicesFound": [
    "service/product 1",
    "service/product 2"
  ],
  "targetAudience": "who the company appears to serve",
  "valueProposition": "main value proposition if clearly stated",
  "usefulFacts": [
    "specific factual observation",
    "specific factual observation"
  ],
  "researchConfidence": "high, medium, or low"
}

RULES:

- Do not invent information.
- Do not guess statistics.
- Do not guess revenue, company size, traffic or rankings.
- Do not infer website performance.
- Do not claim SEO problems.
- Do not claim advertising problems.
- Do not make subjective statements about website quality.
- Extract only information supported by the website.
- Keep companySummary under 50 words.
- Keep usefulFacts concise.
- Return at most 3 usefulFacts.
- Return at most 5 services/products.
- If information cannot be determined, use "unknown".
- If the website contains very little useful information,
  set researchConfidence to "low".
- Do not wrap the JSON in markdown.
`;


  const response =
    callGeminiWithUrlContext(prompt);


  let cleaned =
    response.text.trim();


  // Defensive cleanup
  cleaned = cleaned
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/, "")
    .replace(/```$/, "")
    .trim();


  let research;


  try {

    research =
      JSON.parse(cleaned);

  } catch (error) {

    throw new Error(
      "Website research returned invalid JSON: " +
      response.text
    );
  }


  // -------------------------------
  // BASIC VALIDATION
  // -------------------------------

  if (!research.companySummary) {
    research.companySummary = "unknown";
  }

  if (!Array.isArray(research.servicesFound)) {
    research.servicesFound = [];
  }

  if (!research.targetAudience) {
    research.targetAudience = "unknown";
  }

  if (!research.valueProposition) {
    research.valueProposition = "unknown";
  }

  if (!Array.isArray(research.usefulFacts)) {
    research.usefulFacts = [];
  }


  const allowedConfidence = [
    "high",
    "medium",
    "low"
  ];


  const confidence =
    String(
      research.researchConfidence || "low"
    )
      .toLowerCase()
      .trim();


  research.researchConfidence =
    allowedConfidence.includes(confidence)
      ? confidence
      : "low";


  return research;
}
function testWebsiteResearch() {

  const testLead = {

    company: "YOUR TEST COMPANY",

    email: "test@example.com",

    website: "dynamisers.com"

  };


  try {

    console.log(
      "=== WEBSITE RESEARCH STARTED ==="
    );


    const research =
      researchWebsite(testLead);


    console.log(
      JSON.stringify(
        research,
        null,
        2
      )
    );


    console.log(
      "=== WEBSITE RESEARCH COMPLETE ==="
    );


  } catch (error) {

    console.error(
      "Website research failed: " +
      error.message
    );

  }
}

function testResearchPersonalisation() {

  const testLead = {

    name: "Divyan",

    company: "YOUR TEST COMPANY",

    email: "test@example.com",

    website: "YOUR REAL TEST WEBSITE",

    industry: "Technology",

    service: "SEO"

  };


  try {

    console.log(
      "=== RESEARCHING WEBSITE ==="
    );


    const research =
      researchWebsite(testLead);


    console.log(
      "Research:"
    );

    console.log(
      JSON.stringify(
        research,
        null,
        2
      )
    );


    console.log(
      "=== GENERATING PERSONALISATION ==="
    );


    const intro =
      generatePersonalisation(
        testLead,
        research
      );


    console.log(
      "FINAL INTRO:"
    );

    console.log(intro);


  } catch (error) {

    console.error(
      "Research personalisation test failed: " +
      error.message
    );

  }
}
function testResearchPersonalisation() {

  const testLead = {

    name: "Divyan",

    company: "YOUR TEST COMPANY",

    email: "test@example.com",

    website: "dynamisers.com",

    industry: "Technology",

    service: "SEO"

  };


  try {

    console.log(
      "=== RESEARCHING WEBSITE ==="
    );


    const research =
      researchWebsite(testLead);


    console.log(
      "Research:"
    );

    console.log(
      JSON.stringify(
        research,
        null,
        2
      )
    );


    console.log(
      "=== GENERATING PERSONALISATION ==="
    );


    const intro =
      generatePersonalisation(
        testLead,
        research
      );


    console.log(
      "FINAL INTRO:"
    );

    console.log(intro);


  } catch (error) {

    console.error(
      "Research personalisation test failed: " +
      error.message
    );

  }
}
function saveWebsiteResearch(lead, research) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("AIResearch");


  if (!sheet) {
    throw new Error(
      "AIResearch sheet not found."
    );
  }


  sheet.appendRow([

    lead.leadId,

    lead.company,

    lead.website,

    research.companySummary,

    JSON.stringify(
      research.servicesFound || []
    ),

    research.targetAudience,

    research.valueProposition,

    JSON.stringify(
      research.usefulFacts || []
    ),

    research.researchConfidence,

    new Date()

  ]);


  console.log(
    "Website research cached for " +
    lead.email
  );
}
function getCachedWebsiteResearch(lead) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("AIResearch");


  if (!sheet) {
    throw new Error(
      "AIResearch sheet not found."
    );
  }


  const data =
    sheet.getDataRange().getValues();


  // Search backwards so newest research wins
  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    const cachedLeadId =
      String(data[i][0]).trim();

    const currentLeadId =
      String(lead.leadId).trim();


    if (
      cachedLeadId === currentLeadId
    ) {

      let servicesFound = [];
      let usefulFacts = [];


      try {

        servicesFound =
          JSON.parse(
            data[i][4] || "[]"
          );

      } catch (error) {

        servicesFound = [];

      }


      try {

        usefulFacts =
          JSON.parse(
            data[i][7] || "[]"
          );

      } catch (error) {

        usefulFacts = [];

      }


      return {

        companySummary:
          data[i][3] || "unknown",

        servicesFound:
          servicesFound,

        targetAudience:
          data[i][5] || "unknown",

        valueProposition:
          data[i][6] || "unknown",

        usefulFacts:
          usefulFacts,

        researchConfidence:
          data[i][8] || "low",

        researchedAt:
          data[i][9]

      };
    }
  }


  return null;
}
function getWebsiteResearch(lead) {

  // --------------------------------
  // CHECK CACHE
  // --------------------------------

  const cached =
    getCachedWebsiteResearch(lead);


  if (cached) {

    console.log(
      "Using cached website research for " +
      lead.email
    );

    return cached;
  }


  // --------------------------------
  // NO CACHE → RESEARCH
  // --------------------------------

  console.log(
    "No cached research found for " +
    lead.email
  );


  console.log(
    "Researching website: " +
    lead.website
  );


  const research =
    researchWebsite(lead);


  // --------------------------------
  // SAVE
  // --------------------------------

  saveWebsiteResearch(
    lead,
    research
  );


  return research;
}
function isWorkingDay(config) {

  const timezone =
    String(
      config.TIMEZONE || "Europe/London"
    ).trim();


  const configuredDays =
    String(
      config.WORKING_DAYS ||
      "MON,TUE,WED,THU,FRI"
    )
      .toUpperCase()
      .split(",")
      .map(function(day) {
        return day.trim();
      });


  const now = new Date();


  // EEE produces values such as:
  // MON, TUE, WED...
  const today =
    Utilities.formatDate(
      now,
      timezone,
      "EEE"
    ).toUpperCase();


  const allowed =
    configuredDays.includes(today);


  console.log(
    "Working day check | " +
    "Timezone: " + timezone +
    " | Today: " + today +
    " | Allowed: " + allowed
  );


  return allowed;
}
function testWorkingDay() {

  const config =
    getConfig();


  const result =
    isWorkingDay(config);


  console.log(
    "Automation allowed today: " +
    result
  );
}
function isWithinSendingWindow(config) {

  const timezone =
    String(
      config.TIMEZONE || "Europe/London"
    ).trim();


  const startHour =
    Number(config.SEND_START_HOUR);


  const endHour =
    Number(config.SEND_END_HOUR);


  if (
    isNaN(startHour) ||
    isNaN(endHour)
  ) {

    throw new Error(
      "Invalid SEND_START_HOUR or SEND_END_HOUR."
    );
  }


  const now =
    new Date();


  const currentHour =
    Number(
      Utilities.formatDate(
        now,
        timezone,
        "H"
      )
    );


  const allowed =
    currentHour >= startHour &&
    currentHour < endHour;


  console.log(
    "Sending window check | " +
    "Current hour: " +
    currentHour +
    " | Window: " +
    startHour +
    ":00-" +
    endHour +
    ":00" +
    " | Allowed: " +
    allowed
  );


  return allowed;
}

function testSendingWindow() {

  const config =
    getConfig();


  const result =
    isWithinSendingWindow(config);


  console.log(
    "Inside sending window: " +
    result
  );
}
function canAutomationRunNow() {

  const config =
    getConfig();


  if (!isWorkingDay(config)) {

    console.log(
      "Automation blocked: Non-working day."
    );

    return false;
  }


  if (!isWithinSendingWindow(config)) {

    console.log(
      "Automation blocked: Outside sending window."
    );

    return false;
  }


  console.log(
    "Automation schedule checks passed."
  );


  return true;
}
function testAutomationSchedule() {

  const allowed =
    canAutomationRunNow();


  console.log(
    "Final schedule result: " +
    allowed
  );
}
function isFastTestMode(config) {

  return (
    config.FAST_TEST_MODE === true ||
    String(config.FAST_TEST_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
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
function testFollowUpTiming() {

  const config =
    getConfig();


  console.log(
    "FAST_TEST_MODE: " +
    isFastTestMode(config)
  );


  for (
    let followUp = 1;
    followUp <= 3;
    followUp++
  ) {

    const delay =
      getFollowUpDelay(
        config,
        followUp
      );


    console.log(
      "Follow-up " +
      followUp +
      ": " +
      delay.value +
      " " +
      delay.mode
    );
  }
}
function isAutomationScheduleAllowed(config) {

  // ==============================
  // TEST MODE BYPASS
  // ==============================

  if (isFastTestMode(config)) {

    console.log(
      "FAST TEST MODE enabled."
    );

    console.log(
      "Working-day and sending-window checks bypassed."
    );

    return true;
  }


  // ==============================
  // PRODUCTION CHECKS
  // ==============================

  if (!isWorkingDay(config)) {

    console.log(
      "Automation blocked: Non-working day."
    );

    return false;
  }


  if (!isWithinSendingWindow(config)) {

    console.log(
      "Automation blocked: Outside sending window."
    );

    return false;
  }


  console.log(
    "Production schedule checks passed."
  );


  return true;
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
function isPreviewMode(config) {

  return (
    config.PREVIEW_MODE === true ||
    String(config.PREVIEW_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
}
function buildEmailPreview(lead, step) {

  const template =
    getTemplate(
      lead.service,
      step
    );

  if (!template) {
    throw new Error(
      "Template not found: " +
      lead.service +
      " / " +
      step
    );
  }


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


  return {
  subject: subject,
  body: body
};
}
function generatePreview() {

  const config = getConfig();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const leadsSheet =
    ss.getSheetByName("Leads");

  const previewSheet =
    ss.getSheetByName("Preview");


  if (!leadsSheet) {
    throw new Error("Leads sheet not found.");
  }

  if (!previewSheet) {
    throw new Error("Preview sheet not found.");
  }


  // Clear old previews but preserve header
  const lastRow =
    previewSheet.getLastRow();

  if (lastRow > 1) {

    previewSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        previewSheet.getLastColumn()
      )
      .clearContent();
  }


  // Generate missing AI intros first.
  generateMissingPersonalizations();


  // Reload because AI may have updated Leads.
  const data =
    leadsSheet
      .getDataRange()
      .getValues();


  let previewCount = 0;


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
      status: data[i][14]

    };


    if (lead.status !== "NEW") {
      continue;
    }


    if (!lead.email) {
      continue;
    }


    if (!lead.service) {
      continue;
    }


    try {

      const email =
        buildEmailPreview(
          lead,
          "EMAIL_1"
        );


      const previewText =
  "SUBJECT:\n" +
  email.subject +
  "\n\nBODY:\n" +
  email.body;


      previewSheet.appendRow([

        new Date(),
        lead.leadId,
        lead.email,
        lead.company,
        lead.service,
        "EMAIL_1",
        previewText

      ]);


      previewCount++;


      console.log(
        "Preview generated for " +
        lead.email +
        " | " +
        lead.service
      );


    } catch (error) {

      console.error(
        "Preview failed for " +
        lead.email +
        ": " +
        error.message
      );

    }

  }


  console.log(
    "Preview generation complete. " +
    previewCount +
    " emails ready."
  );
}
function isTestMode(config) {

  return (
    config.TEST_MODE === true ||
    String(config.TEST_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
}
function getActualRecipient(lead, config) {

  if (isTestMode(config)) {

    const testEmail =
      String(config.TEST_EMAIL || "")
        .trim();


    if (!testEmail) {
      throw new Error(
        "TEST_MODE is enabled but TEST_EMAIL is empty."
      );
    }


    console.log(
      "TEST MODE | Original recipient: " +
      lead.email +
      " | Redirected to: " +
      testEmail
    );


    return testEmail;
  }


  return lead.email;
}
/**
 * Builds the final MVP operational dashboard.
 *
 * SAFE TO RE-RUN:
 * - Only clears/rebuilds the Dashboard sheet.
 * - Does NOT modify Leads, ActivityLog, Services, Config, Templates,
 *   AIResearch, Preview or AIConfig.
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dashboard = ss.getSheetByName("Dashboard");
  const leads = ss.getSheetByName("Leads");
  const activity = ss.getSheetByName("ActivityLog");
  const servicesSheet = ss.getSheetByName("Services");
  const configSheet = ss.getSheetByName("Config");

  if (!dashboard) throw new Error("Dashboard sheet not found.");
  if (!leads) throw new Error("Leads sheet not found.");
  if (!activity) throw new Error("ActivityLog sheet not found.");
  if (!servicesSheet) throw new Error("Services sheet not found.");
  if (!configSheet) throw new Error("Config sheet not found.");

  const services = getServices();

  // ---------------------------------------------------------
  // RESET DASHBOARD ONLY
  // ---------------------------------------------------------

  dashboard.clear();
  dashboard.clearConditionalFormatRules();

  // Remove existing charts.
  dashboard.getCharts().forEach(function(chart) {
    dashboard.removeChart(chart);
  });

  dashboard.setHiddenGridlines(true);

  // ---------------------------------------------------------
  // COLOURS
  // ---------------------------------------------------------

  const NAVY = "#0F172A";
  const SLATE = "#334155";
  const MUTED = "#64748B";
  const LIGHT = "#F8FAFC";
  const BORDER = "#E2E8F0";
  const GREEN = "#DCFCE7";
  const RED = "#FEE2E2";
  const AMBER = "#FEF3C7";
  const BLUE = "#DBEAFE";
  const WHITE = "#FFFFFF";

  // ---------------------------------------------------------
  // COLUMN WIDTHS
  // ---------------------------------------------------------

  dashboard.setColumnWidth(1, 175); // A
  dashboard.setColumnWidth(2, 110); // B
  dashboard.setColumnWidth(3, 110); // C
  dashboard.setColumnWidth(4, 110); // D
  dashboard.setColumnWidth(5, 110); // E
  dashboard.setColumnWidth(6, 110); // F
  dashboard.setColumnWidth(7, 110); // G
  dashboard.setColumnWidth(8, 135); // H
  dashboard.setColumnWidth(9, 135); // I
  dashboard.setColumnWidth(10, 160); // J

  // ---------------------------------------------------------
  // HEADER
  // ---------------------------------------------------------

  dashboard.getRange("A1:J2").merge();

  dashboard.getRange("A1")
    .setValue("AI Cold Email Automation")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontSize(20)
    .setFontWeight("bold")
    .setVerticalAlignment("middle");

  dashboard.setRowHeight(1, 30);
  dashboard.setRowHeight(2, 30);

  dashboard.getRange("A3:J3").merge();

  dashboard.getRange("A3")
    .setValue("Campaign performance, pipeline health and automation activity")
    .setFontColor(MUTED)
    .setFontSize(10)
    .setBackground(LIGHT);

  // ---------------------------------------------------------
  // SYSTEM STATUS
  // ---------------------------------------------------------

  dashboard.getRange("A5:J5")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  dashboard.getRange("A5").setValue("SYSTEM STATUS");

  const configData = configSheet.getDataRange().getValues();
  const configMap = {};

  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0]) {
      configMap[String(configData[i][0]).trim()] = configData[i][1];
    }
  }

  dashboard.getRange("A6").setValue("Automation");
  dashboard.getRange("B6").setValue(
    configMap.AUTOMATION_ENABLED === true ? "ON" : "OFF"
  );

  dashboard.getRange("C6").setValue("Preview Mode");
  dashboard.getRange("D6").setValue(
    configMap.PREVIEW_MODE === true ? "ON" : "OFF"
  );

  dashboard.getRange("E6").setValue("Test Mode");
  dashboard.getRange("F6").setValue(
    configMap.TEST_MODE === true ? "ON" : "OFF"
  );

  dashboard.getRange("G6").setValue("Fast Test");
  dashboard.getRange("H6").setValue(
    configMap.FAST_TEST_MODE === true ? "ON" : "OFF"
  );

dashboard.getRange("I6").setValue("Daily Usage");

const dailyLimit = Number(configMap.DAILY_LIMIT) || 0;

dashboard.getRange("J6").setFormula(
  '=(' +
    'COUNTIFS(ActivityLog!A2:A,">="&TODAY(),' +
    'ActivityLog!A2:A,"<"&TODAY()+1,' +
    'ActivityLog!J2:J,"SUCCESS",' +
    'ActivityLog!D2:D,"EMAIL_1")+' +

    'COUNTIFS(ActivityLog!A2:A,">="&TODAY(),' +
    'ActivityLog!A2:A,"<"&TODAY()+1,' +
    'ActivityLog!J2:J,"SUCCESS",' +
    'ActivityLog!D2:D,"FOLLOWUP_1")+' +

    'COUNTIFS(ActivityLog!A2:A,">="&TODAY(),' +
    'ActivityLog!A2:A,"<"&TODAY()+1,' +
    'ActivityLog!J2:J,"SUCCESS",' +
    'ActivityLog!D2:D,"FOLLOWUP_2")+' +

    'COUNTIFS(ActivityLog!A2:A,">="&TODAY(),' +
    'ActivityLog!A2:A,"<"&TODAY()+1,' +
    'ActivityLog!J2:J,"SUCCESS",' +
    'ActivityLog!D2:D,"FOLLOWUP_3")' +
  ')&" / ' + dailyLimit + '"'
);

  dashboard.getRange("A6:J6")
    .setBackground(WHITE)
    .setBorder(true, true, true, true, true, true, BORDER, null);

  dashboard.getRangeList(["A6", "C6", "E6", "G6", "I6"])
    .setFontWeight("bold")
    .setFontColor(SLATE);

  // ---------------------------------------------------------
  // KPI SECTION
  // ---------------------------------------------------------

  dashboard.getRange("A8:J8")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  dashboard.getRange("A8").setValue("CAMPAIGN OVERVIEW");

  const kpis = [
    ["A9", "Total Leads", "A10", '=COUNTIF(Leads!D2:D,"<>")'],
    ["C9", "Active Leads", "C10",
      '=COUNTIFS(Leads!D2:D,"<>",Leads!O2:O,"<>COMPLETED",Leads!O2:O,"<>INVALID",Leads!O2:O,"<>DO_NOT_CONTACT",Leads!O2:O,"<>REPLIED")'
    ],
    ["E9", "Emails Sent", "E10",
  '=COUNTIFS(ActivityLog!J2:J,"SUCCESS",ActivityLog!D2:D,"EMAIL_1")+' +
  'COUNTIFS(ActivityLog!J2:J,"SUCCESS",ActivityLog!D2:D,"FOLLOWUP_1")+' +
  'COUNTIFS(ActivityLog!J2:J,"SUCCESS",ActivityLog!D2:D,"FOLLOWUP_2")+' +
  'COUNTIFS(ActivityLog!J2:J,"SUCCESS",ActivityLog!D2:D,"FOLLOWUP_3")'
],
    ["G9", "Replies", "G10",
      '=COUNTIF(Leads!O2:O,"REPLIED")'
    ],
    ["I9", "Reply Rate", "I10",
      '=IFERROR(COUNTIF(Leads!O2:O,"REPLIED")/COUNTIF(Leads!D2:D,"<>"),0)'
    ]
  ];

  kpis.forEach(function(kpi) {
    dashboard.getRange(kpi[0])
      .setValue(kpi[1])
      .setFontWeight("bold")
      .setFontColor(MUTED)
      .setHorizontalAlignment("center");

    dashboard.getRange(kpi[2])
      .setFormula(kpi[3])
      .setFontSize(18)
      .setFontWeight("bold")
      .setFontColor(NAVY)
      .setHorizontalAlignment("center");
  });

  dashboard.getRange("I10").setNumberFormat("0.0%");

  dashboard.getRange("A9:J10")
    .setBackground(WHITE)
    .setBorder(true, true, true, true, true, true, BORDER, null);

  // ---------------------------------------------------------
  // SECONDARY KPIs
  // ---------------------------------------------------------

const secondary = [
  ["A12", "New Leads", "A13",
    '=COUNTIF(Leads!O2:O,"NEW")'
  ],

  ["C12", "Completed", "C13",
    '=COUNTIF(Leads!O2:O,"COMPLETED")'
  ],

  ["E12", "Invalid / Bounced", "E13",
    '=COUNTIF(Leads!O2:O,"INVALID")'
  ],

  ["G12", "Do Not Contact", "G13",
    '=COUNTIF(Leads!O2:O,"DO_NOT_CONTACT")'
  ],

  ["I12", "Failed Actions", "I13",
    '=COUNTIF(ActivityLog!J2:J,"FAILED")'
  ]
];

  secondary.forEach(function(kpi) {
    dashboard.getRange(kpi[0])
      .setValue(kpi[1])
      .setFontWeight("bold")
      .setFontColor(MUTED)
      .setHorizontalAlignment("center");

    dashboard.getRange(kpi[2])
      .setFormula(kpi[3])
      .setFontSize(16)
      .setFontWeight("bold")
      .setFontColor(NAVY)
      .setHorizontalAlignment("center");
  });

  dashboard.getRange("A12:J13")
    .setBackground(LIGHT)
    .setBorder(true, true, true, true, true, true, BORDER, null);

  // ---------------------------------------------------------
  // SERVICE PERFORMANCE
  // ---------------------------------------------------------

  dashboard.getRange("A15:F15")
    .merge()
    .setValue("SERVICE PERFORMANCE")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  dashboard.getRange("A16:F16")
    .setValues([[
      "Service",
      "Current Leads",
      "Campaigns Started",
      "Replies",
      "Service Completed",
      "Reply Rate"
    ]])
    .setBackground(SLATE)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  let serviceRow = 17;

  services.forEach(function(service) {
    const safeService = service.name.replace(/"/g, '""');

    dashboard.getRange(serviceRow, 1).setValue(service.name);

dashboard.getRange(serviceRow, 2)
  .setFormula(
    '=COUNTIFS(' +
    'Leads!H2:H,"' + safeService + '",' +
    'Leads!O2:O,"<>COMPLETED",' +
    'Leads!O2:O,"<>INVALID",' +
    'Leads!O2:O,"<>DO_NOT_CONTACT",' +
    'Leads!O2:O,"<>REPLIED")'
  );

dashboard.getRange(serviceRow, 3)
  .setFormula(
    '=COUNTIFS(' +
    'ActivityLog!F2:F,"' + safeService + '",' +
    'ActivityLog!D2:D,"EMAIL_1",' +
    'ActivityLog!J2:J,"SUCCESS")'
  );

    dashboard.getRange(serviceRow, 4)
      .setFormula(
        '=COUNTIFS(Leads!H2:H,"' +
        safeService +
        '",Leads!O2:O,"REPLIED")'
      );

    dashboard.getRange(serviceRow, 5)
      .setFormula(
        '=COUNTIF(Leads!O2:O,"NOT_INTERESTED_SERVICE_' +
        service.order +
        '")'
      );

    dashboard.getRange(serviceRow, 6)
      .setFormula(
        '=IFERROR(D' +
        serviceRow +
        '/B' +
        serviceRow +
        ',0)'
      )
      .setNumberFormat("0.0%");

    serviceRow++;
  });

  if (services.length > 0) {
    dashboard.getRange(
      17,
      1,
      services.length,
      6
    )
      .setBorder(true, true, true, true, true, true, BORDER, null);
  }

  // ---------------------------------------------------------
  // PIPELINE
  // ---------------------------------------------------------

  dashboard.getRange("H15:J15")
    .merge()
    .setValue("CAMPAIGN PIPELINE")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  dashboard.getRange("H16:I16")
    .setValues([["Stage", "Leads"]])
    .setBackground(SLATE)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  const pipeline = [
    ["NEW", '=COUNTIF(Leads!O2:O,"NEW")'],
    ["Email 1 Sent", '=COUNTIF(Leads!O2:O,"EMAIL_1_SENT")'],
    ["Follow-up 1", '=COUNTIF(Leads!O2:O,"FOLLOWUP_1_SENT")'],
    ["Follow-up 2", '=COUNTIF(Leads!O2:O,"FOLLOWUP_2_SENT")'],
    ["Follow-up 3", '=COUNTIF(Leads!O2:O,"FOLLOWUP_3_SENT")'],
    ["Replied", '=COUNTIF(Leads!O2:O,"REPLIED")'],
    ["Completed", '=COUNTIF(Leads!O2:O,"COMPLETED")']
  ];

  pipeline.forEach(function(row, index) {
    const r = 17 + index;

    dashboard.getRange(r, 8).setValue(row[0]);
    dashboard.getRange(r, 9).setFormula(row[1]);
  });

  dashboard.getRange("H17:I23")
    .setBorder(true, true, true, true, true, true, BORDER, null);

// ---------------------------------------------------------
// DELIVERY OVERVIEW
// ---------------------------------------------------------

const deliveryRow = Math.max(serviceRow + 1, 22);

dashboard.getRange(deliveryRow, 1, 1, 6)
  .merge()
  .setValue("DELIVERY OVERVIEW")
  .setBackground(NAVY)
  .setFontColor(WHITE)
  .setFontWeight("bold");

dashboard.getRange(deliveryRow + 1, 1, 1, 6)
  .setValues([[
    "Sent Today",
    "Sent This Week",
    "Daily Remaining",
    "Follow-ups Due",
    "Daily Limit",
    "Utilisation"
  ]])
  .setBackground(SLATE)
  .setFontColor(WHITE)
  .setFontWeight("bold")
  .setHorizontalAlignment("center");

const sendActions =
  '((ActivityLog!D2:D="EMAIL_1")+' +
  '(ActivityLog!D2:D="FOLLOWUP_1")+' +
  '(ActivityLog!D2:D="FOLLOWUP_2")+' +
  '(ActivityLog!D2:D="FOLLOWUP_3"))';

dashboard.getRange(deliveryRow + 2, 1).setFormula(
  '=SUMPRODUCT(' +
  '--(ActivityLog!A2:A>=TODAY()),' +
  '--(ActivityLog!A2:A<TODAY()+1),' +
  '--(ActivityLog!J2:J="SUCCESS"),' +
  '--(' + sendActions + '>0))'
);

dashboard.getRange(deliveryRow + 2, 2).setFormula(
  '=SUMPRODUCT(' +
  '--(ActivityLog!A2:A>=TODAY()-WEEKDAY(TODAY(),2)+1),' +
  '--(ActivityLog!A2:A<TODAY()+1),' +
  '--(ActivityLog!J2:J="SUCCESS"),' +
  '--(' + sendActions + '>0))'
);

dashboard.getRange(deliveryRow + 2, 3).setFormula(
  '=MAX(0,' + dailyLimit + '-A' + (deliveryRow + 2) + ')'
);

dashboard.getRange(deliveryRow + 2, 5)
  .setValue(dailyLimit);

dashboard.getRange(deliveryRow + 2, 4)
  .setValue("-");

dashboard.getRange(deliveryRow + 2, 6)
  .setFormula(
    '=IFERROR(A' +
    (deliveryRow + 2) +
    '/E' +
    (deliveryRow + 2) +
    ',0)'
  )
  .setNumberFormat("0.0%");

dashboard.getRange(
  deliveryRow + 2,
  1,
  1,
  6
)
  .setFontSize(14)
  .setFontWeight("bold")
  .setHorizontalAlignment("center")
  .setBackground(LIGHT)
  .setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    BORDER,
    null
  );
  // ---------------------------------------------------------
  // RECENT ACTIVITY
  // ---------------------------------------------------------

  const activityStartRow = Math.max(deliveryRow + 5, 28);

  dashboard.getRange(
    activityStartRow,
    1,
    1,
    10
  )
    .merge()
    .setValue("RECENT ACTIVITY")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  dashboard.getRange(
    activityStartRow + 1,
    1,
    1,
    7
  )
    .setValues([[
      "Timestamp",
      "Email",
      "Action",
      "Campaign",
      "Service",
      "Status",
      "Result"
    ]])
    .setBackground(SLATE)
    .setFontColor(WHITE)
    .setFontWeight("bold");

  /*
   * Pull the 10 most recent ActivityLog records.
   * QUERY sorts newest first and ignores blank rows.
   */
  dashboard.getRange(
    activityStartRow + 2,
    1
  ).setFormula(
    '=QUERY(ActivityLog!A2:K,' +
    '"select A,C,D,E,F,G,J ' +
    'where A is not null ' +
    'order by A desc limit 10",0)'
  );

  // ---------------------------------------------------------
  // CONDITIONAL FORMATTING
  // ---------------------------------------------------------

  const rules = [];

  // System ON states
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ON")
      .setBackground(GREEN)
      .setFontColor("#166534")
      .setRanges([
        dashboard.getRange("B6"),
        dashboard.getRange("D6"),
        dashboard.getRange("F6"),
        dashboard.getRange("H6")
      ])
      .build()
  );

  // System OFF states
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OFF")
      .setBackground(RED)
      .setFontColor("#991B1B")
      .setRanges([
        dashboard.getRange("B6"),
        dashboard.getRange("D6"),
        dashboard.getRange("F6"),
        dashboard.getRange("H6")
      ])
      .build()
  );

  // Success
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SUCCESS")
      .setBackground(GREEN)
      .setFontColor("#166534")
      .setRanges([
        dashboard.getRange(
          activityStartRow + 2,
          7,
          20,
          1
        )
      ])
      .build()
  );

  // Failed
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("FAILED")
      .setBackground(RED)
      .setFontColor("#991B1B")
      .setRanges([
        dashboard.getRange(
          activityStartRow + 2,
          7,
          20,
          1
        )
      ])
      .build()
  );

  dashboard.setConditionalFormatRules(rules);

  // ---------------------------------------------------------
  // GENERAL FORMATTING
  // ---------------------------------------------------------

  dashboard.getRange("A1:J50")
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");

  dashboard.getRange("A1:J50")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  dashboard.getRange("A1:J50")
    .setFontSize(10);

  // Restore larger title/KPIs after global formatting.
  dashboard.getRange("A1").setFontSize(20);
  dashboard.getRangeList([
    "A10", "C10", "E10", "G10", "I10"
  ]).setFontSize(18);

  dashboard.getRangeList([
    "A13", "C13", "E13", "G13", "I13"
  ]).setFontSize(16);

  dashboard.setFrozenRows(3);

  // Give activity rows breathing room.
  for (
    let r = activityStartRow + 1;
    r <= activityStartRow + 12;
    r++
  ) {
    dashboard.setRowHeight(r, 28);
  }

  SpreadsheetApp.flush();

  console.log(
    "Dashboard setup complete. Active services: " +
    services.length
  );
}
/**
 * Returns true when the same email address has already started
 * the same service under a different Lead ID.
 *
 * This does NOT block legitimate progression to another service.
 */
function hasDuplicateCampaign(email, service, leadId) {

  const normalizedEmail = String(email || "")
    .trim()
    .toLowerCase();

  const normalizedService = String(service || "")
    .trim();

  const normalizedLeadId = String(leadId || "")
    .trim();

  if (!normalizedEmail || !normalizedService) {
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activitySheet = ss.getSheetByName("ActivityLog");

  if (!activitySheet || activitySheet.getLastRow() < 2) {
    return false;
  }

  const data = activitySheet.getDataRange().getValues();

  // IMPORTANT:
  // Adjust these indexes ONLY if your ActivityLog columns differ.
  //
  // A = Timestamp
  // B = Lead ID
  // C = Email
  // D = Action
  // F = Service
  // J = Result

  for (let i = 1; i < data.length; i++) {

    const loggedLeadId = String(data[i][1] || "").trim();

    const loggedEmail = String(data[i][2] || "")
      .trim()
      .toLowerCase();

    const action = String(data[i][3] || "").trim();
    const loggedService = String(data[i][5] || "").trim();
    const result = String(data[i][9] || "").trim();

    if (
      loggedEmail === normalizedEmail &&
      loggedService === normalizedService &&
      action === "EMAIL_1" &&
      result === "SUCCESS" &&
      loggedLeadId !== normalizedLeadId
    ) {
      return true;
    }
  }

  return false;
}