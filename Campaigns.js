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
      isFlagTrue(active);

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