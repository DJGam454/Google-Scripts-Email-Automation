function testReadSheet() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  const data = sheet.getDataRange().getValues();

  console.log(data);
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

function testDailyCount() {

  const count = getEmailsSentToday();

  console.log(
    "Emails sent today: " + count
  );
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

function testAutomationSchedule() {

  const allowed =
    canAutomationRunNow();


  console.log(
    "Final schedule result: " +
    allowed
  );
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
