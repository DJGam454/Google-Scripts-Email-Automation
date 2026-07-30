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

function testReadSheet() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  const data = sheet.getDataRange().getValues();

  console.log(data);
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