// ============================================================
// BOUNCE MANAGEMENT
// ============================================================
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


        const lead =
          buildLeadFromRow(data, i);


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