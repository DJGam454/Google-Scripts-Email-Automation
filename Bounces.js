// ============================================================
// BOUNCE MANAGEMENT
// ============================================================
// Classifies DSNs into 4 categories:
//   HARD_BOUNCE       - mailbox not found / disabled (permanent 5.1.x)
//   SOFT_BOUNCE       - temporary 4xx / mailbox full / greylisted
//   POLICY_REJECTION  - reputation / content / policy 5.7.x
//   UNKNOWN           - unparseable, needs manual review
//
// Only HARD_BOUNCE is permanently suppressed. SOFT retries with
// capped count, POLICY is not retried blindly, UNKNOWN is logged.

function checkBounces() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName("Leads");

  if (!leadsSheet) {
    throw new Error("Leads sheet not found.");
  }

  console.log("=== CHECKING FOR BOUNCES ===");

  // Wider window (7d vs 2d) and more subjects: Yahoo/Outlook can
  // delay DSNs. The search is cheap; classification filters noise.
  const threads = GmailApp.search(
    'newer_than:7d (from:mailer-daemon OR from:mail-daemon OR subject:"Delivery Status Notification" OR subject:"Delivery incomplete" OR subject:"Undelivered Mail Returned" OR subject:"Mail delivery failed" OR subject:"failure notice" OR subject:"Returned mail")'
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

  // Ensure S-V columns exist before first write.
  ensureLeadsBounceColumns();
  _loadSuppressionCache();

  // Dedup: never process the same Gmail bounce Message ID twice.
  // Stored in Script Properties as JSON { msgId: timestamp } and pruned >7d.
  let seenIds = {};
  try {
    const raw = PropertiesService.getScriptProperties().getProperty("BOUNCE_SEEN_IDS");
    if (raw) {
      seenIds = JSON.parse(raw);
    }
  } catch (e) {
    seenIds = {};
  }

  const nowMs = Date.now();
  const pruneMs = nowMs - 7 * 24 * 60 * 60 * 1000;
  let pruned = 0;
  for (let k in seenIds) {
    if (seenIds[k] < pruneMs) {
      delete seenIds[k];
      pruned++;
    }
  }
  if (pruned > 0) {
    console.log("Pruned " + pruned + " expired bounce IDs from dedup cache.");
  }

  // Summary counters.
  let messagesScanned = 0;
  let identifiedBounces = 0;
  let unidentifiedMessages = 0;
  let leadsMarkedInvalid = 0;
  let alreadyInvalid = 0;
  let softBounced = 0;
  let policyRejected = 0;
  let unknownBounces = 0;
  let suppressedWrites = 0;
  let dedupSkipped = 0;


  for (let t = 0; t < threads.length; t++) {

    const messages =
      threads[t].getMessages();


    for (let m = 0; m < messages.length; m++) {

      messagesScanned++;

      const message =
        messages[m];

      const msgId = message.getId();

      if (msgId && seenIds[msgId]) {
        dedupSkipped++;
        continue;
      }

      const body =
        message.getPlainBody() || "";

      // Fallback: some Gmail delivery-status parts have no plain
      // body but carry payload in the raw MIME. getBody() includes
      // HTML part where the diagnostic may still be visible.
      const fallbackBody =
        !body
          ? (message.getBody() || "")
          : "";

      const combinedBody = body || fallbackBody;

      const result =
        classifyBounce(combinedBody);


      if (!result || !result.email) {

        unidentifiedMessages++;

        continue;
      }


      identifiedBounces++;

      console.log(
        "Bounce classified | Email: " + result.email +
        " | Category: " + result.category +
        " | Code: " + (result.code || "-") +
        " | Diag: " + (result.diagnostic || "").slice(0, 120)
      );


      // Find bounced email in Leads (case-insensitive exact match).
      for (let i = 1; i < data.length; i++) {

        const leadEmail =
          String(data[i][3] || "")
            .trim()
            .toLowerCase();


        if (
          leadEmail !==
          result.email.toLowerCase()
        ) {
          continue;
        }


        const row = i + 1;

        const currentStatus =
          String(data[i][14] || "").trim();

        const currentCategory =
          String(
            data[i].length > LEADS_COL.BOUNCE_CATEGORY
              ? data[i][LEADS_COL.BOUNCE_CATEGORY]
              : ""
          ).trim();

        const currentDiag =
          String(
            data[i].length > LEADS_COL.BOUNCE_DIAGNOSTIC
              ? data[i][LEADS_COL.BOUNCE_DIAGNOSTIC]
              : ""
          ).trim();

        // Already hard-suppressed - just ensure SuppressionList is
        // consistent and count as already-invalid for the summary.
        if (
          currentStatus === "INVALID" &&
          currentCategory === "HARD_BOUNCE"
        ) {
          // Keep suppression list warm (idempotent).
          addToSuppressionList(
            result.email,
            "HARD_BOUNCE",
            result.diagnostic,
            "DSN"
          );

          alreadyInvalid++;

          // Still mark this Gmail message as seen so future runs skip it entirely.
          if (msgId) {
            seenIds[msgId] = nowMs;
          }

          break;
        }

        // For non-HARD: if Leads S/T already exactly matches this bounce,
        // skip the re-log (idempotent second guard). Keeps ActivityLog clean.
        if (
          result.category !== "HARD_BOUNCE" &&
          currentCategory === result.category &&
          currentDiag &&
          currentDiag.slice(0, 80) === String(result.diagnostic || "").slice(0, 80)
        ) {
          console.log(
            "Bounce already logged for " + result.email +
            " | Category: " + result.category + " - skipping duplicate."
          );

          if (msgId) {
            seenIds[msgId] = nowMs;
          }

          dedupSkipped++;

          break;
        }

        const lead =
          buildLeadFromRow(data, i);


        // =================================
        // ROUTE BY CATEGORY
        // =================================

        if (result.category === "HARD_BOUNCE") {

          // O = Status
          leadsSheet
            .getRange(row, 15)
            .setValue("INVALID");

          // R = Last Updated
          leadsSheet
            .getRange(row, 18)
            .setValue(new Date());

          writeLeadBounceState(row, "HARD_BOUNCE", result.diagnostic, false);
          addToSuppressionList(result.email, "HARD_BOUNCE", result.diagnostic, "DSN");

          logActivity(
            lead,
            "EMAIL_BOUNCED",
            "INVALID",
            "FAILED",
            "",
            lead.threadId,
            "HARD_BOUNCE | " + (result.code || "") + " | " + (result.diagnostic || "").slice(0, 180)
          );

          leadsMarkedInvalid++;
          suppressedWrites++;

          console.log(
            result.email +
            " marked INVALID / HARD_BOUNCE: " + (result.diagnostic || "").slice(0, 120)
          );

        } else if (result.category === "SOFT_BOUNCE") {

          const alreadyRetries = Number(
            data[i].length > LEADS_COL.RETRY_COUNT
              ? data[i][LEADS_COL.RETRY_COUNT]
              : 0
          );

          const SOFT_MAX_RETRIES = 3;

          if (alreadyRetries >= SOFT_MAX_RETRIES) {
            // Promote to UNKNOWN for manual review after 3 softs.
            writeLeadBounceState(row, "UNKNOWN", result.diagnostic + " [soft cap reached]", true);

            logActivity(
              lead,
              "EMAIL_BOUNCED",
              lead.status,
              "FAILED",
              "",
              lead.threadId,
              "UNKNOWN | soft cap " + SOFT_MAX_RETRIES + " | " + (result.code || "") + " | " + (result.diagnostic || "").slice(0, 150)
            );

            unknownBounces++;

            console.log(
              result.email + " soft cap reached -> UNKNOWN"
            );

          } else {
            writeLeadBounceState(row, "SOFT_BOUNCE", result.diagnostic, true);

            logActivity(
              lead,
              "EMAIL_BOUNCED",
              lead.status,
              "FAILED",
              "",
              lead.threadId,
              "SOFT_BOUNCE | " + (result.code || "") + " | " + (result.diagnostic || "").slice(0, 180)
            );

            softBounced++;

            console.log(
              result.email + " soft bounce logged (retry " + (alreadyRetries + 1) + "/" + SOFT_MAX_RETRIES + ")"
            );
          }

        } else if (result.category === "POLICY_REJECTION") {

          writeLeadBounceState(row, "POLICY_REJECTION", result.diagnostic, false);

          logActivity(
            lead,
            "EMAIL_BOUNCED",
            lead.status,
            "FAILED",
            "",
            lead.threadId,
            "POLICY_REJECTION | " + (result.code || "") + " | " + (result.diagnostic || "").slice(0, 180)
          );

          policyRejected++;

          console.log(
            result.email + " policy rejection logged (no retry)"
          );

        } else {

          // UNKNOWN
          writeLeadBounceState(row, "UNKNOWN", result.diagnostic, false);

          logActivity(
            lead,
            "EMAIL_BOUNCED",
            lead.status,
            "FAILED",
            "",
            lead.threadId,
            "UNKNOWN | " + (result.code || "") + " | " + (result.diagnostic || "").slice(0, 180)
          );

          unknownBounces++;

          console.log(
            result.email + " unknown bounce logged for manual review"
          );
        }

        // Mark this Gmail bounce message as seen (all categories).
        if (msgId) {
          seenIds[msgId] = nowMs;
        }

        break;
      }
    }
  }


  // =================================
  // BOUNCE SCAN SUMMARY
  // =================================

  // Persist dedup cache.
  try {
    PropertiesService.getScriptProperties().setProperty(
      "BOUNCE_SEEN_IDS",
      JSON.stringify(seenIds)
    );
  } catch (e) {
    console.log("Could not persist bounce dedup cache: " + e.message);
  }

  console.log(
    "Bounce scan summary | " +
    "Threads: " +
    threads.length +
    " | Messages: " +
    messagesScanned +
    " | Bounces identified: " +
    identifiedBounces +
    " | Unrecognized: " +
    unidentifiedMessages +
    " | Dedup skipped: " + dedupSkipped +
    " | HARD -> INVALID: " +
    leadsMarkedInvalid +
    " | Already HARD: " +
    alreadyInvalid +
    " | Soft: " + softBounced +
    " | Policy: " + policyRejected +
    " | Unknown: " + unknownBounces +
    " | Suppression writes: " + suppressedWrites
  );


  console.log(
    "=== BOUNCE CHECK COMPLETE ==="
  );
}

// ============================================================
// BOUNCE CLASSIFIER
// ============================================================
// Returns { email, category, code, diagnostic } or null.
// Category: HARD_BOUNCE | SOFT_BOUNCE | POLICY_REJECTION | UNKNOWN
//
// Signals are extracted from Status, Diagnostic-Code, Action,
// Remote-MTA plus body heuristics and legacy header patterns.

function classifyBounce(body) {

  if (!body) {
    return null;
  }

  const text = String(body);

  // --------------------------------
  // EMAIL EXTRACTION
  // --------------------------------

  const emailPatterns = [
    /Final-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i,
    /Original-Recipient:\s*rfc822;\s*([^\s<>]+@[^\s<>]+)/i,
    /Recipient:\s*([^\s<>]+@[^\s<>]+)/i,
    /Your message wasn't delivered to\s+([^\s<>]+@[^\s<>]+)/i,
    /Diagnostic-Code:\s*[^;]+;\s*[^\n]*?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/i
  ];

  let email = null;

  for (let i = 0; i < emailPatterns.length; i++) {
    const m = text.match(emailPatterns[i]);
    if (m && m[1]) {
      email = m[1].replace(/[>,.;]+$/, "").trim();
      break;
    }
  }

  if (!email) {
    return null;
  }

  // --------------------------------
  // SIGNAL EXTRACTION
  // --------------------------------

  const diagMatch = text.match(/Diagnostic-Code:\s*([^\n]+)/i);
  const statusMatch = text.match(/Status:\s*([245]\.\d+\.\d+)/i);
  const actionMatch = text.match(/Action:\s*(\w+)/i);

  const diagnostic = diagMatch
    ? String(diagMatch[1] || "").trim().slice(0, 250)
    : "";

  const enhancedStatus = statusMatch
    ? String(statusMatch[1] || "").trim()
    : "";

  const action = actionMatch
    ? String(actionMatch[1] || "").trim().toLowerCase()
    : "";

  // Extract SMTP reply code (3 digits at start of a line or after ;).
  let smtpCode = "";
  const codeMatch = text.match(/\b([245]\d\d)\s*(?:[#\-\.]\d\.\d\.\d+)?\b/);
  if (codeMatch) {
    smtpCode = codeMatch[1];
  }
  // Prefer code from Diagnostic-Code line if present.
  if (diagnostic) {
    const diagCode = diagnostic.match(/\b([245]\d\d)\b/);
    if (diagCode) {
      smtpCode = diagCode[1];
    }
  }

  const lower = (text + " " + diagnostic).toLowerCase();

  // --------------------------------
  // POLICY / REPUTATION (check before HARD - 5.7.x beats 5.1.x)
  // --------------------------------

  const isPolicy =
    /5\.7\.\d/.test(text) ||
    /policy|reputation|blocked|blacklisted|blacklist|spam|bulk|rate limit.*policy|reverse dns|ptr record|dmarc|spf|dkim|content rejected|message rejected|client host.*blocked|ip.*blocked|sender.*blocked|domain.*blocked/.test(lower) ||
    (/5\.7\.1/.test(enhancedStatus) || /554\s*5\.7/.test(text));

  if (isPolicy) {
    return {
      email: email,
      category: "POLICY_REJECTION",
      code: smtpCode || enhancedStatus || "5.7.x",
      diagnostic: diagnostic || lower.slice(0, 200)
    };
  }

  // --------------------------------
  // HARD BOUNCE (permanent 5.1.x / mailbox disabled / not found)
  // --------------------------------

  const hardSignals =
    /mailbox (is )?disabled|mailbox not found|recipient.*does not exist|user unknown|no such user|mailbox unavailable|address rejected.*not found|recipient address rejected|invalid recipient|unknown recipient|undeliverable.*mailbox|recipient rejected.*not found/i.test(lower) ||
    (smtpCode && /^5[0-9][0-9]$/.test(smtpCode) && /5\.1\.1|5\.1\.2|5\.1\.10|5\.2\.1|5\.4\.7.*mailbox|mailbox/.test(lower)) ||
    (enhancedStatus && /^5\.1\./.test(enhancedStatus) && /mailbox|recipient|user/i.test(lower)) ||
    (/^5/.test(smtpCode) && action === "failed" && /5\.1\./.test(enhancedStatus));

  // Explicit hard codes without policy overlap: 550/551/552/553 with mailbox language,
  // or 554.30 (Yahoo mailbox disabled), or 552 mailbox not found plateau.
  if (hardSignals) {
    return {
      email: email,
      category: "HARD_BOUNCE",
      code: smtpCode || enhancedStatus || "5.1.1",
      diagnostic: diagnostic || lower.slice(0, 200)
    };
  }

  // --------------------------------
  // SOFT BOUNCE (4xx temporary)
  // --------------------------------

  const isSoft =
    /^4/.test(smtpCode) ||
    /^4\./.test(enhancedStatus) ||
    /temporarily|try again later|deferred|greylisted|server busy|mailbox.*temporarily unavailable|over quota|mailbox full|quota exceeded.*temporarily|4\.7\.0.*too many|421.*too many/i.test(lower) ||
    (action === "delayed" || action === "transient");

  if (isSoft) {
    return {
      email: email,
      category: "SOFT_BOUNCE",
      code: smtpCode || enhancedStatus || "4.x.x",
      diagnostic: diagnostic || lower.slice(0, 200)
    };
  }

  // --------------------------------
  // FALLBACK: UNKNOWN
  // --------------------------------

  return {
    email: email,
    category: "UNKNOWN",
    code: smtpCode || enhancedStatus || "",
    diagnostic: diagnostic || lower.slice(0, 200)
  };
}

function extractBouncedEmail(body) {

  const result = classifyBounce(body);
  return result ? result.email : null;
}

function testBounceParser() {

  const samples = [
    {
      body: "Final-Recipient: rfc822; yahoo_user@yahoo.com\nAction: failed\nStatus: 5.1.1\nDiagnostic-Code: smtp; 554 5.1.1 <yahoo_user@yahoo.com>: Recipient address rejected: mailbox not found",
      expect: "HARD_BOUNCE"
    },
    {
      body: "Final-Recipient: rfc822; disabled@yahoo.com\nAction: failed\nStatus: 5.2.1\nDiagnostic-Code: smtp; 554 30 Mailbox is disabled",
      expect: "HARD_BOUNCE"
    },
    {
      body: "Final-Recipient: rfc822; victim@yahoo.com\nAction: failed\nStatus: 5.2.1\nDiagnostic-Code: smtp; 552 1 Requested mail action aborted, mailbox not found",
      expect: "HARD_BOUNCE"
    },
    {
      body: "Final-Recipient: rfc822; temp@yahoo.com\nAction: delayed\nStatus: 4.7.0\nDiagnostic-Code: smtp; 421 4.7.0 [TSS04] Messages from 1.2.3.4 temporarily deferred",
      expect: "SOFT_BOUNCE"
    },
    {
      body: "Final-Recipient: rfc822; blocked@provider.example\nAction: failed\nStatus: 5.7.1\nDiagnostic-Code: smtp; 550 5.7.1 Your IP has been blocked",
      expect: "POLICY_REJECTION"
    },
    {
      body: "Final-Recipient: rfc822; generic@unknown.com\nAction: failed\nStatus: 5.0.0\nDiagnostic-Code: smtp; 550 Unknown error",
      expect: "UNKNOWN"
    },
    {
      body: "Your message wasn't delivered to broken@example.com because the address couldn't be found.",
      expect: "HARD_BOUNCE"
    }
  ];

  let passed = 0;

  for (let i = 0; i < samples.length; i++) {
    const r = classifyBounce(samples[i].body);
    const got = r ? r.category : "null";
    const ok = got === samples[i].expect;
    console.log(
      "Sample " + (i + 1) + " | Expected " + samples[i].expect +
      " | Got " + got + " | " + (ok ? "PASS" : "FAIL") +
      " | Email: " + (r ? r.email : "-")
    );
    if (ok) {
      passed++;
    }
  }

  console.log("testBounceParser: " + passed + "/" + samples.length + " passed");
}

function testBounceClassification() {
  return testBounceParser();
}
