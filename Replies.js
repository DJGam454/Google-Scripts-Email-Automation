// ============================================================
// REPLY DETECTION
// ============================================================
// Scans every lead in an active campaign state (with a stored
// Gmail Thread ID) for a human reply.
//
// How it works:
//   1. Fetch the lead's thread via the Gmail API in metadata
//      format (headers only, no bodies).
//   2. A message counts as a reply when ALL of these hold:
//        - it is NOT a message we sent (no SENT/DRAFT label)
//        - it is NOT an auto-reply, bounce or delivery notice
//        - it did NOT come from our own mailbox
//        - its timestamp is newer than the lead's
//          Last Email Date (filters out old inbound mail)
//   3. Replies sent through the mailto "reply to this email"
//      link start a brand-new thread (no In-Reply-To headers),
//      so a second pass searches the mailbox for any message
//      FROM the lead's address that sits outside the tracked
//      thread and is newer than the last send.
//   4. On match: Status -> REPLIED, Last Updated -> now, and a
//      REPLY_DETECTED row is appended to the ActivityLog. The
//      ActivityLog timestamp is the recorded reply date.
//
// Idempotency: REPLIED is a terminal stop-status, so a lead is
// flipped exactly once — later runs skip it automatically.
// Reply detection runs before follow-ups, so a lead that just
// replied never receives another email.

// Run-scoped cache for the mailbox pass. Each unique lead email
// is searched once per scan, not once per row.
var _mailboxReplyCache = {};

// Run-scoped cache of the most recent INBOX messages (no query),
// loaded once per scan. Reading this list is near-real-time,
// unlike the `q:` search index which can lag 30-60s behind newly
// received mail.
var _recentInboxCache = null;

// ============================================================
// MAIN SCAN
// ============================================================

function checkReplies(dryRun) {

  // Standalone trigger (every 5 min) shares the script lock with
  // runAutomation so a REPLIED is not written while a follow-up
  // batch is scanning a stale in-memory snapshot. The lock is
  // tryLock(0) - never wait, just skip this scan if automation
  // is mid-send; the main run will catch replies on its next cycle.
  let _replyLock = null;
  let _gotLock = false;

  try {
    _replyLock = LockService.getScriptLock();
    _gotLock = _replyLock.tryLock(0);
  } catch (e) {
    _gotLock = false;
  }

  if (!dryRun && !_gotLock) {
    console.log("checkReplies skipped: automation holds the lock (sending).");
    return 0;
  }

  try {

  const sheet = getLeadsSheet();

  const data =
    sheet.getDataRange().getValues();

  console.log("=== CHECKING FOR REPLIES ===");

  // Reset the mailbox caches for this run.
  _mailboxReplyCache = {};
  _recentInboxCache = null;

  let scanned = 0;
  let replied = 0;
  let noThread = 0;
  let threadErrors = 0;
  let mailboxReplies = 0;

  for (let i = 1; i < data.length; i++) {

    const lead =
      buildLeadFromRow(data, i);

    // --------------------------------
    // ONLY ACTIVE CAMPAIGNS
    // --------------------------------

    if (!ACTIVE_CAMPAIGN_STATUSES.includes(lead.status)) {
      continue;
    }

    // --------------------------------
    // THREAD REQUIRED
    // --------------------------------

    if (!lead.threadId) {

      noThread++;

      continue;
    }

    scanned++;

    try {

      let reply =
        _findReplyInThread(lead);

      // Mailto replies live in a NEW thread, so the thread scan
      // can't see them. Fall back to a mailbox search for
      // anything inbound from the lead's address that arrived
      // after the last send and outside the tracked thread.
      if (!reply) {

        reply =
          _findReplyInMailbox(lead);

        if (reply) {

          mailboxReplies++;
        }
      }

      if (!reply) {
        continue;
      }

      // =========================
      // REPLY FOUND
      // =========================

      const row = i + 1;

      const now = new Date();

      console.log(
        "Reply detected for " +
        lead.email +
        " | Thread: " +
        lead.threadId +
        " | Source: " +
        reply.source +
        " | From: " +
        reply.from +
        " | Date: " +
        new Date(reply.timestamp).toISOString()
      );

      if (dryRun) {

        console.log(
          "[DRY RUN] Would set status REPLIED and log REPLY_DETECTED."
        );

        replied++;

        continue;
      }

      // M - Status
      sheet
        .getRange(row, LEADS_COL.STATUS + 1)
        .setValue("REPLIED");

      // P - Last Updated
      sheet
        .getRange(row, LEADS_COL.LAST_UPDATED + 1)
        .setValue(now);

      logActivity(
        lead,
        "REPLY_DETECTED",
        "REPLIED",
        "SUCCESS",
        reply.messageId,
        lead.threadId,
        ""
      );

      replied++;

    } catch (error) {

      threadErrors++;

      console.error(
        "Reply check failed for " +
        lead.email +
        ": " +
        error.message
      );
    }
  }

  console.log(
    "Reply scan summary | Active leads with threads: " +
    scanned +
    " | Replies detected: " +
    replied +
    " (thread: " +
    (replied - mailboxReplies) +
    ", mailbox: " +
    mailboxReplies +
    ") | No thread: " +
    noThread +
    " | Thread errors: " +
    threadErrors
  );

  console.log("=== REPLY CHECK COMPLETE ===");

  return replied;

  } finally {
    if (_gotLock && _replyLock) {
      try { _replyLock.releaseLock(); } catch (e) {}
    }
  }
}

// ============================================================
// THREAD ANALYSIS
// ============================================================
// Returns the newest qualifying inbound message, or null.
// Per-lead isolation is handled by the caller.

function _findReplyInThread(lead) {

  const thread =
    Gmail.Users.Threads.get(
      "me",
      lead.threadId,
      {
        format: "metadata",
        metadataHeaders: [
          "From",
          "Subject",
          "Auto-Submitted",
          "Precedence",
          "X-Auto-Response-Suppress"
        ]
      }
    );

  if (
    !thread ||
    !thread.messages ||
    thread.messages.length === 0
  ) {
    return null;
  }

  const lastSentTime =
    _getLastSentTime(lead);

  const myEmail = _getOwnEmail();

  let found = null;

  for (let m = 0; m < thread.messages.length; m++) {

    const message = thread.messages[m];

    // --------------------------------
    // SKIP OUR OWN OUTBOUND MESSAGES
    // --------------------------------

    const labels = message.labelIds || [];

    if (
      labels.indexOf("SENT") !== -1 ||
      labels.indexOf("DRAFT") !== -1
    ) {
      continue;
    }

    const headers = message.payload.headers || [];

    const fromRaw = _getHeader(headers, "From");
    const subject = _getHeader(headers, "Subject");

    // --------------------------------
    // SKIP NON-HUMAN MESSAGES
    // --------------------------------

    if (_isAutomatedMessage(headers, fromRaw, subject)) {
      continue;
    }

    const fromEmail = _extractEmail(fromRaw);

    if (!fromEmail) {
      continue;
    }

    if (
      myEmail &&
      String(fromEmail).toLowerCase() ===
      String(myEmail).toLowerCase()
    ) {
      continue;
    }

    // --------------------------------
    // MUST BE NEWER THAN OUR LAST EMAIL
    // --------------------------------

    const timestamp = Number(message.internalDate);

    if (
      isNaN(timestamp) ||
      timestamp <= lastSentTime
    ) {
      continue;
    }

    // Newest qualifying message wins.
    if (
      !found ||
      timestamp > found.timestamp
    ) {
      found = {
        messageId: message.id,
        from: fromRaw,
        timestamp: timestamp,
        source: "thread"
      };
    }
  }

  return found;
}

// ============================================================
// MAILBOX-SCAN REPLY DETECTION
// ============================================================
// Catches replies that bypass the original thread — e.g. when
// the recipient replies via the mailto "reply to this email"
// link. Those messages cannot carry In-Reply-To/References, so
// they start a fresh conversation. This pass searches the
// mailbox for anything inbound FROM the lead after the last
// send that is outside the tracked thread.
//
// Two tiers:
//   1. Recent INBOX list (no query, maxResults 100, cached per
//      run) — near-real-time read that catches a reply within
//      seconds of arrival.
//   2. `q:` search (maxResults 5) — fallback for replies older
//      than the recent window; the search index can lag 30-60s.
//
// Results are cached per lead email so multiple rows for the
// same address trigger only one search per run.

function _findReplyInMailbox(lead) {

  const key =
    String(lead.email || "")
      .trim()
      .toLowerCase();

  if (!key) {
    return null;
  }

  if (
    _mailboxReplyCache.hasOwnProperty(key)
  ) {
    return _mailboxReplyCache[key];
  }

  const lastSentTime =
    _getLastSentTime(lead);

  let result = null;

  try {

    // --------------------------------
    // TIER 1 — RECENT INBOX (FAST)
    // --------------------------------
    // Loaded once per run; no query means no search-index lag.

    if (_recentInboxCache === null) {

      _recentInboxCache = [];

      const list =
        Gmail.Users.Messages.list(
          "me",
          {
            labelIds: ["INBOX"],
            maxResults: 100
          }
        );

      const messages =
        (list && list.messages) || [];

      for (let i = 0; i < messages.length; i++) {

        const message =
          Gmail.Users.Messages.get(
            "me",
            messages[i].id,
            {
              format: "metadata",
              metadataHeaders: [
                "From",
                "Subject",
                "Message-ID"
              ]
            }
          );

        if (message) {
          _recentInboxCache.push(message);
        }
      }
    }

    for (let i = 0; i < _recentInboxCache.length; i++) {

      const match =
        _matchMailboxReply(
          _recentInboxCache[i],
          lead,
          key,
          lastSentTime
        );

      if (
        match &&
        (!result ||
         match.timestamp > result.timestamp)
      ) {
        result = match;
      }
    }

    // --------------------------------
    // TIER 2 — SEARCH FALLBACK
    // --------------------------------
    // Covers replies older than the recent-window list (e.g.
    // arrived before the previous run).

    if (!result) {

      const list =
        Gmail.Users.Messages.list(
          "me",
          {
            q: "from:" + key,
            maxResults: 5
          }
        );

      const messages =
        (list && list.messages) || [];

      for (let i = 0; i < messages.length; i++) {

        const message =
          Gmail.Users.Messages.get(
            "me",
            messages[i].id,
            {
              format: "metadata",
              metadataHeaders: [
                "From",
                "Subject",
                "Message-ID"
              ]
            }
          );

        if (!message) {
          continue;
        }

        const match =
          _matchMailboxReply(
            message,
            lead,
            key,
            lastSentTime
          );

        if (
          match &&
          (!result ||
           match.timestamp > result.timestamp)
        ) {
          result = match;
        }
      }
    }

  } catch (error) {

    result = null;
  }

  _mailboxReplyCache[key] = result;

  return result;
}

// Shared filter: does this mailbox message count as a reply to
// the lead? Returns the reply record, or null.
function _matchMailboxReply(message, lead, key, lastSentTime) {

  // Anything still inside the tracked thread is handled by
  // the thread scan — only NEW threads count here.
  if (
    message.threadId === lead.threadId
  ) {
    return null;
  }

  // --------------------------------
  // SKIP OUR OWN OUTBOUND MESSAGES
  // --------------------------------

  const labels = message.labelIds || [];

  if (
    labels.indexOf("SENT") !== -1 ||
    labels.indexOf("DRAFT") !== -1
  ) {
    return null;
  }

  const headers =
    message.payload.headers || [];

  const fromRaw =
    _getHeader(headers, "From");

  const subject =
    _getHeader(headers, "Subject");

  // --------------------------------
  // SKIP NON-HUMAN MESSAGES
  // --------------------------------

  if (_isAutomatedMessage(headers, fromRaw, subject)) {
    return null;
  }

  const fromEmail =
    _extractEmail(fromRaw);

  if (!fromEmail) {
    return null;
  }

  // Strict sender match — the message must really come from
  // the lead's own address to be treated as a reply.
  if (
    String(fromEmail)
      .toLowerCase() !== key
  ) {
    return null;
  }

  // --------------------------------
  // MUST BE NEWER THAN OUR LAST EMAIL
  // --------------------------------

  const timestamp =
    Number(message.internalDate);

  if (
    isNaN(timestamp) ||
    timestamp <= lastSentTime
  ) {
    return null;
  }

  return {
    messageId: message.id,
    from: fromRaw,
    timestamp: timestamp,
    source: "mailbox"
  };
}

// Millisecond timestamp of the lead's last sent email (0 when
// unset) — the cutoff for what counts as a new inbound reply.
function _getLastSentTime(lead) {

  const lastEmailDate =
    new Date(lead.lastEmailDate);

  return isNaN(lastEmailDate.getTime())
    ? 0
    : lastEmailDate.getTime();
}

// ============================================================
// HEADER HELPERS
// ============================================================

function _getHeader(headers, name) {

  for (let i = 0; i < headers.length; i++) {

    if (
      String(headers[i].name || "")
        .toLowerCase() ===
      name.toLowerCase()
    ) {
      return headers[i].value || "";
    }
  }

  return "";
}

function _extractEmail(fromRaw) {

  const match = String(fromRaw || "").match(
    /[\w.+-]+@[\w-]+\.[\w.]+/
  );

  return match ? match[0] : "";
}

// Identifies bounces, out-of-office and other automated
// messages that must never be treated as a human reply.
function _isAutomatedMessage(headers, fromRaw, subject) {

  const autoSubmitted = _getHeader(
    headers,
    "Auto-Submitted"
  );

  if (
    autoSubmitted &&
    String(autoSubmitted)
      .trim()
      .toLowerCase() !== "no"
  ) {
    return true;
  }

  const precedence = _getHeader(
    headers,
    "Precedence"
  );

  if (
    /auto_reply|bulk|junk/i.test(
      String(precedence)
    )
  ) {
    return true;
  }

  const responseSuppress = _getHeader(
    headers,
    "X-Auto-Response-Suppress"
  );

  if (responseSuppress) {
    return true;
  }

  const text =
    String(fromRaw || "") + " " +
    String(subject || "");

  return /mailer-daemon|postmaster|delivery status|undeliverable|delivery failure|returned mail|out of office|auto[- ]?reply|automatic reply|vacation|failure notice|recipient address rejected/i
    .test(text);
}

function _getOwnEmail() {

  try {

    return Session
      .getEffectiveUser()
      .getEmail() || "";

  } catch (error) {

    return "";
  }
}

// ============================================================
// TESTS
// ============================================================

// Dry run: scans active leads and reports what would be
// marked as replied, without writing to the sheet.
function testReplyDetection() {

  const replied =
    checkReplies(true);

  console.log(
    "Dry-run replies that would be marked: " +
    replied
  );
}
