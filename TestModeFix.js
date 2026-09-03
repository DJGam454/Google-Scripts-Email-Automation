// ============================================================
// TEST MODE FIXERS (one-time, dryRun by default)
// ============================================================
// 1) Fix dates that FAST_TEST_MODE wrote as minutes instead of
//    days: reconcile Leads H-K/L/R against Gmail internalDate.
//    Bursts are detected inside a caller-supplied time window,
//    not by email address.
// 2) Purge test ActivityLog rows from that window (bursts of
//    2-3 within the same minute, not by TEST_EMAIL).
//
// Both are read-only when dryRun=true. Live writes are batched.
// The scan window defaults to the last 48 hours when no window
// is supplied; pass explicit ISO strings for precision.

function _parseWindow(startIstStr, endIstStr) {

  // Default window: last 48h to 1h ahead.
  const fallbackStart = new Date(Date.now() - 48 * 60 * 60 * 1000);
  const fallbackEnd = new Date(Date.now() + 60 * 60 * 1000);

  const startIst = startIstStr
    ? new Date(startIstStr)
    : fallbackStart;

  const endIst = endIstStr
    ? new Date(endIstStr)
    : fallbackEnd;

  if (isNaN(startIst.getTime()) || isNaN(endIst.getTime())) {
    throw new Error("Invalid window: " + startIstStr + " / " + endIstStr);
  }

  return { startMs: startIst.getTime(), endMs: endIst.getTime(), startIst: startIst, endIst: endIst };
}

function _getGmailThreadDate(threadId) {

  try {
    const thread = Gmail.Users.Threads.get("me", threadId, { format: "minimal" });
    if (!thread || !thread.messages || thread.messages.length === 0) return null;
    // First message is EMAIL_1; its internalDate is the true send ms.
    const first = thread.messages[0];
    if (first.internalDate) return Number(first.internalDate);
  } catch (e) {}
  try {
    const thread2 = GmailApp.getThreadById(threadId);
    if (thread2) {
      const msgs = thread2.getMessages();
      if (msgs.length > 0) return msgs[0].getDate().getTime();
    }
  } catch (e2) {}
  return null;
}

function _isBurstActivityRow(timestampMs, mapMinuteToCount) {

  const d = new Date(timestampMs);
  const key = Utilities.formatDate(d, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm");
  return (mapMinuteToCount[key] || 0) >= 2 && (mapMinuteToCount[key] || 0) <= 3;
}

// ------------------------------------------------------------
// FIX DATES: reconcile Leads H-K/L/R against Gmail
// ------------------------------------------------------------

function reconcileTestDates(dryRun, startIstStr, endIstStr) {

  const isDry = dryRun !== false;
  const win = _parseWindow(startIstStr, endIstStr);

  console.log("=== RECONCILE TEST DATES | dryRun=" + isDry + " | window " + win.startIst.toISOString() + " to " + win.endIst.toISOString() + " ===");

  const sheet = getLeadsSheet();
  const data = sheet.getDataRange().getValues();
  const activity = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ActivityLog");

  // Build minute burst map from ActivityLog inside window (not by email).
  const minuteMap = {};
  if (activity && activity.getLastRow() >= 2) {
    const adata = activity.getDataRange().getValues();
    for (let i = 1; i < adata.length; i++) {
      const ts = adata[i][0];
      if (!ts) continue;
      const ms = new Date(ts).getTime();
      if (isNaN(ms) || ms < win.startMs || ms > win.endMs) continue;
      const action = String(adata[i][3] || "").trim();
      const result = String(adata[i][9] || "").trim();
      if (!["EMAIL_1", "FOLLOWUP_1", "FOLLOWUP_2", "FOLLOWUP_3"].includes(action)) continue;
      if (result !== "SUCCESS") continue;
      const key = Utilities.formatDate(new Date(ms), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm");
      minuteMap[key] = (minuteMap[key] || 0) + 1;
    }
  }

  let toFix = [];
  let checked = 0;

  for (let i = 1; i < data.length; i++) {
    const lead = buildLeadFromRow(data, i);
    if (!lead.threadId) continue;
    // Only rows that had activity inside window and look burst-like.
    const lastMs = lead.lastEmailDate ? new Date(lead.lastEmailDate).getTime() : 0;
    if (isNaN(lastMs) || (lastMs >= win.startMs && lastMs <= win.endMs)) {
      // fall through to Gmail check
    } else {
      // Also check if any of H-K falls in window
      let inWin = false;
      [lead.email1Sent, lead.followup1Sent, lead.followup2Sent, lead.followup3Sent].forEach(function(v) {
        const m = v ? new Date(v).getTime() : 0;
        if (!isNaN(m) && m >= win.startMs && m <= win.endMs) inWin = true;
      });
      if (!inWin) continue;
    }
    checked++;
    const gmailMs = _getGmailThreadDate(lead.threadId);
    if (!gmailMs || isNaN(gmailMs)) continue;

    // Detect minute-scale vs day-scale: compare first send (Leads H)
    // against the Gmail first message date.
    const sheetFirstMs = lead.email1Sent ? new Date(lead.email1Sent).getTime() : 0;
    if (!isNaN(sheetFirstMs) && !isNaN(gmailMs) && Math.abs(sheetFirstMs - gmailMs) > 60000) {
      toFix.push({ row: i + 1, leadId: lead.leadId, email: lead.email, status: lead.status, sheetL: lead.lastEmailDate, gmailL: new Date(gmailMs), threadId: lead.threadId, type: "L_MISMATCH" });
    } else if (lead.followup1Sent && lead.email1Sent) {
      const gap = new Date(lead.followup1Sent).getTime() - new Date(lead.email1Sent).getTime();
      if (!isNaN(gap) && gap < 24 * 60 * 60 * 1000 && gap >= 0) {
        // Prod gap should be days; minutes = test artifact.
        toFix.push({ row: i + 1, leadId: lead.leadId, email: lead.email, status: lead.status, sheetL: lead.lastEmailDate, gmailL: new Date(gmailMs), threadId: lead.threadId, type: "FAST_GAP" });
      }
    }
  }

  console.log("Leads checked with thread in window: " + checked + " | to fix (H-K/L/R): " + toFix.length);
  for (let i = 0; i < Math.min(toFix.length, 10); i++) {
    const f = toFix[i];
    console.log("  row " + f.row + " lead " + f.leadId + " " + f.email + " " + f.type + " sheetL=" + f.sheetL + " gmailL=" + f.gmailL);
  }

  if (isDry || toFix.length === 0) {
    console.log("Dry run - no rows written. Call reconcileTestDates(false) to fix dates.");
    console.log("=== RECONCILE DATES COMPLETE (dry) ===");
    return { checked: checked, toFix: toFix.length, sample: toFix.slice(0, 5) };
  }

  // Live: batch update L and R to Gmail date; H is derived from the
  // thread's first message date when available.
  let updated = 0;
  for (let i = 0; i < toFix.length; i++) {
    const f = toFix[i];
    const gMs = f.gmailL.getTime();
    // N - Last Email Date, R - Last Updated
    sheet.getRange(f.row, LEADS_COL.LAST_EMAIL_DATE + 1).setValue(new Date(gMs));
    sheet.getRange(f.row, LEADS_COL.LAST_UPDATED + 1).setValue(new Date());
    // Also align H if it was test - set to same Gmail first message date.
    const gmailFirst = _getGmailThreadDate(f.threadId);
    if (gmailFirst) {
      sheet.getRange(f.row, LEADS_COL.EMAIL_1_SENT + 1).setValue(new Date(gmailFirst));
    }
    updated++;
    if (updated % 20 === 0) Utilities.sleep(200);
  }

  console.log("Live fixed: " + updated + " leads L/R (and H) from Gmail.");
  console.log("=== RECONCILE DATES COMPLETE (live) ===");
  return { updated: updated };
}

function dryRunFixTestDates() { return reconcileTestDates(true); }
function liveFixTestDates() { return reconcileTestDates(false); }

// ------------------------------------------------------------
// PURGE TEST ACTIVITY LOG ROWS (burst detection, not email)
// ------------------------------------------------------------

function purgeTestActivityLog(dryRun, startIstStr, endIstStr, archiveName) {

  const isDry = dryRun !== false;
  const win = _parseWindow(startIstStr, endIstStr);
  const archiveSheetName = String(archiveName || "TestArchive").trim() || "TestArchive";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ActivityLog");
  if (!sheet) throw new Error("ActivityLog not found.");

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  // Build minute burst map inside window for EMAIL_1/FOLLOWUP SUCCESS only.
  const minuteMap = {};
  for (let i = 0; i < rows.length; i++) {
    const ts = rows[i][0];
    if (!ts) continue;
    const ms = new Date(ts).getTime();
    if (isNaN(ms) || ms < win.startMs || ms > win.endMs) continue;
    const action = String(rows[i][3] || "").trim();
    const result = String(rows[i][9] || "").trim();
    if (!["EMAIL_1", "FOLLOWUP_1", "FOLLOWUP_2", "FOLLOWUP_3"].includes(action)) continue;
    if (result !== "SUCCESS") continue;
    const key = Utilities.formatDate(new Date(ms), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd HH:mm");
    minuteMap[key] = (minuteMap[key] || 0) + 1;
  }

  let keep = [];
  let toDelete = [];
  let toArchive = [];

  for (let i = 0; i < rows.length; i++) {
    const ts = rows[i][0];
    const ms = ts ? new Date(ts).getTime() : 0;
    const inWin = !isNaN(ms) && ms >= win.startMs && ms <= win.endMs;
    const action = String(rows[i][3] || "").trim();
    const result = String(rows[i][9] || "").trim();
    const isBurst = inWin && ["EMAIL_1", "FOLLOWUP_1", "FOLLOWUP_2", "FOLLOWUP_3"].includes(action) && result === "SUCCESS" && _isBurstActivityRow(ms, minuteMap);

    if (isBurst) {
      toDelete.push(rows[i]);
      toArchive.push(rows[i]);
    } else {
      keep.push(rows[i]);
    }
  }

  console.log(
    "ActivityLog purge scan | total rows: " + rows.length +
    " | window " + win.startIst.toISOString() + " to " + win.endIst.toISOString() +
    " | to delete (bursts 2-3/min): " + toDelete.length +
    " | to keep: " + keep.length
  );

  for (let i = 0; i < Math.min(toDelete.length, 5); i++) {
    console.log("  sample delete: " + new Date(toDelete[i][0]).toISOString() + " | " + toDelete[i][2] + " | " + toDelete[i][3] + " | " + toDelete[i][9]);
  }

  if (isDry) {
    console.log("Dry run - no rows deleted. Call purgeTestActivityLog(false) to archive+delete.");
    return { total: rows.length, toDelete: toDelete.length, toKeep: keep.length };
  }

  // Archive first.
  let archive = ss.getSheetByName(archiveSheetName);
  if (!archive) {
    archive = ss.insertSheet(archiveSheetName);
    archive.appendRow(header);
    archive.setFrozenRows(1);
  }
  if (toArchive.length > 0) {
    archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, toArchive[0].length).setValues(toArchive);
    console.log("Archived " + toArchive.length + " rows to " + archiveSheetName + ".");
  }

  // Rewrite ActivityLog with header + keep.
  sheet.clearContents();
  const newData = [header].concat(keep);
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  console.log("Deleted " + toDelete.length + " rows from ActivityLog. Now " + keep.length + " data rows remain.");

  try { setupDashboard(); console.log("Dashboard rebuilt."); } catch (e) { console.log("Dashboard rebuild skipped: " + e.message); }

  return { deleted: toDelete.length, kept: keep.length };
}

function dryRunPurgeTestLog() { return purgeTestActivityLog(true); }
function livePurgeTestLog() { return purgeTestActivityLog(false); }

// ------------------------------------------------------------
// LOCAL-TIME DISPLAY HELPER (no storage change)
// ------------------------------------------------------------

function ensureLocalTimeDisplayHelper(timezoneId, columnLabel) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = String(timezoneId || "Asia/Kolkata").trim();
  const label = String(columnLabel || "Local Timestamp").trim();

  // Instruction only - the sheet timezone itself is changed in
  // File > Settings. This helper adds a display column derived
  // from the stored UTC timestamps.
  const offsetMinutes = _timezoneOffsetMinutes(tz);

  const sheet = ss.getSheetByName("ActivityLog");
  if (!sheet) throw new Error("ActivityLog not found.");

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, lastCol).getValue();
  if (String(header).trim() === label) {
    console.log(label + " column already exists at col " + lastCol);
    return;
  }

  sheet.getRange(1, lastCol + 1).setValue(label);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Stored dates are UTC numbers; offset converts to local display.
    const offsetHours = offsetMinutes / 60;
    sheet.getRange(2, lastCol + 1, lastRow - 1, 1).setFormulaR1C1("=RC[-" + lastCol + "]+TIME(" + offsetHours + ",0,0)");
    sheet.getRange(2, lastCol + 1, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }
  console.log("Added " + label + " column at " + (lastCol + 1) + " (formula A+offset " + offsetMinutes + "m).");
}

// Fixed offsets for common zones (DST-free); falls back to 0.
function _timezoneOffsetMinutes(timezoneId) {

  const offsets = {
    "Asia/Kolkata": 330,
    "Europe/London": 0,
    "America/New_York": -300,
    "America/Los_Angeles": -480,
    "Asia/Dubai": 240,
    "Asia/Singapore": 480,
    "Australia/Sydney": 600
  };

  return offsets[timezoneId] !== undefined ? offsets[timezoneId] : 0;
}
