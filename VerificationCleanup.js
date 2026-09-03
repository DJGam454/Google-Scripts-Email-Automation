// ============================================================
// VERIFICATION CLEANUP (bulk verification export)
// ============================================================
// Applies a bulk email-verification export to the sheet:
// undeliverable addresses -> HARD_BOUNCE suppression, risky
// addresses -> UNKNOWN suppression. Both are written to the
// SuppressionList so the pre-send gate blocks them forever.
//
// Populate the two lists below from your verification vendor's
// export (one entry per address). Leave them empty and the
// helper becomes a no-op - it is safe to run at any time.
//
// Dry run logs every decision, live run writes S/T/U +
// SuppressionList + Dashboard rebuild. No Leads rows are deleted.

var VERIFICATION_UNDELIVERABLE = [
  // { "email": "", "reason": "Invalid username" }
];

var VERIFICATION_RISKY = [
  // { "email": "", "reason": "Accept all" }
];

function applyVerificationCleanup(dryRun) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName("Leads");
  if (!leadsSheet) throw new Error("Leads sheet not found.");

  ensureLeadsBounceColumns();
  _loadSuppressionCache();

  const data = leadsSheet.getDataRange().getValues();
  const emailToRow = {};
  for (let i = 1; i < data.length; i++) {
    const e = String(data[i][LEADS_COL.EMAIL] || "").trim().toLowerCase();
    if (e) emailToRow[e] = i + 1; // 1-based row
  }

  let suppressedNew = 0;
  let alreadySuppressed = 0;
  let notFound = 0;
  let markedInvalid = 0;

  function processList(list, category) {
    for (let j = 0; j < list.length; j++) {
      const item = list[j];
      const email = String(item.email || "").trim().toLowerCase();
      const reason = String(item.reason || "").trim().slice(0, 200);
      if (!email) continue;

      const row = emailToRow[email];
      const already = isSuppressed(email);

      if (!row) {
        notFound++;
        if (!dryRun) {
          // Still add to SuppressionList even if not in Leads (future import protection)
          addToSuppressionList(email, category, reason + " | bulk verification", "PRE_VERIFY");
        } else {
          console.log("[DRY RUN] Not in Leads but would suppress: " + email + " | " + category);
        }
        continue;
      }

      if (already) {
        alreadySuppressed++;
        continue;
      }

      if (dryRun) {
        console.log("[DRY RUN] Would suppress: " + email + " | " + category + " | Reason: " + reason + " | Row: " + row);
      } else {
        // Leads S/T/U
        writeLeadBounceState(row, category, reason + " | bulk verification", false);
        // For clear INVALID signal on undeliverable, also set INVALID status if currently NEW
        const currentStatus = String(data[row - 1][LEADS_COL.STATUS] || "").trim();
        if (currentStatus === "NEW" || currentStatus === "") {
          leadsSheet.getRange(row, LEADS_COL.STATUS + 1).setValue("INVALID");
          leadsSheet.getRange(row, LEADS_COL.LAST_UPDATED + 1).setValue(new Date());
          markedInvalid++;
        }
        addToSuppressionList(email, category, reason + " | bulk verification", "PRE_VERIFY");
        suppressedNew++;
      }
    }
  }

  console.log("=== VERIFICATION CLEANUP START ===");
  console.log("Mode: " + (dryRun ? "DRY RUN" : "LIVE") + " | Undeliverable: " + VERIFICATION_UNDELIVERABLE.length + " | Risky: " + VERIFICATION_RISKY.length);
  console.log("Leads rows: " + (data.length - 1) + " | Already suppressed: " + Object.keys(_loadSuppressionCache()).length);

  // Undeliverable -> HARD_BOUNCE, Risky -> UNKNOWN (both suppressed via SuppressionList)
  processList(VERIFICATION_UNDELIVERABLE, "HARD_BOUNCE");
  processList(VERIFICATION_RISKY, "UNKNOWN");

  if (!dryRun) {
    // Also ensure dashboard reflects new INVALID count
    try { setupDashboard(); } catch(e) {}
  }

  console.log("Cleanup summary | Suppressed new: " + suppressedNew + " | Already suppressed: " + alreadySuppressed + " | Marked INVALID: " + markedInvalid + " | Not found in Leads: " + notFound);
  console.log("=== VERIFICATION CLEANUP COMPLETE ===");
  return { suppressedNew, alreadySuppressed, markedInvalid, notFound };
}

function testVerificationCleanup() {
  return applyVerificationCleanup(true);
}
