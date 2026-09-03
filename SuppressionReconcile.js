// ============================================================
// SUPPRESSION RECONCILE - HISTORICAL + 30D MAILBOX BACKFILL
// ============================================================
// One-time helpers to catch bounces missed before the S-V /
// SuppressionList layer landed. Does not delete any Leads rows -
// writes only to SuppressionList + S/T/U/V.
//
// Dry run (default) never sends mail, never writes sheets.
// Live run (dryRun=false) backfills SuppressionList + S/T/U so
// the existing pre-send gate (EmailEngine isSuppressed) blocks
// the next send automatically.
//
// Domain blocking is config-driven: SUPPRESSED_DOMAINS in the
// Config sheet holds a comma-separated list (e.g.
// "example.com,example.org"). isBlockedDomainEmail() is the
// single gate reused by the pre-send checks and the backfill.

function normalizeSuppressionEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function getSuppressedDomains(config) {

  const raw = String(
    (config || getConfig()).SUPPRESSED_DOMAINS || ""
  ).trim();

  if (!raw) {
    return [];
  }

  return raw
    .split(",")
    .map(function(domain) {
      return String(domain || "").trim().toLowerCase();
    })
    .filter(Boolean);
}

function isBlockedDomainEmail(email) {

  const n = normalizeSuppressionEmail(email);

  if (!n || n.indexOf("@") === -1) {
    return false;
  }

  const domain = n.split("@").pop();

  const domains = getSuppressedDomains();

  for (let i = 0; i < domains.length; i++) {
    if (domain === domains[i]) {
      return true;
    }
  }

  return false;
}

function _buildCurrentLeadIndex() {

  const data = getLeadsData();
  const index = {};

  for (let i = 1; i < data.length; i++) {
    const email = normalizeSuppressionEmail(data[i][LEADS_COL.EMAIL]);
    if (!email) {
      continue;
    }
    index[email] = {
      row: i + 1,
      status: String(data[i][LEADS_COL.STATUS] || "").trim(),
      bounceCategory: String(data[i].length > LEADS_COL.BOUNCE_CATEGORY ? data[i][LEADS_COL.BOUNCE_CATEGORY] : "").trim(),
      retryCount: Number(data[i].length > LEADS_COL.RETRY_COUNT ? data[i][LEADS_COL.RETRY_COUNT] : 0) || 0,
      leadId: data[i][LEADS_COL.LEAD_ID]
    };
  }

  return { data: data, index: index };
}

function _buildSuppressionIndexNormalized() {

  const cache = _loadSuppressionCache();
  const out = {};

  Object.keys(cache).forEach(function(k) {
    const nk = normalizeSuppressionEmail(k);
    out[nk] = cache[k];
  });

  return out;
}

function _isHardKeywordNormalized(lowerText) {

  const t = String(lowerText || "").toLowerCase();
  const hasCode = /550|551|552|553|554/.test(t);
  const hasPhrase = /mailbox not found|user unknown|address rejected|domain does not exist|recipient address rejected|no such user|unknown recipient/.test(t);
  return hasCode || hasPhrase;
}

// ------------------------------------------------------------
// 30D MAILBOX SCAN (Gmail) + RE-VERIFY AGAINST CURRENT LIST
// ------------------------------------------------------------
// Scans Gmail for DSNs in the last `daysBack` days (default 30),
// classifies via existing classifyBounce(), then reconciles
// against Leads + SuppressionList.
// dryRun=true logs what would be suppressed; false writes.

function reconcileMailboxBounces(dryRun, daysBack) {

  // Default: dry run, 30 days.
  const isDry = dryRun !== false;
  const days = Number(daysBack) || 30;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsInfo = _buildCurrentLeadIndex();
  const leadIndex = leadsInfo.index;
  const leadData = leadsInfo.data;
  const suppressionIndex = _buildSuppressionIndexNormalized();

  console.log("=== RECONCILE MAILBOX BOUNCES | dryRun=" + isDry + " | daysBack=" + days + " ===");
  console.log("Leads total: " + (leadData.length - 1) + " | Already in SuppressionList: " + Object.keys(suppressionIndex).length);

  // --------------------------------
  // 1. SCAN GMAIL FOR BOUNCE THREADS
  // --------------------------------

  const query =
    'newer_than:' + days + 'd (from:mailer-daemon OR from:mail-daemon OR from:postmaster@outlook.com OR from:googlemail.com OR subject:"Delivery Status Notification" OR subject:"Delivery incomplete" OR subject:"Undelivered Mail Returned" OR subject:"Mail delivery failed" OR subject:"failure notice" OR subject:"Returned mail" OR subject:"Delivery has failed")';

  const threads = GmailApp.search(query);

  console.log("Gmail bounce threads found (last " + days + "d): " + threads.length);

  const bounceMap = {}; // normalized email -> {category, code, diagnostic, sources:[], count}
  let messagesScanned = 0;
  let unidentified = 0;

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (let m = 0; m < messages.length; m++) {
      messagesScanned++;
      const msg = messages[m];
      const body = msg.getPlainBody() || msg.getBody() || "";
      const result = classifyBounce(body);
      if (!result || !result.email) {
        unidentified++;
        continue;
      }

      const norm = normalizeSuppressionEmail(result.email);
      if (!bounceMap[norm]) {
        bounceMap[norm] = {
          email: norm,
          origEmail: result.email,
          category: result.category,
          code: result.code,
          diagnostic: result.diagnostic,
          count: 0,
          sources: []
        };
      }

      // Keep the hardest category seen for this address.
      const rank = { "HARD_BOUNCE": 4, "POLICY_REJECTION": 3, "SOFT_BOUNCE": 2, "UNKNOWN": 1 };
      if ((rank[result.category] || 0) > (rank[bounceMap[norm].category] || 0)) {
        bounceMap[norm].category = result.category;
        bounceMap[norm].code = result.code;
        bounceMap[norm].diagnostic = result.diagnostic;
      }

      bounceMap[norm].count++;
      bounceMap[norm].sources.push(result.category + "|" + (result.code || ""));
    }
  }

  console.log("Messages scanned: " + messagesScanned + " | Identified bounces (unique emails): " + Object.keys(bounceMap).length + " | Unidentified: " + unidentified);

  // Include ActivityLog historical bounces as well (for repeat count).
  try {
    const activity = ss.getSheetByName("ActivityLog");
    if (activity && activity.getLastRow() >= 2) {
      const adata = activity.getDataRange().getValues();
      for (let i = 1; i < adata.length; i++) {
        const action = String(adata[i][3] || "").trim();
        const err = String(adata[i][10] || "");
        const emailRaw = String(adata[i][2] || "").trim();
        if (action !== "EMAIL_BOUNCED" || !emailRaw) {
          continue;
        }
        const norm = normalizeSuppressionEmail(emailRaw);
        if (!bounceMap[norm]) {
          // Derive category from ActivityLog error prefix if present.
          let cat = "UNKNOWN";
          if (err.indexOf("HARD_BOUNCE") !== -1) cat = "HARD_BOUNCE";
          else if (err.indexOf("SOFT_BOUNCE") !== -1) cat = "SOFT_BOUNCE";
          else if (err.indexOf("POLICY_REJECTION") !== -1) cat = "POLICY_REJECTION";
          else if (_isHardKeywordNormalized(err)) cat = "HARD_BOUNCE";
          bounceMap[norm] = {
            email: norm,
            origEmail: emailRaw,
            category: cat,
            code: "",
            diagnostic: err.slice(0, 250),
            count: 1,
            sources: ["ActivityLog:" + cat]
          };
        } else {
          bounceMap[norm].count++;
        }
      }
    }
  } catch (e) {
    console.log("ActivityLog scan skipped: " + e.message);
  }

  // --------------------------------
  // 2. RE-VERIFY AGAINST CURRENT LIST
  // --------------------------------

  let hardToSuppress = 0;
  let softToNote = 0;
  let policyToNote = 0;
  let unknownToNote = 0;
  let alreadySuppressed = 0;
  let wouldBeSkipped = 0;
  const toAdd = []; // list of {email, category, diagnostic}

  const keys = Object.keys(bounceMap);
  for (let i = 0; i < keys.length; i++) {
    const norm = keys[i];
    const rec = bounceMap[norm];
    const inSuppression = Object.prototype.hasOwnProperty.call(suppressionIndex, norm);
    if (inSuppression) {
      alreadySuppressed++;
      continue;
    }

    // Promote repeated POLICY_REJECTION to HARD for backfill.
    let effectiveCat = rec.category;
    if (rec.category === "POLICY_REJECTION" && rec.count >= 2) {
      effectiveCat = "HARD_BOUNCE";
    }

    // SOFT only promotes if Leads V >=3 or mailbox count >=3.
    if (rec.category === "SOFT_BOUNCE") {
      const leadInfo = leadIndex[norm];
      const retries = leadInfo ? leadInfo.retryCount : 0;
      if (retries >= 3 || rec.count >= 3) {
        effectiveCat = "UNKNOWN";
      }
    }

    if (effectiveCat === "HARD_BOUNCE") {
      hardToSuppress++;
      toAdd.push({ email: norm, category: "HARD_BOUNCE", diagnostic: rec.diagnostic, code: rec.code, count: rec.count });
      const li = leadIndex[norm];
      if (li && ACTIVE_CAMPAIGN_STATUSES.concat(["NEW"]).includes(li.status)) {
        wouldBeSkipped++;
      }
    } else if (rec.category === "SOFT_BOUNCE") {
      softToNote++;
    } else if (rec.category === "POLICY_REJECTION") {
      policyToNote++;
    } else {
      unknownToNote++;
    }
  }

  console.log(
    "Re-verify summary | mailbox uniques: " + keys.length +
    " | hard to suppress: " + hardToSuppress +
    " | soft: " + softToNote +
    " | policy: " + policyToNote +
    " | unknown: " + unknownToNote +
    " | already in SuppressionList: " + alreadySuppressed +
    " | leads that would be skipped (active): " + wouldBeSkipped
  );

  if (toAdd.length > 0) {
    console.log("Top hard addresses to suppress (max 10):");
    for (let i = 0; i < Math.min(toAdd.length, 10); i++) {
      console.log("  " + toAdd[i].email + " | " + toAdd[i].code + " | count=" + toAdd[i].count + " | " + String(toAdd[i].diagnostic).slice(0, 100));
    }
  }

  if (isDry) {
    console.log("Dry run - no rows written. Call reconcileMailboxBounces(false, " + days + ") to backfill SuppressionList.");
    console.log("=== RECONCILE COMPLETE (dry) ===");
    return {
      totalLeadsChecked: leadData.length - 1,
      mailboxBouncedUniques: keys.length,
      hardToSuppress: hardToSuppress,
      softToNote: softToNote,
      policyToNote: policyToNote,
      unknownToNote: unknownToNote,
      alreadySuppressed: alreadySuppressed,
      leadsThatWouldBeSkipped: wouldBeSkipped,
      toAdd: toAdd
    };
  }

  // --------------------------------
  // 3. LIVE BACKFILL (SuppressionList + S/T/U)
  // --------------------------------

  ensureLeadsBounceColumns();
  _loadSuppressionCache();

  let added = 0;
  let sUpdated = 0;

  for (let i = 0; i < toAdd.length; i++) {
    const item = toAdd[i];
    addToSuppressionList(item.email, item.category, item.diagnostic, "MAILBOX_REVERIFY");
    added++;
    const li = leadIndex[item.email];
    if (li) {
      writeLeadBounceState(li.row, item.category, item.diagnostic, false);
      const status = String(leadData[li.row - 1][LEADS_COL.STATUS] || "").trim();
      if (status !== "INVALID" && status !== "DO_NOT_CONTACT" && status !== "REPLIED" && status !== "COMPLETED") {
        getLeadsSheet().getRange(li.row, LEADS_COL.STATUS + 1).setValue("INVALID");
        getLeadsSheet().getRange(li.row, LEADS_COL.LAST_UPDATED + 1).setValue(new Date());
      }
      sUpdated++;
    }
  }

  // Also ensure any HARD already in Leads S but missing from SuppressionList gets added (legacy rows).
  const data = getLeadsData();
  let legacyAdded = 0;
  for (let i = 1; i < data.length; i++) {
    const norm = normalizeSuppressionEmail(data[i][LEADS_COL.EMAIL]);
    if (!norm || isSuppressed(norm)) {
      continue;
    }
    const cat = String(data[i].length > LEADS_COL.BOUNCE_CATEGORY ? data[i][LEADS_COL.BOUNCE_CATEGORY] : "").trim();
    if (cat === "HARD_BOUNCE") {
      const diag = String(data[i].length > LEADS_COL.BOUNCE_DIAGNOSTIC ? data[i][LEADS_COL.BOUNCE_DIAGNOSTIC] : "").trim();
      addToSuppressionList(norm, "HARD_BOUNCE", diag || "Backfilled legacy HARD", "LEGACY_S");
      legacyAdded++;
    }
  }

  try { setupDashboard(); } catch (e) { console.log("Dashboard rebuild skipped: " + e.message); }

  console.log("Backfill live | added to SuppressionList: " + added + " | S/T updated: " + sUpdated + " | legacy S -> Suppression: " + legacyAdded);
  console.log("=== RECONCILE COMPLETE (live) ===");

  return {
    totalLeadsChecked: leadData.length - 1,
    mailboxBouncedUniques: keys.length,
    hardToSuppress: hardToSuppress,
    added: added,
    sUpdated: sUpdated,
    legacyAdded: legacyAdded,
    leadsThatWouldBeSkipped: wouldBeSkipped
  };
}

// Shorthand wrappers for the editor Run menu.

function dryRunReconcileMailbox30d() {
  return reconcileMailboxBounces(true, 30);
}

function liveReconcileMailbox30d() {
  return reconcileMailboxBounces(false, 30);
}

// ------------------------------------------------------------
// DOMAIN BLOCK (backfill + pre-send, config-driven)
// ------------------------------------------------------------
// SUPPRESSED_DOMAINS lists domains that are globally blocked
// before every send. Dry/live helpers backfill existing Leads.

function dryRunDomainSuppression() {

  const leadsInfo = _buildCurrentLeadIndex();
  const leadData = leadsInfo.data;
  const leadIndex = leadsInfo.index;
  const suppressionIndex = _buildSuppressionIndexNormalized();

  let totalBlocked = 0;
  let alreadySuppressed = 0;
  let wouldBeAdded = 0;
  let wouldBeSkipped = 0;
  const samples = [];

  Object.keys(leadIndex).forEach(function(norm) {
    if (!isBlockedDomainEmail(norm)) return;
    totalBlocked++;
    const inSupp = Object.prototype.hasOwnProperty.call(suppressionIndex, norm);
    if (inSupp) {
      alreadySuppressed++;
      return;
    }
    wouldBeAdded++;
    const st = leadIndex[norm].status;
    if (ACTIVE_CAMPAIGN_STATUSES.concat(["NEW"]).includes(st)) {
      wouldBeSkipped++;
    }
    if (samples.length < 10) samples.push(norm + " | " + st + " | row " + leadIndex[norm].row);
  });

  console.log(
    "Domain-block dry-run | total blocked-domain leads: " + totalBlocked +
    " | already in SuppressionList: " + alreadySuppressed +
    " | would be added: " + wouldBeAdded +
    " | leads that would be skipped (active): " + wouldBeSkipped
  );
  if (samples.length > 0) {
    console.log("Sample blocked-domain leads to add (max 10):");
    for (let i = 0; i < samples.length; i++) console.log("  " + samples[i]);
  }

  return {
    totalBlocked: totalBlocked,
    alreadySuppressed: alreadySuppressed,
    wouldBeAdded: wouldBeAdded,
    wouldBeSkipped: wouldBeSkipped
  };
}

function liveDomainSuppression() {

  const leadsInfo = _buildCurrentLeadIndex();
  const leadIndex = leadsInfo.index;
  ensureLeadsBounceColumns();
  _loadSuppressionCache();

  let added = 0;
  let sUpdated = 0;
  const norms = Object.keys(leadIndex);
  for (let i = 0; i < norms.length; i++) {
    const norm = norms[i];
    if (!isBlockedDomainEmail(norm)) continue;
    if (isSuppressed(norm)) continue;
    addToSuppressionList(norm, "POLICY_REJECTION", "Globally suppressed domain (SUPPRESSED_DOMAINS)", "DOMAIN_BLOCK");
    added++;
    const li = leadIndex[norm];
    writeLeadBounceState(li.row, "POLICY_REJECTION", "Globally suppressed domain", false);
    const status = leadInfoSafeStatus(leadIndex, norm);
    if (status !== "INVALID" && status !== "DO_NOT_CONTACT" && status !== "REPLIED" && status !== "COMPLETED") {
      getLeadsSheet().getRange(li.row, LEADS_COL.STATUS + 1).setValue("INVALID");
      getLeadsSheet().getRange(li.row, LEADS_COL.LAST_UPDATED + 1).setValue(new Date());
    }
    sUpdated++;
  }

  try { setupDashboard(); } catch (e) { console.log("Dashboard rebuild skipped: " + e.message); }

  console.log("Domain-block live | added to SuppressionList: " + added + " | S/T updated: " + sUpdated);
  console.log("=== DOMAIN BLOCK COMPLETE ===");
  return { added: added, sUpdated: sUpdated };
}

function leadInfoSafeStatus(leadIndex, norm) {
  const v = leadIndex[norm];
  return v ? String(v.status || "").trim() : "";
}

// ------------------------------------------------------------
// DRY-RUN: HISTORICAL vs CURRENT LEADS (no Gmail, just sheets)
// ------------------------------------------------------------
// Compares all historical bounce sources (SuppressionList,
// Leads S, ActivityLog) against current Leads.
// No Gmail search, so faster for a quick sheet-only audit.

function dryRunHistoricalSuppression() {

  const leadsInfo = _buildCurrentLeadIndex();
  const leadData = leadsInfo.data;
  const leadIndex = leadsInfo.index;
  const suppressionIndex = _buildSuppressionIndexNormalized();

  const histSet = {}; // normalized -> {category, diagnostic}
  let histHard = 0;
  let histSoft = 0;
  let histPolicy = 0;
  let histUnknown = 0;

  // Leads S
  for (let i = 1; i < leadData.length; i++) {
    const norm = normalizeSuppressionEmail(leadData[i][LEADS_COL.EMAIL]);
    const cat = String(leadData[i].length > LEADS_COL.BOUNCE_CATEGORY ? leadData[i][LEADS_COL.BOUNCE_CATEGORY] : "").trim();
    if (!cat) continue;
    if (!histSet[norm]) {
      const diag = String(leadData[i].length > LEADS_COL.BOUNCE_DIAGNOSTIC ? leadData[i][LEADS_COL.BOUNCE_DIAGNOSTIC] : "").trim();
      histSet[norm] = { category: cat, diagnostic: diag };
      if (cat === "HARD_BOUNCE") histHard++;
      else if (cat === "SOFT_BOUNCE") histSoft++;
      else if (cat === "POLICY_REJECTION") histPolicy++;
      else histUnknown++;
    }
  }

  // SuppressionList (superset)
  Object.keys(suppressionIndex).forEach(function(nk) {
    if (!histSet[nk]) {
      histSet[nk] = { category: suppressionIndex[nk].category, diagnostic: "" };
      if (suppressionIndex[nk].category === "HARD_BOUNCE") histHard++;
    }
  });

  // ActivityLog EMAIL_BOUNCED hard
  try {
    const activity = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ActivityLog");
    if (activity && activity.getLastRow() >= 2) {
      const adata = activity.getDataRange().getValues();
      for (let i = 1; i < adata.length; i++) {
        const action = String(adata[i][3] || "").trim();
        const err = String(adata[i][10] || "");
        const emailRaw = String(adata[i][2] || "").trim();
        if (action !== "EMAIL_BOUNCED" || !emailRaw) continue;
        const norm = normalizeSuppressionEmail(emailRaw);
        if (histSet[norm]) continue;
        let cat = "UNKNOWN";
        if (err.indexOf("HARD_BOUNCE") !== -1) cat = "HARD_BOUNCE";
        else if (_isHardKeywordNormalized(err)) cat = "HARD_BOUNCE";
        histSet[norm] = { category: cat, diagnostic: err.slice(0, 250) };
        if (cat === "HARD_BOUNCE") histHard++;
      }
    }
  } catch (e) {}

  let wouldBeSkipped = 0;
  const hardList = [];
  Object.keys(histSet).forEach(function(norm) {
    if (histSet[norm].category === "HARD_BOUNCE" && leadIndex[norm]) {
      const st = leadIndex[norm].status;
      if (ACTIVE_CAMPAIGN_STATUSES.concat(["NEW"]).includes(st)) {
        wouldBeSkipped++;
      }
      hardList.push(norm);
    }
  });

  console.log(
    "Historical dry-run | total leads checked: " + (leadData.length - 1) +
    " | historical bounced addresses found: " + Object.keys(histSet).length +
    " | hard to suppress: " + histHard +
    " | soft: " + histSoft +
    " | policy: " + histPolicy +
    " | unknown: " + histUnknown +
    " | leads that would be skipped (active): " + wouldBeSkipped
  );

  if (hardList.length > 0) {
    console.log("Sample hard (max 10): " + hardList.slice(0, 10).join(", "));
  }

  return {
    totalLeadsChecked: leadData.length - 1,
    historicalBouncedFound: Object.keys(histSet).length,
    hardToSuppress: histHard,
    soft: histSoft,
    policy: histPolicy,
    unknown: histUnknown,
    leadsThatWouldBeSkipped: wouldBeSkipped
  };
}
