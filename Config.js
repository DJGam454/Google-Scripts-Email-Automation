function getConfig() {

  const cache = CacheService.getScriptCache();
  const cached = cache.get("app_config");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Fall through to fresh read.
    }
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config");

  if (!sheet) {
    throw new Error("Config sheet not found.");
  }

  const data = sheet.getDataRange().getValues();

  const config = {};

  for (let i = 1; i < data.length; i++) {

    const key = data[i][0];
    const value = data[i][1];

    if (key) {
      config[key] = value;
    }
  }

  // Fail-safe defaults so a missing key never opens the floodgate.
  if (config.DAILY_LIMIT === undefined || config.DAILY_LIMIT === "") {
    config.DAILY_LIMIT = 0;
  }
  if (config.HOURLY_LIMIT === undefined || config.HOURLY_LIMIT === "") {
    // Global hourly cap: if not set, derive from DAILY_LIMIT to
    // keep bursts safe without admin action.
    const daily = Number(config.DAILY_LIMIT);
    config.HOURLY_LIMIT = (!isNaN(daily) && daily > 0) ? Math.ceil(daily / 4) : 15;
  }

  cache.put("app_config", JSON.stringify(config), 300);

  return config;
}

function clearConfigCache() {
  try {
    CacheService.getScriptCache().remove("app_config");
  } catch (e) {}
}

function testConfig() {

  const config = getConfig();

  console.log(config);

  console.log(
    "Automation enabled: " +
    config.AUTOMATION_ENABLED
  );

  console.log(
    "Follow-up 1: " +
    config.FOLLOWUP_1_MINUTES +
    " minutes"
  );
}

function getAIConfig() {

  const cache = CacheService.getScriptCache();
  const cached = cache.get("ai_config");
  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("AIConfig");

  if (!sheet) {
    throw new Error("AIConfig sheet not found.");
  }

  const data = sheet
    .getDataRange()
    .getValues();

  const config = {};

  for (let i = 1; i < data.length; i++) {

    const key = data[i][0];
    const value = data[i][1];

    if (key) {
      config[key] = value;
    }
  }

  cache.put("ai_config", JSON.stringify(config), 300);

  return config;
}

function testAIConfig() {

  const config = getAIConfig();

  console.log(
    "AI Enabled: " +
    config.AI_ENABLED
  );

  console.log(
    "Model: " +
    config.AI_MODEL
  );

  console.log(
    "Max Intro Words: " +
    config.MAX_INTRO_WORDS
  );

  console.log(
    "Temperature: " +
    config.TEMPERATURE
  );
}

// ============================================================
// HARDENING MIGRATION HELPER (one-click live sheet setup)
// ============================================================
// Creates the S-V headers, SuppressionList sheet, and the new
// throttling / bounce / unsubscribe config keys. Existing values
// are never overwritten - only missing keys are appended. Safe to
// run multiple times. Run once after deploying the hardening push.

function ensureHardeningMigration() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Leads S-V columns.
  try {
    ensureLeadsBounceColumns();
    console.log("Leads S-V columns verified.");
  } catch (e) {
    console.error("S-V migration failed: " + e.message);
  }

  // 2. SuppressionList sheet.
  try {
    _getSuppressionSheet(true);
    console.log("SuppressionList sheet verified.");
  } catch (e) {
    console.error("SuppressionList migration failed: " + e.message);
  }

  // 3. Backfill SuppressionList from existing INVALID / DO_NOT_CONTACT.
  try {
    const leadsSheet = ss.getSheetByName("Leads");
    if (leadsSheet && leadsSheet.getLastRow() >= 2) {
      const data = leadsSheet.getDataRange().getValues();
      let backfilled = 0;
      for (let i = 1; i < data.length; i++) {
        const email = String(data[i][LEADS_COL.EMAIL] || "").trim().toLowerCase();
        const status = String(data[i][LEADS_COL.STATUS] || "").trim();
        const cat = String(
          data[i].length > LEADS_COL.BOUNCE_CATEGORY
            ? data[i][LEADS_COL.BOUNCE_CATEGORY]
            : ""
        ).trim();

        if (!email || isSuppressed(email)) {
          continue;
        }

        if (status === "INVALID" && (cat === "HARD_BOUNCE" || !cat)) {
          const diag = String(
            data[i].length > LEADS_COL.BOUNCE_DIAGNOSTIC
              ? data[i][LEADS_COL.BOUNCE_DIAGNOSTIC]
              : ""
          ).trim() || "Backfilled from INVALID";
          addToSuppressionList(email, "HARD_BOUNCE", diag, "MIGRATION");
          backfilled++;
        } else if (status === "DO_NOT_CONTACT") {
          addToSuppressionList(email, "UNSUBSCRIBED", "Backfilled from DO_NOT_CONTACT", "MIGRATION");
          backfilled++;
        }
      }
      if (backfilled > 0) {
        console.log("Backfilled " + backfilled + " suppressed addresses from Leads (INVALID/DO_NOT_CONTACT).");
      }
    }
  } catch (e) {
    console.error("Backfill failed: " + e.message);
  }

  // 4. Config keys for throttling / send windows.
  try {
    refreshHardeningConfig();
  } catch (e) {
    console.error("Config hardening failed: " + e.message);
  }
  try {
    refreshSendWindowsConfig();
  } catch (e) {
    console.error("Send windows config failed: " + e.message);
  }

  // 5. Warm token secret.
  try {
    _getUnsubSecret();
    console.log("Unsubscribe token secret verified.");
  } catch (e) {
    console.error("Token secret warm failed: " + e.message);
  }

  clearConfigCache();
  _resetSuppressionCache();

  console.log("=== HARDENING MIGRATION COMPLETE ===");
}

function refreshSendWindowsConfig() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config");

  if (!sheet) {
    throw new Error("Config sheet not found.");
  }

  const updates = {
    "SEND_WINDOWS": ""
  };

  const data = sheet.getDataRange().getValues();
  let appended = 0;
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || "").trim();
    if (Object.prototype.hasOwnProperty.call(updates, key)) {
      delete updates[key];
    }
  }
  const pending = Object.keys(updates);
  if (pending.length > 0) {
    const rows = pending.map(function(k) { return [k, updates[k]]; });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
    appended = rows.length;
    console.log("Send windows appended " + appended + " keys: " + pending.join(", "));
  } else {
    console.log("Send windows keys already present.");
  }
  return appended;
}

function refreshHardeningConfig() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config");

  if (!sheet) {
    throw new Error("Config sheet not found.");
  }

  const updates = {
    "HOURLY_LIMIT": "30",
    "HOURLY_LIMIT_YAHOO": "8",
    "HOURLY_LIMIT_OUTLOOK": "8",
    "HOURLY_LIMIT_ICLOUD": "5",
    "MAX_SENDS_PER_RUN": "50",
    "NEW_LEADS_DAILY_FLOOR": "",
    "SUPPRESSED_DOMAINS": "",
    "SKIP_REPLIES_IN_MAIN_RUN": "FALSE"
  };

  const data = sheet.getDataRange().getValues();
  let appended = 0;

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || "").trim();
    if (Object.prototype.hasOwnProperty.call(updates, key)) {
      delete updates[key];
    }
  }

  const pending = Object.keys(updates);
  if (pending.length > 0) {
    const rows = pending.map(function(k) { return [k, updates[k]]; });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
    appended = rows.length;
    console.log("Hardening config appended " + appended + " keys: " + pending.join(", "));
  } else {
    console.log("Hardening config keys already present.");
  }

  return appended;
}