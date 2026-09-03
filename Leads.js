// ============================================================
// LEADS DATA LAYER
// ============================================================
// Single source of truth for the Leads sheet contract:
// column positions, statuses, follow-up step mapping and the
// shared lead-building helper.
//
// The Leads sheet column map is load-bearing. If you change a
// column here, every consumer (EmailEngine, AI, Bounces,
// Campaigns, Dashboard) must be audited, and the excel-sheet
// skill's contract documentation must be updated.

// ============================================================
// LEADS COLUMN MAP (1-based sheet columns, 0-based indexes)
// ============================================================

var LEADS_COL = {
  LEAD_ID: 0,            // A
  COMPANY: 1,            // B
  NAME: 2,               // C
  EMAIL: 3,              // D
  WEBSITE: 4,            // E
  INDUSTRY: 5,           // F
  INTRO: 6,              // G - Personalised Intro (AI)
  SERVICE: 7,            // H - Service Assigned
  CAMPAIGN: 8,           // I - Current Campaign
  EMAIL_1_SENT: 9,       // J
  FOLLOWUP_1_SENT: 10,   // K
  FOLLOWUP_2_SENT: 11,   // L
  FOLLOWUP_3_SENT: 12,   // M
  LAST_EMAIL_DATE: 13,   // N
  STATUS: 14,            // O
  THREAD_ID: 15,         // P - Gmail Thread ID
  NOTES: 16,             // Q
  LAST_UPDATED: 17,      // R
  BOUNCE_CATEGORY: 18,   // S - HARD_BOUNCE | SOFT_BOUNCE | POLICY_REJECTION | UNKNOWN
  BOUNCE_DIAGNOSTIC: 19, // T - Diagnostic snippet
  SUPPRESSED_AT: 20,     // U - Suppression timestamp
  RETRY_COUNT: 21        // V - Soft-bounce retry count
};

// ============================================================
// LEAD STATUSES (col O)
// ============================================================

var LEAD_STATUS = {
  NEW: "NEW",
  EMAIL_1_SENT: "EMAIL_1_SENT",
  FOLLOWUP_1_SENT: "FOLLOWUP_1_SENT",
  FOLLOWUP_2_SENT: "FOLLOWUP_2_SENT",
  FOLLOWUP_3_SENT: "FOLLOWUP_3_SENT",
  REPLIED: "REPLIED",
  INVALID: "INVALID",
  DO_NOT_CONTACT: "DO_NOT_CONTACT",
  COMPLETED: "COMPLETED"
};

// Statuses that halt all further automation.
var STOP_STATUSES = [
  LEAD_STATUS.REPLIED,
  LEAD_STATUS.INVALID,
  LEAD_STATUS.DO_NOT_CONTACT
];

// Statuses in which a reply can still arrive.
var ACTIVE_CAMPAIGN_STATUSES = [
  LEAD_STATUS.EMAIL_1_SENT,
  LEAD_STATUS.FOLLOWUP_1_SENT,
  LEAD_STATUS.FOLLOWUP_2_SENT,
  LEAD_STATUS.FOLLOWUP_3_SENT
];

// ============================================================
// FOLLOW-UP STEP MAP
// ============================================================
// Collapses the three identical follow-up branches into data:
//   status            -> which template step to send,
//                        the next status after sending,
//                        the timestamp column (K/L/M) to write,
//                        and the follow-up number for timing.
// ============================================================

var FOLLOWUP_STEPS = {
  EMAIL_1_SENT: {
    step: "FOLLOWUP_1",
    nextStatus: "FOLLOWUP_1_SENT",
    column: LEADS_COL.FOLLOWUP_1_SENT + 1, // K (1-based)
    number: 1
  },
  FOLLOWUP_1_SENT: {
    step: "FOLLOWUP_2",
    nextStatus: "FOLLOWUP_2_SENT",
    column: LEADS_COL.FOLLOWUP_2_SENT + 1, // L (1-based)
    number: 2
  },
  FOLLOWUP_2_SENT: {
    step: "FOLLOWUP_3",
    nextStatus: "FOLLOWUP_3_SENT",
    column: LEADS_COL.FOLLOWUP_3_SENT + 1, // M (1-based)
    number: 3
  }
};

// ============================================================
// SHEET ACCESS (per-execution in-memory cache)
// ============================================================
// getLeadsSheet() / getLeadsData() read the sheet at most once
// per execution instead of once per lead row.

var _leadsSheetCache = null;

function getLeadsSheet() {

  if (_leadsSheetCache) {
    return _leadsSheetCache;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Leads");

  if (!sheet) {
    throw new Error("Leads sheet not found.");
  }

  _leadsSheetCache = sheet;

  return sheet;
}

// Full 2D array (header + rows) of the Leads sheet.
function getLeadsData() {
  return getLeadsSheet().getDataRange().getValues();
}

// ============================================================
// LEAD BUILDER
// ============================================================
// The one place a spreadsheet row becomes a lead object.
// All pipeline files must use this instead of parsing rows
// inline (previously duplicated in five files).

function buildLeadFromRow(data, index) {

  const row = data[index];

  return {
    leadId: row[LEADS_COL.LEAD_ID],
    company: row[LEADS_COL.COMPANY],
    name: row[LEADS_COL.NAME],
    email: row[LEADS_COL.EMAIL],
    website: row[LEADS_COL.WEBSITE],
    industry: row[LEADS_COL.INDUSTRY],
    personalisedIntro: row[LEADS_COL.INTRO],
    service: row[LEADS_COL.SERVICE],
    campaign: row[LEADS_COL.CAMPAIGN],

    email1Sent: row[LEADS_COL.EMAIL_1_SENT],
    followup1Sent: row[LEADS_COL.FOLLOWUP_1_SENT],
    followup2Sent: row[LEADS_COL.FOLLOWUP_2_SENT],
    followup3Sent: row[LEADS_COL.FOLLOWUP_3_SENT],

    lastEmailDate: row[LEADS_COL.LAST_EMAIL_DATE],
    status: row[LEADS_COL.STATUS],
    threadId: row[LEADS_COL.THREAD_ID],
    notes: row[LEADS_COL.NOTES],
    lastUpdated: row[LEADS_COL.LAST_UPDATED],
    bounceCategory: row.length > LEADS_COL.BOUNCE_CATEGORY ? row[LEADS_COL.BOUNCE_CATEGORY] : "",
    bounceDiagnostic: row.length > LEADS_COL.BOUNCE_DIAGNOSTIC ? row[LEADS_COL.BOUNCE_DIAGNOSTIC] : "",
    suppressedAt: row.length > LEADS_COL.SUPPRESSED_AT ? row[LEADS_COL.SUPPRESSED_AT] : "",
    retryCount: row.length > LEADS_COL.RETRY_COUNT ? row[LEADS_COL.RETRY_COUNT] : ""
  };
}

// ============================================================
// SUPPRESSION LAYER (S-V + SuppressionList sheet)
// ============================================================
// Hard-bounced / unsubscribed addresses live permanently in the
// SuppressionList sheet and in Leads S/U. Every send path checks
// this before touching Gmail, so a deleted-and-reimported row
// with the same email still stays blocked.

var SUPPRESSION_SHEET_NAME = "SuppressionList";

var _suppressionCache = null;

function _getSuppressionSheet(createIfMissing) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUPPRESSION_SHEET_NAME);

  if (sheet) {
    return sheet;
  }

  if (!createIfMissing) {
    return null;
  }

  const created = ss.insertSheet(SUPPRESSION_SHEET_NAME);

  created.appendRow([
    "Email",
    "Category",
    "Diagnostic",
    "First Seen",
    "Last Seen",
    "Count",
    "Source"
  ]);

  created.setFrozenRows(1);
  created.autoResizeColumns(1, 7);

  // Data validation for Category (B) on the new sheet.
  try {
    const catRange = created.getRange("B2:B");
    const catRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["HARD_BOUNCE", "SOFT_BOUNCE", "POLICY_REJECTION", "UNKNOWN", "UNSUBSCRIBED"],
        true
      )
      .setAllowInvalid(true)
      .build();
    catRange.setDataValidation(catRule);
  } catch (e) {
    console.log("SuppressionList validation skipped: " + e.message);
  }

  return created;
}

function _loadSuppressionCache() {

  if (_suppressionCache !== null) {
    return _suppressionCache;
  }

  _suppressionCache = {};

  const sheet = _getSuppressionSheet(false);

  if (!sheet || sheet.getLastRow() < 2) {
    return _suppressionCache;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][0] || "").trim().toLowerCase();
    if (email) {
      _suppressionCache[email] = {
        category: String(data[i][1] || "").trim(),
        row: i + 1
      };
    }
  }

  return _suppressionCache;
}

function _resetSuppressionCache() {
  _suppressionCache = null;
}

function isSuppressed(email) {

  const key = String(email || "").trim().toLowerCase();
  if (!key) {
    return false;
  }

  const cache = _loadSuppressionCache();
  return Object.prototype.hasOwnProperty.call(cache, key);
}

function getSuppressionCategory(email) {

  const key = String(email || "").trim().toLowerCase();
  const cache = _loadSuppressionCache();
  const hit = cache[key];
  return hit ? hit.category : "";
}

function addToSuppressionList(email, category, diagnostic, source) {

  const key = String(email || "").trim().toLowerCase();
  if (!key) {
    return;
  }

  const sheet = _getSuppressionSheet(true);
  const now = new Date();
  const cat = String(category || "UNKNOWN").trim() || "UNKNOWN";
  const diag = String(diagnostic || "").slice(0, 250);
  const src = String(source || "DSN").trim() || "DSN";

  // Update in place if already present (idempotent).
  const cache = _loadSuppressionCache();

  if (Object.prototype.hasOwnProperty.call(cache, key)) {
    const row = cache[key].row;
    sheet.getRange(row, 2).setValue(cat);
    sheet.getRange(row, 3).setValue(diag);
    sheet.getRange(row, 5).setValue(now);
    const countCell = sheet.getRange(row, 6);
    const current = Number(countCell.getValue());
    countCell.setValue((isNaN(current) ? 0 : current) + 1);
    _suppressionCache[key].category = cat;
    return;
  }

  sheet.appendRow([
    key,
    cat,
    diag,
    now,
    now,
    1,
    src
  ]);

  _suppressionCache[key] = {
    category: cat,
    row: sheet.getLastRow()
  };
}

function ensureLeadsBounceColumns() {

  const sheet = getLeadsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = [
    "Bounce Category",
    "Bounce Diagnostic",
    "Suppressed At",
    "Retry Count"
  ];

  let lastCol = sheet.getLastColumn();

  for (let i = 0; i < needed.length; i++) {
    const targetIndex = LEADS_COL.BOUNCE_CATEGORY + i;
    const headerPos = targetIndex + 1;

    // Already has header at that position.
    if (headers[targetIndex] === needed[i]) {
      continue;
    }

    // Sheet not wide enough yet - append.
    if (headerPos > lastCol) {
      sheet.getRange(1, headerPos).setValue(needed[i]);
      lastCol = headerPos;
      continue;
    }

    // Header mismatch but column exists - fix label if blank.
    if (!headers[targetIndex]) {
      sheet.getRange(1, headerPos).setValue(needed[i]);
    }
  }

  // Light formatting for the new columns.
  if (sheet.getLastColumn() >= LEADS_COL.BOUNCE_CATEGORY + 1) {
    sheet.getRange(1, LEADS_COL.BOUNCE_CATEGORY + 1, 1, 4)
      .setFontWeight("bold")
      .setBackground("#F8FAFC");
  }

  // Data validation for S: Bounce Category dropdown. Prevents typo'd
  // categories from silently missing HARD_BOUNCE suppression.
  try {
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const sRange = sheet.getRange(2, LEADS_COL.BOUNCE_CATEGORY + 1, lastRow - 1, 1);
    const sRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["HARD_BOUNCE", "SOFT_BOUNCE", "POLICY_REJECTION", "UNKNOWN", "UNSUBSCRIBED"],
        true
      )
      .setAllowInvalid(true)
      .setHelpText("Select: HARD_BOUNCE (permanent), SOFT_BOUNCE (retry), POLICY_REJECTION (no retry), UNKNOWN (review), UNSUBSCRIBED")
      .build();
    sRange.setDataValidation(sRule);
  } catch (e) {
    // Validation is best-effort - never block header creation.
    console.log("S validation skipped: " + e.message);
  }

  return sheet.getLastColumn();
}

function writeLeadBounceState(row, category, diagnostic, incrementRetry) {

  const now = new Date();
  const cat = String(category || "").trim();
  const diag = String(diagnostic || "").slice(0, 250);

  ensureLeadsBounceColumns();

  if (cat) {
    getLeadsSheet().getRange(row, LEADS_COL.BOUNCE_CATEGORY + 1).setValue(cat);
  }

  if (diag) {
    getLeadsSheet().getRange(row, LEADS_COL.BOUNCE_DIAGNOSTIC + 1).setValue(diag);
  }

  if (cat === "HARD_BOUNCE" || cat === "UNSUBSCRIBED") {
    getLeadsSheet().getRange(row, LEADS_COL.SUPPRESSED_AT + 1).setValue(now);
  }

  if (incrementRetry) {
    const cell = getLeadsSheet().getRange(row, LEADS_COL.RETRY_COUNT + 1);
    const current = Number(cell.getValue());
    const next = (isNaN(current) ? 0 : current) + 1;
    cell.setValue(next);
  }
}

function testSuppression() {

  console.log("Suppression cache size: " + Object.keys(_loadSuppressionCache()).length);
  console.log("isSuppressed('test@example.com'): " + isSuppressed("test@example.com"));
  console.log("Leads columns: " + getLeadsSheet().getLastColumn());
}
