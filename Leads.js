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
  LAST_UPDATED: 17       // R
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
    lastUpdated: row[LEADS_COL.LAST_UPDATED]
  };
}
