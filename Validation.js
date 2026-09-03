// ============================================================
// EMAIL VALIDATION & SENDING LIMITS
// ============================================================

// ============================================================
// EMAIL VALIDATION & DAILY LIMIT
// ============================================================
// The daily count is read from the ActivityLog once per run and
// kept in a run-scoped cache, so per-send checks are O(1)
// instead of re-reading the whole log before every email.

var _dailyCountCache = null;
var _hourlyCountCache = null;
var _hourlyDomainCache = null;

function _resetDailyCountCache() {
  _dailyCountCache = null;
  _hourlyCountCache = null;
  _hourlyDomainCache = null;
}

function _incrementDailyCountCache() {
  if (_dailyCountCache !== null) {
    _dailyCountCache++;
  }
  if (_hourlyCountCache !== null) {
    _hourlyCountCache++;
  }
}

// Force hourly domain counters to refresh on next check (e.g.
// after a quota block so the next run sees fresh counts).
function _resetHourlyDomainCache() {
  _hourlyDomainCache = null;
}

function getEmailsSentToday() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("ActivityLog");

  if (!sheet) {
    throw new Error("ActivityLog sheet not found.");
  }

  const data = sheet.getDataRange().getValues();

  const today = new Date();

  let count = 0;

  for (let i = 1; i < data.length; i++) {

    const timestamp = data[i][0];
    const action = data[i][3];
    const result = data[i][9];

    if (!timestamp) {
      continue;
    }

    const logDate = new Date(timestamp);

    const isToday =
      logDate.getFullYear() === today.getFullYear() &&
      logDate.getMonth() === today.getMonth() &&
      logDate.getDate() === today.getDate();


    const isEmailAction = [
      "EMAIL_1",
      "FOLLOWUP_1",
      "FOLLOWUP_2",
      "FOLLOWUP_3"
    ].includes(action);


    if (
      isToday &&
      isEmailAction &&
      result === "SUCCESS"
    ) {
      count++;
    }
  }

  return count;
}

function canSendEmail() {

  const config = getConfig();

  const dailyLimit =
    Number(config.DAILY_LIMIT);

  let sentToday;

  if (_dailyCountCache === null) {

    sentToday = getEmailsSentToday();

    _dailyCountCache = sentToday;

  } else {

    sentToday = _dailyCountCache;
  }

  const remaining =
    dailyLimit - sentToday;


  console.log(
    "Daily limit: " + dailyLimit +
    " | Sent today: " + sentToday +
    " | Remaining: " + Math.max(remaining, 0)
  );


  return sentToday < dailyLimit;
}

// ============================================================
// HOURLY + PER-DOMAIN LIMITS
// ============================================================
// Global hourly cap (HOURLY_LIMIT) protects against bursts that
// would trigger Gmail quota ("Service invoked too many times").
// Per-domain caps protect Yahoo/Outlook/iCloud which defer at
// ~8-10/hour for a cold domain. Gmail stays uncapped at this
// layer beyond the global hourly cap.

var DOMAIN_GROUP = {
  YAHOO: ["yahoo.com", "yahoo.co.in", "yahoo.in", "yahoo.co.uk", "ymail.com", "rocketmail.com", "aol.com"],
  OUTLOOK: ["outlook.com", "hotmail.com", "hotmail.co.uk", "live.com", "live.co.uk", "msn.com"],
  ICLOUD: ["icloud.com", "me.com", "mac.com"]
};

function _getDomainGroup(domain) {

  const d = String(domain || "").trim().toLowerCase();

  for (let group in DOMAIN_GROUP) {
    if (DOMAIN_GROUP[group].indexOf(d) !== -1) {
      return group;
    }
  }

  return null;
}

function _getHourlyLimitForGroup(group, config) {

  // Per-group overrides fall back to global HOURLY_LIMIT.
  const globalLimit = Number(config.HOURLY_LIMIT);
  const hasGlobal = !isNaN(globalLimit) && globalLimit > 0;

  if (!group) {
    return hasGlobal ? globalLimit : null;
  }

  const key = "HOURLY_LIMIT_" + group;
  const perGroup = Number(config[key]);

  if (!isNaN(perGroup) && perGroup > 0) {
    return perGroup;
  }

  // No per-group cap configured - fall back to global.
  return hasGlobal ? globalLimit : null;
}

function _countSendsSince(sinceMs) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("ActivityLog");

  if (!sheet || sheet.getLastRow() < 2) {
    return { total: 0, byDomain: {} };
  }

  const data = sheet.getDataRange().getValues();
  const since = Number(sinceMs);
  let total = 0;
  const byDomain = {};

  for (let i = 1; i < data.length; i++) {
    const ts = data[i][0];
    const action = String(data[i][3] || "").trim();
    const result = String(data[i][9] || "").trim();

    if (result !== "SUCCESS") {
      continue;
    }

    if (["EMAIL_1", "FOLLOWUP_1", "FOLLOWUP_2", "FOLLOWUP_3"].indexOf(action) === -1) {
      continue;
    }

    if (!ts) {
      continue;
    }

    const t = new Date(ts).getTime();
    if (isNaN(t) || t < since) {
      continue;
    }

    total++;

    const email = String(data[i][2] || "").trim().toLowerCase();
    const domain = email.indexOf("@") !== -1 ? email.split("@").pop() : "";
    const group = _getDomainGroup(domain) || domain || "_unknown";
    byDomain[group] = (byDomain[group] || 0) + 1;
  }

  return { total: total, byDomain: byDomain };
}

function canSendEmailHourly() {

  const config = getConfig();
  const globalLimit = Number(config.HOURLY_LIMIT);

  // No hourly limit configured - open.
  if (isNaN(globalLimit) || globalLimit <= 0) {
    return true;
  }

  const now = Date.now();
  const hourAgo = now - 60 * 60 * 1000;

  let counts;

  if (_hourlyCountCache !== null && _hourlyDomainCache !== null) {
    counts = { total: _hourlyCountCache, byDomain: _hourlyDomainCache };
  } else {
    counts = _countSendsSince(hourAgo);
    _hourlyCountCache = counts.total;
    _hourlyDomainCache = counts.byDomain;
  }

  if (counts.total >= globalLimit) {
    console.log(
      "Hourly limit reached | Global: " + globalLimit +
      " | Sent last hour: " + counts.total
    );
    return false;
  }

  return true;
}

function canSendToDomain(email) {

  const config = getConfig();
  const rawDomain = String(email || "").split("@").pop() || "";
  const domain = String(rawDomain).trim().toLowerCase();
  const group = _getDomainGroup(domain);

  // No domain-specific cap for this address.
  const limit = _getHourlyLimitForGroup(group, config);
  if (limit === null) {
    return true;
  }

  const now = Date.now();
  const hourAgo = now - 60 * 60 * 1000;

  let counts;

  if (_hourlyDomainCache !== null) {
    counts = { total: _hourlyCountCache || 0, byDomain: _hourlyDomainCache };
  } else {
    counts = _countSendsSince(hourAgo);
    _hourlyCountCache = counts.total;
    _hourlyDomainCache = counts.byDomain;
  }

  // For grouped domains (YAHOO/OUTLOOK) count the group key.
  const key = group || domain;
  const sentForKey = counts.byDomain[key] || 0;

  if (sentForKey >= limit) {
    console.log(
      "Per-domain hourly limit reached | Domain/group: " + key +
      " | Limit: " + limit + " | Sent last hour: " + sentForKey
    );
    return false;
  }

  return true;
}

// ============================================================
// GMAIL QUOTA CIRCUIT BREAKER
// ============================================================
// Gmail throws "Service invoked too many times for one day: email"
// when the Workspace Gmail quota is hit. Retrying in a tight loop
// makes it worse. This breaker stops the run and blocks the rest
// of the day.

var _quotaBlocked = false;
var QUOTA_BLOCK_PROP = "QUOTA_BLOCK_UNTIL";

function isQuotaBlocked() {

  if (_quotaBlocked) {
    return true;
  }

  try {
    const raw = PropertiesService.getScriptProperties().getProperty(QUOTA_BLOCK_PROP);
    if (!raw) {
      return false;
    }
    const until = Number(raw);
    if (isNaN(until)) {
      return false;
    }
    if (Date.now() < until) {
      _quotaBlocked = true;
      return true;
    }
    // Expired - clear.
    PropertiesService.getScriptProperties().deleteProperty(QUOTA_BLOCK_PROP);
    return false;
  } catch (e) {
    return false;
  }
}

function tripQuotaBreaker() {

  _quotaBlocked = true;

  try {
    const tomorrow = new Date();
    tomorrow.setHours(24, 0, 0, 0);
    PropertiesService.getScriptProperties().setProperty(
      QUOTA_BLOCK_PROP,
      String(tomorrow.getTime())
    );
  } catch (e) {
    // Properties write is best-effort.
  }

  console.error(
    "QUOTA BREAKER TRIPPED: Gmail daily quota exceeded. " +
    "All sends blocked until midnight. " +
    "Check DAILY_LIMIT and MAX_SENDS_PER_RUN."
  );
}

function _isQuotaError(message) {

  const m = String(message || "").toLowerCase();
  return (
    m.indexOf("service invoked too many times") !== -1 ||
    m.indexOf("daily sending quota exceeded") !== -1 ||
    m.indexOf("user-rate limit exceeded") !== -1 ||
    m.indexOf("quota exceeded") !== -1 ||
    m.indexOf("rate limit exceeded") !== -1
  );
}

function clearQuotaBlock() {

  _quotaBlocked = false;
  try {
    PropertiesService.getScriptProperties().deleteProperty(QUOTA_BLOCK_PROP);
  } catch (e) {}
  console.log("Quota block cleared.");
}

function isValidEmail(email) {

  if (!email) {
    return false;
  }

  email = String(email).trim();

  // Simple practical validation for MVP
  const emailPattern =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return emailPattern.test(email);
}
