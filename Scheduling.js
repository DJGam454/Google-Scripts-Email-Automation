// ============================================================
// AUTOMATION SCHEDULING
// ============================================================

function isWorkingDay(config) {

  const timezone =
    String(
      config.TIMEZONE || "Europe/London"
    ).trim();


  const configuredDays =
    String(
      config.WORKING_DAYS ||
      "MON,TUE,WED,THU,FRI"
    )
      .toUpperCase()
      .split(",")
      .map(function(day) {
        return day.trim();
      });


  const now = new Date();


  // EEE produces values such as:
  // MON, TUE, WED...
  const today =
    Utilities.formatDate(
      now,
      timezone,
      "EEE"
    ).toUpperCase();


  const allowed =
    configuredDays.includes(today);


  console.log(
    "Working day check | " +
    "Timezone: " + timezone +
    " | Today: " + today +
    " | Allowed: " + allowed
  );


  return allowed;
}

function _parseTimeToMinutes(value) {

  const raw = String(value || "").trim();

  if (!raw) {
    return null;
  }

  // Support "9:30" or "09:30" or "9.5" (9.5 -> 9:30)
  if (raw.indexOf(":") !== -1) {
    const parts = raw.split(":");
    const h = Number(parts[0]);
    const m = Number(parts[1]);
    if (isNaN(h) || isNaN(m) || h < 0 || h > 23 || m < 0 || m > 59) {
      return null;
    }
    return h * 60 + m;
  }

  const f = Number(raw);
  if (isNaN(f) || f < 0 || f >= 24) {
    return null;
  }
  return Math.round(f * 60);
}

function isWithinSendingWindow(config) {

  const timezone =
    String(
      config.TIMEZONE || "Europe/London"
    ).trim();

  const sendWindowsRaw =
    String(config.SEND_WINDOWS || "").trim();

  const now = new Date();

  const nowMinutes =
    Number(Utilities.formatDate(now, timezone, "H")) * 60 +
    Number(Utilities.formatDate(now, timezone, "m"));

  // Custom windows like "9:30-12:25,14:00-17:00" or "9-12,14-17" or "9.5-12.5"
  // takes precedence over legacy SEND_START_HOUR/END_HOUR.
  if (sendWindowsRaw) {
    const parts = sendWindowsRaw.split(",");
    let anyValid = false;
    let allowed = false;

    for (let i = 0; i < parts.length; i++) {
      const part = String(parts[i] || "").trim();
      if (!part) {
        continue;
      }
      const dash = part.indexOf("-");
      if (dash === -1) {
        // Single point like "9:30" - allow only that minute
        const t = _parseTimeToMinutes(part);
        if (t === null) {
          console.log("Invalid SEND_WINDOWS entry: " + part);
          continue;
        }
        anyValid = true;
        if (nowMinutes === t) {
          allowed = true;
        }
        continue;
      }
      const startStr = String(part.slice(0, dash) || "").trim();
      const endStr = String(part.slice(dash + 1) || "").trim();
      const startMin = _parseTimeToMinutes(startStr);
      const endMin = _parseTimeToMinutes(endStr);
      if (startMin === null || endMin === null) {
        console.log("Invalid SEND_WINDOWS range: " + part);
        continue;
      }
      anyValid = true;
      // Half-open [start, end)
      if (nowMinutes >= startMin && nowMinutes < endMin) {
        allowed = true;
      }
    }

    if (!anyValid) {
      throw new Error("Invalid SEND_WINDOWS: " + sendWindowsRaw);
    }

    console.log(
      "Sending window check (custom) | Now: " +
      Math.floor(nowMinutes / 60) + ":" + String(nowMinutes % 60).padStart(2, "0") +
      " | Windows: " + sendWindowsRaw + " | Allowed: " + allowed
    );

    return allowed;
  }

  // Legacy fallback: SEND_START_HOUR / SEND_END_HOUR (hour precision)
  const startHour =
    Number(config.SEND_START_HOUR);

  const endHour =
    Number(config.SEND_END_HOUR);

  if (
    isNaN(startHour) ||
    isNaN(endHour)
  ) {
    throw new Error(
      "Invalid SEND_START_HOUR or SEND_END_HOUR."
    );
  }

  const currentHour =
    Number(
      Utilities.formatDate(
        now,
        timezone,
        "H"
      )
    );

  const allowed =
    currentHour >= startHour &&
    currentHour < endHour;

  console.log(
    "Sending window check (legacy) | " +
    "Current hour: " +
    currentHour +
    " | Window: " +
    startHour +
    ":00-" +
    endHour +
    ":00" +
    " | Allowed: " +
    allowed
  );

  return allowed;
}

function canAutomationRunNow() {

  // Delegates to the same schedule gate used by runAutomation()
  // (working-day + sending-window, with the FAST_TEST_MODE bypass).
  return isAutomationScheduleAllowed(
    getConfig()
  );
}

function isAutomationScheduleAllowed(config) {

  // ==============================
  // TEST MODE BYPASS
  // ==============================

  if (isFastTestMode(config)) {

    console.log(
      "FAST TEST MODE enabled."
    );

    console.log(
      "Working-day and sending-window checks bypassed."
    );

    return true;
  }


  // ==============================
  // PRODUCTION CHECKS
  // ==============================

  if (!isWorkingDay(config)) {

    console.log(
      "Automation blocked: Non-working day."
    );

    return false;
  }


  if (!isWithinSendingWindow(config)) {

    console.log(
      "Automation blocked: Outside sending window."
    );

    return false;
  }


  console.log(
    "Production schedule checks passed."
  );


  return true;
}