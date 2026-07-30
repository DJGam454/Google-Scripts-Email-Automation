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

function isWithinSendingWindow(config) {

  const timezone =
    String(
      config.TIMEZONE || "Europe/London"
    ).trim();


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


  const now =
    new Date();


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
    "Sending window check | " +
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

  const config =
    getConfig();


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
    "Automation schedule checks passed."
  );


  return true;
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