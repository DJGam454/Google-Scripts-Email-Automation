// ============================================================
// EMAIL VALIDATION & SENDING LIMITS
// ============================================================

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

  const sentToday =
    getEmailsSentToday();

  const remaining =
    dailyLimit - sentToday;


  console.log(
    "Daily limit: " + dailyLimit +
    " | Sent today: " + sentToday +
    " | Remaining: " + Math.max(remaining, 0)
  );


  return sentToday < dailyLimit;
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