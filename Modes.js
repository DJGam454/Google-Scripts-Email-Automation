// ============================================================
// AUTOMATION MODES & RECIPIENT RESOLUTION
// ============================================================

// ============================================================
// BOOLEAN PARSING
// ============================================================
// Sheets may return booleans as real booleans or as the strings
// "TRUE"/"FALSE". This is the single canonical parser for
// config and sheet flag cells. Never inline new boolean parsing.

function isFlagTrue(value) {

  return (
    value === true ||
    String(value)
      .trim()
      .toUpperCase() === "TRUE"
  );
}

function isFastTestMode(config) {

  return isFlagTrue(config.FAST_TEST_MODE);
}

function isPreviewMode(config) {

  return isFlagTrue(config.PREVIEW_MODE);
}

function isTestMode(config) {

  return isFlagTrue(config.TEST_MODE);
}

function getActualRecipient(lead, config) {

  if (isTestMode(config)) {

    const testEmail =
      String(config.TEST_EMAIL || "")
        .trim();


    if (!testEmail) {
      throw new Error(
        "TEST_MODE is enabled but TEST_EMAIL is empty."
      );
    }


    console.log(
      "TEST MODE | Original recipient: " +
      lead.email +
      " | Redirected to: " +
      testEmail
    );


    return testEmail;
  }


  return lead.email;
}
