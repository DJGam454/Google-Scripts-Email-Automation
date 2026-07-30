function isFastTestMode(config) {

  return (
    config.FAST_TEST_MODE === true ||
    String(config.FAST_TEST_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
}

function isPreviewMode(config) {

  return (
    config.PREVIEW_MODE === true ||
    String(config.PREVIEW_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
}

function isTestMode(config) {

  return (
    config.TEST_MODE === true ||
    String(config.TEST_MODE)
      .trim()
      .toUpperCase() === "TRUE"
  );
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