// ============================================================
// EMAIL TEMPLATES
// ============================================================

// Existing getTemplate() function — unchanged


// ============================================================
// TEMPLATE PERSONALISATION
// ============================================================

// Existing personaliseTemplate() function — unchanged

function getTemplate(service, step) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Templates");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const templateService = data[i][0];
    const templateStep = data[i][1];

    if (
      templateService === service &&
      templateStep === step
    ) {
      return {
        subject: data[i][2],
        body: data[i][3]
      };
    }
  }

  return null;
}

function personaliseTemplate(text, lead) {

  if (!text) {
    return "";
  }

  const config = getConfig();

  const senderName =
    config.SENDER_NAME || "";

  const replyToEmail =
    config.REPLY_TO_EMAIL || "";

  let emailSignature =
    config.EMAIL_SIGNATURE || "";

  // Allow the signature itself to contain {{SenderName}}
  // and {{ReplyToEmail}}.
  emailSignature = String(emailSignature)
    .replaceAll("{{SenderName}}", senderName)
    .replaceAll("{{ReplyToEmail}}", replyToEmail);

  return String(text)
    .replaceAll("{{FirstName}}", lead.name || "")
    .replaceAll("{{Company}}", lead.company || "")
    .replaceAll("{{Website}}", lead.website || "")
    .replaceAll("{{Industry}}", lead.industry || "")
    .replaceAll(
      "{{PersonalisedIntro}}",
      lead.personalisedIntro || ""
    )
    .replaceAll(
      "{{SenderName}}",
      senderName
    )
    .replaceAll(
      "{{ReplyToEmail}}",
      replyToEmail
    )
    .replaceAll(
      "{{EmailSignature}}",
      emailSignature
    );
}