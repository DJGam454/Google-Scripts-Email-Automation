function getConfig() {

  const cache = CacheService.getScriptCache();
  const cached = cache.get("app_config");
  if (cached) {
    return JSON.parse(cached);
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

  cache.put("app_config", JSON.stringify(config), 300);

  return config;
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