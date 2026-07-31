// ============================================================
// AI RATE LIMITER
// ============================================================

var _aiLastCallTime = 0;

function _waitForAIRateLimit() {
  var config = getAIConfig();
  var rpm = Number(config.AI_RPM_LIMIT) || 10;
  var minInterval = Math.ceil(60000 / rpm);
  var jitter = Math.floor(Math.random() * 2000);

  var now = Date.now();
  var elapsed = now - _aiLastCallTime;

  if (elapsed < minInterval && _aiLastCallTime > 0) {
    var wait = minInterval - elapsed + jitter;
    console.log(
      "Rate limiter: waiting " +
      Math.round(wait / 1000) +
      "s before next AI call (RPM limit: " +
      rpm +
      ")"
    );
    Utilities.sleep(wait);
  }

  _aiLastCallTime = Date.now();
}

function _normaliseWebsiteUrl(url) {
  if (!url) return "";
  url = String(url).trim();
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) {
    url = "https://" + url;
  }
  return url;
}

function testGeminiKey() {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    console.log("ERROR: GEMINI_API_KEY not found.");
    return;
  }

  console.log("Gemini API key found successfully.");
}

function callGemini(prompt) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");


  if (!apiKey) {
    throw new Error(
      "GEMINI_API_KEY not found in Script Properties."
    );
  }


  const config = getAIConfig();


  // Check AI toggle
  const aiEnabled =
    isFlagTrue(config.AI_ENABLED);


  if (!aiEnabled) {
    throw new Error(
      "AI generation is currently disabled."
    );
  }


  const model =
    String(config.AI_MODEL).trim();


  if (!model) {
    throw new Error(
      "AI_MODEL is missing from AIConfig."
    );
  }


  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent";


  const payload = {

    contents: [
      {
        role: "user",

        parts: [
          {
            text: prompt
          }
        ]
      }
    ],

    generationConfig: {

      temperature:
        Number(config.TEMPERATURE) || 0.4

    }
  };


  const options = {

    method: "post",

    contentType: "application/json",

    headers: {
      "x-goog-api-key": apiKey
    },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true

  };


  // =================================
  // RETRY LOOP
  // =================================

  const maxRetries = 3;
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {

    try {

      _waitForAIRateLimit();

      const response =
        UrlFetchApp.fetch(
          url,
          options
        );

      const statusCode =
        response.getResponseCode();

      const responseText =
        response.getContentText();

      if (
        statusCode < 200 ||
        statusCode >= 300
      ) {

        // Retry on rate limits (429) and server errors (5xx)
        if (
          (statusCode === 429 || statusCode >= 500) &&
          attempt < maxRetries
        ) {

          const backoff =
            Math.pow(2, attempt) * 1000 +
            Math.floor(Math.random() * 2000);

          console.log(
            "Gemini " + statusCode +
            " on attempt " + (attempt + 1) +
            "/" + maxRetries +
            ". Retrying in " +
            Math.round(backoff / 1000) + "s..."
          );

          Utilities.sleep(backoff);
          continue;
        }

        throw new Error(
          "Gemini API error " +
          statusCode +
          ": " +
          responseText
        );
      }


      const result =
        JSON.parse(responseText);


      if (
        !result.candidates ||
        result.candidates.length === 0
      ) {

        throw new Error(
          "Gemini returned no candidates."
        );
      }


      const candidate =
        result.candidates[0];


      if (
        !candidate.content ||
        !candidate.content.parts ||
        candidate.content.parts.length === 0
      ) {

        throw new Error(
          "Gemini returned no text content."
        );
      }


      const text =
        candidate.content.parts
          .map(function(part) {
            return part.text || "";
          })
          .join("")
          .trim();


      if (!text) {

        throw new Error(
          "Gemini returned an empty response."
        );
      }


      return text;

    } catch (error) {

      lastError = error;
      if (attempt >= maxRetries) throw error;

      // Don't retry JSON parse errors
      if (error instanceof SyntaxError) throw error;

      // Don't retry client errors (4xx except 429 which is already handled)
      if (
        error.message &&
        error.message.indexOf("Gemini API error 4") === 0 &&
        error.message.indexOf("Gemini API error 429") !== 0
      ) {
        throw error;
      }

      const backoff =
        Math.pow(2, attempt) * 1000 +
        Math.floor(Math.random() * 2000);

      console.log(
        "Gemini call failed on attempt " +
        (attempt + 1) + ": " +
        error.message +
        ". Retrying in " +
        Math.round(backoff / 1000) + "s..."
      );

      Utilities.sleep(backoff);
    }
  }

  throw lastError || new Error(
    "Gemini API max retries exceeded."
  );
}
function testGeminiConnection() {

  try {

    console.log(
      "Testing Gemini connection..."
    );


    const response =
      callGemini(
        "Return exactly this text and nothing else: " +
        "Gemini connection successful"
      );


    console.log(
      "Gemini response: " +
      response
    );


  } catch (error) {

    console.error(
      "Gemini test failed: " +
      error.message
    );

  }
}

function generatePersonalisation(lead, research = null) {

  const config = getAIConfig();

  const maxWords =
    Number(config.MAX_INTRO_WORDS) || 35;

    const researchContext =
  research
    ? `
VERIFIED WEBSITE RESEARCH:

Company Summary:
${research.companySummary}

Services / Products Found:
${research.servicesFound.join(", ") || "Unknown"}

Target Audience:
${research.targetAudience}

Value Proposition:
${research.valueProposition}

Useful Facts:
${research.usefulFacts.join("\n") || "None"}

Research Confidence:
${research.researchConfidence}

The information above was obtained from the company's
supplied public website.

You may use these facts for personalisation, but do not
invent anything beyond them.
`
    : `
NO VERIFIED WEBSITE RESEARCH IS AVAILABLE.

Use only the basic lead information.
Do not claim to have inspected the website.
`;


  const prompt = `
You are writing a highly concise opening line for a B2B cold email.

Your job is NOT to sell the service.
Your job is to create a natural transition into a later sales message.

SERVICE BEING OFFERED:
${lead.service}

LEAD INFORMATION:
Company: ${lead.company || "Unknown"}
Website: ${lead.website || "Not provided"}
Industry: ${lead.industry || "Not provided"}

${researchContext}

TASK:
Write one short personalised opening sentence relevant to the company
and the service being offered.

IMPORTANT GROUNDING RULES:
- Use ONLY the information supplied above.
- Never invent company facts, statistics, rankings, traffic, customers,
  website problems, advertising activity, revenue, growth, or performance.
- The website URL is provided only as an identifier.
- Do NOT claim you visited, reviewed, analysed, or inspected the website.
- Do NOT claim the company has a specific problem unless that information
  was explicitly supplied.
- If information is limited, keep the observation broad but relevant.

WRITING STYLE:
- Write ONLY the opening observation.
- NEVER include the contact's name.
- NEVER include "Hi", "Hello", "Hey", or any greeting.
- The email template will add the greeting separately.
- Maximum ${maxWords} words.
- Target 12-25 words.
- Professional but conversational.
- Write like a human salesperson, not marketing copy.
- Avoid generic statements that could apply to every company.
- Do not use phrases such as:
  "for a company like"
  "in today's"
  "fast-paced"
  "competitive landscape"
  "robust digital presence"
  "high-intent buyers"
  "unlock"
  "leverage"
  "essential"
  "key to"
  "game-changing"
- Do not exaggerate.
- Do not pitch the service.
- Do not include a CTA.
- Do not introduce the sender.
- Do not repeat the company name unnecessarily.
- Use the company or industry context naturally.

SERVICE GUIDANCE:

Website Development:
Focus on the importance of a company's website or digital presence
without claiming its current website is poor.

SEO:
Focus on discoverability or organic search relevance without claiming
the company currently has poor rankings or traffic.

Google Ads:
Focus on reaching relevant potential customers through paid search
without claiming the company currently does or does not run ads.

Return ONLY valid JSON in exactly this structure:

{
  "intro": "one personalised sentence",
  "confidence": "high, medium, or low"
}
CONFIDENCE RULES:

"high":
Only when specific verified company information has been supplied.

"medium":
When useful company-specific context beyond basic name,
industry and URL has been supplied.

"low":
When only basic fields such as company name, industry,
website URL and service are available.

For the information provided in this prompt, do not use
"high" unless specific verified company information exists.
`;


  const response = callGemini(prompt);

  let result;


  // -------------------------------
  // PARSE JSON
  // -------------------------------

  try {

    let cleanedResponse =
      _cleanGeminiJson(response);


    result =
      JSON.parse(cleanedResponse);

  } catch (error) {

    throw new Error(
      "Gemini returned invalid JSON: " +
      response
    );
  }


  // -------------------------------
  // VALIDATE INTRO
  // -------------------------------

  if (
    !result.intro ||
    typeof result.intro !== "string"
  ) {

    throw new Error(
      "Gemini response does not contain a valid intro."
    );
  }


  const intro =
    result.intro.trim();

// Reject greetings because the email template
// already adds Hi {{FirstName}}
const greetingPattern =
  /^(hi|hello|hey|dear)\b/i;

if (greetingPattern.test(intro)) {

  throw new Error(
    "AI personalisation incorrectly included a greeting: " +
    intro
  );
}

  const wordCount =
    countWords(intro);


  if (wordCount === 0) {

    throw new Error(
      "Gemini generated an empty intro."
    );
  }


  if (wordCount > maxWords) {

    throw new Error(
      "AI personalisation exceeded word limit. " +
      "Generated " +
      wordCount +
      " words. Maximum: " +
      maxWords
    );
  }


  // -------------------------------
  // VALIDATE CONFIDENCE
  // -------------------------------

  const allowedConfidence = [
    "high",
    "medium",
    "low"
  ];


  const confidence =
    String(result.confidence || "low")
      .toLowerCase()
      .trim();


  if (!allowedConfidence.includes(confidence)) {

    throw new Error(
      "Invalid AI confidence value: " +
      confidence
    );
  }


  console.log(
    "AI confidence: " +
    confidence +
    " | Words: " +
    wordCount
  );


  return intro;
}

function countWords(text) {

  if (!text) {
    return 0;
  }

  return String(text)
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .length;
}

// ============================================================
// GEMINI JSON CLEANUP
// ============================================================
// Gemini sometimes wraps JSON in markdown code fences.
// Strips them so JSON.parse can succeed.

function _cleanGeminiJson(response) {

  if (!response) {
    return "";
  }

  return String(response)
    .trim()
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/, "")
    .replace(/```$/, "")
    .trim();
}

function generateMissingPersonalizations() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Leads");

  if (!sheet) {
    throw new Error("Leads sheet not found.");
  }


  const config = getAIConfig();

  const aiEnabled =
    isFlagTrue(config.AI_ENABLED);


  if (!aiEnabled) {

    console.log(
      "AI personalisation is disabled."
    );

    return;
  }


  const data =
    sheet.getDataRange().getValues();


  console.log(
    "=== AI PERSONALISATION STARTED ==="
  );


  for (let i = 1; i < data.length; i++) {

    const row = i + 1;


    const lead =
      buildLeadFromRow(data, i);


    // --------------------------------
    // ONLY NEW LEADS
    // --------------------------------

    if (lead.status !== "NEW") {
      continue;
    }


    // --------------------------------
    // ALREADY PERSONALISED
    // --------------------------------

    if (
      lead.personalisedIntro &&
      String(lead.personalisedIntro).trim() !== ""
    ) {

      console.log(
        "Skipping " +
        lead.email +
        ": Personalisation already exists."
      );

      continue;
    }


    // --------------------------------
    // SERVICE REQUIRED
    // --------------------------------

    if (
      !lead.service ||
      String(lead.service).trim() === ""
    ) {

      console.log(
        "Skipping " +
        lead.email +
        ": No service assigned."
      );

      continue;
    }


    try {

      console.log(
        "Generating personalisation for " +
        lead.email +
        " | Service: " +
        lead.service
      );


      let research = null;


      // =================================
      // WEBSITE RESEARCH
      // =================================

      if (
        lead.website &&
        String(lead.website).trim() !== ""
      ) {

        try {

          console.log(
            "Researching website: " +
            lead.website
          );


          research =
            getWebsiteResearch(lead);


          console.log(
            "Website research completed."
          );


          console.log(
            "Research confidence: " +
            research.researchConfidence
          );


          console.log(
            "Company summary: " +
            research.companySummary
          );


        } catch (researchError) {

          // Website research failure should NOT
          // stop the entire lead.

          console.error(
            "Website research failed for " +
            lead.email +
            ": " +
            researchError.message
          );


          console.log(
            "Falling back to basic lead personalisation."
          );


          research = null;
        }

      } else {

        console.log(
          "No website available for " +
          lead.email +
          ". Using basic personalisation."
        );

      }


      // =================================
      // GENERATE INTRO
      // =================================

      const intro =
        generatePersonalisation(
          lead,
          research
        );


      // =================================
      // SAVE RESULT
      // =================================

      // G - Personalised Intro
      sheet
        .getRange(row, 7)
        .setValue(intro);


      // R - Last Updated
      sheet
        .getRange(row, 18)
        .setValue(new Date());


      console.log(
        "Personalisation generated for " +
        lead.email +
        ": " +
        intro
      );


    } catch (error) {

      console.error(
        "AI personalisation failed for " +
        lead.email +
        ": " +
        error.message
      );


      // Don't allow one lead to kill
      // processing for every other lead.
      continue;
    }
  }


  console.log(
    "=== AI PERSONALISATION COMPLETED ==="
  );
}

function callGeminiWithUrlContext(prompt) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("GEMINI_API_KEY not found.");
  }

  const config = getAIConfig();

  const model = String(config.AI_MODEL).trim();

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent";

  const payload = {

    contents: [
      {
        role: "user",
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],

    // This enables Gemini URL Context
    tools: [
      {
        url_context: {}
      }
    ]

  };

  const options = {

    method: "post",

    contentType: "application/json",

    headers: {
      "x-goog-api-key": apiKey
    },

    payload: JSON.stringify(payload),

    muteHttpExceptions: true
  };


  // =================================
  // RETRY LOOP (same as callGemini)
  // =================================

  const maxRetries = 3;
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {

    try {

      _waitForAIRateLimit();

      const response = UrlFetchApp.fetch(
        url,
        options
      );

      const statusCode =
        response.getResponseCode();

      const responseText =
        response.getContentText();


      if (
        statusCode < 200 ||
        statusCode >= 300
      ) {

        if (
          (statusCode === 429 || statusCode >= 500) &&
          attempt < maxRetries
        ) {

          const backoff =
            Math.pow(2, attempt) * 1000 +
            Math.floor(Math.random() * 2000);

          console.log(
            "Gemini URL Context " + statusCode +
            " on attempt " + (attempt + 1) +
            "/" + maxRetries +
            ". Retrying in " +
            Math.round(backoff / 1000) + "s..."
          );

          Utilities.sleep(backoff);
          continue;
        }

        throw new Error(
          "Gemini URL Context API error " +
          statusCode +
          ": " +
          responseText
        );
      }


      const result =
        JSON.parse(responseText);


      if (
        !result.candidates ||
        result.candidates.length === 0
      ) {

        throw new Error(
          "Gemini returned no candidates."
        );
      }


      const parts =
        result.candidates[0]
          .content
          .parts;


      const text = parts
        .map(function(part) {
          return part.text || "";
        })
        .join("")
        .trim();


      if (!text) {
        throw new Error(
          "Gemini returned no text."
        );
      }


      return {
        text: text,
        raw: result
      };

    } catch (error) {

      lastError = error;
      if (attempt >= maxRetries) throw error;

      if (error instanceof SyntaxError) throw error;

      if (
        error.message &&
        error.message.indexOf("Gemini URL Context API error 4") === 0 &&
        error.message.indexOf("Gemini URL Context API error 429") !== 0
      ) {
        throw error;
      }

      const backoff =
        Math.pow(2, attempt) * 1000 +
        Math.floor(Math.random() * 2000);

      console.log(
        "Gemini URL Context failed on attempt " +
        (attempt + 1) + ": " +
        error.message +
        ". Retrying in " +
        Math.round(backoff / 1000) + "s..."
      );

      Utilities.sleep(backoff);
    }
  }

  throw lastError || new Error(
    "Gemini URL Context max retries exceeded."
  );
}
function testWebsiteContext() {

  const website =
    "https://www.dynamisers.com/";


  const prompt = `
Analyse the following public company website:

${website}

Use information retrieved from the supplied URL.

Return ONLY valid JSON using this structure:

{
  "companySummary": "brief description",
  "servicesFound": ["service 1", "service 2"],
  "targetAudience": "brief description",
  "evidence": [
    "specific fact found on the website",
    "another specific fact found on the website"
  ]
}

Rules:
- Use only information supported by the supplied website.
- Do not invent facts.
- Keep the company summary under 50 words.
- Keep evidence concise.
- If something cannot be determined, use "unknown".
- Do not wrap the JSON in markdown.
`;


  try {

    console.log(
      "=== WEBSITE CONTEXT TEST ==="
    );

    console.log(
      "Website: " + website
    );


    const result =
      callGeminiWithUrlContext(
        prompt
      );


    console.log(
      "Gemini result:"
    );

    console.log(
      result.text
    );


  } catch (error) {

    console.error(
      "Website context test failed: " +
      error.message
    );

  }
}
function researchWebsite(lead) {

  if (
    !lead.website ||
    String(lead.website).trim() === ""
  ) {
    throw new Error(
      "Website URL missing for " + lead.email
    );
  }


  const website =
    _normaliseWebsiteUrl(lead.website);


  const prompt = `
Research this company's public website:

${website}

COMPANY PROVIDED BY LEAD DATA:
${lead.company || "Unknown"}

Your job is to extract factual business context that can later
be used to personalise a B2B cold email.

Use information retrieved from the supplied website only.

Return ONLY valid JSON in exactly this structure:

{
  "companySummary": "short factual summary",
  "servicesFound": [
    "service/product 1",
    "service/product 2"
  ],
  "targetAudience": "who the company appears to serve",
  "valueProposition": "main value proposition if clearly stated",
  "usefulFacts": [
    "specific factual observation",
    "specific factual observation"
  ],
  "researchConfidence": "high, medium, or low"
}

RULES:

- Do not invent information.
- Do not guess statistics.
- Do not guess revenue, company size, traffic or rankings.
- Do not infer website performance.
- Do not claim SEO problems.
- Do not claim advertising problems.
- Do not make subjective statements about website quality.
- Extract only information supported by the website.
- Keep companySummary under 50 words.
- Keep usefulFacts concise.
- Return at most 3 usefulFacts.
- Return at most 5 services/products.
- If information cannot be determined, use "unknown".
- If the website contains very little useful information,
  set researchConfidence to "low".
- Do not wrap the JSON in markdown.
`;


  const response =
    callGeminiWithUrlContext(prompt);


  let cleaned =
    _cleanGeminiJson(response.text);


  let research;


  try {

    research =
      JSON.parse(cleaned);

  } catch (error) {

    throw new Error(
      "Website research returned invalid JSON: " +
      response.text
    );
  }


  // -------------------------------
  // BASIC VALIDATION
  // -------------------------------

  if (!research.companySummary) {
    research.companySummary = "unknown";
  }

  if (!Array.isArray(research.servicesFound)) {
    research.servicesFound = [];
  }

  if (!research.targetAudience) {
    research.targetAudience = "unknown";
  }

  if (!research.valueProposition) {
    research.valueProposition = "unknown";
  }

  if (!Array.isArray(research.usefulFacts)) {
    research.usefulFacts = [];
  }


  const allowedConfidence = [
    "high",
    "medium",
    "low"
  ];


  const confidence =
    String(
      research.researchConfidence || "low"
    )
      .toLowerCase()
      .trim();


  research.researchConfidence =
    allowedConfidence.includes(confidence)
      ? confidence
      : "low";


  return research;
}
function testWebsiteResearch() {

  const testLead = {

    company: "YOUR TEST COMPANY",

    email: "test@example.com",

    website: "dynamisers.com"

  };


  try {

    console.log(
      "=== WEBSITE RESEARCH STARTED ==="
    );


    const research =
      researchWebsite(testLead);


    console.log(
      JSON.stringify(
        research,
        null,
        2
      )
    );


    console.log(
      "=== WEBSITE RESEARCH COMPLETE ==="
    );


  } catch (error) {

    console.error(
      "Website research failed: " +
      error.message
    );

  }
}

function testResearchPersonalisation() {

  const testLead = {

    name: "Divyan",

    company: "YOUR TEST COMPANY",

    email: "test@example.com",

    website: "YOUR REAL TEST WEBSITE",

    industry: "Technology",

    service: "SEO"

  };


  try {

    console.log(
      "=== RESEARCHING WEBSITE ==="
    );


    const research =
      researchWebsite(testLead);


    console.log(
      "Research:"
    );

    console.log(
      JSON.stringify(
        research,
        null,
        2
      )
    );


    console.log(
      "=== GENERATING PERSONALISATION ==="
    );


    const intro =
      generatePersonalisation(
        testLead,
        research
      );


    console.log(
      "FINAL INTRO:"
    );

    console.log(intro);


  } catch (error) {

    console.error(
      "Research personalisation test failed: " +
      error.message
    );

  }
}
function saveWebsiteResearch(lead, research) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("AIResearch");


  if (!sheet) {
    throw new Error(
      "AIResearch sheet not found."
    );
  }


  sheet.appendRow([

    lead.leadId,

    lead.company,

    lead.website,

    research.companySummary,

    JSON.stringify(
      research.servicesFound || []
    ),

    research.targetAudience,

    research.valueProposition,

    JSON.stringify(
      research.usefulFacts || []
    ),

    research.researchConfidence,

    new Date()

  ]);


  console.log(
    "Website research cached for " +
    lead.email
  );
}
function getCachedWebsiteResearch(lead) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("AIResearch");


  if (!sheet) {
    throw new Error(
      "AIResearch sheet not found."
    );
  }


  const data =
    sheet.getDataRange().getValues();


  // Search backwards so newest research wins
  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    const cachedLeadId =
      String(data[i][0]).trim();

    const currentLeadId =
      String(lead.leadId).trim();


    if (
      cachedLeadId === currentLeadId
    ) {

      let servicesFound = [];
      let usefulFacts = [];


      try {

        servicesFound =
          JSON.parse(
            data[i][4] || "[]"
          );

      } catch (error) {

        servicesFound = [];

      }


      try {

        usefulFacts =
          JSON.parse(
            data[i][7] || "[]"
          );

      } catch (error) {

        usefulFacts = [];

      }


      return {

        companySummary:
          data[i][3] || "unknown",

        servicesFound:
          servicesFound,

        targetAudience:
          data[i][5] || "unknown",

        valueProposition:
          data[i][6] || "unknown",

        usefulFacts:
          usefulFacts,

        researchConfidence:
          data[i][8] || "low",

        researchedAt:
          data[i][9]

      };
    }
  }


  return null;
}
function getWebsiteResearch(lead) {

  // --------------------------------
  // CHECK CACHE
  // --------------------------------

  const cached =
    getCachedWebsiteResearch(lead);


  if (cached) {

    console.log(
      "Using cached website research for " +
      lead.email
    );

    return cached;
  }


  // --------------------------------
  // NO CACHE → RESEARCH
  // --------------------------------

  console.log(
    "No cached research found for " +
    lead.email
  );


  console.log(
    "Researching website: " +
    lead.website
  );


  const research =
    researchWebsite(lead);


  // --------------------------------
  // SAVE
  // --------------------------------

  saveWebsiteResearch(
    lead,
    research
  );


  return research;
}
