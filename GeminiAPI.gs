// ============================================================
//  GeminiAPI.gs — Gemini API integration
//
//  Used for:
//  1. Email parsing fallback (when regex parsers fail)
//  2. Extracting structured job data from ambiguous emails
//
//  Uses Gemini 2.5 Flash (free tier: 15 req/min, 1500 req/day)
// ============================================================


// ─────────────────────────────────────────────
//  EMAIL PARSING FALLBACK
// ─────────────────────────────────────────────

/**
 * Uses Gemini to extract job listings from an email when the
 * native regex parsers couldn't parse the format.
 *
 * Returns an array of job objects (same shape as other parsers).
 */
function parseEmailWithGemini(subject, body, source, emailDate) {
  const config = getConfig();
  const jobs   = [];

  const prompt = `Extract all job listings from this ${source} job alert email. Return ONLY a valid JSON array, no other text, no markdown backticks.

EMAIL SUBJECT: ${subject}

EMAIL BODY:
${body.substring(0, 4000)}

For each job listing found, return an object with these exact fields:
{
  "title": "Job Title",
  "company": "Company Name",
  "location": "City, State or Remote",
  "url": "https://...",
  "isRemote": true or false,
  "isRepost": true or false,
  "ageText": "X hours ago or X days ago or empty string"
}

Rules:
- If location contains "remote" (case-insensitive), set isRemote to true
- If the job title or nearby text contains "repost" or "re-post", set isRepost to true  
- If no URL is found for a listing, use an empty string for url
- If you find no job listings, return an empty array []
- Return only the raw JSON array, nothing else
- Do not include any explanation, apology, or commentary — only the JSON array
- Do not wrap the JSON in markdown code fences`;

  try {
    const responseText = callGemini(prompt, 2048);
    Logger.log('  Gemini raw response (' + responseText.length + ' chars): ' + responseText.substring(0, 300));

    // Strip markdown code fences if Gemini wrapped the JSON
    let cleaned = responseText.trim();
    cleaned = cleaned.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

    // If response is too short to be valid JSON array, skip parsing
    if (cleaned.length < 10) {
      Logger.log('  Gemini parser: response too short (' + cleaned.length + ' chars) -- skipping');
      return [];
    }

    // Ensure it starts with [ -- if not, try to extract JSON array from response
    if (!cleaned.startsWith('[')) {
      const arrayMatch = cleaned.match(/\[[\s\S]*\]/);
      if (arrayMatch) {
        cleaned = arrayMatch[0];
        Logger.log('  Gemini parser: extracted JSON array from response');
      } else {
        Logger.log('  Gemini parser: response does not contain a JSON array -- skipping');
        return [];
      }
    }

    const parsed = JSON.parse(cleaned);

    if (!Array.isArray(parsed)) {
      Logger.log('  Gemini parser: response was not an array');
      return [];
    }

    for (const item of parsed) {
      if (!item.title || !item.company) continue;

      const ageDate = parseDateFromAgeText(item.ageText || '') || emailDate;

      jobs.push({
        source:     source,
        company:    cleanText(item.company),
        title:      cleanText(item.title),
        location:   cleanText(item.location) || 'See posting',
        url:        item.url || '',
        isRemote:   !!item.isRemote,
        isRepost:   !!item.isRepost,
        postedDate: ageDate,
        emailBody:  body,
        subject:    subject,
      });
    }

    Logger.log(`  Gemini fallback parsed ${jobs.length} job(s)`);
  } catch (e) {
    Logger.log(`  Gemini parser error: ${e.message}`);
  }

  return jobs;
}


// ─────────────────────────────────────────────
//  GEMINI API HTTP CALL
// ─────────────────────────────────────────────

/**
 * Makes a call to the Google Gemini API.
 * Retries up to 3 times on transient errors (503, 429, 500).
 * Returns the text response string.
 */
function callGemini(prompt, maxTokens) {
  const config = getConfig();
  const apiKey = config.GEMINI_API_KEY;

  if (!apiKey || apiKey === 'YOUR_GEMINI_API_KEY_HERE') {
    throw new Error('Gemini API key not configured. Set GEMINI_API_KEY in Config.gs or Script Properties.');
  }

  const model   = config.GEMINI_MODEL || 'gemini-2.5-flash';
  const url     = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      maxOutputTokens: maxTokens || 1024,
      temperature:     0.1,
    }
  };

  const options = {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  // Transient error codes worth retrying
  const RETRYABLE = [429, 500, 503, 529];
  const MAX_ATTEMPTS = 3;
  const RETRY_DELAYS = [3000, 8000, 15000]; // ms between attempts

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    Logger.log('  Calling Gemini API (model: ' + model + ', attempt ' + attempt + '/' + MAX_ATTEMPTS + ')...');

    let response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (e) {
      if (attempt < MAX_ATTEMPTS) {
        Logger.log('  Gemini network error: ' + e.message + ' — retrying in ' + (RETRY_DELAYS[attempt-1]/1000) + 's');
        Utilities.sleep(RETRY_DELAYS[attempt - 1]);
        continue;
      }
      throw new Error('Gemini API network error: ' + e.message);
    }

    const statusCode = response.getResponseCode();
    const bodyText   = response.getContentText();

    // Success
    if (statusCode === 200) {
      const parsed = JSON.parse(bodyText);
      try {
        const text = parsed.candidates[0].content.parts[0].text;
        Logger.log('  Gemini response received (' + text.length + ' chars)');
        return text;
      } catch (e) {
        throw new Error('Could not extract text from Gemini response');
      }
    }

    // Transient error — retry if attempts remain
    if (RETRYABLE.indexOf(statusCode) !== -1 && attempt < MAX_ATTEMPTS) {
      Logger.log('  Gemini ' + statusCode + ' (transient) — retrying in ' + (RETRY_DELAYS[attempt-1]/1000) + 's...');
      Utilities.sleep(RETRY_DELAYS[attempt - 1]);
      continue;
    }

    // Non-retryable error or out of attempts
    Logger.log('  Gemini API error ' + statusCode + ': ' + bodyText.substring(0, 300));
    throw new Error('Gemini API returned ' + statusCode);
  }

  throw new Error('Gemini API failed after ' + MAX_ATTEMPTS + ' attempts');
}
