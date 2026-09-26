// ============================================================
//  ClaudeAPI.gs — Claude API integration
//
//  Generates tailored resumes and cover letters.
//  Em dashes are explicitly forbidden in all prompts.
//  Output uses # / ## / ### / - markers for Doc formatting.
//
//  JD PREPROCESSOR: preprocessJobDescription() restructures
//  the job description before sending to Claude, ensuring
//  signal-rich sections (responsibilities, qualifications,
//  skills) are always included within the character budget,
//  with company boilerplate deprioritized or dropped.
// ============================================================


// ─────────────────────────────────────────────
//  JOB DESCRIPTION PREPROCESSOR
// ─────────────────────────────────────────────

/**
 * Restructures a raw job description to maximize signal quality
 * within the MAX_JD_CHARS budget before sending to Claude.
 *
 * Priority order (highest to lowest):
 *   Tier 1 — Responsibilities, Duties, What You'll Do, Day-to-Day
 *   Tier 2 — Qualifications, Requirements, Skills, Experience
 *   Tier 3 — About the Role, Role Overview, Position Summary
 *   Tier 4 — About the Team, About the Company, Benefits, Culture
 *   Drop   — EEO statements, legal boilerplate, hashtags, footers
 *
 * Returns a restructured string within maxChars.
 * If section detection fails, falls back to smart truncation
 * that favors the middle/end of the document (where quals live)
 * over the top (where boilerplate lives).
 */
function preprocessJobDescription(rawJd, maxChars) {
  if (!rawJd) return '';
  if (rawJd.length <= maxChars) return rawJd; // no processing needed

  // Clean up the raw text first
  let text = rawJd
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  // ── Drop boilerplate sections entirely ───────────────────────
  // These patterns match common footer/legal blocks that add zero
  // signal for resume tailoring.
  const dropPatterns = [
    /equal opportunity employer[\s\S]{0,800}/gi,
    /eeo[\s\S]{0,400}/gi,
    /we are committed to diversity[\s\S]{0,600}/gi,
    /reasonable accommodation[\s\S]{0,400}/gi,
    /applicants.*disability[\s\S]{0,400}/gi,
    /#[A-Za-z0-9]+/g,                          // hashtags
    /\bsalary range\b.*\n?.*\n?/gi,            // salary lines (often noise)
    /apply now[\s\S]{0,200}$/gi,               // apply CTA at end
  ];

  for (const pattern of dropPatterns) {
    text = text.replace(pattern, '');
  }
  text = text.replace(/\n{3,}/g, '\n\n').trim();

  // If cleaning alone brought it under the limit, return now
  if (text.length <= maxChars) {
    Logger.log(`  JD preprocessor: cleaned to ${text.length} chars (under limit, no sectioning needed)`);
    return text;
  }

  // ── Section detection ─────────────────────────────────────────
  // Split text into labeled sections by scanning for heading lines.
  const sections = detectJdSections(text);

  if (sections.length === 0) {
    // No sections detected — fall back to smart truncation
    Logger.log(`  JD preprocessor: no sections detected, using smart truncation`);
    return smartTruncate(text, maxChars);
  }

  // ── Tier assignment ───────────────────────────────────────────
  const tier1Keywords = [
    'responsibilit', 'duties', "what you'll do", 'what you will do',
    'day-to-day', 'day to day', 'in this role', 'your role',
    'what we need', 'key responsibilities', 'core responsibilities',
    'job duties', 'position duties', 'essential functions',
  ];
  const tier2Keywords = [
    'qualifications', 'requirements', 'required', 'skills',
    'experience', 'education', 'preferred', 'minimum qualifications',
    'basic qualifications', 'what you bring', 'what we are looking for',
    'what we look for', 'you have', 'you bring', 'must have',
    'nice to have', 'technical skills', 'competencies',
  ];
  const tier3Keywords = [
    'about the role', 'about this role', 'role overview', 'position summary',
    'about the position', 'the opportunity', 'overview', 'summary',
    'about this job', 'job summary', 'position overview',
  ];
  // Tier 4 = everything else (about company, benefits, culture, etc.)

  const tiered = { 1: [], 2: [], 3: [], 4: [] };

  for (const section of sections) {
    const heading = section.heading.toLowerCase();
    if (tier1Keywords.some(k => heading.includes(k))) {
      tiered[1].push(section);
    } else if (tier2Keywords.some(k => heading.includes(k))) {
      tiered[2].push(section);
    } else if (tier3Keywords.some(k => heading.includes(k))) {
      tiered[3].push(section);
    } else {
      tiered[4].push(section);
    }
  }

  // ── Assemble within budget ────────────────────────────────────
  // Fill tiers in priority order until the budget is consumed.
  let assembled = '';
  const tierOrder = [1, 2, 3, 4];

  for (const tier of tierOrder) {
    for (const section of tiered[tier]) {
      const candidate = (assembled ? '\n\n' : '') +
                        (section.heading ? section.heading + '\n' : '') +
                        section.body;
      if ((assembled + candidate).length <= maxChars) {
        assembled += candidate;
      } else {
        // Try to fit a truncated version of this section
        const remaining = maxChars - assembled.length - 10;
        if (remaining > 200 && section.body.length > 0) {
          const truncatedBody = section.body.substring(0, remaining);
          // Don't cut mid-word
          const lastSpace = truncatedBody.lastIndexOf(' ');
          const safeBody  = lastSpace > 100 ? truncatedBody.substring(0, lastSpace) + '...' : truncatedBody;
          assembled += (assembled ? '\n\n' : '') +
                       (section.heading ? section.heading + '\n' : '') +
                       safeBody;
        }
        // Budget exhausted — stop adding sections
        if (assembled.length >= maxChars * 0.9) break;
      }
    }
    if (assembled.length >= maxChars * 0.9) break;
  }

  if (!assembled) {
    // Sectioning produced nothing usable — fall back
    Logger.log(`  JD preprocessor: sectioning failed, using smart truncation`);
    return smartTruncate(text, maxChars);
  }

  Logger.log(`  JD preprocessor: ${rawJd.length} chars → ${assembled.length} chars (T1:${tiered[1].length} T2:${tiered[2].length} T3:${tiered[3].length} T4:${tiered[4].length} sections)`);
  return assembled.trim();
}

/**
 * Splits a job description into labeled sections by detecting
 * heading lines (short lines followed by body content).
 * Returns array of { heading, body } objects.
 */
function detectJdSections(text) {
  const lines    = text.split('\n');
  const sections = [];
  let currentHeading = '';
  let currentBody    = [];

  for (let i = 0; i < lines.length; i++) {
    const line      = lines[i];
    const trimmed   = line.trim();
    const nextLine  = i + 1 < lines.length ? lines[i + 1].trim() : '';

    // A heading is: short (under 80 chars), not ending in punctuation
    // (except colon), and followed by a blank line or indented content
    const isHeading = (
      trimmed.length > 0 &&
      trimmed.length < 80 &&
      !/[.!?]$/.test(trimmed) &&
      (trimmed.endsWith(':') || nextLine === '' || /^\s{2,}/.test(lines[i + 1] || ''))
    );

    if (isHeading && currentBody.join('').trim().length > 20) {
      // Save previous section
      sections.push({
        heading: currentHeading,
        body:    currentBody.join('\n').trim(),
      });
      currentHeading = trimmed;
      currentBody    = [];
    } else if (isHeading && currentBody.join('').trim().length === 0) {
      currentHeading = trimmed;
    } else {
      currentBody.push(line);
    }
  }

  // Save final section
  if (currentBody.join('').trim().length > 0) {
    sections.push({
      heading: currentHeading,
      body:    currentBody.join('\n').trim(),
    });
  }

  return sections;
}

/**
 * Smart truncation fallback — when section detection fails.
 * Rather than always taking the top of the document (which is
 * often boilerplate), this attempts to favor the middle portion
 * where responsibilities and qualifications typically live.
 */
function smartTruncate(text, maxChars) {
  if (text.length <= maxChars) return text;

  // If the document is only slightly over, just trim the end
  if (text.length <= maxChars * 1.3) {
    const cut = text.substring(0, maxChars);
    const lastBreak = cut.lastIndexOf('\n\n');
    return lastBreak > maxChars * 0.7 ? cut.substring(0, lastBreak) : cut;
  }

  // For longer documents, take the first third (intro + responsibilities)
  // and the middle third (qualifications), skip the end (legal/boilerplate)
  const firstThird  = Math.floor(text.length / 3);
  const firstChunk  = text.substring(0, Math.floor(maxChars * 0.45));
  const midStart    = firstThird;
  const midChunk    = text.substring(midStart, midStart + Math.floor(maxChars * 0.50));

  const combined = firstChunk + '\n\n[...]\n\n' + midChunk;
  Logger.log(`  JD smart truncation: took first 45% + mid 50% of document`);
  return combined.substring(0, maxChars);
}


// ─────────────────────────────────────────────
//  RESUME GENERATION
// ─────────────────────────────────────────────

/**
 * Calls Claude to generate a tailored resume.
 * Returns formatted text using heading markers.
 */
function generateResume(job, jobDescription, baseResume) {
  const config        = getConfig();
  const processedJd   = preprocessJobDescription(jobDescription, config.MAX_JD_CHARS);

  const prompt = `You are an expert resume writer. Your task is to create a tailored, ATS-optimized resume for the job below.

## CANDIDATE'S BASE RESUME
${baseResume}

## TARGET JOB
**Company:** ${job.company}
**Title:** ${job.title}
**Location:** ${job.location} ${job.isRemote ? '(Remote)' : ''}
**Source:** ${job.source}

## JOB DESCRIPTION
${processedJd}

## JOB DESCRIPTION QUALITY NOTE
${processedJd.length < 800 ? 'WARNING: The job description above is very short and may be incomplete. Tailor conservatively — prioritize exact language from the JD that IS present, and avoid inferring requirements that are not explicitly stated.' : processedJd.length < 1500 ? 'NOTE: The job description is moderately short. Mirror the language present closely and do not embellish with assumed requirements.' : 'The job description is complete. Full tailoring is appropriate.'}

## INSTRUCTIONS

1. Tailor the resume specifically to this role. Mirror keywords and phrases from the job description naturally — only mirror what is explicitly stated, never infer or fabricate requirements not present in the JD. Mirror keywords and phrases from the job description naturally throughout.
2. Reorder, strengthen, or reframe bullet points from the base resume to best match what this employer is looking for.
3. Prioritize and surface the most relevant experience, skills, and achievements for this specific role.
4. Quantify achievements wherever the base resume allows for it.
5. The resume should be 2 pages. Do not compress or omit relevant experience to fit one page. The candidate has 20 years of experience across QA Engineering and IT/Technical Support — both domains are relevant and should be represented appropriately. Lead with QA experience for engineering-focused roles and lead with support/troubleshooting experience for technical support or customer-facing roles.
6. Use clean, professional language. Be specific and results-oriented.
7. ABSOLUTELY FORBIDDEN: Do not use the em dash character (—) anywhere in the resume. Not once. Use commas, periods, colons, or rephrase instead.
8. Do not fabricate experience, titles, companies, dates, or skills not present in the base resume.
9. Do not include an objective statement. Start with a strong professional summary (3 sentences max).

## OUTPUT FORMAT
Use these markers exactly for structure — the document formatter depends on them:
- # for the candidate's name (first line only)
- ## for section headings (Summary, Experience, Skills, Education, etc.)
- ### for job title + company lines within Experience
- - for bullet points
- **text** for bold inline text (company names, degree names, etc.)
- Separate sections with a blank line
- Use plain hyphens for date ranges, not em dashes

Output the resume only. No preamble, no explanation, no closing note.`;

  const response = callClaude(prompt, config.CLAUDE_MAX_TOKENS);
  return removeEmDashes(response);
}


// ─────────────────────────────────────────────
//  COVER LETTER GENERATION
// ─────────────────────────────────────────────

/**
 * Calls Claude to generate a tailored cover letter.
 * Returns formatted text using heading markers.
 */
function generateCoverLetter(job, jobDescription, baseResume, workExamples) {
  const config      = getConfig();
  const processedJd = preprocessJobDescription(jobDescription, config.MAX_JD_CHARS);

  const workExamplesSection = workExamples
    ? `\n## ADDITIONAL WORK EXPERIENCE & STORY EXAMPLES\nUse these to add authentic, specific detail to the cover letter body paragraphs:\n${workExamples.substring(0, 2000)}`
    : '';

  const prompt = `You are an expert cover letter writer. Your task is to write a compelling, personalized cover letter for the job below.

## CANDIDATE'S BASE RESUME
${baseResume.substring(0, 3000)}
${workExamplesSection}

## TARGET JOB
**Company:** ${job.company}
**Title:** ${job.title}
**Location:** ${job.location} ${job.isRemote ? '(Remote)' : ''}

## JOB DESCRIPTION
${processedJd}

## JOB DESCRIPTION QUALITY NOTE
${processedJd.length < 800 ? 'WARNING: The job description above is very short and may be incomplete. Write the cover letter based only on what is explicitly stated — do not assume or invent requirements.' : processedJd.length < 1500 ? 'NOTE: The job description is moderately short. Base the cover letter strictly on what is present.' : 'The job description is complete. Full tailoring is appropriate.'}

## INSTRUCTIONS

1. Write a genuine, compelling cover letter — not a generic template.
2. Opening paragraph: Hook the reader immediately. Reference something specific about the company or role that connects to the candidate's genuine interest or background. Do not reference generic mission statements, company values language, or boilerplate culture descriptions as the "specific" detail. Reference the actual work, technical environment, product, scale of infrastructure, or challenge described in the responsibilities and qualifications sections instead.
3. Middle paragraphs (2-3): Highlight 2-3 specific achievements or experiences from the resume/work examples that directly address what the job description asks for. Be concrete and specific.
4. Closing paragraph: Express clear enthusiasm, include a call to action, and keep it professional but human.
5. Tone: Confident, professional, and authentic. Not stiff or robotic.
6. Length: 3-4 paragraphs. No longer than one page.
7. ABSOLUTELY FORBIDDEN: Do not use the em dash character (—) anywhere. Not once. Rephrase or use commas/colons instead.
8. Do not repeat the resume verbatim. Tell a story, don't list bullets.
9. Do not use filler phrases like "I am writing to express my interest" or "Please find attached."
10. Address it to the Hiring Manager unless a specific name is in the job description.

## OUTPUT FORMAT
Use these markers for structure:
- # for the date line (e.g., # June 5, 2026)
- ## for the salutation line (e.g., ## Dear Hiring Manager,)
- Normal paragraphs (no prefix) for the body
- ## for the closing line (e.g., ## Sincerely,)
- # for the candidate's name at the end

Output the cover letter only. No preamble, no explanation.`;

  const response = callClaude(prompt, config.CLAUDE_MAX_TOKENS);
  return removeEmDashes(response);
}


// ─────────────────────────────────────────────
//  CLAUDE API HTTP CALL
// ─────────────────────────────────────────────

/**
 * Makes a call to the Anthropic Claude API.
 * Returns the text response string.
 */
function callClaude(prompt, maxTokens) {
  const config = getConfig();
  const apiKey = config.CLAUDE_API_KEY;

  if (!apiKey || apiKey === 'YOUR_CLAUDE_API_KEY_HERE') {
    throw new Error('Claude API key not configured. Set CLAUDE_API_KEY in Config.gs or Script Properties.');
  }

  const payload = {
    model:      config.CLAUDE_MODEL,
    max_tokens: maxTokens || 4096,
    messages: [
      {
        role:    'user',
        content: prompt,
      }
    ],
  };

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  Logger.log(`  Calling Claude API (model: ${config.CLAUDE_MODEL}, max_tokens: ${maxTokens})...`);

  let response;
  try {
    response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  } catch (e) {
    throw new Error(`Claude API network error: ${e.message}`);
  }

  const statusCode = response.getResponseCode();
  const bodyText   = response.getContentText();

  if (statusCode !== 200) {
    Logger.log(`  Claude API error ${statusCode}: ${bodyText.substring(0, 500)}`);
    throw new Error(`Claude API returned ${statusCode}: ${JSON.parse(bodyText).error?.message || bodyText.substring(0, 200)}`);
  }

  const parsed = JSON.parse(bodyText);

  if (parsed.content && parsed.content.length > 0) {
    const textBlocks = parsed.content.filter(b => b.type === 'text').map(b => b.text);
    const result     = textBlocks.join('\n').trim();
    Logger.log(`  Claude response received (${result.length} chars, ${parsed.usage?.output_tokens || '?'} tokens)`);
    return result;
  }

  throw new Error('Claude API returned an empty response.');
}
