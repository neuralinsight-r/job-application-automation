// ============================================================
//  JOB APPLICATION AUTOMATION SYSTEM
//  Google Apps Script — Main Orchestrator
//
//  FILES IN THIS PROJECT:
//    Code.gs          <- this file (orchestration + email parsing)
//    SheetManager.gs  <- Google Sheet read/write logic
//    DriveManager.gs  <- Google Drive folder + Doc creation
//    ClaudeAPI.gs     <- Claude API calls (resume + cover letter)
//    GeminiAPI.gs     <- Gemini API calls (email parsing assist)
//    Config.gs        <- All user-configurable settings
// ============================================================


// ─────────────────────────────────────────────
//  ENTRY POINTS
// ─────────────────────────────────────────────

/**
 * TRIGGER 1 — Run every 15 minutes via time-based trigger.
 * Scans Gmail for new job alert emails from LinkedIn, Indeed, Glassdoor.
 */
function scanJobAlerts() {
  const config = getConfig();
  Logger.log('=== scanJobAlerts started ===');

  const sources = [
    { label: 'LinkedIn',  query: config.GMAIL_LABEL_LINKEDIN  },
    { label: 'Indeed',    query: config.GMAIL_LABEL_INDEED    },
    { label: 'Glassdoor', query: config.GMAIL_LABEL_GLASSDOOR },
  ];

  let newJobCount = 0;

  for (const source of sources) {
    try {
      const threads = GmailApp.search(source.query + ' is:unread', 0, 20);
      Logger.log(source.label + ': found ' + threads.length + ' unread thread(s)');

      for (const thread of threads) {
        const messages = thread.getMessages();
        for (const message of messages) {
          if (message.isUnread()) {
            const jobs = parseEmailForJobs(message, source.label);
            for (const job of jobs) {
              const added = processNewJob(job);
              if (added) newJobCount++;
            }
            message.markRead();
          }
        }
      }
    } catch (e) {
      Logger.log('ERROR scanning ' + source.label + ': ' + e.message);
    }
  }

  Logger.log('=== scanJobAlerts complete. ' + newJobCount + ' new job(s) added. ===');
}

/**
 * TRIGGER 2 — Manual or menu-driven.
 * Generates resume + cover letter for all rows marked "Pending".
 */
function generatePendingDocs() {
  Logger.log('=== generatePendingDocs started ===');
  const pendingJobs = getJobsByStatus('Pending');
  Logger.log('Found ' + pendingJobs.length + ' pending job(s)');
  for (const job of pendingJobs) {
    generateDocsForJob(job);
  }
  Logger.log('=== generatePendingDocs complete ===');
}

/**
 * TRIGGER 3 — Adds custom menu to Sheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Automation')
    .addItem('Scan Job Alert Emails Now', 'scanJobAlerts')
    .addItem('Generate Docs for Pending (P3, P5, P10, P15)', 'generatePendingDocs')
    .addSeparator()
    .addItem('Generate Docs for Selected Row', 'generateDocsForSelectedRow')
    .addItem('Re-process Selected Row (force)', 'reprocessSelectedRow')
    .addSeparator()
    .addItem('Process Manual Entry (Selected Row)', 'processManualEntry')
    .addItem('Validate Manual Entry Row', 'validateManualEntryRow')
    .addSeparator()
    .addItem('Backfill Missing LinkedIn URLs', 'backfillMissingLinkedInUrls')
    .addSeparator()
    .addItem('Open Parent Drive Folder', 'openParentFolder')
    .addItem('View Setup Guide', 'showSetupGuide')
    .addToUi();
}


// ─────────────────────────────────────────────
//  CORE ORCHESTRATION
// ─────────────────────────────────────────────

function processNewJob(job) {
  // Layer 0: Company blocklist — skip entirely, no row written
  const config0 = getConfig();
  const blocked  = config0.BLOCKED_COMPANIES || [];
  const companyLower = (job.company || '').toLowerCase();
  for (const blockedName of blocked) {
    if (companyLower.includes(blockedName.toLowerCase())) {
      Logger.log('SKIP blocked company: ' + job.company + ' - ' + job.title);
      return false;
    }
  }

  // Layer 1: Skip reposts flagged in email body
  if (job.isRepost) {
    Logger.log('SKIP repost (email flag): ' + job.company + ' - ' + job.title);
    return false;
  }

  // Layer 2: LinkedIn page-level repost check
  if (job.source === 'LinkedIn' && job.url) {
    const repostResult = isLinkedInRepost(job.url);
    if (repostResult === true) {
      Logger.log('SKIP repost (page confirmed): ' + job.company + ' - ' + job.title);
      return false;
    }
    if (repostResult === 'unverified') {
      Logger.log('WARN repost unverified: ' + job.company + ' - ' + job.title);
      job.repostUnverified = true;
    }
  }

  // Layer 3: Automation title filter
  const automationFlag = isAutomationHeavyTitle(job.title);
  if (automationFlag) {
    job.automationFlagged = true;
    Logger.log('FLAGGED automation-heavy title: ' + job.company + ' - ' + job.title + ' (matched: "' + automationFlag + '")');
  }

  // Layer 4: Pass-1 location filter
  const locationFilter = applyLocationFilterPass1(job);
  if (locationFilter.ignore) {
    Logger.log('SKIP ' + locationFilter.reason + ': ' + job.company + ' - ' + job.title);
    return false;
  }
  if (locationFilter.priority) job.priorityOverride = locationFilter.priority;
  if (locationFilter.status)   job.locationStatus   = locationFilter.status;

  // Deduplicate
  if (isDuplicate(job)) {
    Logger.log('SKIP duplicate: ' + job.company + ' - ' + job.title);
    return false;
  }

  // Pass-1 priority
  job.priority = job.priorityOverride || calculatePriority(job);

  // Write row to correct tab (P1/P2 to Job Tracker, all others to Other)
  const appendResult = appendJobToSheet(job);
  job.rowIndex  = appendResult.rowIndex;
  job.sheetName = appendResult.sheetName;

  // Handle special status flags — all require manual review
  if (job.repostUnverified) {
    updateJobStatus(job.rowIndex, 'Verify repost?', job.sheetName);
    Logger.log('Added row ' + job.rowIndex + ' [' + job.sheetName + ']: [P' + job.priority + '] ' + job.company + ' - ' + job.title + ' (repost unverified)');
    return true;
  }

  if (job.automationFlagged) {
    updateJobStatus(job.rowIndex, 'Review: automation-heavy title', job.sheetName);
    Logger.log('Added row ' + job.rowIndex + ' [' + job.sheetName + ']: [P' + job.priority + '] ' + job.company + ' - ' + job.title + ' (automation flagged)');
    return true;
  }

  if (job.locationStatus) {
    updateJobStatus(job.rowIndex, job.locationStatus, job.sheetName);
    Logger.log('Added row ' + job.rowIndex + ' [' + job.sheetName + ']: [P' + job.priority + '] ' + job.company + ' - ' + job.title + ' (' + job.locationStatus + ')');
    return true;
  }

  Logger.log('Added row ' + job.rowIndex + ' [' + job.sheetName + ']: [P' + job.priority + '] ' + job.company + ' - ' + job.title);

  // AUTO_GENERATE toggle — when false, nothing generates automatically.
  // All generation requires manual trigger via the Job Automation menu.
  const config2 = getConfig();
  if (config2.AUTO_GENERATE === true && job.priority <= 2) {
    Logger.log('Auto-generating docs for P' + job.priority + ' job...');
    generateDocsForJob(job);
  } else if (job.priority <= 2) {
    Logger.log('P' + job.priority + ' job queued in Job Tracker (AUTO_GENERATE is off)');
  }

  return true;
}

function generateDocsForJob(job) {
  const config    = getConfig();
  const sheetName = job.sheetName || null; // pass to all sheet updates

  try {
    updateJobStatus(job.rowIndex, 'Generating...', sheetName);

    // 1. Fetch job description
    const jobDescription = fetchJobDescription(job);
    const MIN_JD = config.MIN_JD_CHARS || 2500;

    if (!jobDescription || jobDescription.length < 50) {
      Logger.log('WARNING: Empty job description for ' + job.company + ' - ' + job.title);
      updateJobStatus(job.rowIndex, 'Review: thin JD — could not fetch description', sheetName);
      return false;
    }

    if (jobDescription.length < MIN_JD) {
      Logger.log('WARNING: Thin JD (' + jobDescription.length + ' chars, need ' + MIN_JD + ') for ' + job.company + ' - ' + job.title);
      updateJobStatus(job.rowIndex, 'Review: thin JD (' + jobDescription.length + ' chars)', sheetName);
      return false;
    }

    Logger.log('  JD accepted: ' + jobDescription.length + ' chars (threshold: ' + MIN_JD + ')');

    // 2. Pass-2 location check (state restriction / TX)
    const pass2 = applyLocationFilterPass2(job, jobDescription);
    if (pass2.ignore) {
      Logger.log('SKIP (Pass-2) ' + pass2.reason + ': ' + job.company + ' - ' + job.title);
      updateJobStatus(job.rowIndex, 'Ignored: ' + pass2.reason, sheetName);
      return false;
    }
    if (pass2.priority && pass2.priority !== job.priority) {
      job.priority = pass2.priority;
      updateJobRow(job.rowIndex, { priority: pass2.priority }, sheetName);
      Logger.log('  Pass-2 priority updated to P' + job.priority + ': ' + pass2.reason);
    }
    if (pass2.status) {
      updateJobStatus(job.rowIndex, pass2.status, sheetName);
      Logger.log('  Pass-2 flagged for manual review (' + pass2.status + ') — aborting auto-gen');
      return false;
    }

    // 3. Load base resume + work examples
    const baseResume   = loadTextFileFromDrive(config.BASE_RESUME_FILE_ID);
    const workExamples = loadTextFileFromDrive(config.WORK_EXAMPLES_FILE_ID);

    // 4. Generate resume via Claude
    const resumeText = generateResume(job, jobDescription, baseResume);

    // 5. Generate cover letter via Claude (if enabled)
    let coverLetterText = null;
    if (config.GENERATE_COVER_LETTERS !== false) {
      coverLetterText = generateCoverLetter(job, jobDescription, baseResume, workExamples);
    }

    // 6. Create Drive folder and Google Docs
    const folderName = sanitizeFolderName(job.company + ' - ' + job.title);
    const folder     = createJobFolder(folderName);
    createGoogleDoc(folder, 'Resume - ' + job.company + ' - ' + job.title, resumeText);
    if (coverLetterText) {
      createGoogleDoc(folder, 'Cover Letter - ' + job.company + ' - ' + job.title, coverLetterText);
    }

    // 7. Update Sheet
    const folderUrl = folder.getUrl();
    updateJobRow(job.rowIndex, {
      status:        'Done',
      folderUrl:     folderUrl,
      processedDate: new Date(),
    }, sheetName);

    Logger.log('SUCCESS: Docs created for ' + job.company + ' - ' + job.title);
    return true;

  } catch (e) {
    Logger.log('ERROR generating docs for ' + job.company + ' - ' + job.title + ': ' + e.message);
    updateJobStatus(job.rowIndex, 'Error: ' + e.message.substring(0, 80), sheetName);
    return false;
  }
}


// ─────────────────────────────────────────────
//  URL BACKFILL
// ─────────────────────────────────────────────

function backfillMissingLinkedInUrls() {
  const config  = getConfig();
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ui = SpreadsheetApp.getUi();
  let fixed = 0, fallback = 0, skipped = 0;

  const data = sheet.getRange(2, 1, lastRow - 1, config.TOTAL_COLUMNS).getValues();

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const rowIndex = i + 2;
    const source   = (row[config.COL_SOURCE - 1] || '').toString().trim();

    if (source !== 'LinkedIn') continue;

    const urlCell     = sheet.getRange(rowIndex, config.COL_JOB_URL);
    const cellValue   = urlCell.getValue().toString().trim();
    const cellFormula = urlCell.getFormula();
    if (cellValue && cellValue !== '' && cellValue !== 'View Posting') { skipped++; continue; }
    if (cellFormula && cellFormula.includes('HYPERLINK')) { skipped++; continue; }

    let url = '';
    try {
      const note = sheet.getRange(rowIndex, config.COL_COMPANY).getNote();
      if (note) {
        const meta = JSON.parse(note);
        if (meta.url && meta.url.startsWith('http')) {
          url = meta.url;
        }
        if (!url && meta.emailBody) {
          const m1 = meta.emailBody.match(/View job[:\s]+(https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/[^\s\n?]+)/i);
          if (m1) url = m1[1].trim();
        }
        if (!url && meta.emailBody) {
          const m2 = meta.emailBody.match(/https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/[^\s\n?&]+/);
          if (m2) url = m2[0].trim();
        }
      }
    } catch (e) {
      Logger.log('Backfill note parse error row ' + rowIndex + ': ' + e.message);
    }

    const company = (row[config.COL_COMPANY - 1] || '').toString().trim();
    const title   = (row[config.COL_TITLE   - 1] || '').toString().trim();

    if (url) {
      urlCell.setFormula('=HYPERLINK("' + url.replace(/"/g, '') + '", "View Posting")');
      Logger.log('Backfilled URL row ' + rowIndex + ': ' + url);
      fixed++;
    } else if (company && title) {
      const query     = encodeURIComponent('"' + company + '" "' + title + '" careers site:linkedin.com');
      const searchUrl = 'https://www.google.com/search?q=' + query;
      urlCell.setFormula('=HYPERLINK("' + searchUrl + '", "Search Google")');
      Logger.log('Google search fallback row ' + rowIndex);
      fallback++;
    }
  }

  SpreadsheetApp.flush();
  ui.alert(
    'URL Backfill Complete',
    'LinkedIn URLs recovered: ' + fixed + '\n' +
    'Google search fallbacks written: ' + fallback + '\n' +
    'Rows already had URLs (skipped): ' + skipped,
    ui.ButtonSet.OK
  );
}


// ─────────────────────────────────────────────
//  MANUAL ENTRY PROCESSING
// ─────────────────────────────────────────────

function processManualEntry() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig().SHEET_NAME);
  const row   = sheet.getActiveCell().getRow();

  if (row <= 1) {
    ui.alert('Please select a data row (not the header row).');
    return;
  }

  let job = getJobFromRow(row);

  if (!job.url) {
    ui.alert('Missing Job URL', 'Please paste the job posting URL into column H (Job URL) before processing.\n\nThe URL is required to fetch the job description.', ui.ButtonSet.OK);
    return;
  }

  Logger.log('=== processManualEntry: row ' + row + ', URL: ' + job.url + ' ===');

  const activeTabName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  updateJobStatus(row, 'Manual: Fetching description...', activeTabName);
  const jobDescription = fetchJobDescription(job);

  if (!jobDescription || jobDescription.length < 50) {
    updateJobStatus(row, 'Manual: Could not fetch description', activeTabName);
    ui.alert('Description Fetch Failed', 'Could not retrieve the job description from the URL.\n\nThis may be due to a login wall or JavaScript-rendered page.\n\nYou can paste the job description text directly into a Google Doc, save it to Drive, and use "Generate Docs for Selected Row" instead.', ui.ButtonSet.OK);
    return;
  }

  job = enrichJobFromDescription(job, jobDescription);

  const locationFilter = applyLocationFilterPass1(job);
  if (locationFilter.ignore) {
    updateJobStatus(row, 'Ignored: ' + locationFilter.reason);
    ui.alert('Job Filtered Out', 'This posting was filtered out: ' + locationFilter.reason + '\n\nThe row has been marked "Ignored". You can delete it or override by using "Generate Docs for Selected Row" directly.', ui.ButtonSet.OK);
    return;
  }

  job.priority = locationFilter.priority || calculatePriority(job);
  job.rowIndex = row;

  const pass2 = applyLocationFilterPass2(job, jobDescription);
  if (pass2.ignore) {
    updateJobStatus(row, 'Ignored: ' + pass2.reason);
    return;
  }
  if (pass2.priority) job.priority = pass2.priority;

  updateManualEntryRow(row, job, pass2.status || locationFilter.status || 'Manual: Pending Review');

  const response = ui.alert(
    'Manual Entry Ready',
    'Job details confirmed:\nCompany:  ' + (job.company || '(not detected)') + '\nTitle:    ' + (job.title || '(not detected)') + '\nLocation: ' + (job.location || '(not detected)') + '\nRemote:   ' + (job.isRemote ? 'Yes' : 'No') + '\nPriority: P' + job.priority + '\n\nGenerate resume and cover letter now?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    generateDocsForJob(job);
  }
}

function validateManualEntryRow() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig().SHEET_NAME);
  const row   = sheet.getActiveCell().getRow();

  if (row <= 1) {
    ui.alert('Please select a data row (not the header row).');
    return;
  }

  const job    = getJobFromRow(row);
  const issues = [];
  const info   = [];

  if (!job.url)      issues.push('No URL found in column H -- required for description fetch');
  if (!job.company)  info.push('Company: will attempt to detect from page');
  if (!job.title)    info.push('Title: will attempt to detect from page');
  if (!job.location) info.push('Location: will default to "See posting"');

  const remoteVal = sheet.getRange(row, getConfig().COL_REMOTE).getValue();
  if (!remoteVal) info.push('Remote: not set -- defaulting to No');

  let message = '';
  if (issues.length > 0) {
    message += 'ISSUES (must fix before processing):\n';
    message += issues.map(function(i) { return '  - ' + i; }).join('\n') + '\n\n';
  }
  if (info.length > 0) {
    message += 'NOTES (will be handled automatically):\n';
    message += info.map(function(i) { return '  - ' + i; }).join('\n') + '\n\n';
  }
  if (issues.length === 0 && info.length === 0) {
    message = 'Row looks complete. Ready to process.\n\n';
  }

  message += 'URL detected: ' + (job.url || 'none');
  ui.alert('Manual Entry Validation', message, ui.ButtonSet.OK);
}

function enrichJobFromDescription(job, description) {
  const lines = description.split(/[\n\r]+/).map(function(l) { return l.trim(); }).filter(Boolean);

  // ── Title detection ───────────────────────────────────────────
  // For Greenhouse and similar pages, the job title is typically the
  // first meaningful line of the extracted description — a short
  // heading before "Remote - United States" or the company blurb.
  // Falls back to explicit label patterns if heuristic fails.
  if (!job.title) {
    // Heuristic: first line that looks like a job title
    // (not too short, not too long, not a URL, not a location string)
    for (var i = 0; i < Math.min(lines.length, 5); i++) {
      const line = lines[i];
      if (line.length > 5 && line.length < 120 &&
          !line.startsWith('http') &&
          !/^[A-Z][a-z]+,\s*[A-Z]{2}$/.test(line) &&
          !/^remote/i.test(line) &&
          !/^(about|who we|what we|overview|summary)/i.test(line)) {
        job.title = cleanText(line);
        break;
      }
    }
    // Fallback: explicit label patterns
    if (!job.title) {
      const m = description.match(/job title[\s:]+([^\n]+)/i);
      if (m) job.title = cleanText(m[1]);
    }
  }

  // ── Company detection ─────────────────────────────────────────
  // Try to extract from the URL first (most reliable for Greenhouse)
  // boards.greenhouse.io/COMPANY/jobs/ID or job-boards.greenhouse.io/COMPANY/jobs/ID
  if (!job.company && job.url) {
    const ghMatch = job.url.match(/greenhouse\.io\/([^\/\?]+)\/jobs/i);
    if (ghMatch) {
      // Convert slug to readable name: "honeycomb" -> "Honeycomb", "my-company" -> "My Company"
      job.company = ghMatch[1]
        .replace(/-/g, ' ')
        .replace(/\b\w/g, function(c) { return c.toUpperCase(); });
    }
  }
  // Fallback: scan description for "at CompanyName" pattern near the title
  if (!job.company) {
    const m = description.match(/\bat\s+([A-Z][A-Za-z0-9\s&.,]+?)(?:\s*[\n\r|\-]|$)/m);
    if (m && m[1].length < 60) job.company = cleanText(m[1]);
  }
  // Last resort: explicit label
  if (!job.company) {
    const m = description.match(/company[\s:]+([^\n]+)/i) || description.match(/employer[\s:]+([^\n]+)/i);
    if (m) job.company = cleanText(m[1]);
  }

  // ── Location detection ────────────────────────────────────────
  if (!job.location) {
    // Look for "Remote - United States" or "City, ST" patterns near the top
    const m = description.match(/\b(remote\s*[-–]\s*united states|remote\s*[-–]\s*us|remote,?\s*united states)/i) ||
              description.match(/location[\s:]+([^\n]+)/i) ||
              description.match(/\b([A-Z][a-zA-Z\s]+,\s*[A-Z]{2})\b/);
    if (m) job.location = cleanText(m[1] || m[0]);
  }

  // ── Remote flag ───────────────────────────────────────────────
  if (!job.isRemote) {
    job.isRemote = detectRemote(job.location || '', job.title || '', description);
  }

  return job;
}

function updateManualEntryRow(rowIndex, job, status) {
  const config = getConfig();
  const sheet  = getSheet();
  const row    = sheet.getRange(rowIndex, 1, 1, config.TOTAL_COLUMNS).getValues()[0];

  if (!row[config.COL_COMPANY  - 1] && job.company)  sheet.getRange(rowIndex, config.COL_COMPANY).setValue(job.company);
  if (!row[config.COL_TITLE    - 1] && job.title)    sheet.getRange(rowIndex, config.COL_TITLE).setValue(job.title);
  if (!row[config.COL_LOCATION - 1] && job.location) sheet.getRange(rowIndex, config.COL_LOCATION).setValue(job.location);
  if (!row[config.COL_REMOTE   - 1])                 sheet.getRange(rowIndex, config.COL_REMOTE).setValue(job.isRemote ? 'Yes' : 'No');

  sheet.getRange(rowIndex, config.COL_PRIORITY).setValue(job.priority);
  sheet.getRange(rowIndex, config.COL_STATUS).setValue(status);
  if (!job.sheetName) job.sheetName = sheet.getName();

  if (!row[config.COL_SOURCE - 1]) sheet.getRange(rowIndex, config.COL_SOURCE).setValue(job.source || 'Manual');
  if (!row[config.COL_DATE_POSTED - 1]) {
    sheet.getRange(rowIndex, config.COL_DATE_POSTED).setValue(new Date()).setNumberFormat('MM/dd/yyyy');
  }

  const jobMeta = JSON.stringify({
    company:    job.company,
    title:      job.title,
    url:        job.url,
    source:     job.source || 'Manual',
    isRemote:   job.isRemote,
    postedDate: new Date().toISOString(),
    emailBody:  '',
  });
  sheet.getRange(rowIndex, config.COL_COMPANY).setNote(jobMeta);
  SpreadsheetApp.flush();
}


// ─────────────────────────────────────────────
//  EMAIL PARSING
// ─────────────────────────────────────────────

function parseEmailForJobs(message, source) {
  const subject  = message.getSubject();
  const body     = message.getPlainBody();
  const htmlBody = message.getBody();
  const date     = message.getDate();

  Logger.log('Parsing [' + source + '] email: "' + subject + '"');

  let jobs = [];

  try {
    switch (source) {
      case 'LinkedIn':
        jobs = parseLinkedInEmail(subject, body, htmlBody, date);
        break;
      case 'Indeed':
        jobs = parseIndeedEmail(subject, body, htmlBody, date);
        break;
      case 'Glassdoor':
        if (/apply now|recently viewed|pick up your job search|jobs you might like/i.test(subject + ' ' + body.substring(0, 200))) {
          jobs = parseGlassdoorApplyNowEmail(subject, body, htmlBody, date);
        } else {
          jobs = parseGlassdoorEmail(subject, body, htmlBody, date);
        }
        if (jobs.length === 0) {
          jobs = /apply now|recently viewed/i.test(subject)
            ? parseGlassdoorEmail(subject, body, htmlBody, date)
            : parseGlassdoorApplyNowEmail(subject, body, htmlBody, date);
        }
        break;
    }
  } catch (e) {
    Logger.log('Parser error for ' + source + ': ' + e.message + '. Falling back to Gemini.');
  }

  if (jobs.length === 0) {
    Logger.log('Falling back to Gemini parser for: "' + subject + '"');
    jobs = parseEmailWithGemini(subject, body, source, date);
  }

  return jobs;
}


// ── LinkedIn parser ──────────────────────────
//
// Confirmed format:
//   Company Name [TAB]
//   Job Title
//   Company Name · Location
//   [noise lines: alumni counts, "Actively recruiting", etc.]
//   URLs appear in HTML body as href values

function parseLinkedInEmail(subject, body, htmlBody, date) {
  const jobs = [];

  const noisePatterns = [
    /^\d+\s+school alumni/i,
    /^\d+\s+company alum/i,
    /^actively recruiting/i,
    /^see all jobs/i,
    /^premium icon/i,
    /^your job alert for/i,
    /^view all jobs/i,
    /^manage alert/i,
    /^unsubscribe/i,
    /^sign in/i,
    /^just now$/i,
    /^\d+[mhd]$/i,
    /^\d+ (minute|hour|day)/i,
  ];

  function isNoiseLine(line) {
    return noisePatterns.some(function(p) { return p.test(line.trim()); });
  }

  // Extract LinkedIn job URLs from both HTML and plain text bodies
  // Matches both /jobs/view/ and /comm/jobs/view/ formats
  const urlPattern = /https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/[^"'\s>)]+/g;
  const htmlUrls = [];
  let urlMatch;
  while ((urlMatch = urlPattern.exec(htmlBody || '')) !== null) {
    const url = urlMatch[0].replace(/&amp;/g, '&').split('?')[0];
    if (!htmlUrls.includes(url)) htmlUrls.push(url);
  }
  // Also scan plain text for "View job:" links
  urlPattern.lastIndex = 0;
  while ((urlMatch = urlPattern.exec(body || '')) !== null) {
    const url = urlMatch[0].replace(/&amp;/g, '&').split('?')[0];
    if (!htmlUrls.includes(url)) htmlUrls.push(url);
  }

  const lines = body.split('\n').map(function(l) { return l.trim(); });
  let urlIndex = 0;

  for (let i = 0; i < lines.length - 1; i++) {
    const line     = lines[i];
    const nextLine = lines[i + 1] || '';
    const afterNext = lines[i + 2] || '';

    if (!line || isNoiseLine(line)) continue;

    const locationLine    = afterNext.includes(' · ') || afterNext.includes(' - ');
    const looksLikeCompany = line.length > 1 && line.length < 80 && !line.includes(' · ') && !isNoiseLine(line);
    const looksLikeTitle   = nextLine.length > 3 && nextLine.length < 120 && !nextLine.includes(' · ') && !isNoiseLine(nextLine);

    if (looksLikeCompany && looksLikeTitle && locationLine) {
      const company  = cleanText(line);
      const title    = cleanText(nextLine);
      const locRaw   = afterNext.includes(' · ')
                       ? afterNext.split(' · ').slice(1).join(' · ').trim()
                       : afterNext.split(' - ').slice(1).join(' - ').trim();
      const location = cleanText(locRaw) || 'See posting';

      const snippet = lines.slice(Math.max(0, i - 1), i + 5).join(' ').toLowerCase();
      if (snippet.includes('repost') || snippet.includes('re-post')) {
        Logger.log('  Skipping LinkedIn repost (email snippet): ' + title + ' at ' + company);
        i += 2;
        continue;
      }

      const url      = htmlUrls[urlIndex] || '';
      if (url) urlIndex++;

      const isRemote   = detectRemote(location, title, body);
      const postedDate = extractLinkedInPostDate(body, i) || date;

      if (company && title) {
        jobs.push({
          source:     'LinkedIn',
          company:    company,
          title:      title,
          location:   location,
          url:        url,
          isRemote:   isRemote,
          postedDate: postedDate,
          isRepost:   false,
          emailBody:  body,
          subject:    subject,
        });
      }
      i += 2;
    }
  }

  if (jobs.length === 0) {
    const singleJob = parseLinkedInSingleJobEmail(subject, body, htmlBody, date);
    if (singleJob) jobs.push(singleJob);
  }

  Logger.log('  LinkedIn: parsed ' + jobs.length + ' job(s) from email');
  return jobs;
}

function parseLinkedInSingleJobEmail(subject, body, htmlBody, date) {
  // LinkedIn subject formats:
  //   "Job alert: Senior Engineer at Acme Corp"
  //   "Senior Quality Engineer at Ladders: up to $175K/year"
  //   "New job for you: Senior Engineer at Acme"

  let title = '', company = '';

  const fmt1 = subject.match(/job alert[:\s]+(.+?)\s+at\s+([^:]+)/i);
  const fmt2 = subject.match(/^(.+?)\s+at\s+([^:]+):/i);
  const fmt3 = subject.match(/new job.*?:\s*(.+?)\s+at\s+(.+)/i);
  const fmt4 = subject.match(/^(.+?)\s+at\s+(.+)$/i);

  const match = fmt1 || fmt3 || fmt2 || fmt4;
  if (!match) return null;

  title   = cleanText(match[1]);
  company = cleanText(match[2]);

  if (!company || /^\$|^up to|^\d/.test(company)) return null;

  // Match both /jobs/view/ and /comm/jobs/view/ and View job: links
  const urlMatch = body.match(/https?:\/\/(?:www\.)?linkedin\.com\/(?:comm\/)?jobs\/view\/[^\s\n?&]+/) ||
                   body.match(/View job:\s*(https?:\/\/[^\s\n]+)/i);
  const url = urlMatch ? urlMatch[0].trim() : '';

  const isRepost = /repost/i.test(body);
  const isRemote = detectRemote('', title, body);
  const locationMatch = body.match(/\b([A-Z][a-zA-Z\s]+,\s*[A-Z]{2})\b/);
  const location = locationMatch ? locationMatch[1] : 'See posting';

  return { source: 'LinkedIn', company, title, location, url, isRemote, isRepost, postedDate: date, emailBody: body, subject };
}


// ── Indeed parser ────────────────────────────

function parseIndeedEmail(subject, body, htmlBody, date) {
  const jobs = [];
  const sections = body.split(/\n{2,}/);

  for (const section of sections) {
    const urlMatch = section.match(/https?:\/\/(?:www\.)?indeed\.com\/(?:viewjob|rc\/clk)[^\s\n"<]+/i);
    if (!urlMatch) continue;

    const url   = urlMatch[0].trim();
    const lines = section.replace(url, '').split('\n').map(function(l) { return l.trim(); }).filter(Boolean);
    if (lines.length < 2) continue;

    const title    = cleanText(lines[0]);
    const company  = cleanText(lines[1]);
    const location = lines[2] ? cleanText(lines[2]) : 'See posting';
    const isRemote = detectRemote(location, title, section);

    jobs.push({
      source:     'Indeed',
      company, title, location, url,
      isRemote,
      isRepost:   false,
      postedDate: extractIndeedPostDate(section) || date,
      emailBody:  section,
      subject,
    });
  }

  Logger.log('  Indeed: parsed ' + jobs.length + ' job(s) from email');
  return jobs;
}


// ── Glassdoor daily digest parser ────────────
//
// Confirmed format:
//   Job Title              (no leading tab)
//   [TAB]City, ST          (tab-prefixed location)
//   avatar                 (noise)
//   [TAB]Company X.X star  (tab-prefixed company with optional rating)

function parseGlassdoorEmail(subject, body, htmlBody, date) {
  const jobs = [];

  const urlPattern = /https?:\/\/(?:www\.)?glassdoor\.com\/[^"'\s>]{10,}/g;
  const htmlUrls   = [];
  let um;
  while ((um = urlPattern.exec(htmlBody || '')) !== null) {
    const u = um[0].replace(/&amp;/g, '&').split('?')[0];
    if (!htmlUrls.includes(u) && u.length < 300 && !u.includes('unsubscribe') && !u.includes('manage') && !u.includes('privacy')) {
      htmlUrls.push(u);
    }
  }

  function isGlassdoorNoise(line) {
    const t = line.trim();
    if (!t) return true;
    if (/^\d{1,2}d$/.test(t)) return true;
    if (/^\$/.test(t)) return true;
    if (/^\(employer est/i.test(t)) return true;
    if (/^\(glassdoor est/i.test(t)) return true;
    if (/^easy apply$/i.test(t)) return true;
    if (/^avatar$/i.test(t)) return true;
    if (/^see more jobs/i.test(t)) return true;
    if (/^want more/i.test(t)) return true;
    if (/^create job alert/i.test(t)) return true;
    if (/^similar jobs/i.test(t)) return true;
    if (/^privacy policy/i.test(t)) return true;
    if (/^manage settings/i.test(t)) return true;
    if (/^you can edit/i.test(t)) return true;
    if (/^sent (daily|weekly)/i.test(t)) return true;
    if (/^edit$/.test(t)) return true;
    if (/^create$/.test(t)) return true;
    if (/^your job (listings|alert)/i.test(t)) return true;
    if (/^job alert:/i.test(t)) return true;
    if (/^looking for something/i.test(t)) return true;
    if (t === '\t' || /^\t+$/.test(t)) return true;
    return false;
  }

  function stripStarRating(text) {
    return text.replace(/\s+[\d.]+\s*\u2605.*$/, '').replace(/\t/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function isTabLine(line) { return line.startsWith('\t') && line.trim().length > 0; }

  function isLocationString(text) {
    const t = text.trim();
    return /^[A-Z][a-zA-Z\s]+,\s*[A-Z]{2}/.test(t) || /^United States/i.test(t) || /^Remote/i.test(t);
  }

  const lines = body.split('\n');
  let urlIndex = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const t    = line.trim();

    if (!line.startsWith('\t') && !isGlassdoorNoise(line) && t.length > 3 && t.length < 120) {
      const title = t;
      let location = 'See posting', company = '', postAge = '';

      for (let j = i + 1; j < Math.min(i + 8, lines.length); j++) {
        const jLine = lines[j];
        const jt    = jLine.trim();
        if (!jt || isGlassdoorNoise(jLine)) continue;

        if (!company && !location.match(/[A-Z]/) && isLocationString(jt)) { location = jt; continue; }
        if (!company && isTabLine(jLine) && isLocationString(jt)) { location = jt; continue; }
        if (isTabLine(jLine) && !isGlassdoorNoise(jLine) && !isLocationString(jt) && jt.length > 1) {
          company = stripStarRating(jt); continue;
        }
        if (/^\d{1,2}d$/.test(jt)) { postAge = jt; continue; }
        if (!jLine.startsWith('\t') && !isGlassdoorNoise(jLine) && jt.length > 3 && company) break;
      }

      if (title && company) {
        let postDate = date;
        if (postAge) {
          const days = parseInt(postAge);
          if (!isNaN(days)) postDate = new Date(date.getTime() - days * 24 * 60 * 60 * 1000);
        }
        const isRemote = detectRemote(location, title, body);
        const url      = htmlUrls[urlIndex] || '';
        if (url) urlIndex++;

        jobs.push({ source: 'Glassdoor', company, title, location, url, isRemote, isRepost: false, postedDate: postDate, emailBody: body, subject });
      }
    }
  }

  Logger.log('  Glassdoor: parsed ' + jobs.length + ' job(s) from email');
  return jobs;
}


// ── Glassdoor "Apply Now / Recently Viewed" parser ───────────
//
// Subject: "Company1,Company2: Apply Now"
// This email type is HTML-only — getPlainBody() returns nothing useful.
// All data is parsed directly from the HTML body using regex patterns
// that match Glassdoor's React email component structure:
//
//   Company: <span style="...display:inline">Company Name</span>
//   Title:   <p style="...font-weight:600">Job Title</p>
//   Location:<p style="...margin-top:4px">City, ST</p>
//   URL:     <a href="https://www.glassdoor.com/partner/jobListing.htm?...">

function parseGlassdoorApplyNowEmail(subject, body, htmlBody, date) {
  const jobs = [];
  const html = htmlBody || body || '';

  if (!html || html.length < 100) {
    Logger.log('  Glassdoor (Apply Now): no HTML body available');
    return jobs;
  }

  // Parse job data at the card level by splitting on each job listing anchor tag.
  // Each <a href="...partner/jobListing.htm..."> block wraps one complete job card
  // containing the company name, title, location, and URL for that listing.
  // This avoids the array misalignment problem that occurs when extracting
  // companies, titles, and locations as three separate global lists.

  // Split HTML into individual job card blocks using the jobListing.htm anchor as delimiter
  const cardSplitPattern = /(<a [^>]*partner\/jobListing\.htm[^>]*>[\s\S]*?<\/a>)/g;
  const cards = [];
  let cardMatch;
  while ((cardMatch = cardSplitPattern.exec(html)) !== null) {
    cards.push(cardMatch[1]);
  }

  Logger.log('  Glassdoor (Apply Now): found ' + cards.length + ' job card(s) in HTML');

  for (let i = 0; i < cards.length; i++) {
    const card = cards[i];

    // Extract URL from the opening anchor tag
    const urlMatch = card.match(/href="(https?:\/\/[^"]+partner\/jobListing\.htm[^"]+)"/);
    const url = urlMatch ? urlMatch[1].replace(/&amp;/g, '&') : '';

    // Extract company name: display:inline span (Glassdoor's company name element)
    const companyMatch = card.match(/display:inline[^"]*"[^>]*>([^<]{2,80})<\/span>/);
    const company = companyMatch ? companyMatch[1].trim() : '';

    // Extract job title: font-weight:600 paragraph
    const titleMatch = card.match(/font-weight:600[^"]*"[^>]*>([^<]{3,150})<\/p>/);
    const title = titleMatch ? titleMatch[1].trim() : '';

    // Extract location: first margin-top:4px paragraph that isn't a salary range
    const locationMatches = card.match(/margin-top:4px[^"]*"[^>]*>([^<]{2,60})<\/p>/g) || [];
    let location = 'See posting';
    for (const lm of locationMatches) {
      const locText = lm.replace(/<[^>]+>/g, '').replace(/margin-top:4px[^>]*>/g, '').trim();
      // Skip salary ranges and parenthetical notes
      if (locText && !locText.startsWith('$') && !locText.startsWith('(') && locText.length < 50) {
        location = locText;
        break;
      }
    }

    // Skip cards where we couldn't extract both company and title
    if (!company || !title) {
      Logger.log('  Card ' + (i+1) + ': skipped (company="' + company + '" title="' + title + '")');
      continue;
    }

    const isRemote = detectRemote(location, title, '');

    Logger.log('  Card ' + (i+1) + ': "' + title + '" at "' + company + '" (' + location + ')');

    jobs.push({
      source:     'Glassdoor',
      company:    cleanText(company),
      title:      cleanText(title),
      location:   cleanText(location),
      url:        url,
      isRemote:   isRemote,
      isRepost:   false,
      postedDate: date,
      emailBody:  '',
      subject:    subject,
    });
  }

  Logger.log('  Glassdoor (Apply Now): parsed ' + jobs.length + ' job(s) from HTML');
  return jobs;
}


// ─────────────────────────────────────────────
//  JOB DESCRIPTION FETCHING
// ─────────────────────────────────────────────

function fetchJobDescription(job) {
  const config     = getConfig();
  const SUFFICIENT = config.MIN_JD_CHARS || 2500;
  const MINIMUM    = 200;
  let best = '';

  // S1: Email body excerpt
  if (job.emailBody && job.emailBody.length > 300) {
    const extracted = extractDescriptionFromEmailBody(job.emailBody, job.title, job.company);
    if (extracted && extracted.length > MINIMUM) {
      Logger.log('  [S1] Email body excerpt: ' + extracted.length + ' chars ' + (extracted.length >= SUFFICIENT ? '(sufficient)' : '(thin — continuing to S2)'));
      if (extracted.length >= SUFFICIENT) return extracted;
      if (extracted.length > best.length) best = extracted;
    }
  }

  // S2: Direct URL fetch
  if (job.url) {
    try {
      Logger.log('  [S2] Attempting direct URL fetch: ' + job.url);
      const response = UrlFetchApp.fetch(job.url, {
        muteHttpExceptions: true,
        followRedirects:    true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
      });

      const statusCode = response.getResponseCode();
      const html       = response.getContentText();
      const isLinkedInBlock = (job.source === 'LinkedIn' && (statusCode !== 200 || (html.toLowerCase().includes('sign in') && html.toLowerCase().includes('linkedin.com/login'))));

      if (isLinkedInBlock) {
        Logger.log('  [S2] LinkedIn login wall — skipping to S3');
      } else if (statusCode === 200) {
        const description = extractDescriptionFromHtml(html, job.source);
        if (description && description.length > MINIMUM) {
          Logger.log('  [S2] Direct URL fetch: ' + description.length + ' chars ' + (description.length >= SUFFICIENT ? '(sufficient)' : '(thin — continuing to S3)'));
          if (description.length >= SUFFICIENT) return description;
          if (description.length > best.length) best = description;
        } else {
          Logger.log('  [S2] URL fetched but description too short — trying S3');
        }
      } else {
        Logger.log('  [S2] URL fetch returned ' + statusCode + ' — trying S3');
      }
    } catch (e) {
      Logger.log('  [S2] URL fetch error: ' + e.message + ' — trying S3');
    }
  }

  // S3: DuckDuckGo career page search
  Logger.log('  [S3] DuckDuckGo career page search: ' + job.title + ' at ' + job.company);
  const careerDescription = fetchDescriptionViaDuckDuckGo(job.title, job.company);
  if (careerDescription && careerDescription.length > MINIMUM) {
    Logger.log('  [S3] DuckDuckGo: ' + careerDescription.length + ' chars ' + (careerDescription.length >= SUFFICIENT ? '(sufficient)' : '(thin — falling back to S4)'));
    if (careerDescription.length >= SUFFICIENT) return careerDescription;
    if (careerDescription.length > best.length) best = careerDescription;
  } else {
    Logger.log('  [S3] DuckDuckGo returned insufficient content — falling back to S4');
  }

  // S4: Full email body fallback
  const s4 = job.emailBody || ('Job Title: ' + job.title + '\nCompany: ' + job.company + '\nLocation: ' + job.location);
  if (s4.length > best.length) best = s4;

  Logger.log('  [S4] Final result: ' + best.length + ' chars ' + (best.length >= SUFFICIENT ? '' : '(THIN JD — will flag for review)'));
  return best;
}


// ─────────────────────────────────────────────
//  DUCKDUCKGO CAREER PAGE SEARCH
// ─────────────────────────────────────────────

function fetchDescriptionViaDuckDuckGo(title, company) {
  const queries = [
    '"' + title + '" "' + company + '" job posting careers',
    title + ' ' + company + ' careers jobs',
    company + ' careers ' + title,
  ];

  for (const query of queries) {
    try {
      Logger.log('    DDG query: "' + query + '"');
      const rawUrls      = searchDuckDuckGoForUrls(query);
      const candidateUrls = extractDDGCandidateUrls({ _rawUrls: rawUrls }, company);

      if (candidateUrls.length === 0) { Logger.log('    DDG: no candidate URLs for this query'); continue; }
      Logger.log('    DDG: found ' + candidateUrls.length + ' candidate URL(s)');

      for (const url of candidateUrls) {
        Logger.log('    Fetching career page: ' + url);
        const description = fetchAndExtractFromCareerPage(url, title);
        if (description && description.length > 200) {
          Logger.log('    Career page fetch succeeded (' + description.length + ' chars): ' + url);
          return description;
        }
      }
    } catch (e) {
      Logger.log('    DDG error for query "' + query + '": ' + e.message);
    }
  }
  return null;
}

function extractDDGCandidateUrls(data, company) {
  const blockedDomains = ['linkedin.com','indeed.com','glassdoor.com','ziprecruiter.com','monster.com','careerbuilder.com','simplyhired.com','dice.com','builtin.com','workday.com','icims.com','taleo.net','myworkdayjobs.com'];
  const careerKeywords = ['career','job','hiring','position','opening','apply','talent'];

  const rawUrls = data._rawUrls || [];

  const scored = rawUrls
    .filter(function(url) {
      if (!url || !url.startsWith('http')) return false;
      const lower = url.toLowerCase();
      return !blockedDomains.some(function(d) { return lower.includes(d); });
    })
    .map(function(url) {
      const lower = url.toLowerCase();
      let score = 0;
      if (careerKeywords.some(function(kw) { return lower.includes(kw); })) score += 3;
      const companySlug = company.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (companySlug.length > 3 && lower.includes(companySlug)) score += 5;
      if (url.length > 200) score -= 1;
      return { url: url, score: score };
    })
    .sort(function(a, b) { return b.score - a.score; })
    .map(function(item) { return item.url; })
    .slice(0, 4);

  return [...new Set(scored)];
}

function fetchAndExtractFromCareerPage(url, jobTitle) {
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects:    true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });

    if (response.getResponseCode() !== 200) return null;

    const description = extractDescriptionFromHtml(response.getContentText(), 'career_page');
    const titleWords   = jobTitle.toLowerCase().split(/\s+/).filter(function(w) { return w.length > 3; });
    const descLower    = description.toLowerCase();
    const titleMatches = titleWords.filter(function(word) { return descLower.includes(word); }).length;
    const matchRatio   = titleWords.length > 0 ? titleMatches / titleWords.length : 0;

    if (matchRatio < 0.5) {
      Logger.log('    Career page relevance check failed (' + Math.round(matchRatio * 100) + '% title match) — skipping');
      return null;
    }
    return description;
  } catch (e) {
    Logger.log('    Career page fetch error for ' + url + ': ' + e.message);
    return null;
  }
}

function searchDuckDuckGoForUrls(query) {
  try {
    const encodedQuery = encodeURIComponent(query);
    const ddgUrl = 'https://api.duckduckgo.com/?q=' + encodedQuery + '&format=json&no_redirect=1&no_html=1&skip_disambig=1';

    const response = UrlFetchApp.fetch(ddgUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
    });

    if (response.getResponseCode() !== 200) return [];

    const data = JSON.parse(response.getContentText());
    const urls = [];

    if (data.AbstractURL) urls.push(data.AbstractURL);
    if (Array.isArray(data.RelatedTopics)) {
      for (const topic of data.RelatedTopics) {
        if (topic.FirstURL) urls.push(topic.FirstURL);
        if (Array.isArray(topic.Topics)) {
          for (const sub of topic.Topics) { if (sub.FirstURL) urls.push(sub.FirstURL); }
        }
      }
    }
    if (Array.isArray(data.Results)) {
      for (const result of data.Results) { if (result.FirstURL) urls.push(result.FirstURL); }
    }

    return [...new Set(urls)].filter(function(u) { return u && u.startsWith('http'); });
  } catch (e) {
    Logger.log('  searchDuckDuckGoForUrls error: ' + e.message);
    return [];
  }
}


// ─────────────────────────────────────────────
//  HTML / EMAIL CONTENT EXTRACTION
// ─────────────────────────────────────────────

function extractDescriptionFromEmailBody(body, title, company) {
  let text = body;
  text = text.replace(/unsubscribe|manage.*alert|view.*browser|privacy policy/gi, '');
  text = text.replace(/https?:\/\/\S+/g, '');
  text = text.replace(/\n{3,}/g, '\n\n');
  const titleIndex = text.toLowerCase().indexOf(title.toLowerCase().substring(0, 20));
  if (titleIndex > -1) text = text.substring(titleIndex);
  return text.trim().substring(0, 8000);
}

function extractDescriptionFromHtml(html, source) {
  let text = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<[^>]+>/g, ' ');
  text = text.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  text = text.replace(/\s{2,}/g, ' ').replace(/\n{3,}/g, '\n\n');

  // ── Step 1: Find where the job description starts ─────────────
  // Scan for common section opening markers and slice from there
  const startMarkers = [
    'job description','about the role','about this role','about this job',
    'the role','what you will do','responsibilities','qualifications',
    'requirements','what we are looking for','about the position',
    "what we're building", 'who we are', 'about us', 'overview',
    'the opportunity', 'position summary', 'role summary',
  ];
  const lower = text.toLowerCase();
  let startIdx = -1;
  for (const marker of startMarkers) {
    const idx = lower.indexOf(marker);
    if (idx > -1 && (startIdx === -1 || idx < startIdx)) {
      startIdx = idx;
    }
  }
  if (startIdx > -1) text = text.substring(startIdx);

  // ── Step 2: Find where the job description ends ───────────────
  // Application forms, privacy notices, and job alert CTAs always
  // appear after the real job content on Greenhouse and similar pages.
  // Truncate at the first termination marker found.
  const endMarkers = [
    'apply for this job',
    'create a job alert',
    'indicates a required field',
    'quick apply',
    'first name',
    'last name',
    'attach resume',
    'resume/cv',
    'privacy notice',
    'equal opportunity employer',
    'voluntary self-identification',
    'we are an equal',
    'eeo statement',
    'applicant information',
    'diversity & accommodations',
    'phishing and recruitment scam',
  ];
  const lowerText = text.toLowerCase();
  let endIdx = text.length;
  for (const marker of endMarkers) {
    const idx = lowerText.indexOf(marker);
    if (idx > -1 && idx < endIdx) {
      endIdx = idx;
    }
  }
  text = text.substring(0, endIdx).trim();

  return text.substring(0, 8000);
}


// ─────────────────────────────────────────────
//  PRIORITY CALCULATION (TWO-PASS)
// ─────────────────────────────────────────────

function calculatePriority(job) {
  const now      = new Date();
  const posted   = job.postedDate instanceof Date ? job.postedDate : new Date(job.postedDate);
  const ageHours = (now - posted) / (1000 * 60 * 60);

  if (job.isRemote && ageHours < 2)  return 1;
  if (job.isRemote && ageHours < 24) return 2;
  if (job.isRemote)                  return 3;
  return 15; // on-site that passed 50-mile check
}


// ─────────────────────────────────────────────
//  LOCATION FILTERING — PASS 1
// ─────────────────────────────────────────────

function applyLocationFilterPass1(job) {
  const location = (job.location || '').toLowerCase();
  const config   = getConfig();

  if (job.isRemote) {
    if (detectIndia(location)) return { ignore: true, reason: 'remote India origin' };
    if (detectNonUsaRemote(location)) return { ignore: false, priority: 10, status: 'Review: non-USA remote origin' };
    return { ignore: false };
  }

  const coords = geocodeLocationString(location);
  if (!coords) return { ignore: false, priority: 15, status: 'Review: on-site location unresolved' };

  const distanceMiles = haversineDistance(config.HOME_ZIP_LAT, config.HOME_ZIP_LNG, coords.lat, coords.lng);
  Logger.log('  On-site distance check: ' + job.location + ' = ' + Math.round(distanceMiles) + ' miles from ' + config.HOME_ZIP);

  if (distanceMiles <= 50) return { ignore: false, priority: 15, status: 'Review: on-site within 50 miles' };
  return { ignore: true, reason: 'on-site outside 50-mile radius' };
}


// ─────────────────────────────────────────────
//  LOCATION FILTERING — PASS 2
// ─────────────────────────────────────────────

function applyLocationFilterPass2(job, jobDescription) {
  if (!job.isRemote || job.priority === 10 || job.priority === 15) return { ignore: false, status: null };

  const jdLower = (jobDescription || '').toLowerCase();
  const hasStateRestriction = detectStateRestrictionLanguage(jdLower);

  if (!hasStateRestriction) {
    Logger.log('  Pass-2: no state restriction detected — proceeding');
    return { ignore: false, status: null };
  }

  Logger.log('  Pass-2: state restriction language detected — checking for Texas');
  const texasListed = detectTexasInJd(jdLower);

  if (texasListed) {
    Logger.log('  Pass-2: Texas found in approved states — proceeding normally');
    return { ignore: false, status: null };
  }

  Logger.log('  Pass-2: Texas NOT in approved states — flagging P5');
  return { ignore: false, priority: 5, status: 'Review: TX not in approved states (P5)' };
}

function detectStateRestrictionLanguage(jdLower) {
  const restrictionPhrases = ['only available in','only open to residents','must reside in','must be located in','limited to the following states','available in the following states','eligible states','approved states','this position is available in','this role is available in','applicants must live in','residents of the following','not available in all states','where permitted by law'];
  return restrictionPhrases.some(function(phrase) { return jdLower.includes(phrase); });
}

function detectTexasInJd(jdLower) {
  return /\btexas\b/.test(jdLower) || /\btx\b/.test(jdLower) || /\b(tx)[,\s)]/.test(jdLower);
}

function detectIndia(locationLower) {
  return locationLower.includes('india') || locationLower.includes('bengaluru') || locationLower.includes('bangalore') || locationLower.includes('hyderabad') || locationLower.includes('mumbai') || locationLower.includes('chennai') || locationLower.includes('pune');
}

function detectNonUsaRemote(locationLower) {
  const nonUsaSignals = ['canada','united kingdom','uk','australia','germany','france','netherlands','spain','brazil','mexico','singapore','philippines','colombia','argentina','poland','ukraine','romania','portugal','ireland',' ca)',' uk)',' au)',' de)',' nl)'];
  return nonUsaSignals.some(function(signal) { return locationLower.includes(signal); });
}


// ─────────────────────────────────────────────
//  ON-SITE DISTANCE / GEOCODING
// ─────────────────────────────────────────────

function geocodeLocationString(locationLower) {
  const dfwLookup = {
    'dallas':{ lat:32.7767,lng:-96.7970 },'fort worth':{ lat:32.7555,lng:-97.3308 },'arlington':{ lat:32.7357,lng:-97.1081 },
    'plano':{ lat:33.0198,lng:-96.6989 },'garland':{ lat:32.9126,lng:-96.6389 },'irving':{ lat:32.8140,lng:-96.9489 },
    'frisco':{ lat:33.1507,lng:-96.8236 },'mckinney':{ lat:33.1972,lng:-96.6397 },'mesquite':{ lat:32.7668,lng:-96.5992 },
    'carrollton':{ lat:32.9537,lng:-96.8903 },'denton':{ lat:33.2148,lng:-97.1331 },'richardson':{ lat:32.9483,lng:-96.7299 },
    'lewisville':{ lat:33.0462,lng:-96.9942 },'allen':{ lat:33.1032,lng:-96.6706 },'grand prairie':{ lat:32.7460,lng:-96.9978 },
    'flower mound':{ lat:33.0146,lng:-97.0969 },'rowlett':{ lat:32.9029,lng:-96.5638 },'wylie':{ lat:33.0151,lng:-96.5388 },
    'mansfield':{ lat:32.5632,lng:-97.1417 },'cedar hill':{ lat:32.5882,lng:-96.9561 },'duncanville':{ lat:32.6518,lng:-96.9083 },
    'desoto':{ lat:32.5896,lng:-96.8572 },'lancaster':{ lat:32.5921,lng:-96.7561 },'euless':{ lat:32.8371,lng:-97.0819 },
    'bedford':{ lat:32.8440,lng:-97.1430 },'hurst':{ lat:32.8232,lng:-97.1883 },'grapevine':{ lat:32.9343,lng:-97.0781 },
    'southlake':{ lat:32.9412,lng:-97.1344 },'keller':{ lat:32.9343,lng:-97.2294 },'north richland hills':{ lat:32.8343,lng:-97.2289 },
    'addison':{ lat:32.9612,lng:-96.8289 },'farmers branch':{ lat:32.9268,lng:-96.8958 },'coppell':{ lat:32.9543,lng:-97.0150 },
    'rockwall':{ lat:32.9290,lng:-96.4597 },'sachse':{ lat:32.9762,lng:-96.5766 },'burleson':{ lat:32.5421,lng:-97.3208 },
    'waxahachie':{ lat:32.3868,lng:-96.8489 },'weatherford':{ lat:32.7596,lng:-97.7975 },'forney':{ lat:32.7482,lng:-96.4697 },
    'haltom city':{ lat:32.7893,lng:-97.2697 },'euless':{ lat:32.8371,lng:-97.0819 },
  };

  for (const city in dfwLookup) {
    if (locationLower.includes(city)) return dfwLookup[city];
  }

  const zipMatch = locationLower.match(/\b(7[5-6]\d{3})\b/);
  if (zipMatch) {
    const prefix = zipMatch[1].substring(0, 3);
    const zipPrefixMap = { '750':{ lat:32.7767,lng:-96.7970 },'751':{ lat:32.7767,lng:-96.7970 },'752':{ lat:32.7767,lng:-96.7970 },'753':{ lat:33.0198,lng:-96.6989 },'754':{ lat:32.9126,lng:-96.6389 },'755':{ lat:33.1507,lng:-96.8236 },'756':{ lat:32.5896,lng:-96.8572 },'757':{ lat:32.7460,lng:-96.9978 },'760':{ lat:32.7555,lng:-97.3308 },'761':{ lat:32.7555,lng:-97.3308 },'762':{ lat:32.8343,lng:-97.2289 },'763':{ lat:33.2148,lng:-97.1331 },'764':{ lat:32.3499,lng:-97.3864 },'765':{ lat:32.5421,lng:-97.3208 },'766':{ lat:32.7596,lng:-97.7975 } };
    if (zipPrefixMap[prefix]) return zipPrefixMap[prefix];
  }

  if (locationLower.includes(', tx') || locationLower.includes(' texas')) return { lat:32.7767, lng:-96.7970 };
  return null;
}

function haversineDistance(lat1, lng1, lat2, lng2) {
  const R    = 3958.8;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a    = Math.sin(dLat/2) * Math.sin(dLat/2) + Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) * Math.sin(dLng/2) * Math.sin(dLng/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}


// ─────────────────────────────────────────────
//  LINKEDIN REPOST PAGE CHECK
// ─────────────────────────────────────────────

function isLinkedInRepost(url) {
  if (!url) return false;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' } });
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      Logger.log('  Repost check: LinkedIn returned ' + statusCode + ' — unverified');
      return 'unverified';
    }

    const html = response.getContentText().toLowerCase();
    if (html.includes('sign in') && html.includes('linkedin.com/login')) {
      Logger.log('  Repost check: LinkedIn login wall detected — unverified');
      return 'unverified';
    }

    const isRepost = html.includes('"reposted"') || html.includes('>reposted<') || html.includes('reposted job') || html.includes('"jobposting-repost"') || /reposted\s*[·•]/.test(html) || /[·•]\s*reposted/.test(html) || /class="[^"]*repost[^"]*"/.test(html);
    Logger.log('  Repost check: ' + (isRepost ? 'REPOST detected' : 'not a repost') + ' — ' + url);
    return isRepost;
  } catch (e) {
    Logger.log('  Repost check: fetch error — unverified: ' + e.message);
    return 'unverified';
  }
}


// ─────────────────────────────────────────────
//  AUTOMATION TITLE FILTER
// ─────────────────────────────────────────────

function isAutomationHeavyTitle(title) {
  if (!title) return false;
  const config   = getConfig();
  const keywords = config.AUTOMATION_TITLE_KEYWORDS || [];
  const lower    = title.toLowerCase();
  for (const keyword of keywords) {
    if (lower.includes(keyword.toLowerCase())) return keyword;
  }
  return false;
}


// ─────────────────────────────────────────────
//  HELPER UTILITIES
// ─────────────────────────────────────────────

function detectRemote(location, title, body) {
  const combined = ((location || '') + ' ' + (title || '') + ' ' + (body || '')).toLowerCase();
  return /\bremote\b/.test(combined);
}

function cleanText(text) {
  if (!text) return '';
  return text.replace(/\s+/g, ' ').replace(/[^\w\s\-.,&()]/g, '').trim();
}

function sanitizeFolderName(name) {
  return name.replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim().substring(0, 100);
}

function extractLinkedInPostDate(body, matchIndex) {
  const snippet = typeof body === 'string' ? body.substring(0, 2000) : '';
  return parseDateFromAgeText(snippet);
}

function extractIndeedPostDate(section) { return parseDateFromAgeText(section); }
function extractGlassdoorPostDate(section) { return parseDateFromAgeText(section); }

function parseDateFromAgeText(text) {
  const now     = new Date();
  const justNow = text.match(/just\s+(?:now|posted)/i);
  const minutes = text.match(/(\d+)\s*(?:minute|min)s?\s+ago/i);
  const hours   = text.match(/(\d+)\s*(?:hour|hr)s?\s+ago/i);
  const days    = text.match(/(\d+)\s*days?\s+ago/i);

  if (justNow)  return new Date(now - 5 * 60 * 1000);
  if (minutes)  return new Date(now - parseInt(minutes[1]) * 60 * 1000);
  if (hours)    return new Date(now - parseInt(hours[1]) * 60 * 60 * 1000);
  if (days)     return new Date(now - parseInt(days[1]) * 24 * 60 * 60 * 1000);
  return null;
}

function loadTextFileFromDrive(fileId) {
  if (!fileId) return '';
  try {
    const file     = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    if (mimeType === 'application/vnd.google-apps.document') {
      return DocumentApp.openById(fileId).getBody().getText();
    }
    return file.getBlob().getDataAsString();
  } catch (e) {
    Logger.log('Could not load file ' + fileId + ': ' + e.message);
    return '';
  }
}

function generateDocsForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig().SHEET_NAME);
  const row   = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('Please select a data row (not the header).'); return; }
  generateDocsForJob(getJobFromRow(row));
}

function reprocessSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig().SHEET_NAME);
  const row   = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('Please select a data row (not the header).'); return; }
  const job = getJobFromRow(row);
  job.isRepost = false;
  generateDocsForJob(job);
}

function openParentFolder() {
  const folder = DriveApp.getFolderById(getConfig().PARENT_FOLDER_ID);
  SpreadsheetApp.getUi().alert('Parent Drive Folder', 'URL: ' + folder.getUrl() + '\n\nCopy this URL to open in Drive.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showSetupGuide() {
  SpreadsheetApp.getUi().alert(
    'Setup Guide',
    'See the SETUP_GUIDE.md file included with this script package.\n\n' +
    'Key steps:\n' +
    '1. Copy Config.gs values and fill in your IDs/keys\n' +
    '2. Upload base resume to Drive, paste its file ID in Config.gs\n' +
    '3. Upload work examples to Drive, paste its file ID in Config.gs\n' +
    '4. Create parent Drive folder, paste its ID in Config.gs\n' +
    '5. Add Claude API key in Config.gs\n' +
    '6. Set up Gmail labels/filters per SETUP_GUIDE.md\n' +
    '7. Set up time-based trigger for scanJobAlerts()',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
