// ============================================================
//  SheetManager.gs — Google Sheet read/write logic
//
//  Handles: header setup, row append, deduplication,
//           status updates, job retrieval, formatting.
// ============================================================


// ─────────────────────────────────────────────
//  SHEET INITIALIZATION
// ─────────────────────────────────────────────

/**
 * Creates and formats both tabs (Job Tracker and Other) if they don't exist.
 * Run once on first setup.
 */
function initializeSheet() {
  const config = getConfig();
  const ss     = getSpreadsheet();

  // Initialize both tabs with identical structure
  const tabNames = [config.SHEET_NAME, config.SHEET_NAME_OTHER];
  let primarySheet;

  for (const tabName of tabNames) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      Logger.log('Created tab: ' + tabName);
    }

    // Write headers if row 1 is empty
    const firstCell = sheet.getRange(1, 1).getValue();
    if (!firstCell) {
      const headers = [
        'Company Name',
        'Job Title',
        'Date Posted',
        'Post Age',
        'Location',
        'Remote',
        'Priority',
        'Job URL',
        'Drive Folder',
        'Status',
        'Source',
        'Processed Date',
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow(sheet, headers.length);
      Logger.log('Headers written to: ' + tabName);
    }

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(config.COL_COMPANY,        160);
    sheet.setColumnWidth(config.COL_TITLE,          220);
    sheet.setColumnWidth(config.COL_DATE_POSTED,    110);
    sheet.setColumnWidth(config.COL_POST_AGE,        90);
    sheet.setColumnWidth(config.COL_LOCATION,       140);
    sheet.setColumnWidth(config.COL_REMOTE,          70);
    sheet.setColumnWidth(config.COL_PRIORITY,        70);
    sheet.setColumnWidth(config.COL_JOB_URL,        220);
    sheet.setColumnWidth(config.COL_DRIVE_FOLDER,   220);
    sheet.setColumnWidth(config.COL_STATUS,         110);
    sheet.setColumnWidth(config.COL_SOURCE,          90);
    sheet.setColumnWidth(config.COL_PROCESSED_DATE, 120);
    applyConditionalFormatting(sheet);

    if (tabName === config.SHEET_NAME) primarySheet = sheet;
  }

  Logger.log('Both tabs initialized.');
  return primarySheet;
}

function formatHeaderRow(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 32);
}

function applyConditionalFormatting(sheet) {
  const config    = getConfig();
  const rules     = [];
  const lastRow   = 1000;
  const priorityCol = config.COL_PRIORITY;

  // P1 — red background (urgent)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#ff4d4d')
    .setFontColor('#ffffff')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // P2 — orange
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground('#ff9933')
    .setFontColor('#ffffff')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // P3 — yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(3)
    .setBackground('#ffdd57')
    .setFontColor('#333333')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // P3 — yellow (was also used for P4 previously, now P4 removed)
  // P5 — purple (state-restricted remote, TX not listed)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(5)
    .setBackground('#e1bee7')
    .setFontColor('#6a1b9a')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // P10 — steel blue (non-USA remote origin)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(10)
    .setBackground('#bbdefb')
    .setFontColor('#1565c0')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // P15 — teal (on-site within 50 miles)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(15)
    .setBackground('#b2dfdb')
    .setFontColor('#00695c')
    .setRanges([sheet.getRange(2, priorityCol, lastRow, 1)])
    .build());

  // Status "Review: automation" — orange text, light yellow background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('Review: automation')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange(2, config.COL_STATUS, lastRow, 1)])
    .build());

  // Status "Review: thin JD" — amber/orange, distinct from automation flag
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('Review: thin JD')
    .setBackground('#ffe8cc')
    .setFontColor('#7d3c00')
    .setRanges([sheet.getRange(2, config.COL_STATUS, lastRow, 1)])
    .build());

  // Status "Done" — green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Done')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([sheet.getRange(2, config.COL_STATUS, lastRow, 1)])
    .build());

  // Status "Error" — pink
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith('Error')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([sheet.getRange(2, config.COL_STATUS, lastRow, 1)])
    .build());

  // Source "Greenhouse" — teal accent to distinguish from email sources
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Greenhouse')
    .setBackground('#d4f1f4')
    .setFontColor('#0c5460')
    .setRanges([sheet.getRange(2, config.COL_SOURCE, lastRow, 1)])
    .build());

  // ── Alternating row text color by Date Posted date ──────────
  // Rows where (TODAY() - DatePosted) is EVEN get blue text.
  // Rows where it's ODD keep default black text.
  // This flips automatically every 24 hours, visually grouping
  // entries by day for easier mobile reading without scrolling
  // to the date column.
  // Applies to the entire row (all columns A through L).
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($C2<>"",ISEVEN(TODAY()-INT($C2)))')
    .setFontColor('#1a73e8')
    .setRanges([sheet.getRange(2, 1, lastRow, config.TOTAL_COLUMNS)])
    .build());

  sheet.setConditionalFormatRules(rules);
}


// ─────────────────────────────────────────────
//  DEDUPLICATION
// ─────────────────────────────────────────────

/**
 * Returns true if a job matching this URL or (Company + Title) already exists.
 */
function isDuplicate(job) {
  const config = getConfig();
  const sheet  = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  function normalizeTitle(t)   { return (t || '').toString().toLowerCase().replace(/\s+/g, ' ').trim(); }
  function normalizeCompany(c) { return (c || '').toString().toLowerCase().replace(/\s+/g, ' ').trim(); }
  function normalizeUrl(u)     { return (u || '').toString().trim().split('?')[0]; }

  const incomingUrl     = normalizeUrl(job.url);
  const incomingTitle   = normalizeTitle(job.title);
  const incomingCompany = normalizeCompany(job.company);

  // Check both tabs for duplicates
  const tabs = [config.SHEET_NAME, config.SHEET_NAME_OTHER];
  for (const tabName of tabs) {
    const s = getSheetByName(tabName);
    if (!s) continue;
    const lr = s.getLastRow();
    if (lr < 2) continue;
    const data = s.getRange(2, 1, lr - 1, config.TOTAL_COLUMNS).getValues();

    for (const row of data) {
      const existingUrl     = normalizeUrl(row[config.COL_JOB_URL - 1]);
      const existingTitle   = normalizeTitle(row[config.COL_TITLE - 1]);
      const existingCompany = normalizeCompany(row[config.COL_COMPANY - 1]);

      if (incomingUrl && existingUrl && incomingUrl === existingUrl) return true;
      if (incomingTitle && existingTitle && incomingCompany && existingCompany &&
          incomingTitle === existingTitle && incomingCompany === existingCompany) return true;
    }
  }

  return false;
}


// ─────────────────────────────────────────────
//  ROW OPERATIONS
// ─────────────────────────────────────────────

/**
 * Appends a new job row to the correct tab based on priority.
 * P1 and P2 go to the primary Job Tracker tab.
 * All other priorities go to the Other tab.
 * Returns an object { rowIndex, sheetName } so callers know which tab was used.
 */
function appendJobToSheet(job) {
  const config    = getConfig();
  const tabName   = (job.priority <= 2) ? config.SHEET_NAME : config.SHEET_NAME_OTHER;
  const sheet     = getOrInitSheetByName(tabName);
  const lastRow = sheet.getLastRow();
  const newRow  = lastRow + 1;

  const postedDate = job.postedDate instanceof Date ? job.postedDate : new Date(job.postedDate || Date.now());

  // Build row array (order must match column map in Config.gs)
  const rowData = new Array(config.TOTAL_COLUMNS).fill('');
  rowData[config.COL_COMPANY        - 1] = job.company   || '';
  rowData[config.COL_TITLE          - 1] = job.title     || '';
  rowData[config.COL_DATE_POSTED    - 1] = postedDate;
  rowData[config.COL_POST_AGE       - 1] = postAgeFormula(newRow, config.COL_DATE_POSTED);
  rowData[config.COL_LOCATION       - 1] = job.location  || '';
  rowData[config.COL_REMOTE         - 1] = job.isRemote  ? 'Yes' : 'No';
  rowData[config.COL_PRIORITY       - 1] = job.priority  || 4;
  rowData[config.COL_JOB_URL        - 1] = job.url       || '';
  rowData[config.COL_DRIVE_FOLDER   - 1] = '';   // filled in after doc generation
  rowData[config.COL_STATUS         - 1] = job.priority <= 2 ? 'Auto-queued' : 'Pending';
  rowData[config.COL_SOURCE         - 1] = job.source    || '';
  rowData[config.COL_PROCESSED_DATE - 1] = '';   // filled in after doc generation

  sheet.getRange(newRow, 1, 1, config.TOTAL_COLUMNS).setValues([rowData]);

  // Format the date posted cell
  sheet.getRange(newRow, config.COL_DATE_POSTED).setNumberFormat('MM/dd/yyyy');

  // Make Job URL a hyperlink
  if (job.url) {
    const urlCell = sheet.getRange(newRow, config.COL_JOB_URL);
    urlCell.setFormula(`=HYPERLINK("${job.url.replace(/"/g, '')}", "View Posting")`);
  }

  // Alternate row shading
  if (newRow % 2 === 0) {
    sheet.getRange(newRow, 1, 1, config.TOTAL_COLUMNS).setBackground('#f8f9fa');
  }

  // Store job data as a note on the company cell for retrieval later
  const jobMeta = JSON.stringify({
    company:    job.company,
    title:      job.title,
    url:        job.url,
    source:     job.source,
    isRemote:   job.isRemote,
    postedDate: postedDate.toISOString(),
    emailBody:  (job.emailBody || '').substring(0, 2000), // truncate for storage
  });
  sheet.getRange(newRow, config.COL_COMPANY).setNote(jobMeta);

  SpreadsheetApp.flush();
  // Return object with both row index and tab name so callers can
  // update the correct sheet when writing status/folder URL later
  return { rowIndex: newRow, sheetName: tabName };
}

/**
 * Updates status + folder URL + processed date after doc generation.
 */
function updateJobRow(rowIndex, updates, sheetName) {
  const config = getConfig();
  const sheet  = sheetName ? getSheetByName(sheetName) : findSheetForRow(rowIndex);

  if (!sheet) { Logger.log('updateJobRow: could not find sheet for row ' + rowIndex); return; }

  if (updates.status) {
    sheet.getRange(rowIndex, config.COL_STATUS).setValue(updates.status);
  }
  if (updates.priority) {
    sheet.getRange(rowIndex, config.COL_PRIORITY).setValue(updates.priority);
  }
  if (updates.folderUrl) {
    sheet.getRange(rowIndex, config.COL_DRIVE_FOLDER)
      .setFormula('=HYPERLINK("' + updates.folderUrl.replace(/"/g, '') + '", "Open Folder")');
  }
  if (updates.processedDate) {
    sheet.getRange(rowIndex, config.COL_PROCESSED_DATE)
      .setValue(updates.processedDate)
      .setNumberFormat('MM/dd/yyyy hh:mm');
  }

  SpreadsheetApp.flush();
}

function updateJobStatus(rowIndex, status, sheetName) {
  if (!rowIndex) return;
  const sheet = sheetName ? getSheetByName(sheetName) : findSheetForRow(rowIndex);
  if (!sheet) { Logger.log('updateJobStatus: could not find sheet for row ' + rowIndex); return; }
  sheet.getRange(rowIndex, getConfig().COL_STATUS).setValue(status);
  SpreadsheetApp.flush();
}

/**
 * Returns all rows matching a given status string across BOTH tabs.
 */
function getJobsByStatus(status) {
  const config = getConfig();
  const jobs   = [];
  const tabs   = [config.SHEET_NAME, config.SHEET_NAME_OTHER];

  for (const tabName of tabs) {
    const sheet   = getSheetByName(tabName);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    const data = sheet.getRange(2, 1, lastRow - 1, config.TOTAL_COLUMNS).getValues();
    data.forEach(function(row, i) {
      const rowStatus = (row[config.COL_STATUS - 1] || '').toString().trim();
      if (rowStatus === status) {
        const job = rowToJob(row, i + 2, sheet);
        job.sheetName = tabName;
        jobs.push(job);
      }
    });
  }

  return jobs;
}

/**
 * Retrieves a job object from a specific sheet row.
 * Searches the active sheet first, then both tabs if not found.
 */
function getJobFromRow(rowIndex) {
  const config = getConfig();
  // Try the active sheet first (most common case when called from menu)
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (activeSheet && activeSheet.getLastRow() >= rowIndex) {
    const row = activeSheet.getRange(rowIndex, 1, 1, config.TOTAL_COLUMNS).getValues()[0];
    if (row[0]) {
      const job = rowToJob(row, rowIndex, activeSheet);
      job.sheetName = activeSheet.getName();
      return job;
    }
  }
  // Fallback: search both tabs
  const tabs = [config.SHEET_NAME, config.SHEET_NAME_OTHER];
  for (const tabName of tabs) {
    const sheet = getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < rowIndex) continue;
    const row = sheet.getRange(rowIndex, 1, 1, config.TOTAL_COLUMNS).getValues()[0];
    if (row[0]) {
      const job = rowToJob(row, rowIndex, sheet);
      job.sheetName = tabName;
      return job;
    }
  }
  return rowToJob([], rowIndex, null);
}

function rowToJob(row, rowIndex, sheet) {
  const config = getConfig();

  // Try to recover full job meta from cell note (auto-parsed rows)
  let meta = {};
  try {
    const s    = sheet || getSheet();
    const note = s.getRange(rowIndex, config.COL_COMPANY).getNote();
    if (note) meta = JSON.parse(note);
  } catch (e) {}

  // For the URL: prefer stored meta (most reliable), then fall back to
  // reading the actual cell value. This supports manually entered rows
  // where no note exists but the user has typed/pasted a URL directly.
  let urlFromCell = '';
  try {
    const urlCell    = getSheet().getRange(rowIndex, config.COL_JOB_URL);
    const cellFormula = urlCell.getFormula();
    const cellValue   = urlCell.getValue();
    if (cellFormula) {
      // Extract URL from HYPERLINK formula: =HYPERLINK("url","label")
      const match = cellFormula.match(/=HYPERLINK\("([^"]+)"/i);
      urlFromCell = match ? match[1] : '';
    } else if (cellValue && cellValue.toString().startsWith('http')) {
      urlFromCell = cellValue.toString().trim();
    }
  } catch (e) {}

  // Determine if this is a manual entry (no cell note present)
  const isManualEntry = !meta.company && !meta.url;

  return {
    rowIndex:      rowIndex,
    company:       row[config.COL_COMPANY     - 1] || meta.company  || '',
    title:         row[config.COL_TITLE       - 1] || meta.title    || '',
    location:      row[config.COL_LOCATION    - 1] || '',
    isRemote:      (row[config.COL_REMOTE     - 1] || '').toString().toLowerCase() === 'yes',
    priority:      row[config.COL_PRIORITY    - 1] || 3,
    url:           meta.url || urlFromCell || '',
    source:        meta.source || row[config.COL_SOURCE - 1] || (isManualEntry ? 'Manual' : ''),
    postedDate:    meta.postedDate ? new Date(meta.postedDate) : new Date(row[config.COL_DATE_POSTED - 1] || Date.now()),
    emailBody:     meta.emailBody || '',
    isRepost:      false,
    isManualEntry: isManualEntry,
  };
}

// ─────────────────────────────────────────────
//  SHEET HELPERS
// ─────────────────────────────────────────────

function getSpreadsheet() {
  const config = getConfig();
  return config.SPREADSHEET_ID
    ? SpreadsheetApp.openById(config.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const config = getConfig();
  return getSpreadsheet().getSheetByName(config.SHEET_NAME);
}

function getSheetByName(name) {
  return getSpreadsheet().getSheetByName(name);
}

function getOrInitSheet() {
  return getSheet() || initializeSheet();
}

function getOrInitSheetByName(name) {
  const sheet = getSheetByName(name);
  if (sheet) return sheet;
  initializeSheet();
  return getSheetByName(name);
}

/**
 * Finds which tab a given row index lives in by checking both tabs.
 * Used when sheetName is not stored on the job object.
 */
function findSheetForRow(rowIndex) {
  const config = getConfig();
  const tabs   = [config.SHEET_NAME, config.SHEET_NAME_OTHER];
  for (const tabName of tabs) {
    const sheet = getSheetByName(tabName);
    if (sheet && sheet.getLastRow() >= rowIndex) {
      const val = sheet.getRange(rowIndex, 1).getValue();
      if (val) return sheet;
    }
  }
  return getSheet(); // fallback to primary
}

// Returns a formula string for the Post Age column
function postAgeFormula(row, dateCol) {
  const colLetter = columnToLetter(dateCol);
  // Returns "X hrs" or "X days"
  return `=IF(${colLetter}${row}="","",IF((NOW()-${colLetter}${row})*24<24,TEXT(ROUND((NOW()-${colLetter}${row})*24,1),"0.0")&" hrs",TEXT(ROUND(NOW()-${colLetter}${row},0),"0")&" days"))`;
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
