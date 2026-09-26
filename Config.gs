// ============================================================
//  Config.gs — ALL USER SETTINGS IN ONE PLACE
//
//  Fill in every value marked  ← REQUIRED
//  before running the system.
// ============================================================

function getConfig() {
  return {

    // ── Google Sheet ──────────────────────────────────────────
    // The spreadsheet this script is bound to (leave blank = active sheet)
    SPREADSHEET_ID:  '',                    // ← leave blank if script is bound to the sheet
    SHEET_NAME:      'Job Tracker',         // ← primary tab: P1 and P2 jobs
    SHEET_NAME_OTHER: 'Other',              // ← secondary tab: P3, P5, P10, P15 and all others

    // ── Auto-generation toggle ────────────────────────────────
    // Set to true to re-enable automatic resume/cover letter generation
    // for P1 and P2 jobs. When false, ALL generation requires manual trigger.
    AUTO_GENERATE:   false,

    // ── Google Drive ─────────────────────────────────────────
    // Create a folder in Drive called "Job Applications" (or any name).
    // Right-click it → Share → Copy link. The ID is the string after /folders/
    // Example: https://drive.google.com/drive/folders/1aBcD2EfGh3... → "1aBcD2EfGh3..."
    PARENT_FOLDER_ID: 'YOUR_PARENT_FOLDER_ID_HERE',  // ← REQUIRED

    // ── Base Resume ───────────────────────────────────────────
    // Upload your resume to Drive as a Google Doc or .txt file.
    // Right-click → Get link → copy the ID portion.
    BASE_RESUME_FILE_ID: 'YOUR_RESUME_FILE_ID_HERE',  // ← REQUIRED

    // ── Work Experience Examples ─────────────────────────────
    // A Google Doc or .txt with bullet points / paragraphs describing
    // past projects, achievements, and responsibilities.
    // Used to craft personalized cover letter body paragraphs.
    WORK_EXAMPLES_FILE_ID: 'YOUR_WORK_EXAMPLES_FILE_ID_HERE',  // ← REQUIRED

    // ── Claude API ────────────────────────────────────────────
    // Get your key at https://console.anthropic.com
    // Recommended: store in Script Properties (more secure) — see SETUP_GUIDE.md
    // For quick start you can paste it here, but rotate it after testing.
    CLAUDE_API_KEY:   PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY') || 'YOUR_CLAUDE_API_KEY_HERE',
    CLAUDE_MODEL:     'claude-sonnet-4-6',  // sonnet is recommended; opus is higher quality but ~4x cost
    CLAUDE_MAX_TOKENS: 4096,

    // ── Gemini API ────────────────────────────────────────────
    // Get your key at https://aistudio.google.com/app/apikey (free)
    GEMINI_API_KEY:   PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || 'YOUR_GEMINI_API_KEY_HERE',
    GEMINI_MODEL:     'gemini-2.5-flash',  // free tier; fast for parsing tasks

    // ── Gmail Labels / Search Queries ────────────────────────
    // Set up Gmail filters to label incoming job alert emails.
    // These are Gmail search queries used to find unread alert emails.
    // Adjust to match your actual labels or sender addresses.
    GMAIL_LABEL_LINKEDIN:  'label:job-alerts-linkedin',
    GMAIL_LABEL_INDEED:    'label:job-alerts-indeed',
    GMAIL_LABEL_GLASSDOOR: 'label:job-alerts-glassdoor',

    // ── Sheet Column Map ──────────────────────────────────────
    // Column numbers (1-based). Change if you reorder columns.
    COL_COMPANY:        1,
    COL_TITLE:          2,
    COL_DATE_POSTED:    3,
    COL_POST_AGE:       4,   // formula-computed
    COL_LOCATION:       5,
    COL_REMOTE:         6,
    COL_PRIORITY:       7,
    COL_JOB_URL:        8,
    COL_DRIVE_FOLDER:   9,
    COL_STATUS:         10,
    COL_SOURCE:         11,
    COL_PROCESSED_DATE: 12,
    TOTAL_COLUMNS:      12,

    // ── Priority Thresholds ───────────────────────────────────
    // Hours — adjust if you want different cutoffs
    PRIORITY_1_REMOTE_HOURS:  2,   // Remote USA + < 2 hrs  → P1 (auto-generate)
    PRIORITY_2_HOURS:         24,  // Remote USA + < 24 hrs → P2 (auto-generate)
    // Remote USA + > 24 hrs                      → P3 (manual)
    // Remote USA + state-restricted, TX not listed → P5 (manual)
    // Remote non-USA (not India)                 → P10 (manual)
    // On-site within 50 miles of HOME_ZIP        → P15 (manual)
    // Remote India / on-site outside 50mi        → Ignored

    // ── Home Location (for on-site proximity check) ───────────
    // Centroid coordinates for zip 75241 (Dallas, TX).
    // Update these if you move — use any lat/lng lookup tool.
    HOME_ZIP:     '75241',
    HOME_ZIP_LAT:  32.6721,   // latitude  for 75241
    HOME_ZIP_LNG: -96.7862,   // longitude for 75241

    // ── Doc Naming ────────────────────────────────────────────
    RESUME_DOC_PREFIX:       'Resume',
    COVER_LETTER_DOC_PREFIX: 'Cover Letter',

    // ── Automation Role Filter ────────────────────────────────
    // Job TITLES containing any of these keywords (case-insensitive) will be
    // flagged in the Status column as "Review: automation-heavy title" and
    // will NOT auto-generate docs regardless of priority.
    // You can review flagged rows and manually trigger generation if desired.
    //
    // Add or remove terms freely. All matching is against the job title only.
    // Partial matches count: "SDET" catches "Senior SDET", "Automation" catches
    // "Test Automation Engineer", etc.
    // ── Company Blocklist ─────────────────────────────────────
    // Companies listed here are ignored completely — no row written,
    // no docs generated. Matching is case-insensitive and partial,
    // so "Sundayy" also catches "Sundayy Inc", "Sundayy LLC", etc.
    // Add as many companies as needed.
    BLOCKED_COMPANIES: [
      'Sundayy',
    ],

    AUTOMATION_TITLE_KEYWORDS: [
      'SDET',
      'Automation Engineer',
      'Automation QA',
      'QA Automation',
      'Test Automation',
      'Automation Tester',
      'Automation Analyst',
      'Automation Specialist',
      'Software Development Engineer in Test',
      'Software Engineer in Test',
    ],

    // ── Misc ──────────────────────────────────────────────────
    // Minimum characters required before the system considers a job description
    // sufficient to generate a tailored resume and cover letter.
    // If all four fetch strategies return less than this, the row is flagged
    // "Review: thin JD" for manual review rather than generating with incomplete info.
    // 2500 chars typically covers a full responsibilities + qualifications section.
    MIN_JD_CHARS: 2500,

    // ── Cover Letter Toggle ─────────────────────────────────
    // Set to false to skip cover letter generation entirely.
    // Resumes are always generated regardless of this setting.
    GENERATE_COVER_LETTERS: true,

    // Max characters of preprocessed job description to send to Claude.
    // The JD preprocessor in ClaudeAPI.gs prioritizes signal-rich sections
    // (responsibilities, qualifications, requirements, skills) over boilerplate
    // before this limit is applied, so the full cap is used on useful content.
    MAX_JD_CHARS: 8000,
  };
}
