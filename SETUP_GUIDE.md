# Job Application Automation System
## Complete Setup Guide

---

## What This System Does

1. **Scans Gmail** every 15 minutes for job alert emails from LinkedIn, Indeed, and Glassdoor
2. **Parses each listing** — company, title, location, URL, post age — and skips reposts and duplicates
3. **Filters every role** through a multi-layer criteria engine: repost detection, automation-heavy title flagging, geographic eligibility, state-level hiring restriction checks, and country-of-origin screening
4. **Assigns priority** using a two-pass scoring system based on post age, remote status, location, and job description content
5. **Auto-generates** tailored resumes and cover letters (via Claude API) for Priority 1 and 2 jobs immediately
6. **Queues** all other priorities for manual review before generation
7. **Saves** each resume and cover letter as separate Google Docs inside a `[Company] - [Role]` folder in Drive
8. **Supports manual entry** — paste any job URL directly into the sheet and trigger generation on demand

> **Greenhouse note:** Greenhouse job search was removed from the automated workflow because Google's Custom Search JSON API is closed to new customers. Greenhouse roles can still be processed by pasting job URLs manually into the tracker.

---

## Priority System

The system uses a two-pass priority model. Pass 1 runs at parse time using location and remote status. Pass 2 runs after the job description is fetched and checks for state-level hiring restrictions.

| Priority | Color | Condition | Auto-generate? |
|----------|-------|-----------|----------------|
| **P1** | Red | Remote USA + posted < 2 hours ago + TX eligible | Yes |
| **P2** | Orange | Remote USA + posted < 24 hours ago + TX eligible | Yes |
| **P3** | Yellow | Remote USA + posted > 24 hours ago + TX eligible | No — manual |
| **P5** | Purple | Remote USA + state-restricted + Texas not listed | No — manual |
| **P10** | Steel blue | Remote + non-USA origin (not India) | No — manual |
| **P15** | Teal | On-site + within 50 miles of zip 75241 | No — manual |
| **Ignored** | — | Remote role originating from India | Never added |
| **Ignored** | — | On-site role outside 50-mile radius | Never added |
| **Ignored** | — | Non-remote, non-local (everything else) | Never added |

**Texas eligibility (Pass 2):** When a remote job description contains state restriction language (e.g. "only available in the following states"), the system checks whether Texas is listed. If not, the role is flagged P5 for manual review regardless of post age.

**Thin JD flag:** If all four description fetch strategies return fewer than 2,500 characters, the row is flagged "Review: thin JD" and generation is skipped until you can review and supplement the description manually.

---

## Filtering Layers

Every job passes through these checks in order before being added to the sheet:

1. **Repost detection** — email body scan + LinkedIn page fetch (fails open if blocked, flags as "Verify repost?")
2. **Automation title filter** — flags titles like "SDET", "QA Automation", "Test Automation Engineer" for manual review; never auto-generates
3. **Location / origin filter (Pass 1)** — ignores India-origin remote roles; flags non-USA remote as P10; checks on-site distance from 75241; ignores out-of-range on-site roles
4. **Deduplication** — URL match first, then Company + Title match
5. **State restriction check (Pass 2)** — runs after description fetch; downgrades to P5 if TX not in approved states list
6. **Thin JD check** — requires 2,500+ chars of job description content before generating; flags short descriptions for manual review

---

## Job Description Preprocessor

Before sending a job description to Claude, the system restructures it to maximize signal quality within the 8,000 character budget:

- **Always included first:** Responsibilities, Duties, What You'll Do
- **Included second:** Qualifications, Requirements, Skills, Experience
- **Included if space allows:** Role overview, position summary
- **Deprioritized:** About the company, benefits, culture descriptions
- **Dropped entirely:** EEO statements, legal boilerplate, hashtags, apply CTAs

This ensures Claude always receives the most actionable content regardless of how much boilerplate a JD contains.

---

## Files in This Package

| File | Purpose |
|------|---------|
| `Code.gs` | Main orchestration: email scanning, job processing, location filtering, manual entry, priority system, URL backfill |
| `Config.gs` | All settings: IDs, API keys, keyword lists, priority thresholds, home zip coordinates, cover letter toggle |
| `SheetManager.gs` | Google Sheet read/write, deduplication, conditional formatting, manual entry row handling |
| `DriveManager.gs` | Drive folder creation, Google Doc writing, em dash removal, inline formatting |
| `ClaudeAPI.gs` | Claude API calls, JD preprocessor, resume + cover letter generation |
| `GeminiAPI.gs` | Gemini API fallback for email parsing (with retry logic for transient errors) |
| `SETUP_GUIDE.md` | This file |

---

## Step-by-Step Setup

### Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet
2. Rename it **"Job Application Tracker"** (or any name you prefer)
3. The tab inside the sheet should match `SHEET_NAME` in `Config.gs` — default is `Job Tracker`
   - Right-click the tab at the bottom → **Rename** → type `Job Tracker`

> **Important:** The script must be created from within this spreadsheet (Extensions → Apps Script), not from script.google.com directly. This binds it to the sheet so `SPREADSHEET_ID` can be left blank in Config.gs.

---

### Step 2 — Open the Apps Script Editor

1. In your spreadsheet, click **Extensions → Apps Script**
2. Delete the default `Code.gs` content
3. Create the following script files by clicking **+** next to "Files":
   - `Code.gs`
   - `Config.gs`
   - `SheetManager.gs`
   - `DriveManager.gs`
   - `ClaudeAPI.gs`
   - `GeminiAPI.gs`
4. Paste the contents of each `.gs` file from this package into the corresponding file
5. Save all files (**Save all** button or Ctrl/Cmd + S)

> **Tip:** Rename the Apps Script project to something meaningful (e.g. "RobertsJobSearch") by clicking the project name at the top left of the editor. This name appears in your Google account's security permissions portal.

---

### Step 3 — Create the Parent Drive Folder

1. Go to [drive.google.com](https://drive.google.com)
2. Create a new folder called **"Job Applications"** (or any name you prefer)
3. Open the folder and copy its ID from the URL:
   - URL format: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`
   - Copy the `FOLDER_ID_HERE` portion

---

### Step 4 — Upload Your Base Resume

1. Open your base resume as a Google Doc in Drive
   - This is your master resume — Claude tailors it for each job
   - Make it comprehensive. Include all experience. Claude will prioritize and reorder.
   - Keep your formatted PDF separately as a submission template — the Google Doc is for content only
2. Copy its file ID from the URL:
   - URL format: `https://docs.google.com/document/d/FILE_ID_HERE/edit`

---

### Step 5 — Upload Work Experience Examples

1. Create a Google Doc called **"Work Experience Examples"**
2. Structure each entry with a bracketed label header describing the entry's purpose and any usage notes
3. Include freeform descriptions of past roles, projects, achievements, and stories
   - Be specific: numbers, outcomes, technologies, team sizes
   - Write in first person — Claude shapes it into cover letter prose
4. Copy its file ID from the URL

**Entry header format:**
```
[Entry type and context]
Context: When and how to use this entry.
Note: Any cross-references to other entries or usage restrictions.

Content begins here...
```

**Special flags:**
- Company-specific entries that should never be reused: add `DO NOT reuse for other employers` to the header
- Entries demonstrating structural patterns: add `Use for STRUCTURE ONLY`
- The opening line style of a sample letter is not representative of preferred tone: add a note saying `Do not mirror this opener`

---

### Step 6 — Get Your API Keys

#### Claude API Key (Anthropic)
1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Sign up or log in → **API Keys → Create Key**
3. Copy the key (shown once only)
4. Estimated cost at 10 applications/day: ~$12–$15/month using claude-sonnet-4-6

#### Gemini API Key (Google)
1. Go to [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
2. Click **Create API Key** and copy it
3. Free tier: 15 requests/minute, 1,500 requests/day
4. Used as fallback email parser when the native parsers can't read a format

---

### Step 7 — Store API Keys Securely

Store keys in Script Properties rather than pasting them directly in `Config.gs`:

1. In the Apps Script editor, click **Project Settings** (gear icon)
2. Scroll to **Script Properties** → **Add script property**
3. Add the following properties:

| Property | Value |
|----------|-------|
| `CLAUDE_API_KEY` | Your Claude API key |
| `GEMINI_API_KEY` | Your Gemini API key |

`Config.gs` reads from Script Properties first — no code changes needed.

---

### Step 8 — Fill in Config.gs

Open `Config.gs` and fill in the three required Drive/file IDs:

```javascript
PARENT_FOLDER_ID:      'paste-your-drive-folder-id-here',
BASE_RESUME_FILE_ID:   'paste-your-resume-file-id-here',
WORK_EXAMPLES_FILE_ID: 'paste-your-work-examples-file-id-here',
```

Review and adjust these optional settings if needed:

```javascript
SHEET_NAME:             'Job Tracker',   // must match your tab name
HOME_ZIP:               '75241',         // your zip for on-site radius check
HOME_ZIP_LAT:            32.6721,        // update if you move
HOME_ZIP_LNG:           -96.7862,        // update if you move
GENERATE_COVER_LETTERS:  true,           // set false to skip cover letters
MIN_JD_CHARS:            2500,           // minimum chars to consider JD sufficient
AUTOMATION_TITLE_KEYWORDS: [...],        // add/remove automation filter terms
```

---

### Step 9 — Set Up Gmail Labels and Filters

#### Create Labels
1. In Gmail: **Settings → See all settings → Labels → Create new label**
2. Create these three labels:
   - `job-alerts-linkedin`
   - `job-alerts-indeed`
   - `job-alerts-glassdoor`

#### Create Filters
For each source, create a filter that applies the label automatically:

**LinkedIn:**
1. Gmail → Settings → Filters and Blocked Addresses → Create new filter
2. From: `jobalerts@linkedin.com`
3. Click "Create filter" → check **Apply the label** → select `job-alerts-linkedin`
4. Also check **Never send to Spam** and **Mark as important**
5. Leave **"Also apply filter to matching conversations"** unchecked to avoid processing years of backlog

**Indeed:**
- From: `alert@indeed.com` → Apply label: `job-alerts-indeed`

**Glassdoor:**
- From: `noreply@glassdoor.com` (or `alerts@glassdoor.com`) → Apply label: `job-alerts-glassdoor`

> **Tip:** If filters don't trigger on new emails, open one alert email, click the three-dot menu → **Filter messages like these** — Gmail pre-fills the exact sender address, eliminating typo risk.

#### Set Up Job Alerts on Each Platform
- **LinkedIn:** Jobs → Job Alerts → Create alert → set to **Daily** frequency + **Email** notification. Daily digest format includes more per-listing content than instant alerts.
- **Indeed:** Run a search → Get email alerts for this search
- **Glassdoor:** Jobs → Set alert

---

### Step 10 — Initialize the Sheet

1. In the Apps Script editor, open `SheetManager.gs`
2. Select `initializeSheet` from the function dropdown
3. Click **Run**
4. Grant permissions when prompted (Gmail, Drive, Sheets, Docs, external URLs)
5. Switch to your spreadsheet — you should see the formatted header row with color-coded columns

Then run `onOpen` once to install the **Job Automation** custom menu in the spreadsheet toolbar.

---

### Step 11 — Set Up the Automation Trigger

1. In the Apps Script editor, click **Triggers** (clock icon in left sidebar)
2. Click **+ Add Trigger**
3. Configure:
   - Function: `scanJobAlerts`
   - Deployment: Head
   - Event source: Time-driven
   - Type: Minutes timer
   - Interval: Every 15 minutes
4. Save

---

### Step 12 — Test the System

1. Find an existing unread job alert email in Gmail and manually apply the appropriate label to it
2. In the Apps Script editor, select `scanJobAlerts` and click **Run**
3. Click **Execution log** to watch it process
4. Check your spreadsheet for the new row
5. Check Drive for the new subfolder (if it was P1 or P2)

If the row appeared but fields are blank, the Gemini fallback parser handled it — check the log for "Falling back to Gemini parser."

---

## Using the Custom Menu

Once set up, a **"Job Automation"** menu appears in your spreadsheet toolbar:

| Menu Item | What It Does |
|-----------|--------------|
| Scan Job Alert Emails Now | Manually trigger an email scan (all three sources) |
| Generate Docs for Pending (P3, P5, P10, P15) | Process all queued manual-review rows |
| Generate Docs for Selected Row | Generate for whichever row your cursor is on |
| Re-process Selected Row (force) | Re-generates docs even if already done |
| Process Manual Entry (Selected Row) | Fetch description + generate docs for a manually added row |
| Validate Manual Entry Row | Preview what the system detected before committing |
| Backfill Missing LinkedIn URLs | Recover job URLs from cell notes for existing LinkedIn rows |
| Open Parent Drive Folder | Shows the URL to your Drive folder |

---

## Adding Jobs Manually

You can add any job posting to the tracker without waiting for an alert email. This is the recommended workflow for Greenhouse roles found via manual Google search (`site:greenhouse.io "remote" "qa engineer" "united states"`).

1. Open the Job Tracker sheet and add a new row
2. Paste the job posting URL into **column H (Job URL)** — this is the only required field
3. Optionally fill in any columns (the system will attempt to detect blank fields from the page):
   - **Column A** — Company Name
   - **Column B** — Job Title
   - **Column C** — Date Posted
   - **Column E** — Location
   - **Column F** — Remote (`Yes` or `No`)
   - **Column K** — Source (e.g. `LinkedIn`, `Greenhouse`, `Company Website`)
4. Select the row
5. Click **Job Automation → Validate Manual Entry Row** to preview detection (optional but recommended)
6. Click **Job Automation → Process Manual Entry (Selected Row)**

### What happens during processing

- The system fetches the job description from the URL using the four-strategy waterfall
- Blank fields are filled from the page content where possible
- Location filters, automation checks, and priority rules run identically to automated entries
- A confirmation dialog shows detected details and asks whether to generate docs now or defer
- If deferred, status is set to "Manual: Pending Review" — generate later via **Generate Docs for Selected Row**

### Job description fetch waterfall

For every job — automated or manual — the system attempts to retrieve the full description in this order, continuing through each strategy until 2,500+ characters are found:

| Strategy | Source | Notes |
|----------|--------|-------|
| S1 | Email body excerpt | Always available for email-sourced jobs; empty for manual entries |
| S2 | Direct URL fetch | Works for Indeed, Glassdoor, Greenhouse, company sites; blocked for LinkedIn |
| S3 | DuckDuckGo career page search | Finds the company's own careers page; bypasses login walls |
| S4 | Full email body fallback | Last resort; always succeeds for email-sourced jobs |

For LinkedIn manual entries, S2 will hit the login wall and S3 automatically searches the company's career site instead. This succeeds for most known companies.

If all four strategies return fewer than 2,500 characters, the row is flagged "Review: thin JD" and generation is paused for your review.

---

## Workflow After Setup

### Automatic (runs in background)
- Email scan fires every 15 minutes — new jobs appear in the sheet with priority colors
- P1 and P2 jobs generate resume + cover letter automatically (subject to thin JD check)
- All other priorities appear as review flags for your attention

### Manual review (when you're ready)
1. Open the sheet
2. Review P3, P5, P10, P15, and any flagged rows
3. Delete any you don't want to apply to
4. For roles you want to pursue: select the row → **Generate Docs for Selected Row**
   - Or process all at once: **Generate Docs for Pending (P3, P5, P10, P15)**

### Greenhouse (manual discovery)
1. Search Google: `site:greenhouse.io "remote" "qa engineer" "united states"`
2. Paste promising job URLs into the tracker (column H)
3. Run **Process Manual Entry (Selected Row)** for each

### Applying
1. Find the job in the sheet
2. Click **Open Folder** in the Drive Folder column
3. Review and lightly edit the resume and cover letter in Google Docs
4. Apply your formatting template (margins, fonts, layout)
5. Download as PDF: **File → Download → PDF Document**

---

## Google Sheet Column Reference

| Column | Name | Notes |
|--------|------|-------|
| A | Company Name | |
| B | Job Title | |
| C | Date Posted | Date the job was posted |
| D | Post Age | Auto-calculated formula (e.g., "4.5 hrs", "3 days") |
| E | Location | City/State or Remote |
| F | Remote | Yes / No |
| G | Priority | 1, 2, 3, 5, 10, or 15 (color-coded) |
| H | Job URL | Clickable "View Posting" link |
| I | Drive Folder | Clickable "Open Folder" link (after generation) |
| J | Status | Pending / Auto-queued / Generating... / Done / Error / Review flags |
| K | Source | LinkedIn / Indeed / Glassdoor / Greenhouse / Manual |
| L | Processed Date | When the resume/cover letter was created |

### Status values reference

| Status | Meaning |
|--------|---------|
| Auto-queued | P1 or P2 — docs being generated automatically |
| Pending | Queued for manual review before generation |
| Generating... | Claude API call in progress |
| Done | Resume and cover letter created successfully |
| Error: ... | Generation failed — see Execution Log for details |
| Verify repost? | LinkedIn page fetch was blocked — manually confirm before generating |
| Review: automation-heavy title | Title matched automation filter — review before generating |
| Review: TX not in approved states (P5) | State restriction detected, TX not listed |
| Review: non-USA remote origin | Remote role from non-USA country |
| Review: on-site within 50 miles | On-site role within your radius |
| Review: on-site location unresolved | On-site role, location could not be geocoded |
| Review: thin JD (X chars) | All fetch strategies returned less than 2,500 chars — supplement manually |
| Review: thin JD — could not fetch description | Fetch returned nothing usable |
| Ignored: ... | Role was filtered out — row kept for reference |
| Manual: Pending Review | Manually added row, deferred generation |
| Manual: Fetching description... | In progress |

---

## Troubleshooting

**"No jobs found" after running scanJobAlerts**
- Verify Gmail labels exist and have unread emails
- Check that Gmail filters are applying labels to incoming alert emails
- Open a real alert email, click three-dot menu → "Filter messages like these" to verify the sender address matches your filter

**"Claude API returned 401"**
- API key is wrong or not set — re-check Script Properties

**"Claude API returned 529" or "overloaded"**
- Anthropic is temporarily at capacity — re-run the failed row manually via the menu

**Gemini API returned 503**
- Temporary capacity spike — the system retries up to 3 times automatically with increasing delays. If all retries fail, the email falls through without Gemini parsing. Run `scanJobAlerts` again manually after a few minutes.

**Docs created but have no content**
- Check Execution Logs for Claude errors
- Verify base resume file ID is correct and the file is not empty

**Sheet rows are duplicated**
- Dedup checks URL first, then Company + Title. If LinkedIn sends slightly different URLs for the same job they may slip through. Delete the duplicate row manually.

**Em dashes appearing in docs**
- The system enforces this at two layers (Claude prompt + DriveManager.gs post-processing). If one appears, use Find & Replace (Ctrl/Cmd + H) to replace `—` with `, `

**P5 flagging a role that should be eligible**
- Open the job description and check for state restriction language. If the JD is ambiguous, the system may have detected a false positive. Use **Re-process Selected Row** or manually change the priority and trigger generation.

**LinkedIn rows have blank Job URL**
- Run **Job Automation → Backfill Missing LinkedIn URLs** — the system scans cell notes for "View job:" links and populates column H. Rows where no URL is found get a Google search fallback link instead.

**"Review: thin JD" on a job you want to apply to**
- Visit the job posting directly, copy the full description text
- Paste it into a new Google Doc saved in Drive
- Change the row status to "Pending" and run **Generate Docs for Selected Row** — the system will use S3/S4 fallback with whatever content it has

**Manual entry: description not fetching**
- The URL may require authentication or JavaScript rendering
- Copy the job description text manually, paste into a Google Doc in Drive
- Use **Generate Docs for Selected Row** directly

---

## Cost Estimate

At 10 applications per day (~300/month):

| Service | Usage | Monthly cost |
|---------|-------|--------------|
| Claude (claude-sonnet-4-6) | 300 resume + cover letter pairs | ~$12–$15 |
| Gemini 2.5 Flash | Email parsing fallback | Free |
| Google Apps Script | All automation | Free |
| Google Drive / Docs / Sheets | All storage | Free |

**Recommended:** `claude-sonnet-4-6` is set as the default in `Config.gs`. To disable cover letter generation and reduce costs by ~40%, set `GENERATE_COVER_LETTERS: false` in Config.gs.

---

## Customization Tips

- **Disable cover letters:** Set `GENERATE_COVER_LETTERS: false` in Config.gs
- **Change JD sufficiency threshold:** Edit `MIN_JD_CHARS` in Config.gs (default 2,500)
- **Change priority thresholds:** Edit `PRIORITY_1_REMOTE_HOURS` and `PRIORITY_2_HOURS` in Config.gs
- **Change home zip:** Update `HOME_ZIP`, `HOME_ZIP_LAT`, and `HOME_ZIP_LNG` in Config.gs
- **Add automation filter terms:** Add to `AUTOMATION_TITLE_KEYWORDS` in Config.gs
- **Adjust JD character budget:** Change `MAX_JD_CHARS` in Config.gs (default 8,000)
- **Add more email sources:** Duplicate a parser function in Code.gs and add a new Gmail label
- **Adjust Claude's writing style:** Edit the prompts in ClaudeAPI.gs — instructions are clearly labeled
- **Add voice/tone to cover letters:** Include a writing style sample in your Work Experience Examples doc with a header note telling Claude to match the tone
