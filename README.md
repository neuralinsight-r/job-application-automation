# Job Application Automation System

A Google Apps Script system that monitors Gmail job alert emails from LinkedIn, Indeed, and Glassdoor, parses listings, applies multi-layer filters, routes by priority into a two-tab Google Sheet, and generates tailored resumes and cover letters via the Claude API.

---

## Features

- **Email parsing** for LinkedIn, Indeed, and Glassdoor job alert digests
- **Multi-layer filtering** — company blocklist, repost detection, automation title flagging, location/remote eligibility checks
- **Two-tab Google Sheet routing** — P1 and P2 (high-priority remote) go to "Job Tracker"; everything else goes to "Other"
- **Priority scoring** — P1 (remote, under 2 hrs old) through P15 (on-site, local)
- **JD fetch waterfall** — four strategies to retrieve the full job description before generation
- **AI-powered document generation** — tailored resume and cover letter via Claude Sonnet, with Gemini 2.5 Flash as a parsing fallback
- **Manual entry support** — paste any Greenhouse, LinkedIn, or job board URL to process it on demand
- **All-manual generation mode** — `AUTO_GENERATE: false` by default; nothing runs without your trigger
- **Mobile-friendly sheet** — alternating row text color by date posted for easy scanning in portrait mode

---

## File Structure

```
job-application-automation/
├── README.md              <- You are here
├── SETUP_GUIDE.md         <- Full setup instructions
├── .gitignore
├── Code.gs                <- Main orchestration, parsing, JD fetch, filtering
├── Config.gs              <- All user settings in one place
├── SheetManager.gs        <- Sheet read/write, two-tab routing, conditional formatting
├── DriveManager.gs        <- Drive folder and Google Doc creation
├── ClaudeAPI.gs           <- Resume and cover letter generation, JD preprocessor
└── GeminiAPI.gs           <- Gemini fallback parser with retry logic
```

---

## Quick Start

1. Create a Google Sheet with two tabs named **Job Tracker** and **Other**
2. Open **Extensions > Apps Script** in the sheet
3. Create one file per `.gs` file in this repo and paste the contents
4. Fill in `Config.gs` with your API keys, Drive folder ID, and resume file IDs
5. Run `initializeSheet()` once to apply headers and formatting to both tabs
6. Run `createTrigger()` to schedule email scanning every 15 minutes
7. Set up Gmail filters to label job alert emails (see `SETUP_GUIDE.md`)

Full step-by-step instructions, Gmail filter setup, and API key configuration are in [SETUP_GUIDE.md](SETUP_GUIDE.md).

---

## Tech Stack

| Component | Technology |
|---|---|
| Runtime | Google Apps Script (V8) |
| Sheet storage | Google Sheets (two-tab routing) |
| Document output | Google Docs via Drive API |
| Resume / cover letter AI | Claude Sonnet (Anthropic API) |
| Email parse fallback | Gemini 2.5 Flash (Google AI Studio) |
| JD fetch | UrlFetchApp + DuckDuckGo career page search |
| Scheduling | Apps Script time-based triggers |

---

## Priority Routing

| Priority | Criteria | Tab |
|---|---|---|
| P1 | Remote USA + posted under 2 hours | Job Tracker |
| P2 | Remote USA + posted under 24 hours | Job Tracker |
| P3 | Remote USA, any age | Other |
| P5 | Remote USA + state restriction flag | Other |
| P10 | On-site within 50 miles | Other |
| P15 | On-site, distance unresolved | Other |
| Ignored | India-origin remote; on-site over 50 miles | Not added |

---

## Generation

All document generation is manual by default (`AUTO_GENERATE: false` in `Config.gs`). Use the **Job Automation** menu in the sheet to trigger generation for any row. To re-enable automatic generation for P1 and P2 jobs, set `AUTO_GENERATE: true`.

---

## Notes

- API keys should be stored in **Script Properties** (not hardcoded). See `SETUP_GUIDE.md`.
- Greenhouse job search is manual-entry only. Google's Custom Search JSON API is closed to new signups. Search `site:greenhouse.io "remote" "qa engineer" "united states"` on Google and paste URLs into the tracker.
- The base resume and work examples document stay as Google Docs in Drive. Claude uses them as source material for every tailored output.

---

## License

Copyright (c) 2026 Robert Ledesma. All rights reserved.

This source code is made available for reference and portfolio purposes only. No part of this project may be copied, modified, distributed, or used in any form without explicit written permission from the author.
