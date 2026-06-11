# snafu-rally-scoring

Automated bonus scoring for rally events via Google Apps Script. Riders submit bonus claims by email — the script validates them, records them in your scoring spreadsheet, and manages everything through Gmail labels. Scorers approve or deny submissions from a Gmail sidebar add-on with one click.

---

## How it works

```
Rider sends email: "42 BP-07"
        │
        ▼
  rally/unprocessed
        │
  Script validates format + sender email
        │
        ├─ bad format ──────────────► rally/format-error
        ├─ unregistered sender ─────► rally/email-error
        ├─ bonus ID not found ──────► rally/processing-error
        │
        ▼
  rally/needs-review   ◄── scorer reviews in Gmail sidebar
        │
        ├─ Deny ────────────────────► rally/denied  (+ X written to sheet immediately)
        │
        ▼
  rally/approved       ◄── scorer clicks Approve in sidebar
        │
  Script picks up on next run
        │
        ▼
  rally/scored         ◄── X + timestamp written to rider's score sheet
```

The `rally` parent label is also created so scorers can click it to see every submission at once regardless of status.

---

## Files

| File | Purpose |
|------|---------|
| `code.js` | Main script — `setup()`, `processEmails()`, spreadsheet logic, email validation |
| `Sidebar.gs` | Gmail Add-on sidebar — Approve / Deny / Flag buttons |
| `appsscript.json` | Apps Script manifest with OAuth scopes and add-on registration |
| `setup.html` | Web-based setup wizard — generates all three files pre-configured for your event |

---

## Spreadsheet structure

The script reads all configuration at runtime from a **Config** sheet. `setup()` writes this sheet automatically — you don't create it manually.

### Rider Master

A tab named **`Rider Master`** (configurable). Row 1 is headers:

| Rider Number | Name | Email |
|---|---|---|
| 42 | Jane Smith | jane@example.com |
| 7 | Bob Jones | bob@example.com |

The sender email on every submission is validated against this sheet.

### Bonus Master

A tab named **`Bonus Master`** (configurable). Column A lists every bonus ID, one per row (row 1 = header):

| Bonus ID |
|----------|
| BP-01    |
| BP-02    |

Used to auto-create rider score sheets on first submission.

### Rider score sheets

One tab per rider, named by their **rider number** (e.g. `42`). Created automatically from Bonus Master on first submission if the tab doesn't exist.

Default column layout (all configurable):

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Bonus ID | Submitted | Submit Time | Approved | Approve Time | Denied | Deny Time |

---

## Gmail labels

All labels are created automatically by `setup()`. The label prefix is configurable (default: `rally`).

| Label | Meaning |
|-------|---------|
| `rally` | Parent label — click to see all submissions |
| `rally/unprocessed` | Incoming submission, not yet processed |
| `rally/needs-review` | Valid submission awaiting scorer decision |
| `rally/approved` | Scorer approved — script will record on next run |
| `rally/scored` | Fully recorded in the spreadsheet |
| `rally/denied` | Scorer denied — X written to sheet immediately |
| `rally/format-error` | Subject didn't match `<number> <bonusID>` |
| `rally/email-error` | Sender not registered for that rider number |
| `rally/processing-error` | Script error — check Executions log |

---

## Installation

### 1. Run the setup wizard

Open `setup.html` in any browser. Fill in:

- Event name and organizer email
- Your Google Sheets ID (from the spreadsheet URL)
- Column layout for rider score sheets
- Label prefix and sub-label names
- How often the script should check for new emails

Click **Generate files** and download all three:
- `code.js`
- `Sidebar.gs`
- `appsscript.json`

### 2. Install in Apps Script

1. Open your Google Sheet → **Extensions → Apps Script**
2. Paste `code.js` into `Code.gs`
3. Click **+** next to Files → New script → name it `Sidebar` → paste `Sidebar.gs`
4. Project Settings → check **Show "appsscript.json" manifest file in editor** → paste `appsscript.json`
5. Save all files

### 3. Run setup()

In the Apps Script editor:

1. Select function `setup` from the dropdown
2. Click **Run ▶**
3. Authorise when prompted

`setup()` will:
- Write the Config sheet with all your settings
- Create the bare `rally` parent label
- Create all sub-labels
- Save the spreadsheet ID to Script Properties (used by the sidebar)
- Set the time-driven trigger

### 4. Activate the Gmail sidebar add-on

1. In Apps Script → **Deploy → Test deployments**
2. Click **Install**
3. Open Gmail — the **Rally Scoring** panel will appear on the right when you open any submission email

### 5. Set up the Gmail filter (recommended)

So incoming submissions are labelled automatically:

1. Gmail → Settings → **See all settings → Filters → Create a new filter**
2. **To**: your scoring email address
3. **Create filter** → Apply label → `rally/unprocessed`
4. Save

---

## Submission format

Riders send email to your scoring address with the subject:

```
<RiderNumber> <BonusID>
```

Examples:

```
42 BP-07
7 checkpoint-alpha
101 BONUS_99
```

- Rider number must be numeric
- Anything after the first space is the Bonus ID — matched against column A of the rider's sheet
- Email must come from the rider's registered address in Rider Master

---

## Config sheet reference

`setup()` writes these keys. All values are editable in the sheet after setup — no need to touch the code.

| Key | Default | Notes |
|-----|---------|-------|
| `event_name` | — | Display only |
| `organizer_email` | — | Display only |
| `spreadsheet_id` | — | Do not change after setup |
| `sheet_rider_master` | `Rider Master` | Tab name of the rider roster |
| `sheet_bonus_master` | `Bonus Master` | Tab with all bonus IDs in column A |
| `master_col_rider_number` | `Rider Number` | Column header in Rider Master |
| `master_col_email` | `Email` | Column header in Rider Master |
| `header_row` | `1` | Header row number in rider sheets |
| `col_bonus_id` | `1` | Column A |
| `col_submitted` | `2` | Column B |
| `col_submitted_time` | `3` | Column C |
| `col_approved` | `4` | Column D |
| `col_approved_time` | `5` | Column E |
| `col_denied` | `6` | Column F |
| `col_denied_time` | `7` | Column G |
| `trigger_interval_min` | `10` | Re-run `setup()` to change |
| `label_parent` | `rally` | Bare parent label |
| `label_unprocessed` | `rally/unprocessed` | |
| `label_format_error` | `rally/format-error` | |
| `label_email_error` | `rally/email-error` | |
| `label_processing_error` | `rally/processing-error` | |
| `label_needs_review` | `rally/needs-review` | |
| `label_approved` | `rally/approved` | |
| `label_denied` | `rally/denied` | |
| `label_scored` | `rally/scored` | |

---

## Changing settings after setup

1. Run the setup wizard again with updated values
2. Download and paste the new `code.js`
3. Re-run `setup()` — it rewrites the Config sheet and resets the trigger

Or edit the Config sheet directly for label name and column changes (no re-run needed — the script reads config live on every execution).

---

## Troubleshooting

| Symptom | Check |
|---------|-------|
| "Config sheet not found" | Run `setup()` first |
| Script stops immediately | One or more Gmail labels are missing — re-run `setup()` |
| Email tagged `format-error` | Subject must be `<number> <space> <bonusID>` — no prefix, no punctuation |
| Email tagged `email-error` | Sender address doesn't match Rider Master — check for typos or alias issues |
| Email tagged `processing-error` | Open Apps Script → Executions → click the failed run for the full error |
| Bonus ID not found | Check Column A of the rider's sheet — spacing and capitalisation must match exactly |
| Rider Master sheet not found | Tab must be named exactly as configured (default: `Rider Master`) |
| Sidebar not appearing | Deploy → Test deployments → Install |
| Trigger not running | Apps Script → Triggers — confirm `processEmails` exists; re-run `setup()` if not |

### Viewing logs

Apps Script → **Executions** in the left sidebar → click any run to expand its log. The script logs every action, validation result, and error with a clear message.

---

## OAuth scopes

| Scope | Used for |
|-------|---------|
| `spreadsheets` | Read/write scoring spreadsheet |
| `gmail.modify` | Read emails, manage labels |
| `gmail.addons.*` | Gmail sidebar add-on |
| `script.scriptapp` | Create time-driven trigger in `setup()` |
| `script.external_request` | External requests from sidebar |
