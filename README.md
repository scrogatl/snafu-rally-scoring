# snafu-rally-scoring

A Google Apps Script that automates bonus scoring for rally events. Riders submit bonus claims by emailing a designated Gmail account, and this script processes those emails, validates them against a rider roster, and updates a Google Sheet scoreboard — all automatically via a time-driven trigger.

---

## How It Works

Emails flow through a label-based state machine in Gmail:

```
rally/unprocessed
      │
      ├─ invalid subject format ──────────► rally/format-error
      │
      ├─ unregistered sender email ───────► rally/email-error
      │
      ├─ spreadsheet/processing error ───► rally/processing-error
      │
      └─ valid ───────────────────────────► rally/needs-review
                                                   │
                                          (manual approval)
                                                   │
                                            rally/approved
                                                   │
                                          (script detects)
                                                   │
                                             rally/scored
```

1. A rider sends an email with the subject `<RiderNumber> <BonusID>` (e.g. `42 BP-07`).
2. The script finds it under `rally/unprocessed`, validates the format and sender, and marks their bonus sheet with an `X` plus a timestamp.
3. A scorer reviews the submission and manually applies the `rally/approved` label.
4. On the next trigger run, the script marks the approved column and timestamps it, then applies `rally/scored`.

---

## Spreadsheet Structure

The script expects a Google Sheet with the following layout:

### Rider Master Sheet

A sheet named exactly **`Rider Master`** with these column headers in row 1:

| Rider Number | Name | Email |
|---|---|---|
| 42 | Jane Smith | jane@example.com |

### Rider Sheets

One sheet per rider, named by their **rider number** (e.g. `42`). Row 1 is a header; data starts at row 2.

| Column | Index | Contents |
|--------|-------|----------|
| A | 1 | Bonus ID |
| B | 2 | Submitted (`X`) |
| C | 3 | Submitted Time |
| D | 4 | Approved (`X`) |
| E | 5 | Approved Time |

---

## Gmail Label Setup

Create all of the following labels in Gmail before running the script. The script will stop and log an error if any are missing.

| Label | Purpose |
|-------|---------|
| `rally/unprocessed` | Applied to incoming bonus submission emails |
| `rally/format-error` | Subject line didn't match `Number Space String` |
| `rally/email-error` | Sender email didn't match the registered rider email |
| `rally/processing-error` | Script hit an unexpected error during processing |
| `rally/needs-review` | Valid submission awaiting scorer approval |
| `rally/approved` | Scorer has approved the submission |
| `rally/scored` | Script has recorded the approval in the spreadsheet |

To create labels in Gmail: Settings → See all settings → Labels → Create new label.

> **Tip:** Set up a Gmail filter to automatically apply `rally/unprocessed` to incoming emails sent to your scoring address.

---

## Deployment

### Prerequisites

- A Google account with Gmail and Google Sheets
- The spreadsheet and Gmail labels set up as described above

### Steps

1. Open your Google Sheet and go to **Extensions → Apps Script**.
2. Delete any placeholder code and paste in the contents of `code.js`.
3. Paste the contents of `appscript.json` into the **appsscript.json** manifest file (enable via Project Settings → Show "appsscript.json" manifest file).
4. If using a specific spreadsheet (not the active one), uncomment the `SpreadsheetApp.openById(...)` line in `processEmails()` and replace with your spreadsheet ID. You can find the ID in the sheet's URL: `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`.
5. Save the project.

### Set Up the Time-Driven Trigger

1. In Apps Script, click **Triggers** (clock icon) → **Add Trigger**.
2. Choose function: `processEmails`
3. Event source: **Time-driven**
4. Type: **Minutes timer** (every 5 or 10 minutes is recommended)
5. Save and authorize the script when prompted.

### Required OAuth Scopes

The script uses these scopes (already declared in `appscript.json`):

- `https://www.googleapis.com/auth/spreadsheets.currentonly` — read/write the active spreadsheet
- `https://www.googleapis.com/auth/gmail.modify` — read emails and manage labels

---

## Email Format

Riders must send emails with subjects in this exact format:

```
<RiderNumber> <BonusID>
```

Examples:
- `42 BP-07`
- `7 checkpoint-alpha`
- `101 BONUS_99`

The rider number must be numeric. Anything after the first space is treated as the Bonus ID and matched against Column A of that rider's sheet.

Emails that don't match this format are labeled `rally/format-error` and skipped.

---

## Logs

The script logs all actions to Apps Script's built-in logger. To view logs:

**Executions** tab in the Apps Script editor → click any run to expand its log output.

Useful for debugging missing labels, sheet mismatches, or email format issues.

---

## Troubleshooting

**Script stops immediately on first run**
→ One or more Gmail labels are missing. Check the log for which ones and create them.

**Email flagged as `email-error`**
→ The sender's email doesn't match what's in the Rider Master sheet. Check for typos in the sheet or that the rider is emailing from their registered address.

**Bonus not found / `processing-error`**
→ The Bonus ID in the subject doesn't match anything in Column A of the rider's sheet. Check for spacing or capitalization differences.

**`Rider Master` sheet not found**
→ The sheet tab must be named exactly `Rider Master` (case-sensitive, with a space).
