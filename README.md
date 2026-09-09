# snafu-rally-scoring

Automated bonus scoring for rally events via Google Apps Script. Riders submit bonus claims by email — the script validates them, records them in your scoring spreadsheet, and manages everything through Gmail labels. Scorers approve or deny submissions from a Gmail sidebar add-on with one click.

---

## How it works

```
Rider sends email: "42 ABCD"
        |
        v
  rally/unprocessed
        |
  Script scans ALL messages in thread
  (any valid message passes validation)
        |
        |-- bad format on all ---------> rally/subject-line-error
        |-- unregistered sender -------> rally/email-error
        |-- spreadsheet error ---------> rally/processing-error
        |
        v
  rally/email-requires-review   <-- scorer reviews in Gmail sidebar
        |
        |-- Deny (any message) --------> rally/denied
        |                                (X written to sheet immediately)
        |                                (revertable from sidebar)
        v
  rally/approved       <-- scorer clicks Approve in sidebar
        |                  (X written to sheet immediately)
        |                  (revertable from sidebar)
        v
  rally/scored
```

Threads can contain up to three submission attempts. Any message in the thread can be approved or denied individually from the sidebar. Approvals and denials can be reverted back to `needs-review` at any time.

The `rally` parent label is also created so scorers can click it to see every submission at once regardless of status.

---

## Files

| File | Purpose |
|------|---------|
| `code.js` | Main script — `setup()`, `processEmails()`, spreadsheet logic, email validation |
| `Sidebar.gs` | Gmail Add-on sidebar — per-message Approve / Deny / Revert controls |
| `utility.js` | Manually-run helper(s) for testing — see [Utility scripts](#utility-scripts) |
| `appsscript.json` | Apps Script manifest with OAuth scopes and add-on registration |
| `test/` | Unit tests (Node's built-in test runner) against mocked Gmail/Sheets services - see [Testing](#testing) |
| `LICENSE` | GNU GPL v3.0 or later — see [License](#license) |

---

## Testing

`test/support/mocks.js` implements minimal in-memory stand-ins for `GmailApp`, `SpreadsheetApp`, `PropertiesService`, `LockService`, `ScriptApp`, and `CardService`, then runs the real `code.js`/`Sidebar.gs` source unmodified against them via Node's `vm` module. No real Gmail/Sheets access, no network calls.

```
npm test
```

Requires Node 18+ (uses the built-in `node:test` runner - no dependencies to install).

---

## Spreadsheet structure

All configuration is read at runtime from a **Config** sheet.

### Config sheet

A tab named exactly `Config`, with `key`, `value`, `notes` columns, populated by hand from the [Config sheet reference](#config-sheet-reference) below. See [Installation](#installation) for setup steps.

### Rider Master

A tab named **`Rider Master`** (configurable). Row 1 is headers:

| Rider Number | Name | Email |
|---|---|---|
| 42 | Jane Smith | jane@example.com |
| 7 | Bob Jones | bob@example.com |

The sender email on every submission is validated against this sheet.

### Bonus Master

A tab named **`Bonus Master`** (configurable). Column A lists every bonus ID, column B its point value (`POINTS`), one per row (row 1 = header):

| Bonus ID | POINTS |
|----------|--------|
| ABCD     | 100    |
| WXYZ     | 50     |

Rider score sheets reference this tab with cell formulas (`='Bonus Master'!A2` etc.) so any updates to Bonus Master propagate automatically. `POINTS` is read by [Master Scoring](#master-scoring) — a blank or non-numeric value there just counts as 0, logged as a warning only when it's non-blank but not a number.

**Combination bonuses go here too** — see [Combination bonuses](#combination-bonuses) below.

### Master Scoring

A tab named **`Master Scoring`** (configurable via `sheet_master_scoring`), created by `setup()` **after** rider score sheets (so its formulas reference sheets that already exist) and positioned **2nd from the left, right after [Leader Board](#leader-board)** — `setup()` explicitly moves it there once both exist, since Leader Board must be leftmost but can't be created until Master Scoring already exists. This only happens the first time either sheet is created; if you drag either tab elsewhere afterward, a later `setup()` run won't move it back. A grid: one row per bonus (starting at row 5), one column per rider (starting at column C):

| | A | B | C (1st rider) | D (2nd rider) |
|---|---|---|---|---|
| 1 | | Name | `=VLOOKUP(C2,'Rider Master'!$A$2:$B$…,2)` | … |
| 2 | | Number | *(cell reference to Rider Master)* | … |
| 3 | | Score | `=SUMIF(C5:C,"X",$B5:$B)` | … |
| 4 | Bonus | POINTS | | |
| 5+ | *(ref to Bonus Master)* | *(ref to Bonus Master)* | *(ref to that rider's own Approved column)* | … |

Every formula is written directly by `setup()`, one per cell — not something you fill in by hand. Requires **Bonus Master** to already exist (`setup()` logs an error for this step and skips it otherwise, but still creates labels, rider sheets, and the trigger); tolerates Rider Master being missing/empty (nothing created, no error).

`Score`'s range starts at row 5, not row 1 — a true whole-column range (`C:C`) would include row 3, the `Score` formula's own cell, which Sheets rejects as a circular reference even though the criteria could never actually match it.

**Re-running `setup()` grows this sheet** as your roster and bonus list grow: a new rider gets a new column (backfilled for every existing bonus row), a new bonus gets a new row (backfilled for every existing rider column). Existing cells are never rewritten.

**Formatting is reapplied over the whole sheet every run** (not just newly-added rows/columns): every cell is center-aligned, the whole sheet gets alternating row colors ("banding"), and the data grid (row 5+, column C+) gets a conditional format rule that turns a cell green when it contains `X` — i.e. whenever that rider is approved for that bonus. Growing the sheet on a later run redraws all three over the new, larger extent rather than leaving old rows/columns formatted differently from new ones.

**The sheet is protected with "Edit with warning"** the first time it's created — anyone can still edit any cell, but gets a confirmation dialog first, since the whole sheet is regenerated by `setup()`. This only happens once, at creation; a Master Scoring sheet that already existed before this was added doesn't get protection retroactively.

### Rider score sheets

One tab per rider, named by their **rider number** (e.g. `42`). Created automatically by `setup()` for all riders in Rider Master, appended after existing tabs. Rider sheets are also created on-the-fly if a submission arrives for a rider without a sheet.

Default column layout (all configurable):

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Bonus ID | Submitted | Submit Time | Approved | Approve Time | Denied | Deny Time |

Bonus IDs in Column A are cell references to Bonus Master — they update automatically if Bonus Master changes.

Submit Time, Approve Time, and Deny Time (columns C, E, G) are formatted as **Date + Time** (`M/d/yyyy h:mm:ss am/pm`) at the moment the rider sheet is created, so a timestamp always displays with both — not just a date. This only applies going forward: a rider sheet that already existed before this was added, or a bonus row added to Bonus Master *after* a rider's sheet was created, won't have the format applied automatically — reformat those columns by hand (select the column → Format → Number → Date time) if needed.

### Leader Board

A tab named **`Leader Board`** (configurable via `sheet_leader_board`), created by `setup()` as the **leftmost tab** in the workbook — one row per rider in Rider Master:

| A | B | C | D |
|---|---|---|---|
| Rider Number | Name | Score | Finish |

- **Rider Number** is a cell reference to Rider Master (not a copied value).
- **Name** is a `VLOOKUP` against Rider Master.
- **Score** is an `HLOOKUP` against [Master Scoring](#master-scoring), which `setup()` now creates automatically (right before Leader Board, specifically so it's already there) — you shouldn't normally hit the "must already exist" case at all any more; it only still applies if Bonus Master is *also* missing, since that cascades into Master Scoring not being created either.
- **Finish** is the rider's rank by Score, highest first (rank 1 = highest score).

Re-running `setup()` adds a row for any rider newly added to Rider Master since the last run — existing rows are never rewritten. If you rename or restructure Master Scoring after Leader Board already has rows referencing it, those existing formulas keep pointing at the old range; only rows added afterward pick up the new layout.

---

## Gmail labels

All labels are created automatically by `setup()`. The label prefix is configurable (default: `rally`).

| Label | Meaning |
|-------|---------|
| `rally` | Parent label — click to see all submissions at once |
| `rally/unprocessed` | Incoming submission, not yet processed |
| `rally/email-requires-review` | Valid submission awaiting scorer decision |
| `rally/approved` | Scorer approved — X written to sheet immediately |
| `rally/scored` | Fully recorded in the spreadsheet |
| `rally/denied` | Scorer denied — X written to sheet immediately |
| `rally/subject-line-error` | No message in thread matched the required subject format |
| `rally/email-error` | Sender not registered for that rider number |
| `rally/processing-error` | Script error — check Executions log |

---

## Installation

### 1. Create the Config sheet

Create a new tab in your Google Sheet named exactly `Config`. Add the following columns in row 1: `key`, `value`, `notes`. Then populate it using the [Config sheet reference](#config-sheet-reference) below — one row per key.

`setup()` will automatically fill in the `spreadsheet_id` value when it runs, so you can leave that blank.

### 2. Install script files in Apps Script

1. Open your Google Sheet → **Extensions → Apps Script**
2. Paste `code.js` into `Code.gs`
3. Click **+** next to Files → New script → name it `Sidebar` → paste `Sidebar.gs`
4. *(Optional, for testing)* Click **+** next to Files → New script → name it `Utility` → paste `utility.js` — see [Utility scripts](#utility-scripts)
5. Project Settings → check **Show "appsscript.json" manifest file in editor** → paste `appsscript.json`
6. Save all files

### 4. Run setup()

In the Apps Script editor:

1. Select function `setup` from the dropdown
2. Click **Run**
3. Authorise when prompted

`setup()` will:
- Read all config from the Config sheet
- Create the bare `rally` parent label and all sub-labels
- Save the spreadsheet ID to Script Properties (used by the sidebar)
- Create rider score sheets for every rider in Rider Master
- Create or grow the [Master Scoring](#master-scoring) sheet (skipped, with a logged error, if `Bonus Master` doesn't exist yet — everything else above/below still runs)
- Create or update the [Leader Board](#leader-board) sheet (skipped, with a logged error, if `Master Scoring` doesn't exist — which itself only happens if `Bonus Master` was also missing)
- Position Leader Board leftmost and Master Scoring right after it — only the first time either one is created; a tab you've since moved by hand stays put on later runs
- Set the time-driven trigger

### 5. Activate the Gmail sidebar add-on

1. In Apps Script → **Deploy → Test deployments**
2. Click **Install**
3. Open Gmail — the **Rally Scoring** panel will appear on the right when you open any submission email

### 6. Set up a Gmail filter (recommended)

So incoming submissions are labelled automatically without manual intervention:

1. In Gmail → Settings → **See all settings → Filters and Blocked Addresses → Create a new filter**
2. In the **To** field, enter your scoring email address
3. Click **Create filter**
4. Check **Apply the label** → select `rally/unprocessed`
5. Save

---

## Gmail sidebar

The Rally Scoring sidebar opens automatically when you view any submission email. It shows every message in the thread individually, so scorers can act on specific attempts.

### Per-message controls

Each message in the thread shows its subject, sender, date, and:

- **Approve this message** — writes X + timestamp to the Approved column immediately and marks the thread `rally/scored`
- **Deny this message** — writes X + timestamp to the Denied column immediately and marks the thread `rally/denied`

### Revert controls

Shown at the bottom of the sidebar when a thread has been approved or denied:

- **Remove approval / revert to email-requires-review** — clears the Approved column in the sheet and moves the thread back to `rally/email-requires-review`
- **Remove denial / revert to email-requires-review** — clears the Denied column in the sheet and moves the thread back to `rally/email-requires-review`

This allows scorers to approve a different attempt after an initial denial, or correct a mistake.

---

## Submission format

Riders send email to your scoring address with the subject:

```
<RiderNumber> <BonusCode>
```

Examples:

```
42 ABCD
7 WXYZ
101 abcd
15 ABC1
```

- Rider number must be numeric
- Bonus code is either exactly 4 alphabetic characters (e.g. `ABCD`), or exactly 3 alphabetic characters followed by exactly 1 digit (e.g. `ABC1`) - case-insensitive. No other combination is valid (not 4 letters + a digit, not 3 letters + 2 digits)
- Zero or more spaces are allowed between the rider number and the bonus code
- Email must come from the rider's registered address in Rider Master
- Up to three emails per bonus can be sent in the same thread - each can be approved or denied independently from the sidebar

---

## Combination bonuses

A **combination bonus** (or "combo") bundles several regular bonuses under one code — for example, a bonus that only counts once a rider has also claimed a specific set of other bonuses. There's nothing special about it as far as this script is concerned: add its code to **Bonus Master** exactly like any other bonus. A rider claims it exactly like a regular bonus too — same `<Rider#> <code>` subject format (see [Submission format](#submission-format)), **no photo required** for the combo email itself (the script never checks for photo attachments on any submission, so this is just a note for riders, not a code difference).

**The script does not know a combo is a combo, does not track which bonuses it depends on, and does not check whether those have been scored.** It's up to the scorer to verify a combo's requirements are actually met — by checking the rider's sheet or however you track it — before clicking Approve in the sidebar. There is no automated gate; the sidebar shows the same Approve/Deny buttons it would for any bonus.

What the combo (and its components) are actually worth in points, and totalling that up, happens entirely in your own scoring spreadsheet (Master Scoring/Leader Board) — this script only tracks submit/approve/deny state, for combos exactly as it does for regular bonuses.

---

## Config sheet reference

All values are editable directly in the Config sheet after setup.

| Key | Default | Notes |
|-----|---------|-------|
| `event_name` | | Display only |
| `organizer_email` | | Display only |
| `spreadsheet_id` | | Written automatically by `setup()` — do not change |
| `sheet_rider_master` | `Rider Master` | Tab name of the rider roster |
| `sheet_bonus_master` | `Bonus Master` | Tab with all bonus IDs in column A |
| `sheet_master_scoring` | `Master Scoring` | Tab name for [Master Scoring](#master-scoring), created automatically |
| `sheet_leader_board` | `Leader Board` | Tab name for the [Leader Board](#leader-board) |
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
| `label_format_error` | `rally/subject-line-error` | |
| `label_email_error` | `rally/email-error` | |
| `label_processing_error` | `rally/processing-error` | |
| `label_needs_review` | `rally/email-requires-review` | |
| `label_approved` | `rally/approved` | |
| `label_denied` | `rally/denied` | |
| `label_scored` | `rally/scored` | |

---

## Changing settings after setup

Edit values directly in the Config sheet — the script reads config live on every execution so most changes take effect immediately with no code changes needed.

To change the trigger interval or label names, edit the Config sheet then re-run `setup()` to apply.

To reconfigure from scratch, edit the Config sheet values directly and re-run `setup()`.

> Note: If you have an existing deployment using the older label names rally/needs-review or rally/format-error, update label_needs_review and label_format_error in the Config sheet to rally/email-requires-review and rally/subject-line-error, then re-run setup(). This creates the new Gmail labels - existing emails keep their old labels until relabelled manually or reprocessed.

---

## Removing and reinstalling the add-on

If you need to reset permissions, remove an old version, or force a fresh auth dialog:

### Remove app permissions

1. Go to **https://myaccount.google.com/connections**
2. Find your Apps Script project (may appear as "Rally Scoring" or the script project name)
3. Click it and select **Delete all connections you have with Rally Scoring**

This fully revokes OAuth access. The next time you run `setup()` or install the add-on, Google will show a fresh authorisation dialog.

### Remove the Gmail sidebar add-on

1. Open Gmail
2. Click the **+** (Get add-ons) icon in the right sidebar
3. Find **Rally Scoring** and click the gear icon
4. Select **Uninstall**

Or from Apps Script:

1. **Deploy → Manage deployments**
2. Delete any active test or production deployments

### Reinstall after removing

1. Re-open your Apps Script project
2. **Deploy → Test deployments → Install**
3. Authorise when prompted

---

## Utility scripts

`utility.js` holds manually-run helpers for testing — not part of the app's normal
operation, and never called by `setup()`, `processEmails()`, or the sidebar. Add it
as its own script file per [Installation step 4](#2-install-script-files-in-apps-script)
so its functions are selectable from the Apps Script editor's function dropdown.

### Delete all rider sheets

`deleteAllRiderSheets()` removes every sheet **except** the ones listed in its
`keepSheets` array (`Config`, `Rider Master`, `Bonus Master`) —
i.e. every rider sheet, **and** `Leader Board`, **and** `Master Scoring`. All three
are deleted right along with the rider sheets, not kept: they're all regenerated
from scratch the next time `setup()` runs, exactly like a missing rider sheet is.
Useful between test runs to reset scoring data without touching your roster/bonus
setup.

**Before running it, open `utility.js` and confirm `keepSheets` actually matches your
spreadsheet's non-rider tab names.** It's an exact-string `.includes()` check with no
confirmation prompt — a tab name that doesn't match (a typo, a rename, an extra
scoring tab you added later) gets deleted right along with the rider sheets, with no
way to undo it beyond Google Sheets' version history.

Run it from the Apps Script editor: select `deleteAllRiderSheets` from the function
dropdown → **Run**.

---

## Troubleshooting

| Symptom | Check |
|---------|-------|
| "Config sheet not found" | Create a tab named exactly `Config` and populate it per [Installation](#installation), then re-run `setup()` |
| Script stops immediately | One or more Gmail labels are missing — re-run `setup()` |
| Email tagged `subject-line-error` | Subject must be a number followed by a bonus code - 4 letters, or 3 letters + 1 digit (any case, any amount of spacing) |
| Email tagged `email-error` | Sender address doesn't match Rider Master — check for typos or alias issues |
| Email tagged `processing-error` | Apps Script → Executions → click the failed run for the full error |
| Bonus ID not found | Check Column A of the rider's sheet — spacing and capitalisation must match exactly |
| `must be at least 1, got: "0"` (or similar) when approving/denying | A `col_*` or `header_row` key in the Config sheet is `0` or negative — columns and rows are numbered starting at 1. Fix the value in Config; no need to re-run `setup()` |
| Rider Master sheet not found | Tab must be named exactly as configured (default: `Rider Master`) |
| `Cannot create Master Scoring - "Bonus Master" not found` | Create your `Bonus Master` tab (with a `POINTS` column) and re-run `setup()` — everything else `setup()` does still completes even with this error |
| `Cannot create Leader Board - "Master Scoring" not found` | Only happens if the above also failed — fix `Bonus Master` first, then re-run `setup()`; Master Scoring and Leader Board both get created in that same run |
| Sidebar not appearing | Deploy → Test deployments → Install |
| Sidebar shows stale labels | Gmail doesn't always refresh the thread view instantly — the labels are updated on the server; reload the page to see the current state |
| Auth dialog not appearing | Go to https://myaccount.google.com/connections, remove the connection, then re-run `setup()` |
| Trigger not running | Apps Script → Triggers — confirm `processEmails` exists; re-run `setup()` if not |

### Viewing logs

Apps Script → **Executions** in the left sidebar → click any run to expand its log. The script logs every action, validation result, and error with a clear message — this includes sidebar button clicks (Approve/Deny/Revert), not just the time-driven trigger: every guard clause and failure in the sidebar's handlers writes a log line naming the function and the thread/message involved, so a scorer-facing error toast always has a matching entry here to dig into.

---

## OAuth scopes

| Scope | Used for |
|-------|---------|
| `spreadsheets` | Read/write scoring spreadsheet |
| `gmail.modify` | Read emails, manage labels |
| `gmail.addons.*` | Gmail sidebar add-on |
| `script.scriptapp` | Create time-driven trigger in `setup()` |
| `script.external_request` | Reserved for future use — nothing currently calls `UrlFetchApp` |
| `script.locale` | Required because the manifest sets `useLocaleFromApp: true` |

If you already installed the add-on before `script.locale` was added to `appsscript.json`, Apps Script will prompt for re-authorization the next time you run a function or open the sidebar — accept it once to pick up the new scope. See [Removing and reinstalling the add-on](#removing-and-reinstalling-the-add-on) if it doesn't prompt on its own.

---

## License

GNU General Public License v3.0 or later — see [LICENSE](LICENSE).