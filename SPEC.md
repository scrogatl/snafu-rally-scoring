# Rally Scoring — Technical Specification

This document specifies the behavior of the Rally Scoring system precisely enough that
implementing it from scratch — in Google Apps Script, against the Gmail and Sheets
APIs — should reproduce the existing codebase (`code.js`, `Sidebar.gs`, `appsscript.json`)
to within cosmetic variation. Where behavior is load-bearing (regexes, config keys,
label names, button text, error messages), the exact value is given rather than a
paraphrase. Treat every fenced code block and every table row as normative.

## 1. Purpose

Riders participating in a rally submit bonus claims by email. The system must:

1. Receive each submission by email and classify it automatically (valid submission,
   malformed subject, unregistered sender, or internal processing error).
2. Let a human scorer approve or deny each submission from inside Gmail, with no
   separate tool to open.
3. Record the outcome (submitted / approved / denied, each with a timestamp) in a
   Google Sheet that already contains the rider roster and the bonus list.
4. Make every state visible and reversible via Gmail labels, so the current status of
   any submission is inspectable without opening the spreadsheet, and any scoring
   decision can be undone.
5. Require no code changes for ordinary operation — column layout, label names, sheet
   names, and timing are all read from a configuration sheet at runtime.

## 2. System Overview

This is a Google Apps Script project bound to a Google Sheet, with a Gmail Add-on
component. There is no external server, database, or network call of any kind other
than the Google APIs Apps Script itself provides (`GmailApp`, `SpreadsheetApp`,
`PropertiesService`, `LockService`, `ScriptApp`, `CardService`).

```
Rider sends email: "42 ABCD"
        |
        v
  rally/unprocessed                          (applied by a Gmail filter, not the script)
        |
  Time-driven trigger runs processEmails() every N minutes
  Script scans ALL messages in the thread (any one valid message is enough to pass)
        |
        |-- bad subject format on every message --------> rally/subject-line-error
        |-- format OK but sender not registered ---------> rally/email-error
        |-- format OK, sender OK, but sheet write fails --> rally/processing-error
        |
        v
  rally/email-requires-review   <-- scorer opens the email; sidebar shows Approve/Deny
        |
        |-- Deny (any message, with confirmation) -------> rally/denied
        |                                                   (X + timestamp written immediately)
        |                                                   (revertable from the sidebar)
        v
  rally/approved  <-- scorer clicks Approve in the sidebar
        |             (X + timestamp written immediately; also marks rally/scored)
        v
  rally/scored
```

A thread may contain more than one submission message (typically retries of the same
bonus). Every message in a thread can be approved or denied **individually** from the
sidebar; approvals and denials can be reverted back to `rally/email-requires-review`
at any time.

## 3. Repository Layout

| File | Role |
|---|---|
| `code.js` | Config loading, `setup()`, rider-sheet creation, the `processEmails()` trigger handler and everything it calls, spreadsheet read/write, email validation. Loaded into the Apps Script project as `Code.gs`. |
| `Sidebar.gs` | The Gmail Add-on: the contextual-trigger entry point, card construction, and every button's click handler. Added as a second script file named `Sidebar`. |
| `appsscript.json` | Manifest: OAuth scopes, add-on registration, runtime version, time zone. |

Apps Script concatenates every script file into one global scope at runtime. Any
top-level `function` declared in `code.js` is directly callable from `Sidebar.gs` and
vice versa — there is no module system and no imports. **A helper used by both files
must be defined in exactly one of them** (see §7.3, `loadLabels_`); do not duplicate a
helper across files, since the two copies can silently drift.

## 4. Data Model (Spreadsheet Structure)

The bound spreadsheet has four kinds of sheets: `Config`, a rider roster, a bonus
list, and one sheet per rider. Sheet names other than `Config` are configurable (§5).

### 4.1 Config sheet

Sheet name: **exactly `Config`**, not configurable. Row 1 is a header row (`key`,
`value`, `notes` — the loader ignores row 1 entirely and does not check its contents).
Every subsequent row is one setting:

| key | value | notes |
|---|---|---|
| `label_parent` | `rally` | Bare parent label |
| `col_bonus_id` | `1` | Column A |
| … | … | … |

Loading rule: trim column A and column B of every row after row 1; if the trimmed key
is non-empty, `config[key] = value` (also trimmed). Rows with a blank key are skipped
entirely (not even stored under an empty-string key). Keys are case-sensitive and
compared verbatim.

### 4.2 Rider Master sheet

Default name `Rider Master` (configurable via `sheet_rider_master`). Row 1 is a header
row whose cell text must match the configured column-name keys exactly
(`master_col_rider_number`, default `Rider Number`; `master_col_email`, default
`Email`). Example:

| Rider Number | Name | Email |
|---|---|---|
| 42 | Jane Smith | jane@example.com |
| 7 | Bob Jones | bob@example.com |

The `Name` column is never read by the script; it exists purely for the spreadsheet
owner's benefit. Every other column in this sheet is likewise ignored — the script
only ever looks up the two configured columns by header name.

### 4.3 Bonus Master sheet

Default name `Bonus Master` (configurable via `sheet_bonus_master`). Column A is a
list of bonus IDs, one per row, row 1 a header (its text is never checked). A bonus ID
must be a value that can be produced by the submission-format rule in §6 — i.e.
effectively 4 letters, since that's what a rider can type and what the matching logic
in §7.5 compares against. Example:

| Bonus ID |
|---|
| ABCD |
| WXYZ |

### 4.4 Rider score sheets

One tab per rider, named by the rider's **exact rider number string** (e.g. `42`).
Created automatically (§7.2) for every rider present in Rider Master, and created
on-the-fly (§7.5) the first time a submission arrives for a rider who doesn't have one
yet. Column layout, with a header row at `header_row` (default row 1) that is frozen:

| Column | Header text | Written by |
|---|---|---|
| A (`col_bonus_id`) | `Bonus ID` | Formula, at sheet-creation time only |
| B (`col_submitted`) | `Submitted` | `handleUnprocessedThread` |
| C (`col_submitted_time`) | `Submit Time` | `handleUnprocessedThread` |
| D (`col_approved`) | `Approved` | `handleApprove` / `addApprovedCheck` |
| E (`col_approved_time`) | `Approve Time` | `handleApprove` / `addApprovedCheck` |
| F (`col_denied`) | `Denied` | `handleDeny` |
| G (`col_denied_time`) | `Deny Time` | `handleDeny` |

Column A of a rider sheet is **not** a copy of Bonus Master's values — it is written
as one formula per data row, `='<Bonus Master sheet name>'!A<n>` (single quotes around
the sheet name, always column A, `n` = the corresponding Bonus Master row), so edits to
Bonus Master after the fact propagate to every rider sheet automatically. New sheets
are inserted **after every existing sheet** (`ss.insertSheet(riderNumber,
ss.getSheets().length)`), so pre-existing tabs (`Leaderboard`, `Master Scoring`, etc.)
are never disturbed.

**Important, easy-to-miss inconsistency to preserve exactly, not "fix":**
`createRiderSheet_` writes the 7-column header row and the bonus-ID formulas at
hardcoded columns 1–7 — `sheet.getRange(headerRow, 1, 1, 7)` for the header and
`sheet.getRange(headerRow + 1, 1, formulas.length, 1)` for the formulas, both with a
literal `1` as the column argument, never `configInt_(config, 'col_bonus_id', ...)` or
any other `col_*` key. Only `header_row` (which *row* the header lands on) is actually
threaded through from config here. The `col_bonus_id`/`col_submitted`/`col_approved`/
`col_denied` keys (and their `_time` counterparts) are consulted **only** by
`updateSpreadsheet` (§7.5) when it later reads or writes a specific cell — they have
no effect on where `createRiderSheet_` places anything. Practical consequence: the
"seven independently configurable columns" story in §5 is only trustworthy if every
`col_*` key is left at its default (1/2/3/4/5/6/7, in that order) or the rider sheet's
actual header row is edited to match by hand — reconfiguring, say, `col_approved` to
`10` does not move the "Approved" header; it just makes `handleApprove` write `X` into
column 10 while the sheet still shows "Approved" at column 4. Reproduce this
disconnect faithfully rather than parameterizing `createRiderSheet_`'s column
placement to "fix" it — that would be a behavior change, not a match to this
codebase.

## 5. Configuration Reference

Every config value is a string; numeric ones are parsed at the point of use (never
cached), so edits to the Config sheet take effect on the next run without redeploying.

Two lookup helpers, both defined in `code.js`:

```js
// Non-numeric keys read with a hardcoded `|| 'Default Name'` fallback inline at
// each call site (loadLabels_, createAllRiderSheets_, createRiderSheet_,
// validateEmailAddress) — there is no shared string-default helper.

// Numeric keys go through this helper everywhere:
function configInt_(config, key, defaultValue) {
  const raw = config[key];
  if (raw === undefined || raw === '') {
    if (defaultValue !== undefined) return defaultValue;
    throw new Error('Config key "' + key + '" is missing.');
  }
  const parsed = parseInt(raw, 10);
  if (isNaN(parsed)) throw new Error('Config key "' + key + '" is not a number: "' + raw + '"');
  return parsed;
}
```

Rule: a blank or absent numeric key silently falls back to its documented default; a
*present but non-numeric* value throws immediately with the exact message shape above
(callers propagate this — they do not catch and re-default it).

| Key | Default (when blank/absent) | Read by |
|---|---|---|
| `event_name` | — (display only, never read by code) | — |
| `organizer_email` | — (display only, never read by code) | — |
| `spreadsheet_id` | written by `setup()` | not read by code at runtime (Script Properties is the source of truth — see §7.1) |
| `sheet_rider_master` | `Rider Master` | `createAllRiderSheets_`, `validateEmailAddress` |
| `sheet_bonus_master` | `Bonus Master` | `createRiderSheet_` |
| `master_col_rider_number` | `Rider Number` | `createAllRiderSheets_`, `validateEmailAddress` |
| `master_col_email` | `Email` | `validateEmailAddress` |
| `header_row` | `1` | `createRiderSheet_` (which row the header/formulas land on), `updateSpreadsheet` (where data rows start) |
| `col_bonus_id` | `1` | `updateSpreadsheet` only — **not** read by `createRiderSheet_`, which always writes bonus-ID formulas to column 1 regardless (see §4.4's callout) |
| `col_submitted` | `2` | `handleUnprocessedThread` |
| `col_submitted_time` | `3` | `handleUnprocessedThread` |
| `col_approved` | `4` | `addApprovedCheck`, `handleApprove`, `handleRevertApproved` |
| `col_approved_time` | `5` | same as above |
| `col_denied` | `6` | `handleDeny`, `handleRevertDenied` |
| `col_denied_time` | `7` | same as above |
| `trigger_interval_min` | `10` | `setup()` |
| `label_parent` | `rally` | `setup()` |
| `label_unprocessed` | `rally/unprocessed` | `loadLabels_` |
| `label_format_error` | `rally/subject-line-error` | `loadLabels_` |
| `label_email_error` | `rally/email-error` | `loadLabels_` |
| `label_processing_error` | `rally/processing-error` | `loadLabels_` |
| `label_needs_review` | `rally/email-requires-review` | `loadLabels_` |
| `label_approved` | `rally/approved` | `loadLabels_` |
| `label_denied` | `rally/denied` | `loadLabels_` |
| `label_scored` | `rally/scored` | `loadLabels_` |

`loadLabels_` treats every one of the eight `label_*` keys as **required**: a missing
config key or a Gmail label that doesn't yet exist both count as "missing," and if
anything is missing the whole function returns `null` (after logging every missing
item in one message) rather than a partial object. Every caller of `loadLabels_` must
treat a `null` return as "not configured yet, abort this operation" — never proceed
with a partially-populated labels object.

## 6. Submission Format & Validation

The subject line is the entire payload. Two regexes are authoritative and must match
exactly (both operate on `subject.trim()`):

```js
// Format check (does this look like a submission at all?)
function isValidSubject(subject) {
  return /^\d+\s*[A-Za-z]{4}$/.test(subject.trim());
}

// Format check + extraction, in one step
const match = subject.trim().match(/^(\d+)\s*([A-Za-z]{4})$/);
// match[1] -> rider number (kept as a string, digits only)
// match[2] -> bonus code, upper-cased before use
```

Rules implied by the regex (state these explicitly in any reimplementation, since the
regex alone is easy to get subtly wrong):

- The rider number is one or more digits, with no other rider-number validation
  (no range check, no leading-zero handling beyond what `\d+` naturally allows).
- Zero or more whitespace characters are allowed between the number and the code
  (`\s*`, so `"42ABCD"` and `"42     ABCD"` are both valid).
- The bonus code is **exactly** 4 alphabetic characters, letters only — no digits,
  hyphens, or underscores anywhere in it. It is upper-cased before being compared or
  stored, so `"abcd"`, `"AbCd"`, and `"ABCD"` are the same bonus.
- Nothing may precede the rider number or follow the bonus code — the regex is
  fully anchored (`^...$`) against the trimmed string, so trailing text
  (`"42 ABCD extra"`) is invalid.

Sender validation (`validateEmailAddress`) is separate from format validation:

```js
const EMAIL_RE = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/;
```

Extract the first email-shaped substring from the raw `From` header (which may be a
display-name-plus-angle-bracket form like `"Jane Smith <jane@example.com>"`), lower-case
it, look up the rider by rider number in Rider Master, extract the first email-shaped
substring from that row's registered email cell the same way, lower-case it, and
require an exact string match. No SPF/DKIM or other authenticity check is performed —
this validates the `From` header Gmail already parsed, nothing more. Any of "no
email-shaped substring in the sender string," "rider number not found in Rider
Master," "no email-shaped substring in the registered cell," or "emails don't match"
all resolve to `false` (with a distinct `Logger.log` line per case for diagnosability,
no thrown errors).

## 7. Functional Specification

### 7.1 `setup()` — one-time initialization, re-runnable

Manually invoked from the Apps Script editor. In order:

1. Save the bound spreadsheet's ID to Script Properties under key `SPREADSHEET_ID`
   (this — not the Config sheet's `spreadsheet_id` value — is the runtime source of
   truth every other entry point uses to find the spreadsheet).
2. Load Config; on failure, log `'FATAL: ' + e.message` plus a hint to create/populate
   the Config sheet, and **return** (do not throw out of `setup()`).
3. If the Config sheet has a row whose key is exactly `spreadsheet_id`, overwrite that
   row's value cell with the real ID (cosmetic — for the human reading the sheet — not
   read back by the code).
4. Ensure the bare parent label (`label_parent`, default `rally`) exists.
5. For each of the eight `label_*` keys, ensure the corresponding Gmail label exists;
   log a per-key line either way; log one summary line
   `'Labels: created=' + created + ', skipped=' + skipped'`. A `label_*` key with no
   configured value is logged (`'Config key missing: ' + key`) and skipped — it does
   **not** abort setup.
6. Call `createAllRiderSheets_` (§7.2).
7. Delete every existing project trigger whose handler function is
   `'processEmails'`, then create exactly one new one:
   `ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(intervalMin).create()`,
   where `intervalMin = configInt_(config, 'trigger_interval_min', 10)`. This makes
   `setup()` idempotent with respect to triggers — re-running it never produces
   duplicate triggers.
8. Log `'Setup complete. Trigger set for every ' + intervalMin + ' minutes.'`.

### 7.2 Rider sheet creation

```js
function createAllRiderSheets_(ss, config) {
  // 1. Look up Rider Master (config['sheet_rider_master'] || 'Rider Master').
  //    Missing sheet, or fewer than 2 rows -> log and return, no error thrown.
  // 2. Find the rider-number column by header text
  //    (config['master_col_rider_number'] || 'Rider Number'); if not found,
  //    log and return.
  // 3. For each data row: trim the rider-number cell; blank -> skip.
  //    Sheet with that exact name already exists -> skip (count as "skipped").
  //    Otherwise call createRiderSheet_, catching and logging any error per rider
  //    (one rider's failure does not stop the loop).
  // 4. Log one summary line: 'Rider sheets: created=' + created + ', skipped=' + skipped.
}
```

```js
function createRiderSheet_(ss, config, riderNumber) {
  // 1. Look up Bonus Master (config['sheet_bonus_master'] || 'Bonus Master');
  //    missing -> throw new Error('Cannot create rider sheet - "' + bonusMasterName +
  //    '" not found. Create a sheet with one bonus ID per row in column A.').
  // 2. Insert a new sheet named exactly `riderNumber`, appended after every existing
  //    sheet: ss.insertSheet(riderNumber, ss.getSheets().length).
  // 3. Write the 7-column header row (exact text, §4.4 table) at
  //    configInt_(config, 'header_row', 1), bold, and freeze that many rows.
  // 4. If Bonus Master has more than 1 row (i.e. at least one bonus), write one
  //    formula per bonus into column A starting at headerRow + 1:
  //    "='" + bonusMasterName + "'!A" + (i + 2)   for i = 0 .. numBonuses-1
  // 5. Auto-resize column 1. Return the created sheet.
}
```

### 7.3 `processEmails()` — the trigger handler

```js
function processEmails() {
  // 1. ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties()
  //         .getProperty('SPREADSHEET_ID'))
  // 2. config = loadConfig(ss); on failure, log 'FATAL: ...' and return (no throw).
  // 3. labels = loadLabels_(config); if null, return silently (loadLabels_ already
  //    logged why).
  // 4. threads = GmailApp.search('label:' + config['label_unprocessed'])
  //    for each thread: handleUnprocessedThread(ss, config, thread, labels)
  // 5. threads = GmailApp.search('label:' + config['label_approved'] +
  //              ' -label:' + config['label_scored'])
  //    for each thread: addApprovedCheck(ss, config, thread, labels)
  // 6. Log a start line, a per-search count line, and a 'Processing complete.' line.
}
```

`loadLabels_` (shared, defined once — see §3's rule about not duplicating helpers):

```js
function loadLabels_(config) {
  const defs = {
    unprocessed: 'label_unprocessed', formatError: 'label_format_error',
    emailError: 'label_email_error', processingError: 'label_processing_error',
    needsReview: 'label_needs_review', approved: 'label_approved',
    denied: 'label_denied', scored: 'label_scored',
  };
  // For each entry: config key missing -> add to `missing`. Otherwise look up the
  // Gmail label by that name; not found -> add 'Gmail label: ' + name to `missing`;
  // found -> labels[shortKey] = the label object.
  // If `missing` is non-empty: Logger.log('ERROR: Missing - ' + missing.join(', ')
  //   + '. Run setup() first.'); return null.
  // Else return the labels object (8 keys, each a real Gmail label object).
}
```

### 7.4 `handleUnprocessedThread(ss, config, thread, labels)`

Scans **every** message in the thread, not just the latest one, looking for the first
one that is both format-valid and sender-registered:

```js
function handleUnprocessedThread(ss, config, thread, labels) {
  // for each message in thread.getMessages():
  //   if !isValidSubject(subject): note formatError=true; continue
  //   data = extractEmailData(message)
  //   if !validateEmailAddress(...): note emailError=true; continue
  //   validMessage = message; break   // first hit wins, stop scanning
  //
  // if validMessage found:
  //   try:
  //     updateSpreadsheet(ss, config, data,
  //       configInt_(config,'col_submitted',2), configInt_(config,'col_submitted_time',3),
  //       /* useEmailTime */ true)
  //     thread.addLabel(needsReview); thread.removeLabel(unprocessed); thread.refresh()
  //   catch (e):
  //     thread.addLabel(processingError); thread.removeLabel(unprocessed); thread.refresh()
  // else if emailError: thread.addLabel(emailError); thread.removeLabel(unprocessed); refresh()
  // else if formatError: thread.addLabel(formatError); thread.removeLabel(unprocessed); refresh()
}
```

Priority when no message is fully valid: **email error beats format error** — i.e. if
any message in the thread had a valid format (even if a different message had a bad
format), the thread is labeled `email-error`, not `subject-line-error`. This is a
direct consequence of `emailError` being checked before `formatError` in the
`else if` chain above; preserve that order exactly.

`useEmailTime = true` here: the submission timestamp is the email's own `getDate()`,
not `new Date()` at processing time — a submission that arrives at 9:00 and isn't
processed until the trigger fires at 9:10 still records 9:00.

### 7.5 `updateSpreadsheet(ss, config, data, columnIndex, timeColumnIndex, useEmailTime, value)`

The single choke point for every write to a rider sheet — called from
`handleUnprocessedThread`, `addApprovedCheck`, `handleApprove`, `handleDeny`,
`handleRevertApproved`, and `handleRevertDenied`. No other function ever calls
`sheet.getRange(...).setValue(...)` on a rider sheet.

```js
function updateSpreadsheet(ss, config, data, columnIndex, timeColumnIndex, useEmailTime, value) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);       // up to 30s; throws if it can't acquire in time
  try {
    let sheet = ss.getSheetByName(data['rider-number']);
    if (!sheet) sheet = createRiderSheet_(ss, config, data['rider-number']);

    const startRow = configInt_(config, 'header_row', 1) + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) throw new Error('No data rows in sheet ' + data['rider-number']);

    const bonusCol = configInt_(config, 'col_bonus_id', 1);
    const values = sheet.getRange(startRow, bonusCol, lastRow - startRow + 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim().toUpperCase() === data['bonus'].trim().toUpperCase()) {
        const row = startRow + i;
        if (value === null) {
          sheet.getRange(row, columnIndex).clearContent();
          sheet.getRange(row, timeColumnIndex).clearContent();
        } else {
          sheet.getRange(row, columnIndex).setValue('X');
          sheet.getRange(row, timeColumnIndex).setValue(useEmailTime ? data.date : new Date());
        }
        return;
      }
    }
    throw new Error('Bonus ID "' + data['bonus'] + '" not found in sheet ' + data['rider-number']);
  } finally {
    lock.releaseLock();       // always runs, including on the throw paths above
  }
}
```

Contract callers must follow:

- Pass `value: null` (the 7th, optional argument) to **revert** — this clears both
  the value and the timestamp cell instead of writing them. Omit it (i.e. pass
  `undefined`) for a normal write, which always writes the literal string `'X'`.
- `columnIndex` and `timeColumnIndex` are always resolved by the caller via
  `configInt_` with the matching default from §5 — this function itself has no
  built-in column defaults and never assumes the timestamp column is
  `columnIndex + 1`.
- Bonus-ID matching is case-insensitive and whitespace-trimmed on both sides.
- If the rider sheet doesn't exist yet, it is created on-the-fly via
  `createRiderSheet_` (inside the lock, so two concurrent first-submissions for the
  same brand-new rider can't both try to insert a sheet with that name).
- The whole read-find-write sequence, including the on-the-fly sheet creation, holds
  one script-wide `LockService` lock so the time-driven trigger and interactive
  sidebar button clicks can never race on the same row. The lock is released in a
  `finally` block, so it is released even when the function throws (missing rider
  sheet rows, bonus ID not found, or a `configInt_` failure).

### 7.6 `addApprovedCheck(ss, config, thread, labels)`

This is the **fallback** path for a thread that was marked `label_approved` some way
other than the sidebar's Approve button — e.g. a scorer manually applying the Gmail
label. (The sidebar's own Approve button already applies `label_scored` in the same
click, so a sidebar-approved thread never matches the search this function is fed
from: `label:approved -label:scored`.)

```js
function addApprovedCheck(ss, config, thread, labels) {
  const validMessages = thread.getMessages().filter(m => isValidSubject(m.getSubject()));
  if (!validMessages.length) { /* log and return, untouched */ return; }

  const data = extractEmailData(validMessages[0]);
  if (!data['rider-number'] || !data['bonus']) { /* log and return */ return; }

  // Ambiguity guard: a manually-applied thread-level label carries no information
  // about which message it was meant for. If the valid messages in this thread
  // don't all agree on the same rider+bonus, do NOT guess — flag it instead.
  const ambiguous = validMessages.some(m => {
    const d = extractEmailData(m);
    return d['rider-number'] !== data['rider-number'] || d['bonus'] !== data['bonus'];
  });
  if (ambiguous) {
    thread.addLabel(labels.processingError);
    thread.refresh();
    return;   // no sheet write; the thread stays approved+unscored and will be
              // re-examined (and re-flagged) on every future run until a human
              // resolves it, e.g. via the sidebar's per-message Approve button.
  }

  try {
    updateSpreadsheet(ss, config, data,
      configInt_(config, 'col_approved', 4), configInt_(config, 'col_approved_time', 5), false);
    thread.removeLabel(labels.needsReview);
    thread.addLabel(labels.scored);
    thread.refresh();
  } catch (e) {
    thread.addLabel(labels.processingError);
    thread.refresh();
  }
}
```

If all valid messages in the thread agree on rider+bonus (the documented normal case —
multiple retries of the same bonus), it proceeds exactly like a single-message thread.

### 7.7 Gmail Add-on entry point — `buildAddOn(e)` (`Sidebar.gs`)

Registered in the manifest as the unconditional Gmail contextual trigger (§9), so it
runs every time a user opens any email while the add-on is installed.

```js
function buildAddOn(e) {
  const message = GmailApp.getMessageById(e.gmail.messageId);
  const thread = message.getThread();
  const messages = thread.getMessages();

  // Resolve ss/config; on failure show errorCard_('Config error', <message> +
  //   ' - make sure setup() has been run.') and stop.
  // If NO message in the thread has a valid-format subject, show
  //   infoCard_('Not a rally submission', <this message's subject>,
  //     'No message in this thread matches the required format: Rider Number ' +
  //     'followed by a 4-letter Bonus Code.') and stop.
  // Otherwise: threadLabels = thread.getLabels().map(name); threadStatus =
  //   getStatus_(threadLabels, config); return buildScoringCard_(messages,
  //   threadStatus, thread.getId(), config, threadLabels, ss).
}
```

`errorCard_(subtitle, msg)` and `infoCard_(subtitle, subject, msg)` are small,
single-section card builders — `infoCard_` additionally shows the subject as a
key/value widget. Both always title the card `'Rally Scoring'`.

### 7.8 `getStatus_(threadLabels, config)` — status precedence

A pure function (no side effects, no API calls) — given the thread's current label
names and the config, returns exactly one of these strings, checked in this order,
first match wins:

```js
function getStatus_(threadLabels, config) {
  const has = n => threadLabels.includes(n);
  if (has(config['label_scored']))           return 'Scored';
  if (has(config['label_approved']))         return 'Approved - pending score';
  if (has(config['label_denied']))           return 'Denied';
  if (has(config['label_needs_review']))     return 'Email Requires Review';
  if (has(config['label_processing_error'])) return 'Bonus Code Error';
  if (has(config['label_email_error']))      return 'Email error';
  if (has(config['label_format_error']))     return 'Subject Line Error';
  if (has(config['label_unprocessed']))      return 'Unprocessed';
  return 'Unknown';
}
```

`scored` outranks `approved` deliberately — a thread that is both (the normal
steady state after a successful approve) reports as `'Scored'`, not
`'Approved - pending score'`.

### 7.9 `buildScoringCard_(messages, threadStatus, threadId, config, threadLabels, ss)`

Builds the card the user actually sees. The card header (not a section) is always
title `'Rally Scoring'`, subtitle `'Thread: ' + messages.length + ' message(s)'` —
same title as `errorCard_`/`infoCard_` (§7.7), but with this card's own subtitle.
Section order, exactly:

1. **"Thread status"** section — one key/value widget, top label `'Status'`,
   content = `threadStatus`.
2. **"Error"** section — present **only** if the thread has one of
   `label_format_error` / `label_email_error` / `label_processing_error`. Exactly one
   paragraph, checked in that same priority order, exact text:
   - format error: `'Subject line does not match the required format: Rider Number followed by a 4-letter Bonus Code (e.g. "42 ABCD"). Ask the rider to resend with the correct subject.'`
   - email error: `'The sender email does not match the registered address for this rider number in Rider Master. Verify the rider's registered email or check for a typo.'`
   - processing error: `'A bonus code error occurred while recording this submission - the bonus code could not be matched in the rider sheet. Check the Apps Script Executions log for details.'`
3. **One section per message**, in thread order, header `'Message ' + (idx+1)`, with
   `' (latest)'` appended for the last one. Each section always shows three key/value
   widgets — Subject, From, Date (`message.getDate().toLocaleString()`) — then exactly
   one of:
   - invalid format: paragraph `'Invalid format - not a rally submission.'`
   - `isScored`: paragraph `'Thread is scored. Use Revert below to undo.'`
   - `isDenied`: paragraph `'Thread is denied. Use Revert below to undo.'`
   - has an error label: paragraph `'Approve/Deny unavailable while an error label is present. See Error section above.'`
   - `isUnprocessed`: paragraph `'Waiting for the next processing run to validate this submission. Approve/Deny will be available once it moves to Email Requires Review.'`
   - otherwise (the normal "awaiting a decision" state): an **Approve this message**
     button (filled style, background `#1A56AA`, action `handleApprove` with
     parameters `{msgId, threadId}`) — omitted if the thread is already `isApproved` —
     followed always by a **Deny this message** button (default style, action
     `showDenyConfirmation` with the same parameters).
4. **"Revert"** section — present only if the thread is approved, denied, or scored.
   Buttons use `messages[0].getId()` as `firstMsgId` regardless of which message was
   actually approved/denied (see §7.10 for why that's safe):
   - if approved or scored: **Remove approval / revert to email-requires-review**
     (action `handleRevertApproved`, parameter `{firstMsgId}`)
   - if denied: **Remove denial / revert to email-requires-review** (action
     `handleRevertDenied`, parameter `{firstMsgId}`)

`showDenyConfirmation(e)` pushes a second card (`CardService...pushCard`, not a
replace) titled `'Confirm Bonus Denial'`, subtitle `'This will mark the bonus as
denied'`, one paragraph — exact text `'Are you sure you want to deny this bonus
submission? This will record a denial in the spreadsheet and apply the denied
label.'` — a filled red (`#B3261E`) **Confirm Bonus Denial** button (action
`handleDeny`) and a plain **Cancel** button (action `cancelDeny`, parameter `{msgId}`
only). On success, `cancelDeny` pops the confirmation card and rebuilds+updates the
underlying scoring card from current state — it makes **no** changes to Gmail or the
spreadsheet. If resolving `ss`/`config` fails inside `cancelDeny` (e.g. `setup()` was
never run), it still pops the confirmation card but returns **without** an
`updateCard` — the card underneath is left showing whatever it displayed before the
confirmation card was pushed, not refreshed with current state.

### 7.10 Action handlers and the actioned-message record

`handleApprove(e)` / `handleDeny(e)` read `msgId` and `threadId` from
`e.parameters`, resolve full context via `loadThreadContext_(msgId)`:

```js
function loadThreadContext_(msgId) {
  // try: message = GmailApp.getMessageById(msgId); thread = message.getThread();
  //      data = extractEmailData(message)         -- swallow any error, leave undefined
  // try: ss = SpreadsheetApp.openById(getSpreadsheetId_()); config = loadConfig(ss);
  //      labels = loadLabels_(config)              -- swallow any error, leave null
  // return { message, thread, data, config, ss, labels };
}
```

If `ctx.config` is falsy, respond with a bare notification `'Config error - run
setup() first.'`. If `ctx.labels` is falsy (`loadLabels_` returned `null`), respond
`'Gmail labels missing - run setup() first.'`. Otherwise:

**`handleApprove`** — write, relabel, and update the sidebar in one round trip:

```js
try {
  updateSpreadsheet(ctx.ss, ctx.config, ctx.data,
    configInt_(ctx.config,'col_approved',4), configInt_(ctx.config,'col_approved_time',5), false);
  ctx.thread.addLabel(labels.approved);
  ctx.thread.addLabel(labels.scored);
  ctx.thread.removeLabel(labels.needsReview);
  ctx.thread.removeLabel(labels.denied);
  ctx.thread.refresh();

  setActionedMessage_(threadId, 'approved', msgId);   // <- see below
  clearActionedMessage_(threadId, 'denied');

  // rebuild the card from fresh thread labels, notification text exactly
  // 'Bonus Accepted', setNavigation(...).updateCard(card)  (in place, no push/pop)
} catch (err) {
  // Still transition toward "approved" even though the write failed, but leave the
  // thread in a CONSISTENT state: add approved, and remove every one of
  // needsReview/denied/scored (not just needsReview) so it can never end up
  // carrying both approved and denied (or a stale scored) at once.
  try {
    ctx.thread.addLabel(labels.approved);
    ctx.thread.removeLabel(labels.needsReview);
    ctx.thread.removeLabel(labels.denied);
    ctx.thread.removeLabel(labels.scored);
    ctx.thread.refresh();
  } catch (_) {}
  // notification text: 'Approved (sheet update failed: ' + err.message + ')',
  // response built with setStateChanged(true) instead of a rebuilt card (there is
  // no guarantee ctx.thread/labels are usable enough to rebuild a full card here).
}
```

**`handleDeny`** is the mirror image: writes to `col_denied`/`col_denied_time`, on
success adds `denied` and removes `needsReview`/`approved`/`scored`, records
`setActionedMessage_(threadId,'denied',msgId)` and clears the `'approved'` record,
notification exactly `'Bonus Denied'`; on failure adds `denied` and removes
`needsReview`/`approved`/`scored` (all three, same reasoning as above), notification
`'Denied (sheet update failed: ' + err.message + ')'`. **One navigation difference
from `handleApprove`, not cosmetic:** `handleDeny`
is only ever reached via the Confirm button on the pushed deny-confirmation card
(§7.9), so its success response is
`setNavigation(CardService.newNavigation().popCard().updateCard(card))` — it must pop
that confirmation card off the stack before updating the card underneath it.
`handleApprove` is invoked directly from the main scoring card (nothing was pushed),
so its success response is `setNavigation(CardService.newNavigation().updateCard(card))`
with no `popCard()`. Getting this backwards leaves an extra card on the navigation
stack (Approve) or fails to dismiss the confirmation card (Deny).

**The actioned-message record** — the reason revert is safe even when a thread has
multiple messages referencing different bonus codes. Gmail labels are thread-scoped,
not message-scoped, so once a thread is `approved`, nothing about the thread itself
says *which message* was approved. To make revert correct, every successful approve
or deny call records exactly which message it acted on, in Script Properties (shared
script-wide, exactly like `SPREADSHEET_ID`):

```js
function setActionedMessage_(threadId, kind, msgId) {   // kind: 'approved' | 'denied'
  PropertiesService.getScriptProperties().setProperty(kind + '_msg_' + threadId, msgId);
}
function getActionedMessage_(threadId, kind) {
  return PropertiesService.getScriptProperties().getProperty(kind + '_msg_' + threadId);
}
function clearActionedMessage_(threadId, kind) {
  PropertiesService.getScriptProperties().deleteProperty(kind + '_msg_' + threadId);
}
function dataForMsgId_(msgId) {
  if (!msgId) return null;
  try {
    const msg = GmailApp.getMessageById(msgId);
    return isValidSubject(msg.getSubject()) ? extractEmailData(msg) : null;
  } catch (e) { return null; }
}
```

**`handleRevertApproved(e)` / `handleRevertDenied(e)`** take only `firstMsgId` (always
`messages[0].getId()` from the button, §7.9) and resolve the row to clear in this
priority order, **not** simply from `firstMsgId`:

```js
const data = dataForMsgId_(getActionedMessage_(threadId, 'approved')) // 1. the message that
  || getFirstValidData_(ctx.thread)                                  //    was actually acted on
  || ctx.data;                                                       // 2. else, first valid
                                                                       //    message in the thread
                                                                       // 3. else, firstMsgId's own
                                                                       //    (possibly invalid) data
if (data) updateSpreadsheet(ctx.ss, ctx.config, data,
  configInt_(ctx.config,'col_approved',4), configInt_(ctx.config,'col_approved_time',5),
  false, /* value */ null);
ctx.thread.removeLabel(labels.approved);
ctx.thread.removeLabel(labels.scored);
ctx.thread.addLabel(labels.needsReview);
ctx.thread.refresh();
clearActionedMessage_(threadId, 'approved');
// rebuild + updateCard (no popCard - reached directly from the main card, same as
// handleApprove); notification 'Reverted to email-requires-review.'
```

Unlike `handleApprove`/`handleDeny`, the revert handlers make **no attempt at partial
label cleanup on failure** — the entire body above is wrapped in one `try`, and the
`catch` is just `return notify_('Error: ' + err.message);`. Do not add a label-cleanup
attempt here to "match" the approve/deny pattern; that is a deliberate asymmetry, not
an oversight to fix. `handleRevertDenied` is the same shape, targeting
`col_denied`/`col_denied_time` and the `'denied'` actioned-message record.

This three-tier fallback matters: a stored actioned-message record only exists for
threads whose approval/denial went through the sidebar buttons; a thread that was
manually labeled `approved` in Gmail has no such record and correctly falls back to
"first valid message" (the same heuristic `addApprovedCheck` uses, §7.6). Get this
priority order backwards — e.g. defaulting to `firstMsgId`'s data before checking the
stored record — and revert will silently target the wrong row whenever a thread has
more than one message with different bonus codes.

`getFirstValidData_(thread)` — returns `extractEmailData` of the first message in the
thread whose subject `isValidSubject`, or `null` if none. `errorCard_`/`notify_`/
`notifyRefresh_` are used for the various failure/refresh responses; `notifyRefresh_`
sets `setStateChanged(true)` instead of rebuilding a card.

## 8. Manifest (`appsscript.json`)

```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.addons.execute",
    "https://www.googleapis.com/auth/gmail.addons.current.message.metadata",
    "https://www.googleapis.com/auth/gmail.addons.current.message.action",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "runtimeVersion": "V8",
  "addOns": {
    "common": {
      "name": "Rally Scoring",
      "logoUrl": "https://www.gstatic.com/images/branding/product/1x/drive_2020q4_32dp.png",
      "useLocaleFromApp": true
    },
    "gmail": {
      "contextualTriggers": [
        { "unconditional": {}, "onTriggerFunction": "buildAddOn" }
      ]
    }
  }
}
```

Request **exactly** this scope set — no more, no less. In particular: do not add
`script.locale` (nothing in the code needs it, regardless of `useLocaleFromApp`), and
do not drop `script.external_request` even though nothing currently calls
`UrlFetchApp` — treat that one as a known, currently-unused scope rather than
justification to add a network call.

## 9. Error-Handling & Logging Philosophy

- **Never let a processing error abort the whole trigger run.** Every per-thread
  operation (`handleUnprocessedThread`, `addApprovedCheck`) catches its own errors and
  converts them into a Gmail label (`processing-error`) plus a log line, so one bad
  thread never prevents the rest of `processEmails()`'s batch from running.
- **Never let a missing prerequisite throw past a public entry point.** `setup()`,
  `processEmails()`, and every sidebar handler check for "Config sheet missing" or
  "Gmail labels missing" up front and return/respond gracefully (a log line or a
  Gmail notification) rather than surface a raw stack trace to the user.
- **Prefer a visible, correctable state over guessing.** When information needed to
  make a correct decision isn't available (which message in a thread an approval was
  really for), stop and flag rather than pick a plausible-looking default — see the
  ambiguity guard in §7.6 and the actioned-message priority order in §7.10.
- **Every label transition also gets a `Logger.log` line** stating what happened and
  why (e.g. `'-> Submitted: Rider ' + riderNumber + ' - ' + bonus`), so the Apps
  Script Executions log is a complete audit trail without needing to read the sheet.
- Configuration and email-registered-address errors are private-only (`Logger.log`);
  user-facing messages (sidebar notifications, card text) are reserved for things a
  scorer using the sidebar can actually act on.

## 10. Testing Requirements

A reimplementation is not complete until it ships an automated test suite meeting
these constraints:

- **No real Gmail/Sheets/network access, ever.** Build minimal in-memory mocks for
  every Apps Script service the code touches (`GmailApp`, `SpreadsheetApp`,
  `PropertiesService`, `LockService`, `ScriptApp`, `CardService`, `Logger`) and run
  the *actual* `code.js`/`Sidebar.gs` source against them unmodified — do not
  reimplement the logic under test in the test harness itself.
- Use a realm-isolation mechanism (e.g. Node's `vm` module) so each test gets a fresh,
  independent copy of every mock service; watch for cross-realm `instanceof`/
  `deepStrictEqual` pitfalls on values the loaded script constructs itself (a `new
  Date()` or a plain object literal built inside the sandboxed script is not
  `instanceof` the host realm's `Date`/`Object` unless those constructors are
  explicitly passed into the sandbox before the context is created).
- The Sheets mock must resolve the one formula pattern the app actually writes
  (`='<Sheet Name>'!A<n>`) so that rider-sheet bonus-ID lookups exercise real,
  observable behavior rather than being special-cased in test setup.
- No test dependencies beyond the language runtime's built-in test tools (this
  project uses Node's `node:test` + `node:assert/strict` — no third-party test
  framework).
- Coverage must include, at minimum, one test per: format/extraction regex edge case
  (§6); every config-default and config-error path (§5); every
  `updateSpreadsheet` write/revert/error/locking behavior (§7.5); rider-sheet
  creation including live formula propagation (§4.4); every `handleUnprocessedThread`
  and `addApprovedCheck` outcome, explicitly including the ambiguity guard (§7.6) and
  the email-error-beats-format-error priority (§7.4); `getStatus_`'s full precedence
  table (§7.8); and — as **explicit regression tests**, not just happy-path coverage
  — the two behaviors in §7.10 that are easy to regress silently: reverting a thread
  whose approved/denied message differs from its first message, and that a failed
  approve/deny leaves labels in a consistent (not contradictory) state.

## 11. Explicit Non-Goals

State these plainly so a reimplementation doesn't "fix" them into scope-creep:

- No SPF/DKIM or other cryptographic sender verification — §6's email match is a
  string comparison against the `From` header Gmail already parsed, nothing more.
- No config.csv / setup wizard. The Config sheet is created and populated by hand,
  directly from §5's table; `setup()` only ever *writes* the `spreadsheet_id` row back
  for the human's benefit, never reads a generated file.
- No rider-number range/format validation beyond "one or more digits" — no check
  against Rider Master happens until the sender-email step (§6), and format validity
  never implies the rider number actually exists in Rider Master.
- No UI beyond the Gmail sidebar and the spreadsheet itself — no web app, no external
  dashboard.
- No run-level mutual exclusion around all of `processEmails()` — only the
  per-write critical section in `updateSpreadsheet` (§7.5) is lock-protected. Two
  overlapping full trigger runs are tolerated because each thread's own label
  transitions make it naturally idempotent (a thread that already lost its
  `unprocessed` label is simply absent from the next run's search results).

## 12. Worked Example

Given this Config (abbreviated to the keys that differ from default, plus the
required label keys):

| key | value |
|---|---|
| `label_unprocessed` | `rally/unprocessed` |
| `label_needs_review` | `rally/email-requires-review` |
| `label_approved` | `rally/approved` |
| `label_scored` | `rally/scored` |
| `label_denied` | `rally/denied` |

Rider Master contains rider `42`, name `Jane Smith`, email `jane@example.com`. Bonus
Master contains `ABCD` and `WXYZ`.

1. Jane emails the scoring address, subject `"42 abcd"`. A Gmail filter applies
   `rally/unprocessed`.
2. Within `trigger_interval_min` minutes, `processEmails()` finds the thread, calls
   `handleUnprocessedThread`. The single message is format-valid
   (`isValidSubject('42 abcd')` → true) and Jane's sender address matches Rider
   Master. `updateSpreadsheet` writes `X` to rider sheet `42`'s Submitted column on
   the `ABCD` row (case-insensitive match), plus the email's own timestamp in the
   Submit Time column. The thread becomes `rally/email-requires-review`.
3. A scorer opens the email. The sidebar shows one message section with an **Approve
   this message** / **Deny this message** pair. They click Approve.
4. `handleApprove` writes `X` + the current time to the Approved column on the same
   row, adds `rally/approved` and `rally/scored`, removes
   `rally/email-requires-review`, and records `approved_msg_<threadId> = <that
   message's id>` in Script Properties. The sidebar rebuilds itself in place, showing
   status `'Scored'` and a **Remove approval / revert to email-requires-review**
   button in a new Revert section.
5. If the scorer instead clicks that Revert button, `handleRevertApproved` looks up
   `approved_msg_<threadId>`, finds the message actually approved in step 4,
   re-derives its rider/bonus data, clears both the Approved and Approve-Time cells on
   the `ABCD` row specifically, removes `rally/approved`/`rally/scored`, re-adds
   `rally/email-requires-review`, and deletes the stored property.
