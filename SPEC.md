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
| `utility.js` | `deleteAllRiderSheets()` — a manually-run, destructive helper for wiping rider sheets (and Leader Board and Master Scoring — both regenerated data, kept exactly as un-preserved as the rider sheets they're deleted alongside) between test runs. Reads sheet names from Config (falling back to defaults if Config can't be loaded) rather than hardcoding a second copy of them. Not called by `setup()`, `processEmails()`, or the sidebar. Optional third script file, named `Utility`. |
| `appsscript.json` | Manifest: OAuth scopes, add-on registration, runtime version, time zone. |
| `LICENSE` | GNU General Public License, version 3 or later — the verbatim, unmodified license text. Every `.js`/`.gs` source file carries a short copyright/license notice referencing it. |

Apps Script concatenates every script file into one global scope at runtime. Any
top-level `function` declared in `code.js` is directly callable from `Sidebar.gs` and
vice versa (and from `utility.js`, if added) — there is no module system and no
imports. **A helper used by both files must be defined in exactly one of them** (see
§7.3, `loadLabels_`); do not duplicate a helper across files, since the two copies can
silently drift — exactly what happened between `utility.js`'s `deleteAllRiderSheets`
and the near-identical copy that used to live inline in `README.md`: their
`keepSheets` arrays disagreed on `Leader Board` vs. `Leaderboard` until reconciled.

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
either 4 letters or 3 letters + 1 digit, since that's what a rider can type and what
the matching logic in §7.5 compares against. **Column B is that bonus's point value
("POINTS")** —
required for `createMasterScoring_` (§4.3a/7.2b) to compute anything meaningful; read
by fixed position like column A, never by header text, and not configurable via any
`col_*` key (Bonus Master's own layout never has been — only rider sheets have
configurable columns, and even there with the caveat in §4.4). A blank or non-numeric
POINTS value defaults to `0` in the Score calculation — logged as a warning only when
the cell is non-blank but not a number, silent when simply blank (the common case
before an organizer has filled every value in). Example:

| Bonus ID | POINTS |
|---|---|
| ABCD | 100 |
| WXYZ | 50 |

**Combination bonuses are just another row here, indistinguishable from a regular
bonus, as far as Bonus Master/rider sheets are concerned.** A combo code goes in
column A exactly like any bonus ID, gets its own row in every rider sheet (§4.4), and
can be submitted/approved/denied through the exact same mechanism as any bonus — a
rider can still email in for it directly, and a scorer can still approve that
submission from the sidebar, independent of anything below. What makes a combo code
special — which other bonus codes it depends on, and automatically crediting it once
those are all approved — lives entirely in the separate, optional **Combo Master**
sheet (§4.3b) and Master Scoring's own formula (§4.3a); Bonus Master itself has no
column, key, or special-case for it.

### 4.3a Master Scoring sheet

Default name `Master Scoring` (configurable via `sheet_master_scoring`), created and
grown by `setup()` via `createMasterScoring_` (§7.2b). Two things about *when* and
*where* this happens are handled by `setup()` itself, not `createMasterScoring_`
(§7.1 steps 6-9): it's created **after** rider sheets (so its per-rider formulas
reference sheets that already exist — see §4.4), and it ends up positioned **2nd
from the left, immediately after Leader Board** (§4.5) — not simply "appended after
existing sheets" the way `createMasterScoring_`'s own `insertSheet` call would place
it on its own; `setup()` explicitly repositions it there afterward, since Leader
Board (which must be leftmost) can't exist yet at the moment Master Scoring itself
needs to already exist. A two-dimensional grid: bonus rows starting at row 5, one
column per rider starting at column C. Rows 1–4 are fixed labels, written once at
creation:

| | A | B | C (1st rider) | D (2nd rider) | … |
|---|---|---|---|---|---|
| 1 | | `Name` | `=VLOOKUP(C2,'Rider Master'!$A$2:$B$<riderMasterLastRow>,2)` | … | |
| 2 | | `Number` | `='<Rider Master>'!A<n>` | … | |
| 3 | | `Score` | `=SUMIF(C5:C,"X",$B5:$B)` | … | |
| 4 | `Bonus` | `POINTS` | | | |
| 5+ | `='<Bonus Master>'!A<n>` | `='<Bonus Master>'!B<n>` | `='<rider sheet>'!<approvedCol><row>` | … | |

Row `5 + i` (0-indexed `i`) corresponds to Bonus Master row `2 + i`. Column `C + j`
(0-indexed `j`, via `columnToLetter_` — §7.2a) corresponds to the `j`-th rider
encountered in Rider Master, in row order. A bonus row's per-rider cell (row 5+)
points at that rider's own sheet, row `configInt_(config, 'header_row', 1, 1) + 1 +
i` — the same row arithmetic `createRiderSheet_` (§7.2) uses to place bonus `i` there
in the first place — column `columnToLetter_(configInt_(config, 'col_approved', 4,
1))`, i.e. whichever column `handleApprove`/`addApprovedCheck` actually *write* to,
not a hardcoded column.

**Every formula here is a direct, fully-resolved cell reference generated once per
cell — never `INDIRECT`.** A hand-built version of this sheet (e.g. dragging one
formula across many rider columns) typically uses `INDIRECT` so the *same* formula
text works after being copied anywhere; since this script generates each cell's exact
formula itself, there's no need for that indirection, and direct references avoid
`INDIRECT`'s volatility (recalculated on every edit, regardless of whether its inputs
changed).

**`Score` (row 3) uses ranges open-ended at the bottom but anchored at row 5
(`C5:C`, `$B5:$B`), not a fixed far bound and not a true whole column.** The
open-ended-at-the-bottom part is exactly the same reasoning as Leader Board's `RANK`
(§4.5): a range bounded there too (`C5:C1003`, say) would go stale the instant more
bonus rows are appended below it. The row-5 anchor at the *top* is not optional the
way `RANK`'s range has no anchor at all — a true whole-column range (`C:C`) includes
row 3, the `Score` formula's own cell, which Sheets flags as a **circular
reference** even though the `"X"` criteria could never actually match that cell's
own content. `$B5:$B` is anchored the same way for the same reason, even though
column B's own row 3 is blank (not a formula) — keeping both range starts aligned is
what SUMIF requires (matching criteria/sum range sizes), not a second circular-risk
avoidance in its own right.

**Growth on re-run mirrors Leader Board's, in both dimensions.** A rider newly added
to Rider Master gets a new column, appended to the right, backfilled with the
Approved-check formula for every bonus row that already exists. A bonus newly added
to Bonus Master gets a new row, appended at the bottom, backfilled with the
Approved-check formula for every rider column that already exists. Existing cells —
in either dimension — are never rewritten. (An existing rider column and a new bonus
row together, or a new rider column and an existing bonus row together, are each
handled by exactly one of those two backfill passes — never both, and never neither.)

**A combo bonus's row (§4.3b) gets a different per-rider formula than the plain
direct-reference shown in the table above.** For rider column with letter `L`, bonus
row `row`, whose Combo Master-resolved member rows are `m1, m2, …` (§4.3b, §7.2b):

```
=IF(OR(AND(L<m1>="X",L<m2>="X",…),'<riderNumber>'!<approvedCol><riderSheetRow>="X"),"X","")
```

— `"X"` if **either** every member bonus's own cell for that same rider column is
`"X"`, **or** the combo's own direct-reference check (identical to the plain formula
every other bonus row gets) is itself `"X"`. The two paths are additive: a scorer can
still directly approve a combo's own submission from the sidebar exactly as before,
completely independent of whether its members have been approved. This formula is
computed once, at the same point in the growth logic a plain bonus row's formula
would be (new rider column, or new bonus row on an existing rider column) — a combo
mapping added to Combo Master *after* a given cell already exists does not retroactively
convert it; see §4.3b for that limitation and its workaround.

**Prerequisites, matching §4.5's own split:** Bonus Master must already exist — a hard
`throw`, like `createRiderSheet_`'s own check — since there's nothing to build rows
from otherwise. Rider Master missing or empty is tolerated (soft skip, no error,
nothing created), matching `createAllRiderSheets_`/`createLeaderBoard_`'s existing
tolerance for that same condition.

**Formatting is reapplied over the sheet's full current extent at the end of every
run** (creation or growth) — not computed incrementally against just the newly-written
cells:

- **Center alignment** (`Range.setHorizontalAlignment('center')`) is applied to the
  entire used range (row 1 through the current last row, column A through the current
  last column) — header labels and data cells alike.
- **Row banding** (`Range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true,
  false)`, i.e. Sheets' "alternating colors") is applied the same way, over the same
  full extent. Any banding already on the sheet is removed first (`sheet.getBandings()
  .forEach(b => b.remove())`) — Sheets refuses to apply a new banding whose range
  overlaps an existing one, so growing the sheet and reapplying without first removing
  the old (now too-small) banding would throw.
- **Conditional formatting** — a single rule, `whenTextEqualTo('X').setBackground(
  '#b7e1cd')` — is scoped to just the data grid (row 5+, column C+), not the header
  rows or the Bonus/POINTS label columns, since those never contain `"X"`.
  `sheet.setConditionalFormatRules([...])` **replaces** the sheet's entire rule set each
  run rather than appending, so re-running `setup()` after growth ends up with exactly
  one rule (covering the new, larger extent), never a second stale one left over from
  before the growth.

All three are naturally idempotent (redrawing the same alignment/banding/rule over
cells that already have it is a no-op), and since they only touch presentation — never
`getValues()`/`setValues()` — none of them interact with the read-scan-then-append
growth logic above.

**The sheet is also given warning-only protection at creation time** —
`sheet.protect().setWarningOnly(true).setDescription(...)` — "Edit with warning" in the
Sheets UI: anyone can still edit any cell, but is shown a confirmation dialog first.
Unlike the range-scoped formatting above, this is *sheet-level* protection, which
automatically covers every future row/column as the sheet grows — no code is needed to
extend it on later runs, and it is applied exactly once, only in the branch where the
sheet is freshly created (the `else` branch, growing an already-existing sheet, never
calls `protect()` again — doing so on every run would stack a second, redundant
protection object rather than being a no-op the way re-drawing alignment/banding is). An
existing Master Scoring sheet created before this was added does not get protection
retroactively — the same creation-time-only limitation as rider sheets' Date+Time
number format (§4.4).

### 4.3b Combo Master sheet (optional)

Default name `Combo Master` (configurable via `sheet_combo_master`) — **read only** by
`createMasterScoring_` (§4.3a/7.2b), never created or written by this script, the same
role Rider/Bonus Master play. Missing entirely is not an error; it just means no combos
exist for this event. Two columns, **one row per (combo code, member bonus code) pair**
— any number of members needs no schema change, since a combo with `N` members is
simply `N` rows sharing the same combo code in column A:

| Combo ID | Member Bonus ID |
|---|---|
| COMB | ABCD |
| COMB | WXYZ |

Row 1 is a header (text never checked, matching Bonus/Rider Master's own convention).
**Both the combo code and every one of its member codes must independently already be
an ordinary Bonus Master row** (§4.3) — Combo Master only maps codes to each other, it
never substitutes for a Bonus Master entry (no separate POINTS value, no separate rider-
sheet row: the combo still needs its own Bonus Master row for both of those). A combo
or member code Combo Master references that isn't found in Bonus Master is logged
(`createMasterScoring_: combo "..." is not a Bonus Master entry - ...` or `... references
member bonus "..." ...`) and that combo's row simply falls back to the plain
direct-reference formula, same as if it weren't listed in Combo Master at all — never a
thrown error, since a Bonus/Combo Master data-entry mistake shouldn't block the rest of
the sheet.

**Historical note, to avoid confusion with an earlier, fully-removed feature of the
same name:** an earlier version of this codebase had a *different* "Combo Master"
concept — a full mirrored sheet that stood in for Bonus Master's own combo rows, paired
with sidebar-side enforcement blocking approval until every component was scored. That
was deliberately removed in full (no sheet, no sidebar logic, no code path) in favor of
combos being fully indistinguishable Bonus Master rows, approved entirely at the
scorer's own judgment. The Combo Master documented here is a different, later, strictly
smaller feature: a pure code-to-code mapping consumed only by Master Scoring's own
formula generation (§4.3a) — it does not reintroduce sidebar gating, does not block
manual approval of anything, and Bonus Master rows remain exactly as indistinguishable
as before.

**No rider-sheet growth mechanism exists to retrofit a bonus/combo row into an
already-existing rider sheet** (§4.4 — a longstanding limitation of
`createRiderSheet_`/`createAllRiderSheets_`, unrelated to Master Scoring and unchanged
by this feature). Practical consequence: a combo (and every one of its members) needs
to already be in Bonus Master (and the combo's mapping already in Combo Master) *before
the very first* `setup()` run for that rider's sheet to end up with a row for it at all
— adding a combo mid-event only ever affects rider sheets/Master Scoring columns
created *after* that point, exactly the same rule that already applies to any regular
bonus added mid-event.

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

Columns C, E, and G (Submit/Approve/Deny Time) are additionally given the number
format `'M/d/yyyy h:mm:ss am/pm'` at sheet-creation time, applied to every data row
that exists at that point (the same row range the bonus-ID formulas are written
into). This is deliberate: writing a `Date` via `setValue()` into a cell that still
has Sheets' default "Automatic" format often displays as date-only depending on the
column's inherited formatting, not date+time. Setting the format once at creation
means every future `handleUnprocessedThread`/`handleApprove`/`handleDeny` write
already lands in a cell that's going to show both. A rider sheet created before this
was added, or any row added to Bonus Master *after* a rider's sheet already exists
(§7.2 - such a row never gets a formula on existing rider sheets in the first place),
won't have this format and would need it applied by hand.

Column A of a rider sheet is **not** a copy of Bonus Master's values — it is written
as one formula per data row, `='<Bonus Master sheet name>'!A<n>` (single quotes around
the sheet name, always column A, `n` = the corresponding Bonus Master row), so edits to
Bonus Master after the fact propagate to every rider sheet automatically. New sheets
are inserted **after every existing sheet** (`ss.insertSheet(riderNumber,
ss.getSheets().length)`), so pre-existing tabs (`Leader Board`, `Master Scoring`, etc.)
are never disturbed.

**Important, easy-to-miss inconsistency to preserve exactly, not "fix":**
`createRiderSheet_` writes the 7-column header row, the bonus-ID formulas, and the
Date+Time number formats at hardcoded columns (1–7 for the header;
`sheet.getRange(headerRow + 1, 1, formulas.length, 1)` for the formulas; `[3, 5, 7]`
for the timestamp formats) — never `configInt_(config, 'col_bonus_id', ...)` or any
other `col_*` key. Only `header_row` (which *row* everything lands on) is actually
threaded through from config here. The `col_bonus_id`/`col_submitted`/`col_approved`/
`col_denied` keys (and their `_time` counterparts) are consulted **only** by
`updateSpreadsheet` (§7.5) when it later reads or writes a specific cell — they have
no effect on where `createRiderSheet_` places anything, headers or formats alike.
Practical consequence: the "seven independently configurable columns" story in §5 is
only trustworthy if every `col_*` key is left at its default (1/2/3/4/5/6/7, in that
order) or the rider sheet's actual header row is edited to match by hand —
reconfiguring, say, `col_approved` to `10` does not move the "Approved" header or its
Date+Time formatting; it just makes `handleApprove` write `X` (with a plain,
unformatted `Date` next to it) into column 10 while the sheet still shows "Approved"
at column 4. Reproduce this disconnect faithfully rather than parameterizing
`createRiderSheet_`'s column placement to "fix" it — that would be a behavior change,
not a match to this codebase.

### 4.5 Leader Board sheet

Default name `Leader Board` (configurable via `sheet_leader_board`), created by
`setup()` via `createLeaderBoard_` (§7.2a) — one row per rider in Rider Master, header
row 1, bold, frozen (not configurable via `header_row`; unlike rider score sheets,
Leader Board's header position is always row 1, matching the same "always row 1"
assumption Rider Master/Bonus Master already make). Inserted as the
**leftmost tab** (`ss.insertSheet(name, 0)`) — the one exception to "every sheet this
script creates is appended after whatever already exists" (§4.4). This only happens
at creation; re-running `setup()` against an already-existing Leader Board only adds
rows (§7.2a) and never repositions the tab. Leader Board's own leftmost position is
never disturbed by anything — it's Master Scoring that gets explicitly repositioned
*relative to* Leader Board afterward (§4.3a, §7.1 step 9), not the other way around.

| Column | Header text | Formula shape |
|---|---|---|
| A | `Rider Number` | `='<Rider Master>'!A<n>` — cell reference, not a copied value |
| B | `Name` | `=VLOOKUP(A<row>,'<Rider Master>'!$A$2:$B$<riderMasterLastRow>,2)` |
| C | `Score` | `=HLOOKUP(A<row>,'<Master Scoring>'!$C$2:$<lastColLetter>$3,2)` |
| D | `Finish` | `=RANK(C<row>,$C:$C,0)` |

`Master Scoring` (§4.3a) is itself created by `setup()` now (via `createMasterScoring_`,
§7.2b, called before `createLeaderBoard_` specifically so it exists by the time Leader
Board needs it) — its layout (rider numbers across row 2, scores across row 3,
starting at column C) is exactly what §4.3a's own generation produces, so the two
stay in sync by construction. `createLeaderBoard_` still independently requires
Master Scoring to exist and throws if it doesn't (§7.2a) — in practice this only
happens if Bonus Master is *also* missing (Master Scoring's own hard prerequisite,
§4.3a), cascading into Leader Board too; Rider Master missing is tolerated at every
level (soft skip, nothing created, no error).

**Ranges are computed fresh from each master sheet's actual current size at the
moment a row is written** — `riderMaster.getLastRow()` for column B's range,
`masterScoring.getLastColumn()` for column C's — not hardcoded bounds. This is safe
for B/C specifically because `VLOOKUP`/`HLOOKUP` only need their row's own rider to
be *somewhere* within the range; a range sized when a row is written stays correct
for that row forever, even as more rows/columns are added to the master sheets later.

**Column D is the one exception, and deliberately not sized the same way.** `RANK`
needs to cover every rider row *in Leader Board itself* — a range computed once at
write time would go stale for already-written rows the instant a later `setup()` run
appends more rows below them. It uses a whole-column reference (`$C:$C`) instead,
which needs no maintenance as rows are added — this is what actually makes "existing
rows are never touched" (§7.2a) true for every column, not just A/B/C.

`RANK(..., 0)` ranks **descending** — the highest score is rank 1. This assumes bonus
points accumulate and higher is strictly better in this event's scoring, which is the
obvious default for this codebase's "bonus points" model but is stated explicitly here
since nothing else in the spec pins down a ranking direction; flip the third argument
to `1` if an event instead wants ascending (lowest-score-wins) ranking.

## 5. Configuration Reference

Every config value is a string; numeric ones are parsed at the point of use (never
cached), so edits to the Config sheet take effect on the next run without redeploying.

Two lookup helpers, both defined in `code.js`:

```js
// Non-numeric keys read with a hardcoded `|| 'Default Name'` fallback inline at
// each call site (loadLabels_, createAllRiderSheets_, createRiderSheet_,
// validateEmailAddress) — there is no shared string-default helper.

// Numeric keys go through this helper everywhere:
function configInt_(config, key, defaultValue, minValue) {
  const raw = config[key];
  if (raw === undefined || raw === '') {
    if (defaultValue !== undefined) return defaultValue;
    throw new Error('Config key "' + key + '" is missing.');
  }
  const parsed = parseInt(raw, 10);
  if (isNaN(parsed)) throw new Error('Config key "' + key + '" is not a number: "' + raw + '"');
  if (minValue !== undefined && parsed < minValue) {
    throw new Error('Config key "' + key + '" must be at least ' + minValue + ', got: "' + raw + '"');
  }
  return parsed;
}
```

Rule: a blank or absent numeric key silently falls back to its documented default (the
default itself is never checked against `minValue` — it's a trusted literal in the
source, not a Config sheet value); a *present but non-numeric* value throws
immediately with the exact message shape above (callers propagate this — they do not
catch and re-default it).

**`minValue` (4th, optional argument):** every call site that resolves a value
ultimately used as a spreadsheet row or column number — `header_row`, `col_bonus_id`,
`col_submitted`, `col_submitted_time`, `col_approved`, `col_approved_time`,
`col_denied`, `col_denied_time` — passes `1` here. `trigger_interval_min` (used only
as a trigger interval, never a range coordinate) does not. Without this check, a
misconfigured value like `col_approved = "0"` would parse as a valid-looking integer
and only fail much later, deep inside a `sheet.getRange(...)` call, as Apps Script's
own opaque `"The starting column of the range is too small"` — which names no Config
key at all. Catching it here instead throws `'Config key "col_approved" must be at
least 1, got: "0"'` immediately, at the point the bad value is read, with the specific
key and value named. This is a real bug class, not hypothetical: it is exactly what
produces "starting column/row of the range is too small" from `handleApprove` (and
every other write path) if a `col_*`/`header_row` key is ever set to `0` or negative.

| Key | Default (when blank/absent) | Read by |
|---|---|---|
| `event_name` | — (display only, never read by code) | — |
| `organizer_email` | — (display only, never read by code) | — |
| `spreadsheet_id` | written by `setup()` | not read by code at runtime (Script Properties is the source of truth — see §7.1) |
| `sheet_rider_master` | `Rider Master` | `createAllRiderSheets_`, `validateEmailAddress` |
| `sheet_bonus_master` | `Bonus Master` | `createRiderSheet_` |
| `sheet_combo_master` | `Combo Master` | `createMasterScoring_` (optional — §4.3b) |
| `sheet_master_scoring` | `Master Scoring` | `createMasterScoring_` (creates/grows it), `createLeaderBoard_` (must exist by the time it runs — see §4.3a/4.5) |
| `sheet_leader_board` | `Leader Board` | `createLeaderBoard_` |
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
  return /^\d+\s*(?:[A-Za-z]{4}|[A-Za-z]{3}\d)$/.test(subject.trim());
}

// Format check + extraction, in one step
const match = subject.trim().match(/^(\d+)\s*([A-Za-z]{4}|[A-Za-z]{3}\d)$/);
// match[1] -> rider number (kept as a string, digits only)
// match[2] -> bonus code, upper-cased before use
```

Rules implied by the regex (state these explicitly in any reimplementation, since the
regex alone is easy to get subtly wrong):

- The rider number is one or more digits, with no other rider-number validation
  (no range check, no leading-zero handling beyond what `\d+` naturally allows).
- Zero or more whitespace characters are allowed between the number and the code
  (`\s*`, so `"42ABCD"` and `"42     ABCD"` are both valid).
- The bonus code is **either** exactly 4 alphabetic characters (letters only — no
  digits, hyphens, or underscores), **or** exactly 3 alphabetic characters followed by
  exactly 1 digit (e.g. `"ABC1"`). No other letter/digit combination is valid — not 4
  letters plus a digit (`"ABCD1"`), not 3 letters plus 2 digits (`"ABC12"`), not fewer
  than 3 letters before a digit (`"AB1"`). The letters are upper-cased before being
  compared or stored, so `"abcd"`, `"AbCd"`, and `"ABCD"` are the same bonus, and
  `"abc1"`/`"ABC1"` are the same bonus.
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
6. Call `createAllRiderSheets_` (§7.2) — *before* Master Scoring, so its
   Approved-check formulas reference rider sheets that already exist (not a
   correctness requirement — a forward reference to a not-yet-existing sheet just
   shows `#REF!` until the sheet appears — but the intended, documented order).
7. Record whether Master Scoring and Leader Board (their configured or default
   names) each already existed, *before* touching either — needed by step 9.
8. Call `createMasterScoring_` (§7.2b), then `createLeaderBoard_` (§7.2a), each
   wrapped in its own `try { ... } catch (e) { Logger.log('FATAL: ' + e.message); }`
   — a thrown error in either must not skip labels/rider sheets/the trigger, so
   each is caught individually rather than propagating out of `setup()` like step
   2's Config failure does. Master Scoring is called first specifically so it
   already exists by the time Leader Board needs it (§4.5) — `createLeaderBoard_`
   still independently throws if it doesn't.
9. If either sheet from step 7 didn't already exist: look both up again by name: if
   both now exist, `ss.setActiveSheet(masterScoringSheet)` then
   `ss.moveActiveSheet(2)` — the real Apps Script pattern for repositioning an
   existing sheet (there's no direct "set index" call) — placing Master Scoring at
   the 2nd position, 1-based, i.e. immediately after Leader Board's own leftmost
   position (§4.5). Skipped entirely when both sheets already existed before this
   run, so a tab a human has since moved by hand is never silently undone on a
   later `setup()` re-run.
10. Delete every existing project trigger whose handler function is
   `'processEmails'`, then create exactly one new one:
   `ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(intervalMin).create()`,
   where `intervalMin = configInt_(config, 'trigger_interval_min', 10)`. This makes
   `setup()` idempotent with respect to triggers — re-running it never produces
   duplicate triggers.
11. Log `'Setup complete. Trigger set for every ' + intervalMin + ' minutes.'`.

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
  //    configInt_(config, 'header_row', 1, 1), bold, and freeze that many rows.
  // 4. Build one formula per Bonus Master row ("='" + bonusMasterName + "'!A" + (i + 2)).
  //    Write them all in one setFormulas call starting at headerRow + 1, only if the
  //    array is non-empty (Bonus Master may legitimately have zero data rows).
  // 5. If that array was non-empty: setNumberFormat('M/d/yyyy h:mm:ss am/pm') on
  //    columns 3, 5, and 7 (Submit/Approve/Deny Time), same row range as step 4,
  //    so every future timestamp write already displays as Date+Time.
  // 6. Auto-resize column 1. Return the created sheet.
}
```

### 7.2a Leader Board creation

```js
function columnToLetter_(col) {
  // Converts a 1-based column number to A1-style letters (3 -> 'C', 98 -> 'CT').
  // The reverse conversion of what formula-resolution needs letters-to-index for;
  // needed here to build a lookup range's ending column reference from Master
  // Scoring's actual width. Shared with createMasterScoring_ (§7.2b), which needs
  // the same conversion for its rider columns and its Approved-column reference -
  // defined once here, not duplicated.
}
```

```js
function createLeaderBoard_(ss, config) {
  // 1. Look up Rider Master; missing or <2 rows -> log and return (soft - same
  //    tolerance createAllRiderSheets_ already has for this exact condition).
  // 2. Find the rider-number column by header text; not found -> log and return.
  // 3. Look up Master Scoring (config['sheet_master_scoring'] || 'Master Scoring');
  //    missing -> throw new Error('Cannot create Leader Board - "' + masterScoringName +
  //    '" not found. ...') - a hard prerequisite, unlike Rider Master above.
  // 4. Resolve or create the Leader Board sheet (config['sheet_leader_board'] ||
  //    'Leader Board'):
  //      - Doesn't exist: ss.insertSheet(name, 0) - leftmost tab, unlike every
  //        other sheet this script creates; write the header row (Rider Number,
  //        Name, Score, Finish) at row 1, bold, frozen.
  //      - Exists: read column A's resolved values (row 2..lastRow) into a Set of
  //        rider numbers already present.
  // 5. riderMasterLastRow = max(riderMaster.getLastRow(), 2)
  //    masterScoringLastCol = max(masterScoring.getLastColumn(), 3) -> letter via
  //    columnToLetter_.
  // 6. For each Rider Master data row, trimmed rider number, skip blank, skip if
  //    already in the existing-riders set (added to the set immediately after its
  //    row is written, guarding against a duplicate rider number producing two
  //    rows): append one row at the next free Leader Board row - see §4.5's table
  //    for the exact formula shape of each of the four columns. Catches and logs
  //    any per-row error without stopping the loop (same resilience pattern as
  //    createAllRiderSheets_).
  // 7. Log one summary line: 'Leader Board: added=' + added + ', skipped=' + skipped.
}
```

Every row this function writes is independent of every other row it has ever
written in a *previous* call — nothing here ever rewrites a row for a rider who
already has one. That's what makes "add rows for new riders, leave existing rows
alone" true across repeated `setup()` runs as the roster grows, and it's also why
column D's range can't follow the same "sized when written" rule as B/C (§4.5).

### 7.2b Master Scoring creation

```js
function createMasterScoring_(ss, config) {
  // 1. Look up Rider Master; missing or <2 rows -> log and return (soft - same
  //    tolerance createAllRiderSheets_/createLeaderBoard_ already have).
  // 2. Find the rider-number column by header text; not found -> log and return.
  // 3. Look up Bonus Master (config['sheet_bonus_master'] || 'Bonus Master');
  //    missing -> throw new Error('Cannot create Master Scoring - "' + bonusMasterName +
  //    '" not found. ...') - a hard prerequisite, matching createRiderSheet_'s own
  //    Bonus Master check. Read its data rows once.
  // 4. Look up Combo Master (config['sheet_combo_master'] || 'Combo Master') -
  //    OPTIONAL, missing is not an error. If present, read every (comboCode,
  //    memberCode) row pair into a Map: comboCode -> [memberCode, ...] (§4.3b).
  // 5. Resolve or create Master Scoring (config['sheet_master_scoring'] || 'Master
  //    Scoring'), appended after existing sheets - NOT leftmost, unlike Leader
  //    Board (§4.5's special case, not this sheet's). This function has no
  //    opinion on final tab position beyond that; setup() (§7.1 step 9)
  //    repositions it to sit right after Leader Board afterward, once Leader
  //    Board exists - something this function can't do itself, since it always
  //    runs before Leader Board does (§7.1 step 8):
  //      - Doesn't exist: create it; write B1/B2/B3 = 'Name'/'Number'/'Score' and
  //        A4/B4 = 'Bonus'/'POINTS'; sheet.protect().setWarningOnly(true)
  //        .setDescription(...) - once, here, never in the "already exists" branch.
  // 6. Scan existing state (both empty if the sheet was just created):
  //      - Row 2 across columns C.. -> Map of riderNumber -> column index.
  //      - Column A rows 5..lastRow -> Map of bonus code -> Master Scoring row
  //        (not just a Set of codes - the row number is needed to resolve combo
  //        member references below).
  // 7. New bonus rows next (column A/B only for now, direct refs to Bonus Master;
  //    POINTS defaulting to 0, logging a warning only if the cell is present but
  //    non-numeric - never for simply blank). Track each new row's corresponding
  //    rider-sheet row (headerRow + 1 + i) for step 9/10, and add its code -> row
  //    to the same Map step 6 built (a combo and its members can all be new in
  //    the same run and still resolve each other).
  // 8. Resolve each Combo Master entry from step 4 against the code -> row Map
  //    from steps 6-7: a combo/member code not found there is logged and that
  //    combo is skipped entirely (falls back to the plain formula in step 9/10,
  //    same as an ordinary bonus row) - never a thrown error. Result: a Map of
  //    combo row -> [member row, ...], only for combos that fully resolved.
  // 9. New rider columns next (rows 1-3: VLOOKUP / direct Rider Master ref / SUMIF
  //    over a range open-ended at the bottom but anchored at row 5, e.g. 'C5:C' -
  //    not a true whole column, which would include row 3 itself and trip Sheets'
  //    circular-reference check), then - for each new column - the Approved-check
  //    formula (or the combo formula from step 8, if this row resolved as one) for
  //    EVERY bonus row that exists by now, old and new. This is why new bonus rows
  //    are handled first: by the time a new column is backfilled, the final bonus
  //    row set (and combo resolution) is already known, so a new column is never
  //    touched twice.
  // 10. For each EXISTING rider column (not one just added in step 9): write the
  //     Approved-check formula (or combo formula) only for the newly-added bonus
  //     rows from step 7. Existing-column/existing-row cells are never written in
  //     any step - this is also why a Combo Master mapping added after a given
  //     cell already exists does not retroactively convert it (§4.3b).
  // 11. Reapply formatting over the sheet's full current extent (§4.3a): remove
  //     any existing banding, then center-align and row-band (LIGHT_GREY) the
  //     whole used range; rebuild (not append to) the sheet's conditional
  //     format rule set with a single "whenTextEqualTo('X') -> green
  //     background" rule scoped to just the data grid (row 5+, column C+).
  // 12. Log one summary line: 'Master Scoring: bonus rows added=' + N + ', rider
  //     columns added=' + M + '.'
}
```

The Approved-check formula written in steps 9-10, for rider `riderNumber`'s cell in
bonus row `row` (`row >= 5`), is `='<riderNumber>'!<approvedColLetter><riderSheetRow>`
where `approvedColLetter = columnToLetter_(configInt_(config, 'col_approved', 4, 1))`
and `riderSheetRow = configInt_(config, 'header_row', 1, 1) + 1 + (row - 5)` — computed
once per call (not per cell) since neither `col_approved` nor `header_row` changes
mid-run — **unless `row` resolved as a combo row in step 8**, in which case it's the
`IF(OR(AND(...` combo formula from §4.3a instead, built from that same direct-reference
string plus the member rows' cells in the same rider column. §4.3a has the full
rationale for every deviation from the reference spreadsheet this behavior was
reverse-engineered from (direct references over `INDIRECT`, `Score`'s
row-5-anchored-but-open-ended range instead of a fixed far bound, reading the
*actual* `col_approved` column instead of a hardcoded one, and the combo formula
itself).

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
  //       configInt_(config,'col_submitted',2,1), configInt_(config,'col_submitted_time',3,1),
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

Row-lookup is factored into its own helper:

```js
function findBonusRow_(sheet, config, bonusCode) {
  const startRow = configInt_(config, 'header_row', 1, 1) + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return null;

  const bonusCol = configInt_(config, 'col_bonus_id', 1, 1);
  const values = sheet.getRange(startRow, bonusCol, lastRow - startRow + 1, 1).getValues();
  const target = bonusCode.trim().toUpperCase();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === target) return startRow + i;
  }
  return null;
}

function updateSpreadsheet(ss, config, data, columnIndex, timeColumnIndex, useEmailTime, value) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);       // up to 30s; throws if it can't acquire in time
  try {
    let sheet = ss.getSheetByName(data['rider-number']);
    if (!sheet) sheet = createRiderSheet_(ss, config, data['rider-number']);

    const startRow = configInt_(config, 'header_row', 1, 1) + 1;
    if (sheet.getLastRow() < startRow) throw new Error('No data rows in sheet ' + data['rider-number']);

    const row = findBonusRow_(sheet, config, data['bonus']);
    if (row === null) throw new Error('Bonus ID "' + data['bonus'] + '" not found in sheet ' + data['rider-number']);

    if (value === null) {
      sheet.getRange(row, columnIndex).clearContent();
      sheet.getRange(row, timeColumnIndex).clearContent();
    } else {
      sheet.getRange(row, columnIndex).setValue('X');
      sheet.getRange(row, timeColumnIndex).setValue(useEmailTime ? data.date : new Date());
    }
  } finally {
    lock.releaseLock();       // always runs, including on the throw paths above
  }
}
```

Note the two distinct thrown messages are still checked separately, in this order —
`updateSpreadsheet` itself still re-checks `sheet.getLastRow() < startRow` before
calling `findBonusRow_`, rather than folding that into `findBonusRow_`'s `null`
return, purely so **"No data rows in sheet X"** and **"Bonus ID ... not found in
sheet X"** stay distinguishable error messages.

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
      configInt_(config, 'col_approved', 4, 1), configInt_(config, 'col_approved_time', 5, 1), false);
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

Registered in the manifest as the unconditional Gmail contextual trigger (§8), so it
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
  //     'followed by a Bonus Code (4 letters, or 3 letters + 1 digit).') and stop.
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
   - format error: `'Subject line does not match the required format: Rider Number followed by a Bonus Code - 4 letters, or 3 letters + 1 digit (e.g. "42 ABCD" or "42 ABC1"). Ask the rider to resend with the correct subject.'`
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
    configInt_(ctx.config,'col_approved',4,1), configInt_(ctx.config,'col_approved_time',5,1), false);
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
  configInt_(ctx.config,'col_approved',4,1), configInt_(ctx.config,'col_approved_time',5,1),
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
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.locale"
  ],
  "runtimeVersion": "V8",
  "addOns": {
    "common": {
      "name": "Rally Scoring",
      "logoUrl": "https://raw.githubusercontent.com/scrogatl/snafu-rally-scoring/main/icons/sidebar-icon.png",
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

Request **exactly** this scope set — no more, no less. Two of these are easy to
mistakenly "clean up," for opposite reasons — don't:

- **`script.locale` is required, not optional, *because* `useLocaleFromApp` is
  `true`.** An earlier version of this doc claimed the opposite (that nothing needed
  it "regardless of `useLocaleFromApp`") — that was wrong, corrected after a real
  deployment failed to authorize without it. `useLocaleFromApp: true` is what
  actually creates the dependency, not anything in `code.js`/`Sidebar.gs`'s own
  logic (neither file reads a locale). Drop `useLocaleFromApp` and this scope stops
  being required — but don't drop just the scope while leaving the manifest flag on.
- **`script.external_request` stays even though nothing currently calls
  `UrlFetchApp`** — treat that one as a known, currently-unused scope rather than
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
- **Every `Sidebar.gs` entry point logs on every failure path, not just the ones that
  write to the sheet.** `handleApprove`/`handleDeny`/`handleRevertApproved`/
  `handleRevertDenied` each log their config/labels guard clauses and their catch
  block (with `err.stack` when available);
  `buildAddOn` and `cancelDeny` log their config-resolution failures; and
  `loadThreadContext_`'s two internal `try/catch`es — previously silent (`catch (_)
  {}`) — now each log what failed to resolve and why. The rule: a scorer-visible
  notification or error card is never the *only* record of a failure — there is
  always a matching `Logger.log` line in the Executions log with enough context
  (function name, threadId/msgId) to diagnose it after the fact, without needing the
  scorer to have copied down the toast text.
- **`configInt_`'s optional `minValue` argument (§5) turns a whole class of
  misconfiguration into an immediate, named error** instead of a much-later, opaque
  one. Any `col_*`/`header_row` key resolves through `configInt_(..., 1)`; if the
  Config sheet value is `0` or negative, the error is `'Config key "col_approved"
  must be at least 1, got: "0"'`, thrown at the point the value is read — not Apps
  Script's own `"The starting column/row of the range is too small"` several calls
  later inside `sheet.getRange(...)`, which names no Config key and gives no hint
  that a Config sheet value is the actual cause.

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
  explicitly passed into the sandbox before the context is created). The same applies
  to arrays: any `Array.prototype.map()`/`.filter()` result built inside the sandbox
  is a vm-realm array and fails `assert.deepStrictEqual` against a host-realm array
  literal even when every element matches — normalize with `Array.from(...)`
  (host-side) before comparing, the same way a `toPlain(...)` copy is needed for
  plain objects.
- The Sheets mock must resolve the one formula pattern the app actually writes
  (`='<Sheet Name>'!A<n>`) so that rider-sheet bonus-ID lookups exercise real,
  observable behavior rather than being special-cased in test setup. The mock does
  **not** need to evaluate `VLOOKUP`/`HLOOKUP`/`RANK`/`SUMIF` (§4.3a/4.5/7.2a/7.2b) —
  those are only ever asserted by exact formula *text*, since real Sheets formula
  evaluation is out of scope for an in-memory mock; only Leader Board's column A and
  Master Scoring's row-2 cells (both plain cell references, same pattern as above)
  get both a formula-text and a resolved-value assertion.
- Leader Board test coverage specifically (§4.5/7.2a): the header/frozen row and
  per-rider formula text on first creation; ranges reflecting each master sheet's
  actual current size (including a Master Scoring wide enough to need a two-letter
  column, e.g. `CT`); re-running adds rows only for new riders and leaves existing
  rows' formulas byte-for-byte untouched; the hard throw when Master Scoring is
  missing; the soft no-op when Rider Master is missing/empty; and — as an
  **integration-level regression test**, not just the unit-level ones above —
  `setup()` itself still creates labels, rider sheets, and the trigger when Master
  Scoring is missing and `createLeaderBoard_` throws (§7.1 step 8's `try`/`catch`).
- Master Scoring test coverage specifically (§4.3a/7.2b): exact formula text for
  every cell shape on first creation (Name/Number/Score per rider column, Bonus/
  POINTS/Approved-check per bonus row); growth in **both** dimensions independently
  (a new rider column backfilled for existing bonus rows; a new bonus row backfilled
  for existing rider columns), each an explicit regression test that existing cells
  stay byte-for-byte untouched; a fully idempotent re-run (nothing added, nothing
  changes); the hard throw when Bonus Master is missing; the soft no-op when Rider
  Master is missing/empty; a blank vs. a non-numeric POINTS value (0 either way, but
  only the latter logs a warning); that `header_row`/`col_approved` are actually
  threaded through the per-cell rider-sheet reference, not assumed to be their
  defaults; and — an **explicit regression test**, since a real Sheets circular-
  reference error is easy to reintroduce without noticing (the mock has no
  circular-reference detection of its own) — that `Score`'s range is never a true
  whole column (`C:C`), only anchored-but-open-ended (`C5:C`). Also cover the
  formatting pass (§4.3a): center alignment reaches header cells as well as data
  cells; row banding covers the full extent and is removed-then-reapplied (not
  stacked) on a growth re-run, an **explicit regression test** since Sheets throws
  if a new banding overlaps an existing one; and the conditional-format rule is
  scoped to row 5+/column C+ only, rebuilt (not appended to) on re-run so growth
  never leaves a second, stale rule behind. Also cover the sheet's warning-only
  protection (§4.3a): set at creation time with `getWarningOnly()` true; and — an
  **explicit regression test**, since `sheet.protect()` creates a new object on
  every call rather than being idempotent the way alignment/banding are — that a
  growth re-run does not add a second protection to an already-protected sheet.
- Combo auto-approval test coverage specifically (§4.3b/4.3a/7.2b): the exact
  `IF(OR(AND(...` formula text for a combo row's cell, with both its member-row
  conditions and its own direct-reference fallback; a combo with a single member
  and a combo with several (no fixed-count assumption anywhere in the formula
  generation); that an ordinary (non-combo) bonus row's formula is completely
  unaffected when Combo Master exists; that no Combo Master sheet at all produces
  the exact same plain-reference formulas as before this feature existed (a
  regression test that adding this feature changed nothing for events that don't
  use it); a combo code, and separately a member code, that Combo Master
  references but Bonus Master doesn't have — each logged and each falling back to
  the plain formula, never a thrown error; growth coverage for both write sites
  (a new rider column added after the combo already exists; a new bonus row that
  is itself a combo, added on a later run, backfilled onto existing rider
  columns); and a re-run-is-idempotent regression test that an already-written
  combo cell is never rewritten.
- The Sheets mock must support `Range.setHorizontalAlignment`/`getHorizontalAlignment`,
  `Range.applyRowBanding`/`Sheet.getBandings`/`Banding.remove` (throwing if a new
  banding's range overlaps an existing one on the same sheet, matching real Sheets),
  `SpreadsheetApp.newConditionalFormatRule`/`Sheet.getConditionalFormatRules`/
  `Sheet.setConditionalFormatRules` (the last **replaces** the sheet's rule set,
  it does not append), and `Sheet.protect()`/`Protection.setWarningOnly`/
  `setDescription`/`Sheet.getProtections` — needed for Master Scoring's formatting
  and protection above.
- The Sheets mock must support `setActiveSheet`/`moveActiveSheet` (§7.1 step 9's
  repositioning) — the same two-call pattern real Apps Script requires, not a
  single hypothetical "set index" call, so a reimplementation can't accidentally
  simplify this into an API real Apps Script doesn't have. Coverage: a fresh
  `setup()` run ends with Leader Board leftmost and Master Scoring immediately
  after it; re-running `setup()` after a human has manually moved either tab does
  **not** silently reposition it back — the repositioning only fires when at least
  one of the two sheets didn't already exist before that run (§7.1 step 7).
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
- No **sidebar-side** special handling for combination bonuses — a combo code is
  submitted, approved, and denied through the exact same sidebar mechanism as any
  bonus (§4.3), and `handleApprove`/`handleDeny`/the sidebar cards have no notion of
  combos at all. A scorer reviewing a combo's own submission in the sidebar is not
  shown its dependencies or blocked from approving it regardless of their status —
  that judgment call remains entirely theirs, unchanged from before combo mapping
  existed. What *did* change: **Master Scoring** (only) can now compute a combo's own
  cell automatically once its component bonuses are all approved, via the optional
  Combo Master sheet (§4.3b) — but this is presentation/scoring math, not a gate on
  anything a human can click; see §4.3a's combo formula for the actual mechanism.
  Combo Master itself has no code path outside `createMasterScoring_` — nothing in
  `processEmails()`, `Sidebar.gs`, or `updateSpreadsheet` reads it, ever.
- No *enforcement* of points, even though the script now *generates* the formulas
  that compute them (Bonus Master's POINTS column, §4.3; Master Scoring's Score row,
  §4.3a). The script writes those formulas once and never reads the result back or
  acts on it — it doesn't know or care what any rider's computed Score is, doesn't
  gate approval on it, and doesn't rank anything itself beyond Leader Board's own
  `RANK` formula (§4.5, itself just more generated-and-forgotten spreadsheet text).
  Whether a bonus (combo or otherwise) is actually worth what its POINTS value says,
  and what if anything is done with the totals, is entirely up to the organizer.

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

**Combination bonus, continuing the same event:** suppose Bonus Master also lists
`ARSN`, a code the organizer treats as a combination bonus requiring `ABCD` and
`WXYZ` to both be approved first — nothing in Bonus Master itself marks it as such
(§4.3). The organizer additionally lists two rows in the optional Combo Master sheet:
`ARSN`/`ABCD` and `ARSN`/`WXYZ` (§4.3b).

6. Jane emails `"42 ARSN"`. Exactly like step 1–2 above — format-valid, sender
   matches — `updateSpreadsheet` writes `X` and a timestamp to rider sheet `42`'s
   Submitted column on the `ARSN` row. The thread becomes `rally/email-requires-review`,
   indistinguishable from any other submission — `handleUnprocessedThread`/
   `updateSpreadsheet` have no idea `ARSN` is a combo at all; only
   `createMasterScoring_` ever reads Combo Master.
7. The scorer opens the email. The sidebar shows the same **Approve this message** /
   **Deny this message** pair it would for any bonus — sidebar behavior for `ARSN` is
   completely unaffected by Combo Master (§11). The scorer can approve it directly
   here exactly as before, independent of `ABCD`/`WXYZ`'s status, if they judge it
   warranted.
8. Separately, in Master Scoring (§4.3a), rider `42`'s `ARSN` row cell is a generated
   formula — `=IF(OR(AND(<ABCD's cell>="X",<WXYZ's cell>="X"),'42'!D<n>="X"),"X","")`
   — so once *both* `ABCD` and `WXYZ` show `"X"` for rider `42` in Master Scoring, that
   cell reads `"X"` too, and rider `42`'s `Score` (§4.3a's `SUMIF`) counts `ARSN`'s
   `POINTS` automatically — with no scorer action on `ARSN`'s own submission required.
   If the scorer directly approves `ARSN`'s own submission instead (step 7), the same
   cell reads `"X"` via that path regardless of `ABCD`/`WXYZ`. Either path (or both) is
   sufficient; neither is required over the other.
