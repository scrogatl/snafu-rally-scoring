/*
 * snafu-rally-scoring
 * Copyright (C) 2026 Scott Rogers
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

/**
 * Rally Scoring - code.js
 * Last edited: 2026-09-08
 * Config is read at runtime from the "Config" sheet.
 * Run setup() once after creating and populating the Config sheet.
 */

// --- Config loader ------------------------------------------------------------
function loadConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) throw new Error('Config sheet not found. Run setup() first.');
  const rows = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]).trim();
    const val = String(rows[i][1]).trim();
    if (key) config[key] = val;
  }
  return config;
}

/**
 * Reads a numeric Config value. Falls back to defaultValue when the key is
 * missing/blank; throws a clear error when the key is present but not a number.
 *
 * Pass minValue for any key that's ultimately used as a spreadsheet row or
 * column number (header_row, col_*) - e.g. configInt_(config, 'col_approved',
 * 4, 1). Without this, a misconfigured value like "0" would sail through as
 * a valid-looking number and only fail much later, deep inside a
 * sheet.getRange(...) call, as Apps Script's own cryptic "The starting
 * column/row of the range is too small" - which gives no hint that a Config
 * sheet value is the actual cause. Catching it here instead names the
 * specific key and its bad value immediately.
 */
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

// --- One-time setup -----------------------------------------------------------
// Before running setup():
//   1. Create a sheet named "Config" with key/value/notes columns and populate
//      it per the Config sheet reference in README.md
//   2. Populate your Rider Master and Bonus Master sheets
// Then run this function once.

function setup() {
  Logger.log('Running setup...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const spreadsheetId = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  Logger.log('Spreadsheet ID saved to Script Properties: ' + spreadsheetId);

  let config;
  try {
    config = loadConfig(ss);
  } catch (e) {
    Logger.log('FATAL: ' + e.message);
    Logger.log('Create and populate a sheet named "Config" before running setup().');
    return;
  }

  // Write the actual spreadsheet ID into the Config sheet
  const configSheet = ss.getSheetByName('Config');
  if (configSheet) {
    const configRows = configSheet.getDataRange().getValues();
    for (let i = 1; i < configRows.length; i++) {
      if (String(configRows[i][0]).trim() === 'spreadsheet_id') {
        configSheet.getRange(i + 1, 2).setValue(spreadsheetId);
        Logger.log('Spreadsheet ID written to Config sheet row ' + (i + 1));
        break;
      }
    }
  }

  const parentLabel = config['label_parent'] || 'rally';
  if (!GmailApp.getUserLabelByName(parentLabel)) {
    GmailApp.createLabel(parentLabel);
    Logger.log('Created label: ' + parentLabel);
  } else {
    Logger.log('Label already exists: ' + parentLabel);
  }

  const labelKeys = [
    'label_unprocessed', 'label_format_error', 'label_email_error',
    'label_processing_error', 'label_needs_review', 'label_approved',
    'label_denied', 'label_scored'
  ];
  let created = 0, skipped = 0;
  for (const key of labelKeys) {
    const name = config[key];
    if (!name) { Logger.log('Config key missing: ' + key); continue; }
    if (GmailApp.getUserLabelByName(name)) {
      Logger.log('Label already exists: ' + name); skipped++;
    } else {
      GmailApp.createLabel(name);
      Logger.log('Created label: ' + name); created++;
    }
  }
  Logger.log('Labels: created=' + created + ', skipped=' + skipped);

  createAllRiderSheets_(ss, config);

  // Master Scoring must be CREATED before Leader Board (createLeaderBoard_
  // requires it to already exist, to size its HLOOKUP range) but POSITIONED
  // right after it (2nd-from-left, Leader Board being leftmost) - opposite
  // requirements that can't both be satisfied by insertion order alone.
  // Resolved by creating both, then explicitly repositioning Master Scoring -
  // only when either sheet was freshly created this run, so an already-set-up
  // spreadsheet where a human moved a tab by hand is never silently undone.
  const scoringName = config['sheet_master_scoring'] || 'Master Scoring';
  const leaderBoardName = config['sheet_leader_board'] || 'Leader Board';
  const scoringExistedBefore = !!ss.getSheetByName(scoringName);
  const leaderBoardExistedBefore = !!ss.getSheetByName(leaderBoardName);

  try {
    createMasterScoring_(ss, config);
  } catch (e) {
    Logger.log('FATAL: ' + e.message);
  }

  try {
    createLeaderBoard_(ss, config);
  } catch (e) {
    Logger.log('FATAL: ' + e.message);
  }

  if (!scoringExistedBefore || !leaderBoardExistedBefore) {
    const scoring = ss.getSheetByName(scoringName);
    const leaderBoard = ss.getSheetByName(leaderBoardName);
    if (scoring && leaderBoard) {
      ss.setActiveSheet(scoring);
      ss.moveActiveSheet(2); // 1-based: right after Leader Board at position 1
      Logger.log('Positioned ' + scoringName + ' 2nd-from-left, after ' + leaderBoardName + '.');
    }
  }

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const intervalMin = configInt_(config, 'trigger_interval_min', 10);
  ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(intervalMin).create();

  Logger.log('Setup complete. Trigger set for every ' + intervalMin + ' minutes.');
}

// --- Rider sheet creation -----------------------------------------------------
function createAllRiderSheets_(ss, config) {
  const masterSheet = ss.getSheetByName(config['sheet_rider_master'] || 'Rider Master');
  if (!masterSheet) { Logger.log('createAllRiderSheets_: Rider Master not found - skipping.'); return; }
  const rows = masterSheet.getDataRange().getValues();
  if (rows.length < 2) { Logger.log('createAllRiderSheets_: No riders found.'); return; }
  const headers = rows[0];
  const rci = headers.indexOf(config['master_col_rider_number'] || 'Rider Number');
  if (rci === -1) { Logger.log('createAllRiderSheets_: Rider Number column not found.'); return; }
  let created = 0, skipped = 0;
  for (let i = 1; i < rows.length; i++) {
    const riderNumber = String(rows[i][rci]).trim();
    if (!riderNumber) continue;
    if (ss.getSheetByName(riderNumber)) {
      Logger.log('Rider sheet already exists: ' + riderNumber); skipped++;
    } else {
      try { createRiderSheet_(ss, config, riderNumber); created++; }
      catch (e) { Logger.log('Error creating sheet for Rider ' + riderNumber + ': ' + e.message); }
    }
  }
  Logger.log('Rider sheets: created=' + created + ', skipped=' + skipped);
}

function createRiderSheet_(ss, config, riderNumber) {
  Logger.log('Creating rider sheet for: ' + riderNumber);
  const bonusMasterName = config['sheet_bonus_master'] || 'Bonus Master';
  const bonusMaster = ss.getSheetByName(bonusMasterName);
  if (!bonusMaster) throw new Error(
    'Cannot create rider sheet - "' + bonusMasterName + '" not found. ' +
    'Create a sheet with one bonus ID per row in column A.'
  );
  const sheet = ss.insertSheet(riderNumber, ss.getSheets().length); // append after existing sheets
  const headerRow = configInt_(config, 'header_row', 1, 1);
  sheet.getRange(headerRow, 1, 1, 7)
    .setValues([['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time']])
    .setFontWeight('bold');
  sheet.setFrozenRows(headerRow);

  // Use cell references instead of copying values so any updates to
  // Bonus Master propagate automatically to all rider sheets.
  const formulas = [];
  const bonusLastRow = bonusMaster.getLastRow();
  for (let i = 0; i < bonusLastRow - 1; i++) {
    formulas.push(['=\'' + bonusMasterName + '\'!A' + (i + 2)]);
  }
  if (formulas.length) {
    sheet.getRange(headerRow + 1, 1, formulas.length, 1).setFormulas(formulas);

    // Format the timestamp columns as Date+Time (not just Date) up front, so
    // it's already in place before handleUnprocessedThread/handleApprove/
    // handleDeny ever write a Date into them - Sheets otherwise displays a
    // freshly-written Date with the column's inherited "Automatic" format,
    // which often renders as date-only. Hardcoded columns 3/5/7 (Submit/
    // Approve/Deny Time), matching the header row above - not the col_*_time
    // config keys, for the same reason createRiderSheet_'s header placement
    // doesn't read them either (see SPEC.md §4.4).
    const dateTimeFormat = 'M/d/yyyy h:mm:ss am/pm';
    [3, 5, 7].forEach((col) => {
      sheet.getRange(headerRow + 1, col, formulas.length, 1).setNumberFormat(dateTimeFormat);
    });
  }

  sheet.autoResizeColumn(1);
  Logger.log('Created rider sheet: ' + riderNumber);
  return sheet;
}

// --- Leader Board --------------------------------------------------------------

/**
 * Converts a 1-based column number to its A1-style letters (3 -> 'C',
 * 98 -> 'CT'). The reverse of the letters-to-index conversion the test mocks
 * need for formula resolution - needed here for real, to build a lookup
 * range's ending column reference from Master Scoring's actual width.
 */
function columnToLetter_(col) {
  let letters = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    col = Math.floor((col - 1) / 26);
  }
  return letters;
}

/**
 * Creates the Master Scoring sheet if it doesn't exist yet, then grows it to
 * cover every rider in Rider Master and every bonus in Bonus Master -
 * appending new rider columns and new bonus rows as needed, never rewriting
 * an existing cell. Fixed layout (rows 1-4 are labels, data starts at row 5 /
 * column C):
 *
 *   B1='Name', B2='Number', B3='Score' (one column per rider, starting at C)
 *   A4='Bonus', B4='POINTS' (one row per bonus, starting at row 5)
 *
 * Each rider's own column: Name is a VLOOKUP against Rider Master; Number is
 * a direct cell reference to Rider Master (not a copied value - same idiom
 * as every other propagated value in this codebase); Score is a SUMIF over
 * an open-ended range anchored at row 5 (e.g. 'D5:D', not 'D:D') - deliberately
 * not bounded at the far end, same reasoning as Leader Board's RANK ($C:$C,
 * see createLeaderBoard_ above): a range bounded there would go stale the
 * moment more bonus rows are appended below it. It IS bounded at the near
 * end, unlike RANK's range - a true whole-column 'D:D' includes row 3 (the
 * Score formula's own cell), which Sheets flags as a circular reference even
 * though the criteria would never actually match that cell.
 *
 * Each bonus's own row: columns A/B are direct cell references to Bonus
 * Master (id, POINTS - defaulting to 0 with a logged warning if present but
 * non-numeric, silently 0 if simply blank). Each rider's cell in that row is
 * a direct reference to that rider's own sheet's Approved column - whichever
 * column configInt_ resolves col_approved to, not hardcoded - at whichever
 * row header_row implies for that bonus's position, mirroring the exact row
 * arithmetic createRiderSheet_ uses to place it there in the first place.
 *
 * Throws if Bonus Master doesn't exist (a hard prerequisite, like
 * createRiderSheet_'s own check) - callers (setup()) must catch this so it
 * doesn't prevent labels/rider sheets/the trigger/Leader Board from being set
 * up. Returns (no throw) if Rider Master is missing/empty, matching
 * createAllRiderSheets_'s tolerance for that same condition.
 *
 * Formatting is reapplied over the sheet's full current extent at the end of
 * every run (creation or growth), not just over newly-written cells: center
 * alignment and row banding are cheap to redraw entirely, and doing so means
 * a sheet that's grown since the last run doesn't end up with old rows/
 * columns formatted differently from new ones. Any existing banding is
 * removed first since Sheets refuses to apply a new banding that overlaps
 * one already on the sheet. The "turn green if it contains X" conditional
 * format rule is rebuilt the same way, scoped to just the data grid (row 5+,
 * column C+) - not the header rows/label columns, which never contain "X" -
 * and replaces the sheet's entire rule set each run (this sheet is fully
 * script-managed, so there's nothing else to preserve).
 *
 * The sheet is also given warning-only protection (`Sheet.protect()` +
 * `setWarningOnly(true)`) at creation time - "Edit with warning" in the
 * Sheets UI: anyone can still edit any cell, but gets a confirmation dialog
 * first, since this whole sheet is regenerated by `setup()`. This is a
 * *sheet-level* protection, which - unlike the range-scoped alignment/
 * banding/conditional-format rule above - automatically covers every future
 * row/column as the sheet grows, with no code needed to extend it on later
 * runs; done once, in the `!scoring` (freshly-created) branch only. An
 * existing Master Scoring sheet from before this was added does not get
 * protection retroactively - same "creation-time-only" limitation as rider
 * sheets' Date+Time number format (§4.4).
 */
function createMasterScoring_(ss, config) {
  const riderMasterName = config['sheet_rider_master'] || 'Rider Master';
  const riderMaster = ss.getSheetByName(riderMasterName);
  if (!riderMaster) { Logger.log('createMasterScoring_: Rider Master not found - skipping.'); return; }
  const riderRows = riderMaster.getDataRange().getValues();
  if (riderRows.length < 2) { Logger.log('createMasterScoring_: No riders found.'); return; }
  const rci = riderRows[0].indexOf(config['master_col_rider_number'] || 'Rider Number');
  if (rci === -1) { Logger.log('createMasterScoring_: Rider Number column not found.'); return; }

  const bonusMasterName = config['sheet_bonus_master'] || 'Bonus Master';
  const bonusMaster = ss.getSheetByName(bonusMasterName);
  if (!bonusMaster) throw new Error(
    'Cannot create Master Scoring - "' + bonusMasterName + '" not found. ' +
    'Create a sheet with one bonus ID per row in column A (and its point value in column B).'
  );
  const bonusRows = bonusMaster.getDataRange().getValues();

  const scoringName = config['sheet_master_scoring'] || 'Master Scoring';
  let scoring = ss.getSheetByName(scoringName);
  const existingRiderCols = new Map(); // riderNumber -> column index
  const existingBonusRows = new Set(); // bonus code

  if (!scoring) {
    scoring = ss.insertSheet(scoringName, ss.getSheets().length);
    scoring.getRange(1, 2, 3, 1).setValues([['Name'], ['Number'], ['Score']]);
    scoring.getRange(4, 1, 1, 2).setValues([['Bonus', 'POINTS']]);
    scoring.protect()
      .setWarningOnly(true)
      .setDescription('Master Scoring is generated and maintained by setup() - edits may be overwritten on the next run.');
  } else {
    const lastCol = scoring.getLastColumn();
    if (lastCol >= 3) {
      scoring.getRange(2, 3, 1, lastCol - 2).getValues()[0]
        .forEach((v, idx) => { const n = String(v).trim(); if (n) existingRiderCols.set(n, 3 + idx); });
    }
    const lastRow = scoring.getLastRow();
    if (lastRow >= 5) {
      scoring.getRange(5, 1, lastRow - 4, 1).getValues()
        .forEach((row) => { const v = String(row[0]).trim(); if (v) existingBonusRows.add(v); });
    }
  }

  const riderMasterLastRow = Math.max(riderMaster.getLastRow(), 2);
  const headerRow = configInt_(config, 'header_row', 1, 1);
  const approvedColLetter = columnToLetter_(configInt_(config, 'col_approved', 4, 1));

  // New bonus rows first: column A/B only for now. Rider cells for these
  // rows come after we know the final column set, so a brand-new rider
  // column (backfilled below for every bonus row) never gets written twice.
  let nextBonusRow = scoring.getLastRow() < 5 ? 5 : scoring.getLastRow() + 1;
  let bonusRowsAdded = 0;
  const newBonusRows = []; // { row, riderSheetRow }
  for (let i = 1; i < bonusRows.length; i++) {
    const bonusCode = String(bonusRows[i][0]).trim();
    if (!bonusCode || existingBonusRows.has(bonusCode)) continue;
    const bonusMasterRow = i + 1;
    const pointsRaw = String(bonusRows[i][1] === undefined ? '' : bonusRows[i][1]).trim();
    if (pointsRaw && isNaN(parseFloat(pointsRaw))) {
      Logger.log('createMasterScoring_: bonus "' + bonusCode + '" has a non-numeric POINTS value ("' +
        pointsRaw + '") - defaulting to 0 in the Score calculation.');
    }
    scoring.getRange(nextBonusRow, 1, 1, 2).setFormulas([[
      '=\'' + bonusMasterName + '\'!A' + bonusMasterRow,
      '=\'' + bonusMasterName + '\'!B' + bonusMasterRow,
    ]]);
    const riderSheetRow = headerRow + 1 + (nextBonusRow - 5);
    newBonusRows.push({ row: nextBonusRow, riderSheetRow });
    existingBonusRows.add(bonusCode);
    nextBonusRow++;
    bonusRowsAdded++;
  }
  const finalLastBonusRow = nextBonusRow - 1;

  // New rider columns: rows 1-3, then the Approved-check formula for every
  // bonus row that will exist by the end of this run (existing and new).
  let nextCol = scoring.getLastColumn() < 3 ? 3 : scoring.getLastColumn() + 1;
  let riderColsAdded = 0;
  const newRiderCols = new Set(); // column indices written in this loop
  for (let i = 1; i < riderRows.length; i++) {
    const riderNumber = String(riderRows[i][rci]).trim();
    if (!riderNumber || existingRiderCols.has(riderNumber)) continue;
    const riderMasterRow = i + 1;
    const colLetter = columnToLetter_(nextCol);

    scoring.getRange(1, nextCol, 1, 1).setFormulas([[
      '=VLOOKUP(' + colLetter + '2,\'' + riderMasterName + '\'!$A$2:$B$' + riderMasterLastRow + ',2)',
    ]]);
    scoring.getRange(2, nextCol, 1, 1).setFormulas([['=\'' + riderMasterName + '\'!A' + riderMasterRow]]);
    scoring.getRange(3, nextCol, 1, 1).setFormulas([[
      '=SUMIF(' + colLetter + '5:' + colLetter + ',"X",$B5:$B)',
    ]]);
    if (finalLastBonusRow >= 5) {
      const formulas = [];
      for (let row = 5; row <= finalLastBonusRow; row++) {
        const riderSheetRow = headerRow + 1 + (row - 5);
        formulas.push(['=\'' + riderNumber + '\'!' + approvedColLetter + riderSheetRow]);
      }
      scoring.getRange(5, nextCol, formulas.length, 1).setFormulas(formulas);
    }
    existingRiderCols.set(riderNumber, nextCol);
    newRiderCols.add(nextCol);
    nextCol++;
    riderColsAdded++;
  }

  // New bonus rows' cells for EXISTING rider columns only - new rider
  // columns were already fully backfilled (including these rows) above.
  if (newBonusRows.length) {
    for (const [riderNumber, col] of existingRiderCols) {
      if (newRiderCols.has(col)) continue;
      for (const { row, riderSheetRow } of newBonusRows) {
        scoring.getRange(row, col, 1, 1).setFormulas([[
          '=\'' + riderNumber + '\'!' + approvedColLetter + riderSheetRow,
        ]]);
      }
    }
  }

  // Reapply formatting over the full current extent - see docstring above.
  const finalLastRow = scoring.getLastRow();
  const finalLastCol = scoring.getLastColumn();
  if (finalLastRow >= 1 && finalLastCol >= 1) {
    scoring.getBandings().forEach((b) => b.remove());
    const fullRange = scoring.getRange(1, 1, finalLastRow, finalLastCol);
    fullRange.setHorizontalAlignment('center');
    fullRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }

  const rules = [];
  if (finalLastRow >= 5 && finalLastCol >= 3) {
    const dataRange = scoring.getRange(5, 3, finalLastRow - 4, finalLastCol - 2);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('X')
        .setBackground('#b7e1cd')
        .setRanges([dataRange])
        .build()
    );
  }
  scoring.setConditionalFormatRules(rules);

  Logger.log('Master Scoring: bonus rows added=' + bonusRowsAdded + ', rider columns added=' + riderColsAdded);
}

/**
 * Creates the Leader Board sheet if it doesn't exist yet (as the leftmost
 * tab - unlike every other sheet this script creates, which is appended
 * after whatever already exists), with a header row, bold, frozen - then
 * appends one row per rider in Rider Master that doesn't already have one -
 * existing rows are never touched. Each row's Name/Score
 * formulas reference Rider Master / Master Scoring with a range sized to
 * those sheets' actual current extent at the moment the row is written, so a
 * later run (with more riders in Rider Master, or a wider Master Scoring)
 * naturally produces a correctly-sized range for the NEW rows it adds,
 * without needing to revisit rows written earlier.
 *
 * The Finish (rank) column is the one exception: RANK needs to cover every
 * row in Leader Board itself, not just its own row, so a range sized at
 * write time would go stale the moment a later run appends more rows below
 * it. It uses a whole-column reference ($C:$C) instead, specifically so it
 * never needs to be revisited - "existing rows are never touched" would
 * otherwise be broken for this one column.
 *
 * Throws if Master Scoring doesn't exist (a hard prerequisite, like
 * createRiderSheet_'s Bonus Master check) - callers (setup()) must catch
 * this so it doesn't prevent labels/rider sheets/the trigger from being set
 * up. Returns (no throw) if Rider Master is missing/empty, matching
 * createAllRiderSheets_'s tolerance for that same condition.
 */
function createLeaderBoard_(ss, config) {
  const riderMasterName = config['sheet_rider_master'] || 'Rider Master';
  const riderMaster = ss.getSheetByName(riderMasterName);
  if (!riderMaster) { Logger.log('createLeaderBoard_: Rider Master not found - skipping.'); return; }
  const riderRows = riderMaster.getDataRange().getValues();
  if (riderRows.length < 2) { Logger.log('createLeaderBoard_: No riders found.'); return; }
  const rci = riderRows[0].indexOf(config['master_col_rider_number'] || 'Rider Number');
  if (rci === -1) { Logger.log('createLeaderBoard_: Rider Number column not found.'); return; }

  const masterScoringName = config['sheet_master_scoring'] || 'Master Scoring';
  const masterScoring = ss.getSheetByName(masterScoringName);
  if (!masterScoring) throw new Error(
    'Cannot create Leader Board - "' + masterScoringName + '" not found. ' +
    'Create your scoring sheet first, or set sheet_master_scoring in Config if it has a different name.'
  );

  const leaderBoardName = config['sheet_leader_board'] || 'Leader Board';
  let leaderBoard = ss.getSheetByName(leaderBoardName);
  const existingRiders = new Set();
  if (!leaderBoard) {
    leaderBoard = ss.insertSheet(leaderBoardName, 0); // leftmost tab, unlike every other created sheet
    leaderBoard.getRange(1, 1, 1, 4)
      .setValues([['Rider Number', 'Name', 'Score', 'Finish']])
      .setFontWeight('bold');
    leaderBoard.setFrozenRows(1);
  } else {
    const lastRow = leaderBoard.getLastRow();
    if (lastRow > 1) {
      leaderBoard.getRange(2, 1, lastRow - 1, 1).getValues()
        .forEach((row) => { const v = String(row[0]).trim(); if (v) existingRiders.add(v); });
    }
  }

  const riderMasterLastRow = Math.max(riderMaster.getLastRow(), 2);
  const masterScoringLastCol = Math.max(masterScoring.getLastColumn(), 3);
  const scoreLastColLetter = columnToLetter_(masterScoringLastCol);

  let nextRow = leaderBoard.getLastRow() + 1;
  let added = 0, skipped = 0;
  for (let i = 1; i < riderRows.length; i++) {
    const riderNumber = String(riderRows[i][rci]).trim();
    if (!riderNumber) continue;
    if (existingRiders.has(riderNumber)) { skipped++; continue; }

    const riderMasterRow = i + 1; // riderRows is 0-indexed; sheet rows are 1-indexed
    try {
      leaderBoard.getRange(nextRow, 1, 1, 4).setFormulas([[
        '=\'' + riderMasterName + '\'!A' + riderMasterRow,
        '=VLOOKUP(A' + nextRow + ',\'' + riderMasterName + '\'!$A$2:$B$' + riderMasterLastRow + ',2)',
        '=HLOOKUP(A' + nextRow + ',\'' + masterScoringName + '\'!$C$2:$' + scoreLastColLetter + '$3,2)',
        '=RANK(C' + nextRow + ',$C:$C,0)',
      ]]);
      existingRiders.add(riderNumber);
      nextRow++;
      added++;
    } catch (e) {
      Logger.log('createLeaderBoard_: error adding row for Rider ' + riderNumber + ': ' + e.message);
    }
  }
  Logger.log('Leader Board: added=' + added + ', skipped=' + skipped);
}

// --- Main processing loop -----------------------------------------------------
function processEmails() {
  Logger.log('Starting email processing...');
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  );
  let config;
  try { config = loadConfig(ss); }
  catch (e) { Logger.log('FATAL: ' + e.message); return; }

  const labels = loadLabels_(config);
  if (!labels) return;

  // Process unprocessed threads - scan ALL messages, pass if any is valid
  let threads = GmailApp.search('label:' + config['label_unprocessed']);
  Logger.log('Unprocessed threads: ' + threads.length);
  for (const thread of threads)
    handleUnprocessedThread(ss, config, thread, labels);

  // Record approved threads that haven't been scored yet
  threads = GmailApp.search(
    'label:' + config['label_approved'] + ' -label:' + config['label_scored']
  );
  Logger.log('Approved/unscored threads: ' + threads.length);
  for (const thread of threads)
    addApprovedCheck(ss, config, thread, labels);

  Logger.log('Processing complete.');
}

// --- Email handlers -----------------------------------------------------------

/**
 * Scans all messages in a thread. Uses the first valid submission message
 * (valid format + registered email) to write the submitted X to the sheet.
 */
function handleUnprocessedThread(ss, config, thread, labels) {
  const messages = thread.getMessages();
  Logger.log('Thread has ' + messages.length + ' message(s): ' + messages[0].getSubject());

  let validMessage = null;
  let formatError = false;
  let emailError  = false;

  for (const message of messages) {
    const subject = message.getSubject();
    if (!isValidSubject(subject)) { formatError = true; continue; }
    const data = extractEmailData(message);
    if (!validateEmailAddress(ss, config, message.getFrom(), data['rider-number'])) {
      emailError = true; continue;
    }
    validMessage = message;
    break; // first valid message is enough to mark submitted
  }

  if (validMessage) {
    const data = extractEmailData(validMessage);
    try {
      updateSpreadsheet(ss, config, data,
        configInt_(config, 'col_submitted', 2, 1),
        configInt_(config, 'col_submitted_time', 3, 1),
        true);
      thread.addLabel(labels.needsReview);
      thread.removeLabel(labels.unprocessed);
      thread.refresh();
      Logger.log('-> Submitted: Rider ' + data['rider-number'] + ' - ' + data['bonus']);
    } catch (e) {
      Logger.log('-> Processing error: ' + e.message);
      thread.addLabel(labels.processingError);
      thread.removeLabel(labels.unprocessed);
      thread.refresh();
    }
  } else if (emailError) {
    Logger.log('-> Email error on all messages.');
    thread.addLabel(labels.emailError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
  } else if (formatError) {
    Logger.log('-> Format error on all messages.');
    thread.addLabel(labels.formatError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
  }
}

/**
 * Records an approval for threads where the "approved" label was applied
 * directly in Gmail rather than via the sidebar's per-message Approve button
 * (the sidebar already marks such threads "scored", so this only runs for
 * the manual-label fallback path). Uses the first valid message in the
 * thread for rider/bonus data - if other valid messages in the thread
 * reference a different rider/bonus, there's no way to know which one the
 * manual label was meant for, so we flag it instead of guessing.
 */
function addApprovedCheck(ss, config, thread, labels) {
  const messages = thread.getMessages();
  Logger.log('Recording approval for thread: ' + messages[0].getSubject());

  const validMessages = messages.filter(m => isValidSubject(m.getSubject()));
  if (!validMessages.length) {
    Logger.log('-> No valid message found in approved thread.');
    return;
  }

  const data = extractEmailData(validMessages[0]);
  if (!data['rider-number'] || !data['bonus']) {
    Logger.log('-> No valid message found in approved thread.');
    return;
  }

  const ambiguous = validMessages.some(m => {
    const d = extractEmailData(m);
    return d['rider-number'] !== data['rider-number'] || d['bonus'] !== data['bonus'];
  });
  if (ambiguous) {
    Logger.log('-> Skipped: thread has multiple valid messages referencing different ' +
      'rider/bonus values, so it is unclear which one the manual "approved" label ' +
      'was meant for. Use the sidebar\'s per-message Approve button instead.');
    thread.addLabel(labels.processingError);
    thread.refresh();
    return;
  }

  try {
    updateSpreadsheet(ss, config, data,
      configInt_(config, 'col_approved', 4, 1),
      configInt_(config, 'col_approved_time', 5, 1),
      false);
    thread.removeLabel(labels.needsReview);
    thread.addLabel(labels.scored);
    thread.refresh();
    Logger.log('-> Approved: Rider ' + data['rider-number'] + ' - ' + data['bonus']);
  } catch (e) {
    Logger.log('-> Error: ' + e.message);
    thread.addLabel(labels.processingError);
    thread.refresh();
  }
}

// --- Spreadsheet update -------------------------------------------------------

/**
 * Finds the row in `sheet` whose bonus-ID column matches bonusCode
 * (case-insensitive, whitespace-trimmed on both sides). Returns the row
 * number, or null if there are no data rows below the header or no row
 * matches.
 */
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

/**
 * Finds the bonus row and writes X + timestamp to columnIndex / timeColumnIndex.
 * Pass null for value to clear both cells (revert).
 *
 * Holds a script-wide lock for the read-find-write sequence so the time-driven
 * trigger and interactive sidebar clicks can't race on the same row.
 */
function updateSpreadsheet(ss, config, data, columnIndex, timeColumnIndex, useEmailTime, value) {
  const riderNumber = data['rider-number'];
  const bonusToFind = data['bonus'];

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let sheet = ss.getSheetByName(riderNumber);
    if (!sheet) sheet = createRiderSheet_(ss, config, riderNumber);

    const startRow = configInt_(config, 'header_row', 1, 1) + 1;
    if (sheet.getLastRow() < startRow) throw new Error('No data rows in sheet ' + riderNumber);

    const row = findBonusRow_(sheet, config, bonusToFind);
    if (row === null) throw new Error('Bonus ID "' + bonusToFind + '" not found in sheet ' + riderNumber);

    if (value === null) {
      // Revert - clear both value and timestamp
      sheet.getRange(row, columnIndex).clearContent();
      sheet.getRange(row, timeColumnIndex).clearContent();
    } else {
      sheet.getRange(row, columnIndex).setValue('X');
      sheet.getRange(row, timeColumnIndex).setValue(useEmailTime ? data.date : new Date());
    }
  } finally {
    lock.releaseLock();
  }
}

// --- Validation ---------------------------------------------------------------
function validateEmailAddress(ss, config, senderString, riderNumber) {
  const re = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/;
  const sm = senderString.match(re);
  if (!sm) { Logger.log('No valid email in: ' + senderString); return false; }
  const senderEmail = sm[0].toLowerCase();

  const masterSheet = ss.getSheetByName(config['sheet_rider_master'] || 'Rider Master');
  if (!masterSheet) { Logger.log('Rider Master sheet not found.'); return false; }
  const rows = masterSheet.getDataRange().getValues();
  if (rows.length < 2) return false;

  const headers = rows[0];
  const rci = headers.indexOf(config['master_col_rider_number'] || 'Rider Number');
  const eci = headers.indexOf(config['master_col_email'] || 'Email');
  if (rci === -1 || eci === -1) { Logger.log('Missing column headers in Rider Master.'); return false; }

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][rci]).trim() !== String(riderNumber).trim()) continue;
    const em = String(rows[i][eci]).trim().match(re);
    if (!em) { Logger.log('Invalid email in row ' + (i + 1)); return false; }
    const reg = em[0].toLowerCase();
    if (reg === senderEmail) { Logger.log('Email validated for Rider ' + riderNumber); return true; }
    Logger.log('Mismatch for Rider ' + riderNumber + ': expected ' + reg + ', got ' + senderEmail);
    return false;
  }
  Logger.log('Rider ' + riderNumber + ' not found.');
  return false;
}

// --- Helpers ------------------------------------------------------------------
// Shared with Sidebar.gs via Apps Script's global scope - don't duplicate there.
function loadLabels_(config) {
  const defs = {
    unprocessed:     'label_unprocessed',
    formatError:     'label_format_error',
    emailError:      'label_email_error',
    processingError: 'label_processing_error',
    needsReview:     'label_needs_review',
    approved:        'label_approved',
    denied:          'label_denied',
    scored:          'label_scored',
  };
  const labels = {}, missing = [];
  for (const [key, ck] of Object.entries(defs)) {
    const name = config[ck];
    if (!name) { missing.push(ck); continue; }
    const label = GmailApp.getUserLabelByName(name);
    if (!label) missing.push('Gmail label: ' + name);
    else labels[key] = label;
  }
  if (missing.length) {
    Logger.log('ERROR: Missing - ' + missing.join(', ') + '. Run setup() first.');
    return null;
  }
  return labels;
}

// Bonus code shape: either 4 letters (e.g. "ABCD"), or 3 letters followed by
// exactly 1 digit (e.g. "ABC1") - no other letter/digit combination is valid
// (not 4 letters + a digit, not 3 letters + 2 digits, not fewer than 3 letters).
function isValidSubject(subject) {
  return /^\d+\s*(?:[A-Za-z]{4}|[A-Za-z]{3}\d)$/.test(subject.trim());
}

function extractEmailData(message) {
  const subject = message.getSubject().trim();
  const match = subject.match(/^(\d+)\s*([A-Za-z]{4}|[A-Za-z]{3}\d)$/);
  return {
    date:          message.getDate(),
    subject:       subject,
    sender:        message.getFrom(),
    'rider-number': match ? match[1] : null,
    'bonus':        match ? match[2].toUpperCase() : null,
  };
}