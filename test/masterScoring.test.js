// snafu-rally-scoring
// Copyright (C) 2026 Scott Rogers
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

// Last edited: 2026-09-08
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, MockSpreadsheet } = require('./support/mocks');

function buildRiderMaster(ss, riders) {
  return ss.addSheet('Rider Master', [['Rider Number', 'Name', 'Email'], ...riders]);
}

function buildBonusMaster(ss, bonuses) {
  // bonuses: [[code, points], ...]
  return ss.addSheet('Bonus Master', [['Bonus ID', 'POINTS'], ...bonuses]);
}

function buildComboMaster(ss, pairs) {
  // pairs: [[comboCode, memberCode], ...] - one row per member, so a combo
  // with N members takes N rows.
  return ss.addSheet('Combo Master', [['Combo ID', 'Member Bonus ID'], ...pairs]);
}

describe('createMasterScoring_', () => {
  test('creates the sheet with header labels, one column per rider, one row per bonus', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100'], ['WXYZ', '50']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.ok(scoring, 'Master Scoring should have been created');
    assert.deepEqual(scoring.getRange(1, 2, 3, 1).getValues(), [['Name'], ['Number'], ['Score']]);
    assert.deepEqual(scoring.getRange(4, 1, 1, 2).getValues(), [['Bonus', 'POINTS']]);
  });

  test('writes the exact expected formula for each rider column (Name/Number/Score)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    // Rider 1 -> column C, rider 2 -> column D.
    assert.equal(scoring._rawCell(1, 3), "=VLOOKUP(C2,'Rider Master'!$A$2:$B$3,2)");
    assert.equal(scoring._rawCell(2, 3), "='Rider Master'!A2");
    assert.equal(scoring._rawCell(3, 3), '=SUMIF(C5:C,"X",$B5:$B)');
    assert.equal(scoring._rawCell(1, 4), "=VLOOKUP(D2,'Rider Master'!$A$2:$B$3,2)");
    assert.equal(scoring._rawCell(2, 4), "='Rider Master'!A3");
    assert.equal(scoring._rawCell(3, 4), '=SUMIF(D5:D,"X",$B5:$B)');

    // Column A/B resolve the rider number (a plain cell reference).
    assert.equal(scoring.getRange(2, 3).getValue(), '1');
    assert.equal(scoring.getRange(2, 4).getValue(), '2');
  });

  test('regression: Score range must not include row 3 (its own cell) - a true whole-column range causes a real Sheets circular-reference error', () => {
    // A range like 'D:D' includes D3, the Score formula's own cell - Sheets
    // flags that as circular even though the SUMIF criteria would never
    // match D3's own content. Anchoring the range at row 5 (D5:D) keeps the
    // "grows automatically, never needs revisiting" property of an open-ended
    // range without including row 3.
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.match(scoring._rawCell(3, 3), /^=SUMIF\(C5:C,"X",\$B5:\$B\)$/);
    assert.doesNotMatch(scoring._rawCell(3, 3), /\(C:C/, 'must not be a true whole-column range');
  });

  test('writes the exact expected formula for each bonus row (Bonus/POINTS/Approved-check)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100'], ['WXYZ', '50']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.equal(scoring._rawCell(5, 1), "='Bonus Master'!A2");
    assert.equal(scoring._rawCell(5, 2), "='Bonus Master'!B2");
    assert.equal(scoring._rawCell(5, 3), "='1'!D2"); // col_approved defaults to 4 (D), header_row defaults to 1 -> rider sheet row 2
    assert.equal(scoring._rawCell(6, 1), "='Bonus Master'!A3");
    assert.equal(scoring._rawCell(6, 2), "='Bonus Master'!B3");
    assert.equal(scoring._rawCell(6, 3), "='1'!D3");
  });

  test('honors configured header_row and col_approved for the Approved-check reference', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, { header_row: '2', col_approved: '10' });
    const scoring = ss.getSheetByName('Master Scoring');

    // header_row=2 -> first bonus lands at rider-sheet row (2+1+0)=3; col_approved=10 -> column J.
    assert.equal(scoring._rawCell(5, 3), "='1'!J3");
  });

  test('growth: adding a rider and re-running adds exactly one new column, backfilled for existing bonus rows, leaving existing cells untouched', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100'], ['WXYZ', '50']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');
    const originalCol3Row1 = scoring._rawCell(1, 3);

    ss.deleteSheet(ss.getSheetByName('Rider Master'));
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    context.createMasterScoring_(ss, {});

    assert.equal(scoring._rawCell(1, 3), originalCol3Row1, 'existing rider column must not be rewritten');
    assert.equal(scoring.getLastColumn(), 4, 'exactly one new column added');
    assert.equal(scoring._rawCell(2, 4), "='Rider Master'!A3");
    // New column backfilled for both existing bonus rows.
    assert.equal(scoring._rawCell(5, 4), "='2'!D2");
    assert.equal(scoring._rawCell(6, 4), "='2'!D3");
  });

  test('growth: adding a bonus and re-running adds exactly one new row, backfilled for existing rider columns, leaving existing cells untouched', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');
    const originalRow5Col3 = scoring._rawCell(5, 3);

    ss.deleteSheet(ss.getSheetByName('Bonus Master'));
    buildBonusMaster(ss, [['ABCD', '100'], ['WXYZ', '50']]);
    context.createMasterScoring_(ss, {});

    assert.equal(scoring._rawCell(5, 3), originalRow5Col3, 'existing bonus row must not be rewritten');
    assert.equal(scoring.getLastRow(), 6, 'exactly one new bonus row added');
    assert.equal(scoring._rawCell(6, 1), "='Bonus Master'!A3");
    // New row backfilled for both existing rider columns.
    assert.equal(scoring._rawCell(6, 3), "='1'!D3");
    assert.equal(scoring._rawCell(6, 4), "='2'!D3");
  });

  test('re-running with nothing new added is a no-op (idempotent)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    context.createMasterScoring_(ss, {});

    assert.equal(scoring.getLastColumn(), 3);
    assert.equal(scoring.getLastRow(), 5);
  });

  test('throws a clear error when Bonus Master is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    // No Bonus Master sheet at all.

    assert.throws(
      () => context.createMasterScoring_(ss, {}),
      /Cannot create Master Scoring.*Bonus Master.*not found/
    );
    assert.equal(ss.getSheetByName('Master Scoring'), null, 'must not create a half-built sheet');
  });

  test('does nothing (no throw) when Rider Master is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildBonusMaster(ss, [['ABCD', '100']]);

    assert.doesNotThrow(() => context.createMasterScoring_(ss, {}));
    assert.equal(ss.getSheetByName('Master Scoring'), null);
  });

  test('a blank POINTS value defaults to 0 with no warning logged', () => {
    const { context, logger } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID', 'POINTS'], ['ABCD', '']]);

    context.createMasterScoring_(ss, {});

    assert.ok(!logger.logs.some((l) => l.includes('non-numeric')));
  });

  test('a non-numeric POINTS value logs a warning and still creates the row', () => {
    const { context, logger } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID', 'POINTS'], ['ABCD', 'oops']]);

    context.createMasterScoring_(ss, {});

    assert.ok(logger.logs.some((l) => l.includes('non-numeric') && l.includes('ABCD')));
    const scoring = ss.getSheetByName('Master Scoring');
    assert.equal(scoring._rawCell(5, 1), "='Bonus Master'!A2");
  });

  test('honors configured sheet_rider_master/sheet_bonus_master/sheet_master_scoring names', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Riders', [['Rider Number', 'Name', 'Email'], ['1', 'Jane', 'jane@example.com']]);
    ss.addSheet('Bonuses', [['Bonus ID', 'POINTS'], ['ABCD', '100']]);
    const config = {
      sheet_rider_master: 'Riders',
      sheet_bonus_master: 'Bonuses',
      sheet_master_scoring: 'Scoring',
    };

    context.createMasterScoring_(ss, config);

    const scoring = ss.getSheetByName('Scoring');
    assert.ok(scoring);
    assert.equal(scoring._rawCell(2, 3), "='Riders'!A2");
    assert.equal(scoring._rawCell(5, 1), "='Bonuses'!A2");
    assert.equal(scoring._rawCell(5, 3), "='1'!D2");
  });

  test('center-aligns the entire sheet, header rows/columns included', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.equal(scoring._rawAlign(1, 2), 'center', 'header label (Name)');
    assert.equal(scoring._rawAlign(4, 1), 'center', 'header label (Bonus)');
    assert.equal(scoring._rawAlign(3, 3), 'center', 'rider Score formula');
    assert.equal(scoring._rawAlign(5, 3), 'center', 'bonus Approved-check cell');
  });

  test('applies row banding over the full sheet, replacing any prior banding rather than stacking it', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    // Grow the sheet and re-run - must not throw "overlaps an existing
    // banding", and must still end up with exactly one banding covering the
    // new full extent.
    ss.deleteSheet(ss.getSheetByName('Rider Master'));
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    assert.doesNotThrow(() => context.createMasterScoring_(ss, {}));

    assert.equal(scoring.getBandings().length, 1, 'old banding must be removed, not stacked');
    const banding = scoring.getBandings()[0];
    assert.equal(banding.row, 1);
    assert.equal(banding.col, 1);
    assert.equal(banding.numRows, scoring.getLastRow());
    assert.equal(banding.numCols, scoring.getLastColumn());
  });

  test('conditional formatting turns the data grid green on "X", scoped to row 5+/column C+ only', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    const rules = scoring.getConditionalFormatRules();
    assert.equal(rules.length, 1);
    assert.deepEqual(rules[0].condition, { type: 'TEXT_EQ', value: 'X' });
    assert.equal(rules[0].background, '#b7e1cd');
    assert.equal(rules[0].ranges.length, 1);
    const r = rules[0].ranges[0];
    assert.equal(r.row, 5, 'data range must start at row 5, not include the header rows');
    assert.equal(r.col, 3, 'data range must start at column C, not include the Bonus/POINTS columns');
  });

  test('conditional formatting is rebuilt (not stacked) on re-run, growing to cover new rows/columns', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    ss.deleteSheet(ss.getSheetByName('Bonus Master'));
    buildBonusMaster(ss, [['ABCD', '100'], ['WXYZ', '50']]);
    context.createMasterScoring_(ss, {});

    const rules = scoring.getConditionalFormatRules();
    assert.equal(rules.length, 1, 'must not accumulate a second rule on re-run');
    const r = rules[0].ranges[0];
    assert.equal(r.numRows, scoring.getLastRow() - 4);
  });

  test('sets warning-only protection on the sheet at creation time', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    const protections = scoring.getProtections();
    assert.equal(protections.length, 1);
    assert.equal(protections[0].getWarningOnly(), true, 'must be edit-with-warning, not access-restricted');
    assert.ok(protections[0].getDescription().length > 0, 'should explain why the sheet is protected');
  });

  test('re-running does not add a second protection on an already-protected sheet', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    ss.deleteSheet(ss.getSheetByName('Rider Master'));
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    context.createMasterScoring_(ss, {});

    assert.equal(scoring.getProtections().length, 1, 'growth re-runs must not stack a second protection');
  });

  test('is appended after existing sheets, not leftmost (unlike Leader Board)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [['key', 'value']]);
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '100']]);

    context.createMasterScoring_(ss, {});

    assert.deepEqual(
      ss.getSheets().map((s) => s.getName()),
      ['Config', 'Rider Master', 'Bonus Master', 'Master Scoring']
    );
  });
});

describe('createMasterScoring_ - combo auto-approval (Combo Master)', () => {
  test('combo cell is "X" once every member bonus is approved for that rider', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['WXYZ', '20'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD'], ['COMB', 'WXYZ']]);
    // Rider 1's Approved column (col_approved default = D), rows matching
    // header_row=1 -> ABCD at rider-sheet row 2, WXYZ at row 3, COMB at row 4.
    ss.addSheet('1', [
      ['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time'],
      ['ABCD', '', '', 'X', '', '', ''],
      ['WXYZ', '', '', '', '', '', ''], // not yet approved
      ['COMB', '', '', '', '', '', ''],
    ]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    // Master Scoring rows: 5=ABCD, 6=WXYZ, 7=COMB (Bonus Master order).
    assert.match(scoring._rawCell(7, 3), /^=IF\(OR\(AND\(C5="X",C6="X"\),'1'!D4="X"\),"X",""\)$/);
  });

  test('combo Master Scoring cell honors col_approved/header_row when computing the direct-approval fallback', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);

    context.createMasterScoring_(ss, { header_row: '2', col_approved: '10' });
    const scoring = ss.getSheetByName('Master Scoring');

    // header_row=2 -> ABCD at rider-sheet row 3, COMB at row 4; col_approved=10 -> column J.
    assert.match(scoring._rawCell(6, 3), /^=IF\(OR\(AND\(C5="X"\),'1'!J4="X"\),"X",""\)$/);
  });

  test('a regular (non-combo) bonus row keeps the plain direct-reference formula, unaffected by Combo Master', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.equal(scoring._rawCell(5, 3), "='1'!D2"); // ABCD's own row - not a combo
  });

  test('no Combo Master sheet at all -> every row is a plain direct reference (no regression)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10']]);
    // No Combo Master sheet added.

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.equal(scoring._rawCell(5, 3), "='1'!D2");
  });

  test('a combo with a single member still produces a valid AND(...) formula', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.match(scoring._rawCell(6, 3), /^=IF\(OR\(AND\(C5="X"\),'1'!D3="X"\),"X",""\)$/);
  });

  test('a combo with 5 members produces a correctly-anded formula (no fixed member-count limit)', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [
      ['AAAA', '1'], ['BBBB', '1'], ['CCCC', '1'], ['DDDD', '1'], ['EEEE', '1'], ['COMB', '100'],
    ]);
    buildComboMaster(ss, [
      ['COMB', 'AAAA'], ['COMB', 'BBBB'], ['COMB', 'CCCC'], ['COMB', 'DDDD'], ['COMB', 'EEEE'],
    ]);

    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    assert.match(
      scoring._rawCell(10, 3),
      /^=IF\(OR\(AND\(C5="X",C6="X",C7="X",C8="X",C9="X"\),'1'!D7="X"\),"X",""\)$/
    );
  });

  test('combo code referenced by Combo Master but missing from Bonus Master logs a warning and falls back to plain reference', () => {
    const { context, logger } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10']]); // no "COMB" entry
    buildComboMaster(ss, [['COMB', 'ABCD']]);

    assert.doesNotThrow(() => context.createMasterScoring_(ss, {}));

    assert.ok(logger.logs.some((l) => l.includes('"COMB"') && l.includes('not a Bonus Master entry')));
    const scoring = ss.getSheetByName('Master Scoring');
    // Only ABCD got a row - COMB was never a real Bonus Master entry.
    assert.equal(scoring.getLastRow(), 5);
  });

  test('member code referenced by Combo Master but missing from Bonus Master logs a warning and the combo falls back to plain reference', () => {
    const { context, logger } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['COMB', '100']]); // no "ABCD" entry - the member is missing
    buildComboMaster(ss, [['COMB', 'ABCD']]);

    assert.doesNotThrow(() => context.createMasterScoring_(ss, {}));

    assert.ok(logger.logs.some((l) =>
      l.includes('"COMB"') && l.includes('"ABCD"') && l.includes('not a Bonus Master entry')));
    const scoring = ss.getSheetByName('Master Scoring');
    // Falls back to a plain direct reference, same as a regular bonus row.
    assert.equal(scoring._rawCell(5, 3), "='1'!D2");
  });

  test('growth: a new rider column added after the combo already exists gets the full combo formula', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');

    ss.deleteSheet(ss.getSheetByName('Rider Master'));
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    context.createMasterScoring_(ss, {});

    assert.match(scoring._rawCell(6, 4), /^=IF\(OR\(AND\(D5="X"\),'2'!D3="X"\),"X",""\)$/);
  });

  test('growth: a new bonus row that is itself a combo (added on a later run) gets the combo formula for existing rider columns', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10']]);
    context.createMasterScoring_(ss, {}); // no combo yet
    const scoring = ss.getSheetByName('Master Scoring');

    ss.deleteSheet(ss.getSheetByName('Bonus Master'));
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);
    context.createMasterScoring_(ss, {});

    assert.match(scoring._rawCell(6, 3), /^=IF\(OR\(AND\(C5="X"\),'1'!D3="X"\),"X",""\)$/);
  });

  test('regression: re-running with nothing new does not rewrite an already-written combo cell', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildBonusMaster(ss, [['ABCD', '10'], ['COMB', '100']]);
    buildComboMaster(ss, [['COMB', 'ABCD']]);
    context.createMasterScoring_(ss, {});
    const scoring = ss.getSheetByName('Master Scoring');
    const before = scoring._rawCell(6, 3);

    context.createMasterScoring_(ss, {});

    assert.equal(scoring._rawCell(6, 3), before);
  });
});
