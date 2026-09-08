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

function buildMasterScoring(ss, numRiderColumns) {
  // Columns A/B are whatever the organizer uses for their own bookkeeping -
  // rider data starts at column C, row 2 (rider numbers) / row 3 (scores).
  const row2 = ['', '', ...Array.from({ length: numRiderColumns }, (_, i) => i + 1)];
  const row3 = ['', '', ...Array.from({ length: numRiderColumns }, () => 0)];
  return ss.addSheet('Master Scoring', [[], row2, row3]);
}

describe('createLeaderBoard_', () => {
  test('creates the sheet with the header row, frozen, and one row per rider', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    buildMasterScoring(ss, 2);

    context.createLeaderBoard_(ss, {});
    const lb = ss.getSheetByName('Leader Board');

    assert.ok(lb, 'Leader Board should have been created');
    assert.deepEqual(lb.getRange(1, 1, 1, 4).getValues(), [['Rider Number', 'Name', 'Score', 'Finish']]);
    assert.equal(lb.frozenRows, 1);
  });

  test('is created as the leftmost tab, unlike every other sheet this script creates', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [['key', 'value']]);
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildMasterScoring(ss, 1);

    context.createLeaderBoard_(ss, {});

    assert.equal(ss.getSheets()[0].getName(), 'Leader Board');
  });

  test('re-running does not move an already-existing Leader Board', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [['key', 'value']]);
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildMasterScoring(ss, 1);
    context.createLeaderBoard_(ss, {});
    ss.addSheet('Some Other Tab', [['x']]); // added after Leader Board already exists

    context.createLeaderBoard_(ss, {}); // no new riders, but should not reposition anything

    assert.equal(ss.getSheets()[0].getName(), 'Leader Board');
  });

  test('writes the exact expected formula in each column, sized to actual sheet extents', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);
    buildMasterScoring(ss, 2); // columns C, D used -> last column D (index 4)

    context.createLeaderBoard_(ss, {});
    const lb = ss.getSheetByName('Leader Board');

    // Row 2 = rider 1 (Rider Master row 2), row 3 = rider 2 (Rider Master row 3).
    assert.equal(lb._rawCell(2, 1), "='Rider Master'!A2");
    assert.equal(lb._rawCell(2, 2), "=VLOOKUP(A2,'Rider Master'!$A$2:$B$3,2)");
    assert.equal(lb._rawCell(2, 3), "=HLOOKUP(A2,'Master Scoring'!$C$2:$D$3,2)");
    assert.equal(lb._rawCell(2, 4), '=RANK(C2,$C:$C,0)');

    assert.equal(lb._rawCell(3, 1), "='Rider Master'!A3");
    assert.equal(lb._rawCell(3, 2), "=VLOOKUP(A3,'Rider Master'!$A$2:$B$3,2)");
    assert.equal(lb._rawCell(3, 3), "=HLOOKUP(A3,'Master Scoring'!$C$2:$D$3,2)");
    assert.equal(lb._rawCell(3, 4), '=RANK(C3,$C:$C,0)');

    // Column A is a plain cell reference - the mock CAN resolve this one.
    assert.equal(lb.getRange(2, 1).getValue(), '1');
    assert.equal(lb.getRange(3, 1).getValue(), '2');
  });

  test('a wider Master Scoring produces a correspondingly wider HLOOKUP range', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildMasterScoring(ss, 96); // columns C..CT (96 rider columns) -> last column CT

    context.createLeaderBoard_(ss, {});
    const lb = ss.getSheetByName('Leader Board');

    assert.match(lb._rawCell(2, 3), /\$CT\$3,2\)$/);
  });

  test('re-running adds rows only for riders not already present, leaving existing rows untouched', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    buildMasterScoring(ss, 1);
    context.createLeaderBoard_(ss, {});
    const lb = ss.getSheetByName('Leader Board');
    const originalRow2Name = lb._rawCell(2, 2);

    // Roster grows - rider 2 added, rider 1 unchanged.
    ss.deleteSheet(ss.getSheetByName('Rider Master'));
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com'], ['2', 'Bob', 'bob@example.com']]);

    context.createLeaderBoard_(ss, {});

    assert.equal(lb._rawCell(2, 2), originalRow2Name, 'existing row must not be rewritten');
    assert.equal(lb._rawCell(3, 1), "='Rider Master'!A3");
    assert.equal(ss.getSheetByName('Leader Board'), lb, 'must not have created a second sheet');
  });

  test('throws a clear error when Master Scoring is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['1', 'Jane', 'jane@example.com']]);
    // No Master Scoring sheet at all.

    assert.throws(
      () => context.createLeaderBoard_(ss, {}),
      /Cannot create Leader Board.*Master Scoring.*not found/
    );
    assert.equal(ss.getSheetByName('Leader Board'), null, 'must not create a half-built sheet');
  });

  test('does nothing (no throw) when Rider Master is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildMasterScoring(ss, 1);

    assert.doesNotThrow(() => context.createLeaderBoard_(ss, {}));
    assert.equal(ss.getSheetByName('Leader Board'), null);
  });

  test('honors configured sheet_rider_master/sheet_master_scoring/sheet_leader_board names', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Riders', [['Rider Number', 'Name', 'Email'], ['1', 'Jane', 'jane@example.com']]);
    ss.addSheet('Scores', [[], ['', '', 1], ['', '', 0]]);
    const config = {
      sheet_rider_master: 'Riders',
      sheet_master_scoring: 'Scores',
      sheet_leader_board: 'Standings',
    };

    context.createLeaderBoard_(ss, config);

    const lb = ss.getSheetByName('Standings');
    assert.ok(lb);
    assert.equal(lb._rawCell(2, 1), "='Riders'!A2");
    assert.match(lb._rawCell(2, 3), /'Scores'!\$C\$2:\$C\$3,2\)$/);
  });
});
