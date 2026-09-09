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
const { loadApp, registerActiveSpreadsheet, MockSpreadsheet } = require('./support/mocks');

const KEEP_SHEETS = ['Config', 'Rider Master', 'Bonus Master'];

describe('deleteAllRiderSheets', () => {
  test('deletes rider sheets and leaves every kept sheet untouched', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    KEEP_SHEETS.forEach((name) => ss.addSheet(name, [['header']]));
    ss.addSheet('42', [['Bonus ID']]);
    ss.addSheet('7', [['Bonus ID']]);
    registerActiveSpreadsheet(env, ss);

    env.context.deleteAllRiderSheets();

    const remaining = ss.getSheets().map((s) => s.getName());
    assert.deepEqual(remaining.sort(), [...KEEP_SHEETS].sort());
    assert.equal(ss.getSheetByName('42'), null);
    assert.equal(ss.getSheetByName('7'), null);
  });

  test('deletes Leader Board along with the rider sheets - it is not a kept sheet', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    KEEP_SHEETS.forEach((name) => ss.addSheet(name, [['header']]));
    ss.addSheet('Leader Board', [['Rider Number', 'Name', 'Score', 'Finish']]);
    ss.addSheet('42', [['Bonus ID']]);
    registerActiveSpreadsheet(env, ss);

    env.context.deleteAllRiderSheets();

    assert.equal(ss.getSheetByName('Leader Board'), null);
    assert.equal(ss.getSheetByName('42'), null);
    KEEP_SHEETS.forEach((name) => assert.ok(ss.getSheetByName(name), name + ' should still exist'));
  });

  test('deletes Master Scoring along with the rider sheets - it is not a kept sheet', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    KEEP_SHEETS.forEach((name) => ss.addSheet(name, [['header']]));
    ss.addSheet('Master Scoring', [['', 'Name'], ['', 'Number'], ['', 'Score'], ['Bonus', 'POINTS']]);
    ss.addSheet('42', [['Bonus ID']]);
    registerActiveSpreadsheet(env, ss);

    env.context.deleteAllRiderSheets();

    assert.equal(ss.getSheetByName('Master Scoring'), null);
    assert.equal(ss.getSheetByName('42'), null);
    KEEP_SHEETS.forEach((name) => assert.ok(ss.getSheetByName(name), name + ' should still exist'));
  });

  test('does nothing (and does not throw) when there are no rider sheets to delete', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    KEEP_SHEETS.forEach((name) => ss.addSheet(name, [['header']]));
    registerActiveSpreadsheet(env, ss);

    assert.doesNotThrow(() => env.context.deleteAllRiderSheets());

    assert.deepEqual(ss.getSheets().map((s) => s.getName()).sort(), [...KEEP_SHEETS].sort());
  });

  test('operates on the active spreadsheet, not one resolved by Script Properties', () => {
    // deleteAllRiderSheets uses SpreadsheetApp.getActiveSpreadsheet(), the
    // same pattern setup() uses - not openById(SPREADSHEET_ID) like
    // processEmails()/the sidebar. Confirm it still works when only
    // "active" is set (registerActiveSpreadsheet sets both, so this also
    // covers the common case, but the assertion below is what actually
    // matters: it didn't need SPREADSHEET_ID at all).
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [['header']]);
    ss.addSheet('99', [['Bonus ID']]);
    env.spreadsheetApp.setActive(ss); // active only - no Script Properties involved

    env.context.deleteAllRiderSheets();

    assert.equal(ss.getSheetByName('99'), null);
    assert.ok(ss.getSheetByName('Config'));
  });
});
