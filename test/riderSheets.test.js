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

describe('createRiderSheet_', () => {
  test('writes the standard header row and freezes it', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Bonus Master', [['Bonus ID']]);

    const sheet = context.createRiderSheet_(ss, { header_row: '1' }, '42');

    assert.deepEqual(sheet.getRange(1, 1, 1, 7).getValues(), [
      ['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time'],
    ]);
    assert.equal(sheet.frozenRows, 1);
  });

  test('writes cell-reference formulas into column A for each Bonus Master row', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD'], ['WXYZ']]);

    const sheet = context.createRiderSheet_(ss, { header_row: '1' }, '42');

    assert.equal(sheet._rawCell(2, 1), "='Bonus Master'!A2");
    assert.equal(sheet._rawCell(3, 1), "='Bonus Master'!A3");
    // and the formulas resolve to the actual bonus IDs when read
    assert.deepEqual(sheet.getRange(2, 1, 2, 1).getValues(), [['ABCD'], ['WXYZ']]);
  });

  test('formats Submit/Approve/Deny Time columns as Date+Time, not just Date', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD'], ['WXYZ']]);

    const sheet = context.createRiderSheet_(ss, { header_row: '1' }, '42');

    // Column C (3) = Submit Time, E (5) = Approve Time, G (7) = Deny Time.
    for (const col of [3, 5, 7]) {
      for (const row of [2, 3]) { // one per bonus data row
        assert.equal(sheet.getRange(row, col).getNumberFormat(), 'M/d/yyyy h:mm:ss am/pm');
      }
    }
    // Untouched columns keep the sheet's default format.
    assert.equal(sheet.getRange(2, 1).getNumberFormat(), 'General');
    assert.equal(sheet.getRange(2, 2).getNumberFormat(), 'General');
  });

  test('bonus IDs stay in sync when Bonus Master changes after the sheet is created', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const bonusMaster = ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);
    const sheet = context.createRiderSheet_(ss, { header_row: '1' }, '42');
    assert.equal(sheet.getRange(2, 1).getValue(), 'ABCD');

    bonusMaster.getRange(2, 1).setValue('ZZZZ');

    assert.equal(sheet.getRange(2, 1).getValue(), 'ZZZZ');
  });

  test('honors a configured header_row and a renamed Bonus Master sheet', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Bonuses', [['Bonus ID'], ['ABCD']]);

    const sheet = context.createRiderSheet_(ss, { header_row: '3', sheet_bonus_master: 'Bonuses' }, '42');

    assert.equal(sheet.getRange(3, 1).getValue(), 'Bonus ID');
    assert.equal(sheet.getRange(4, 1).getValue(), 'ABCD');
  });

  test('throws when the Bonus Master sheet is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    assert.throws(() => context.createRiderSheet_(ss, {}, '42'), /Bonus Master.*not found/);
  });

  test('appends the new sheet after existing sheets', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [['key', 'value']]);
    ss.addSheet('Bonus Master', [['Bonus ID']]);

    context.createRiderSheet_(ss, {}, '42');

    assert.deepEqual(ss.getSheets().map((s) => s.getName()), ['Config', 'Bonus Master', '42']);
  });

});

describe('createAllRiderSheets_', () => {
  function buildRoster(ss, riders) {
    ss.addSheet('Rider Master', [['Rider Number', 'Name', 'Email'], ...riders]);
  }

  test('creates one sheet per rider listed in Rider Master', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRoster(ss, [['42', 'Jane', 'jane@example.com'], ['7', 'Bob', 'bob@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);

    context.createAllRiderSheets_(ss, {});

    assert.ok(ss.getSheetByName('42'));
    assert.ok(ss.getSheetByName('7'));
  });

  test('skips riders that already have a sheet', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRoster(ss, [['42', 'Jane', 'jane@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);
    const existing = ss.addSheet('42', [['Bonus ID'], ['PRESET']]);

    context.createAllRiderSheets_(ss, {});

    assert.equal(ss.getSheetByName('42'), existing); // untouched, not recreated
  });

  test('skips rows with a blank rider number', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRoster(ss, [['', 'No rider number', ''], ['7', 'Bob', 'bob@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);

    context.createAllRiderSheets_(ss, {});

    assert.equal(ss.getSheetByName('7') !== null, true);
    assert.equal(ss.getSheets().filter((s) => s.getName() !== 'Rider Master' && s.getName() !== 'Bonus Master').length, 1);
  });

  test('does nothing when Rider Master is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    // Should not throw even with no Rider Master sheet at all.
    assert.doesNotThrow(() => context.createAllRiderSheets_(ss, {}));
  });

  test('does nothing when the configured Rider Number column header is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Rider Master', [['Name', 'Email'], ['Jane', 'jane@example.com']]);
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);

    assert.doesNotThrow(() => context.createAllRiderSheets_(ss, {}));
    assert.equal(ss.getSheets().length, 2); // no rider sheet created
  });
});
