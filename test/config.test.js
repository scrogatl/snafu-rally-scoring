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

// Last edited: 2026-07-22
'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, MockSpreadsheet, toPlain } = require('./support/mocks');

describe('loadConfig', () => {
  test('reads key/value rows from the Config sheet, skipping the header row', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [
      ['key', 'value', 'notes'],
      ['header_row', '1', ''],
      ['col_bonus_id', '1', ''],
      ['label_parent', 'rally', ''],
    ]);
    const config = context.loadConfig(ss);
    assert.equal(config.header_row, '1');
    assert.equal(config.col_bonus_id, '1');
    assert.equal(config.label_parent, 'rally');
  });

  test('trims whitespace around keys and values', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [
      ['key', 'value'],
      ['  label_parent  ', '  rally  '],
    ]);
    const config = context.loadConfig(ss);
    assert.equal(config.label_parent, 'rally');
  });

  test('skips rows with a blank key', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', [
      ['key', 'value'],
      ['', 'ignored'],
      ['label_parent', 'rally'],
    ]);
    const config = context.loadConfig(ss);
    assert.deepEqual(toPlain(config), { label_parent: 'rally' });
  });

  test('throws when there is no sheet named Config', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    assert.throws(() => context.loadConfig(ss), /Config sheet not found/);
  });
});

describe('validateEmailAddress', () => {
  function buildSpreadsheet() {
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Rider Master', [
      ['Rider Number', 'Name', 'Email'],
      ['42', 'Jane Smith', 'jane@example.com'],
      ['7', 'Bob Jones', 'bob@example.com'],
    ]);
    return ss;
  }

  test('accepts a sender whose email matches the registered rider', () => {
    const { context } = loadApp();
    const ss = buildSpreadsheet();
    const ok = context.validateEmailAddress(ss, {}, 'Jane Smith <jane@example.com>', '42');
    assert.equal(ok, true);
  });

  test('matches case-insensitively', () => {
    const { context } = loadApp();
    const ss = buildSpreadsheet();
    const ok = context.validateEmailAddress(ss, {}, 'JANE@EXAMPLE.COM', '42');
    assert.equal(ok, true);
  });

  test('rejects a sender whose email does not match the registered rider', () => {
    const { context } = loadApp();
    const ss = buildSpreadsheet();
    const ok = context.validateEmailAddress(ss, {}, 'someone-else@example.com', '42');
    assert.equal(ok, false);
  });

  test('rejects an unknown rider number', () => {
    const { context } = loadApp();
    const ss = buildSpreadsheet();
    const ok = context.validateEmailAddress(ss, {}, 'jane@example.com', '999');
    assert.equal(ok, false);
  });

  test('rejects when the sender string has no parseable email address', () => {
    const { context } = loadApp();
    const ss = buildSpreadsheet();
    const ok = context.validateEmailAddress(ss, {}, 'not an email', '42');
    assert.equal(ok, false);
  });

  test('rejects when Rider Master sheet is missing', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const ok = context.validateEmailAddress(ss, {}, 'jane@example.com', '42');
    assert.equal(ok, false);
  });

  test('honors configured sheet/column name overrides', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Riders', [
      ['Num', 'Mail'],
      ['42', 'jane@example.com'],
    ]);
    const config = { sheet_rider_master: 'Riders', master_col_rider_number: 'Num', master_col_email: 'Mail' };
    assert.equal(context.validateEmailAddress(ss, config, 'jane@example.com', '42'), true);
  });
});
