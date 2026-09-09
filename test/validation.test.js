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
const { loadApp } = require('./support/mocks');

describe('isValidSubject', () => {
  const { context } = loadApp();

  test('accepts rider number + 4-letter bonus code with no space', () => {
    assert.equal(context.isValidSubject('42ABCD'), true);
  });

  test('accepts a space between number and code', () => {
    assert.equal(context.isValidSubject('42 ABCD'), true);
  });

  test('accepts multiple spaces', () => {
    assert.equal(context.isValidSubject('42     ABCD'), true);
  });

  test('is case-insensitive for the bonus code', () => {
    assert.equal(context.isValidSubject('101 abcd'), true);
  });

  test('tolerates surrounding whitespace on the whole subject', () => {
    assert.equal(context.isValidSubject('  42 ABCD  '), true);
  });

  test('rejects a bonus code with a digit in the wrong position (not 3 letters + 1 trailing digit)', () => {
    assert.equal(context.isValidSubject('42 AB3D'), false);
  });

  test('rejects a bonus code shorter than 4 letters', () => {
    assert.equal(context.isValidSubject('42 ABC'), false);
  });

  test('rejects a bonus code longer than 4 letters', () => {
    assert.equal(context.isValidSubject('42 ABCDE'), false);
  });

  test('accepts 3 letters followed by exactly 1 digit', () => {
    assert.equal(context.isValidSubject('42 ABC1'), true);
  });

  test('is case-insensitive for the 3-letters-plus-digit form', () => {
    assert.equal(context.isValidSubject('42 abc1'), true);
  });

  test('rejects 4 letters followed by a digit', () => {
    assert.equal(context.isValidSubject('42 ABCD1'), false);
  });

  test('rejects 3 letters followed by 2 digits', () => {
    assert.equal(context.isValidSubject('42 ABC12'), false);
  });

  test('rejects 2 letters followed by 1 digit (too few letters)', () => {
    assert.equal(context.isValidSubject('42 AB1'), false);
  });

  test('rejects a digit-only bonus code', () => {
    assert.equal(context.isValidSubject('42 1234'), false);
  });

  test('rejects a missing rider number', () => {
    assert.equal(context.isValidSubject('ABCD'), false);
  });

  test('rejects trailing garbage after the bonus code', () => {
    assert.equal(context.isValidSubject('42 ABCD extra'), false);
  });

  test('rejects an empty subject', () => {
    assert.equal(context.isValidSubject(''), false);
  });
});

describe('extractEmailData', () => {
  const { context } = loadApp();

  test('splits rider number and upper-cases the bonus code', () => {
    const msg = {
      getSubject: () => '42 abcd',
      getFrom: () => 'Jane Smith <jane@example.com>',
      getDate: () => new Date(2026, 0, 5),
    };
    const data = context.extractEmailData(msg);
    assert.equal(data['rider-number'], '42');
    assert.equal(data['bonus'], 'ABCD');
    assert.equal(data.subject, '42 abcd');
    assert.equal(data.sender, 'Jane Smith <jane@example.com>');
    assert.deepEqual(data.date, new Date(2026, 0, 5));
  });

  test('splits rider number and upper-cases a 3-letter + 1-digit bonus code', () => {
    const msg = {
      getSubject: () => '7 abc1',
      getFrom: () => 'Bob Jones <bob@example.com>',
      getDate: () => new Date(2026, 0, 5),
    };
    const data = context.extractEmailData(msg);
    assert.equal(data['rider-number'], '7');
    assert.equal(data['bonus'], 'ABC1');
  });

  test('returns nulls for rider/bonus when the subject is invalid', () => {
    const msg = { getSubject: () => 'not a submission', getFrom: () => 'x@example.com', getDate: () => new Date() };
    const data = context.extractEmailData(msg);
    assert.equal(data['rider-number'], null);
    assert.equal(data['bonus'], null);
  });
});

describe('configInt_', () => {
  const { context } = loadApp();

  test('parses a present numeric value', () => {
    assert.equal(context.configInt_({ header_row: '3' }, 'header_row', 1), 3);
  });

  test('falls back to the default when the key is missing', () => {
    assert.equal(context.configInt_({}, 'header_row', 1), 1);
  });

  test('falls back to the default when the value is blank', () => {
    assert.equal(context.configInt_({ header_row: '' }, 'header_row', 1), 1);
  });

  test('throws a clear error when missing and no default is given', () => {
    assert.throws(() => context.configInt_({}, 'header_row'), /Config key "header_row" is missing/);
  });

  test('throws a clear error when the value is present but not numeric', () => {
    assert.throws(
      () => context.configInt_({ header_row: 'abc' }, 'header_row', 1),
      /Config key "header_row" is not a number: "abc"/
    );
  });

  test('minValue: accepts a value at or above the minimum', () => {
    assert.equal(context.configInt_({ col_approved: '1' }, 'col_approved', 4, 1), 1);
    assert.equal(context.configInt_({ col_approved: '10' }, 'col_approved', 4, 1), 10);
  });

  test('minValue: throws a clear error when the configured value is below it', () => {
    assert.throws(
      () => context.configInt_({ col_approved: '0' }, 'col_approved', 4, 1),
      /Config key "col_approved" must be at least 1, got: "0"/
    );
  });

  test('minValue: also rejects a negative configured value', () => {
    assert.throws(
      () => context.configInt_({ col_approved: '-3' }, 'col_approved', 4, 1),
      /Config key "col_approved" must be at least 1, got: "-3"/
    );
  });

  test('minValue: does not affect the fallback default when the key is blank/missing', () => {
    // The default itself is trusted (it's always a valid literal in the source),
    // so minValue only ever needs to police an explicit Config sheet value.
    assert.equal(context.configInt_({}, 'col_approved', 4, 1), 4);
  });

  test('minValue: omitted entirely means no lower bound is enforced', () => {
    assert.equal(context.configInt_({ trigger_interval_min: '0' }, 'trigger_interval_min', 10), 0);
  });
});
