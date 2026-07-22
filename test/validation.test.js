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

  test('rejects a bonus code with digits', () => {
    assert.equal(context.isValidSubject('42 AB3D'), false);
  });

  test('rejects a bonus code shorter than 4 letters', () => {
    assert.equal(context.isValidSubject('42 ABC'), false);
  });

  test('rejects a bonus code longer than 4 letters', () => {
    assert.equal(context.isValidSubject('42 ABCDE'), false);
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
});
