'use strict';

const { test, describe } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, MockSpreadsheet } = require('./support/mocks');

function buildRiderSheet(ss, riderNumber, bonusIds) {
  const rows = [['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time']];
  bonusIds.forEach((id) => rows.push([id, '', '', '', '', '', '']));
  return ss.addSheet(riderNumber, rows);
}

describe('updateSpreadsheet', () => {
  const CONFIG = { header_row: '1', col_bonus_id: '1' };

  test('writes X and the given timestamp to the matching bonus row', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    const when = new Date(2026, 5, 1, 10, 30);

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'wxyz', date: when }, 2, 3, true);

    assert.equal(sheet.getRange(3, 2).getValue(), 'X');
    assert.deepEqual(sheet.getRange(3, 3).getValue(), when);
    // untouched row/columns stay empty
    assert.equal(sheet.getRange(2, 2).getValue(), '');
  });

  test('uses the current time (not the email time) when useEmailTime is false', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);
    const before = new Date();

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD', date: new Date(2000, 0, 1) }, 4, 5, false);

    const written = sheet.getRange(2, 5).getValue();
    assert.ok(written instanceof Date);
    assert.ok(written.getTime() >= before.getTime());
  });

  test('matches bonus IDs case-insensitively and ignoring whitespace', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', [' abcd ']);

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 2, 3, true);

    assert.equal(sheet.getRange(2, 2).getValue(), 'X');
  });

  test('reverting (value=null) clears both the value and timestamp cells', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);
    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 4, 5, false);
    assert.equal(sheet.getRange(2, 4).getValue(), 'X');

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD' }, 4, 5, false, null);

    assert.equal(sheet.getRange(2, 4).getValue(), '');
    assert.equal(sheet.getRange(2, 5).getValue(), '');
  });

  test('throws when the bonus ID is not found in the rider sheet', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderSheet(ss, '42', ['ABCD']);

    assert.throws(
      () => context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ZZZZ', date: new Date() }, 2, 3, true),
      /Bonus ID "ZZZZ" not found in sheet 42/
    );
  });

  test('throws when the rider sheet has no data rows below the header', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderSheet(ss, '42', []); // header only

    assert.throws(
      () => context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 2, 3, true),
      /No data rows in sheet 42/
    );
  });

  test('auto-creates the rider sheet (via Bonus Master formulas) when it does not exist yet', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD'], ['WXYZ']]);

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '99', bonus: 'WXYZ', date: new Date() }, 2, 3, true);

    const sheet = ss.getSheetByName('99');
    assert.ok(sheet, 'rider sheet should have been created');
    assert.equal(sheet.getRange(3, 2).getValue(), 'X'); // WXYZ is row 3 (row 2 is ABCD)
  });

  test('respects configured header_row and col_bonus_id instead of the defaults', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    // Header on row 2, bonus IDs in column B instead of A
    const sheet = ss.addSheet('42', [
      ['', ''],
      ['', 'Bonus ID'],
      ['', 'ABCD'],
    ]);
    const config = { header_row: '2', col_bonus_id: '2' };

    context.updateSpreadsheet(ss, config, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 3, 4, true);

    assert.equal(sheet.getRange(3, 3).getValue(), 'X');
  });

  test('falls back to documented defaults when header_row/col_bonus_id are not configured', () => {
    const { context } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);

    // No header_row/col_bonus_id keys at all in config
    context.updateSpreadsheet(ss, {}, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 2, 3, true);

    assert.equal(sheet.getRange(2, 2).getValue(), 'X');
  });

  test('acquires and releases the script lock around the write', () => {
    const { context, lockService } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderSheet(ss, '42', ['ABCD']);

    context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ABCD', date: new Date() }, 2, 3, true);

    assert.equal(lockService.waitCount, 1);
    assert.equal(lockService.releaseCount, 1);
    assert.equal(lockService.lock.hasLock(), false);
  });

  test('still releases the lock when the write fails (bonus not found)', () => {
    const { context, lockService } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderSheet(ss, '42', ['ABCD']);

    assert.throws(() =>
      context.updateSpreadsheet(ss, CONFIG, { 'rider-number': '42', bonus: 'ZZZZ', date: new Date() }, 2, 3, true));

    assert.equal(lockService.releaseCount, 1);
    assert.equal(lockService.lock.hasLock(), false);
  });
});
