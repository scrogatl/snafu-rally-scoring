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

const CONFIG = {
  label_unprocessed: 'rally/unprocessed',
  label_format_error: 'rally/subject-line-error',
  label_email_error: 'rally/email-error',
  label_processing_error: 'rally/processing-error',
  label_needs_review: 'rally/email-requires-review',
  label_approved: 'rally/approved',
  label_denied: 'rally/denied',
  label_scored: 'rally/scored',
  header_row: '1',
  col_bonus_id: '1',
  col_submitted: '2',
  col_submitted_time: '3',
  col_approved: '4',
  col_approved_time: '5',
  col_denied: '6',
  col_denied_time: '7',
};

function configRows(config) {
  return [['key', 'value'], ...Object.entries(config)];
}

function buildLabels(gmail, context, config) {
  Object.values(CONFIG)
    .filter((v) => typeof v === 'string' && v.startsWith('rally/'))
    .forEach((name) => gmail.api.createLabel(name));
  return context.loadLabels_(config);
}

function buildRiderSheet(ss, riderNumber, bonusIds) {
  const rows = [['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time']];
  bonusIds.forEach((id) => rows.push([id, '', '', '', '', '', '']));
  return ss.addSheet(riderNumber, rows);
}

function buildRiderMaster(ss, riders) {
  ss.addSheet('Rider Master', [['Rider Number', 'Name', 'Email'], ...riders]);
}

describe('handleUnprocessedThread', () => {
  function setup() {
    const { context, gmail } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    buildRiderMaster(ss, [['42', 'Jane Smith', 'jane@example.com']]);
    const sheet = buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    const labels = buildLabels(gmail, context, CONFIG);
    return { context, gmail, ss, sheet, labels };
  }

  test('valid, registered submission -> Submitted written, moved to needs-review', () => {
    const { context, gmail, ss, sheet, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'Jane Smith <jane@example.com>' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.equal(sheet.getRange(2, 2).getValue(), 'X');
    assert.ok(thread.hasLabel(labels.needsReview));
    assert.ok(!thread.hasLabel(labels.unprocessed));
  });

  test('bad subject format on every message -> subject-line-error', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: 'not a submission', from: 'jane@example.com' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.formatError));
    assert.ok(!thread.hasLabel(labels.unprocessed));
    assert.ok(!thread.hasLabel(labels.needsReview));
  });

  test('valid format but unregistered sender -> email-error', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'someone-else@example.com' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.emailError));
    assert.ok(!thread.hasLabel(labels.unprocessed));
  });

  test('scans every message and uses the first one that is valid + registered', () => {
    const { context, gmail, ss, sheet, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: 'garbage', from: 'jane@example.com' },
      { subject: '42 WXYZ', from: 'Jane Smith <jane@example.com>' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.equal(sheet.getRange(3, 2).getValue(), 'X'); // WXYZ row
    assert.ok(thread.hasLabel(labels.needsReview));
  });

  test('email error takes priority over format error when no message is fully valid', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: 'garbage', from: 'jane@example.com' },
      { subject: '42 ABCD', from: 'someone-else@example.com' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.emailError));
    assert.ok(!thread.hasLabel(labels.formatError));
  });

  test('valid + registered but bonus ID unknown to the rider sheet -> processing-error', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ZZZZ', from: 'Jane Smith <jane@example.com>' },
    ], ['rally/unprocessed']);

    context.handleUnprocessedThread(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.processingError));
    assert.ok(!thread.hasLabel(labels.unprocessed));
  });
});

describe('addApprovedCheck', () => {
  function setup() {
    const { context, gmail } = loadApp();
    const ss = new MockSpreadsheet('ss1');
    const sheet = buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    const labels = buildLabels(gmail, context, CONFIG);
    return { context, gmail, ss, sheet, labels };
  }

  test('single valid message -> Approved written, scored label applied', () => {
    const { context, gmail, ss, sheet, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
    ], ['rally/approved']);

    context.addApprovedCheck(ss, CONFIG, thread, labels);

    assert.equal(sheet.getRange(2, 4).getValue(), 'X');
    assert.ok(thread.hasLabel(labels.scored));
    assert.ok(!thread.hasLabel(labels.needsReview));
  });

  test('no valid message in the thread -> left untouched', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: 'not a submission', from: 'jane@example.com' },
    ], ['rally/approved']);

    context.addApprovedCheck(ss, CONFIG, thread, labels);

    assert.ok(!thread.hasLabel(labels.scored));
    assert.ok(!thread.hasLabel(labels.processingError));
  });

  test('multiple valid messages referencing different bonuses -> flagged, not guessed', () => {
    const { context, gmail, ss, sheet, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
      { subject: '42 WXYZ', from: 'jane@example.com' },
    ], ['rally/approved']);

    context.addApprovedCheck(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.processingError));
    assert.ok(!thread.hasLabel(labels.scored));
    // neither bonus row was written since we couldn't tell which one was meant
    assert.equal(sheet.getRange(2, 4).getValue(), '');
    assert.equal(sheet.getRange(3, 4).getValue(), '');
  });

  test('multiple valid messages agreeing on the same rider/bonus -> approved normally', () => {
    const { context, gmail, ss, sheet, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
      { subject: '42 abcd', from: 'jane@example.com' }, // retry, same bonus
    ], ['rally/approved']);

    context.addApprovedCheck(ss, CONFIG, thread, labels);

    assert.equal(sheet.getRange(2, 4).getValue(), 'X');
    assert.ok(thread.hasLabel(labels.scored));
  });

  test('bonus ID unknown to the rider sheet -> processing-error', () => {
    const { context, gmail, ss, labels } = setup();
    const thread = gmail.createThread('t1', [
      { subject: '42 ZZZZ', from: 'jane@example.com' },
    ], ['rally/approved']);

    context.addApprovedCheck(ss, CONFIG, thread, labels);

    assert.ok(thread.hasLabel(labels.processingError));
    assert.ok(!thread.hasLabel(labels.scored));
  });

});

describe('processEmails (end-to-end orchestration)', () => {
  function setupFullEnv() {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', configRows(CONFIG));
    buildRiderMaster(ss, [['42', 'Jane Smith', 'jane@example.com']]);
    buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    buildLabels(env.gmail, env.context, CONFIG);
    registerActiveSpreadsheet(env, ss);
    return { ...env, ss };
  }

  test('moves an unprocessed submission to needs-review and writes Submitted', () => {
    const { context, gmail, ss } = setupFullEnv();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
    ], ['rally/unprocessed']);

    context.processEmails();

    assert.equal(ss.getSheetByName('42').getRange(2, 2).getValue(), 'X');
    assert.ok(thread.getLabels().some((l) => l.getName() === 'rally/email-requires-review'));
  });

  test('scores an approved-but-unscored thread', () => {
    const { context, gmail, ss } = setupFullEnv();
    const thread = gmail.createThread('t1', [
      { subject: '42 WXYZ', from: 'jane@example.com' },
    ], ['rally/approved']);

    context.processEmails();

    assert.equal(ss.getSheetByName('42').getRange(3, 4).getValue(), 'X');
    assert.ok(thread.getLabels().some((l) => l.getName() === 'rally/scored'));
  });

  test('does not re-touch a thread that is already approved and scored', () => {
    const { context, gmail, ss } = setupFullEnv();
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
    ], ['rally/approved', 'rally/scored']);
    const sheet = ss.getSheetByName('42');

    context.processEmails();

    // approved/scored search excludes -label:scored, so this thread is never visited
    assert.equal(sheet.getRange(2, 4).getValue(), '');
  });

  test('exits gracefully (no throw) when the Config sheet is missing', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1'); // no Config sheet
    registerActiveSpreadsheet(env, ss);

    assert.doesNotThrow(() => env.context.processEmails());
    assert.ok(env.logger.logs.some((l) => l.includes('FATAL')));
  });

  test('exits gracefully (no throw) when a Gmail label is missing', () => {
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', configRows(CONFIG));
    // Deliberately skip creating the Gmail labels this time.
    registerActiveSpreadsheet(env, ss);

    assert.doesNotThrow(() => env.context.processEmails());
  });
});
