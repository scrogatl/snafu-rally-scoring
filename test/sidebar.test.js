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

function buildRiderSheet(ss, riderNumber, bonusIds) {
  const rows = [['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time']];
  bonusIds.forEach((id) => rows.push([id, '', '', '', '', '', '']));
  return ss.addSheet(riderNumber, rows);
}

function setupFullEnv() {
  const env = loadApp();
  const ss = new MockSpreadsheet('ss1');
  ss.addSheet('Config', configRows(CONFIG));
  Object.values(CONFIG)
    .filter((v) => typeof v === 'string' && v.startsWith('rally/'))
    .forEach((name) => env.gmail.api.createLabel(name));
  registerActiveSpreadsheet(env, ss);
  return { ...env, ss };
}

describe('getStatus_', () => {
  const { context } = loadApp();
  const cases = [
    [['rally/scored', 'rally/approved'], 'Scored'], // scored wins even alongside approved
    [['rally/approved'], 'Approved - pending score'],
    [['rally/denied'], 'Denied'],
    [['rally/email-requires-review'], 'Email Requires Review'],
    [['rally/processing-error'], 'Bonus Code Error'],
    [['rally/email-error'], 'Email error'],
    [['rally/subject-line-error'], 'Subject Line Error'],
    [['rally/unprocessed'], 'Unprocessed'],
    [[], 'Unknown'],
  ];
  for (const [labels, expected] of cases) {
    test(`${JSON.stringify(labels)} -> ${expected}`, () => {
      assert.equal(context.getStatus_(labels, CONFIG), expected);
    });
  }
});

describe('handleApprove', () => {
  test('writes Approved + timestamp, sets labels, records the actioned message', () => {
    const { context, gmail, ss, propertiesService } = setupFullEnv();
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/email-requires-review']);
    const msgId = thread.getMessages()[0].getId();

    const result = context.handleApprove({ parameters: { msgId, threadId: thread.getId() } });

    assert.equal(sheet.getRange(2, 4).getValue(), 'X');
    assert.ok(thread.hasLabel(gmail.labelNamed('rally/approved')));
    assert.ok(thread.hasLabel(gmail.labelNamed('rally/scored')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/email-requires-review')));
    assert.equal(propertiesService.store.get('approved_msg_' + thread.getId()), msgId);
    assert.equal(result.notification._text, 'Bonus Accepted');
  });

  test('clears a previously-recorded denial when a message is approved', () => {
    const { context, gmail, ss, propertiesService } = setupFullEnv();
    buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/denied']);
    const msgId = thread.getMessages()[0].getId();
    propertiesService.store.set('denied_msg_' + thread.getId(), msgId);

    context.handleApprove({ parameters: { msgId, threadId: thread.getId() } });

    assert.equal(propertiesService.store.get('denied_msg_' + thread.getId()), undefined);
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/denied')));
  });

  test('on sheet-write failure, cleans up opposing labels instead of leaving approved+denied+scored together', () => {
    const { context, gmail, ss } = setupFullEnv();
    buildRiderSheet(ss, '42', ['ABCD']); // no ZZZZ - the write below will throw
    const thread = gmail.createThread('t1', [{ subject: '42 ZZZZ', from: 'jane@example.com' }], [
      'rally/denied', 'rally/scored', 'rally/email-requires-review',
    ]);
    const msgId = thread.getMessages()[0].getId();

    const result = context.handleApprove({ parameters: { msgId, threadId: thread.getId() } });

    assert.ok(thread.hasLabel(gmail.labelNamed('rally/approved')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/denied')), 'denied label should be removed, not left alongside approved');
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/scored')), 'scored label should be removed since nothing was actually recorded');
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/email-requires-review')));
    assert.match(result.notification._text, /sheet update failed/);
  });

  test('regression: a misconfigured column (e.g. "0") fails with a clear config error, not a raw range exception', () => {
    // Reproduces a real-world report: col_approved was accidentally set to "0"
    // in the Config sheet, and handleApprove surfaced Apps Script's opaque
    // "The starting column of the range is too small" instead of naming the
    // actual cause. configInt_'s minValue check (code.js) should now catch
    // this immediately, with the bad key and value named in the message.
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', configRows({ ...CONFIG, col_approved: '0' }));
    Object.values(CONFIG)
      .filter((v) => typeof v === 'string' && v.startsWith('rally/'))
      .forEach((name) => env.gmail.api.createLabel(name));
    registerActiveSpreadsheet(env, ss);
    buildRiderSheet(ss, '42', ['ABCD']);
    const thread = env.gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/email-requires-review']);
    const msgId = thread.getMessages()[0].getId();

    const result = env.context.handleApprove({ parameters: { msgId, threadId: thread.getId() } });

    assert.match(result.notification._text, /col_approved.*must be at least 1/);
    assert.ok(!/starting column/.test(result.notification._text), 'should not leak the raw Apps Script range error');
  });
});

describe('handleDeny', () => {
  test('writes Denied + timestamp, sets labels, records the actioned message', () => {
    const { context, gmail, ss, propertiesService } = setupFullEnv();
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/email-requires-review']);
    const msgId = thread.getMessages()[0].getId();

    const result = context.handleDeny({ parameters: { msgId, threadId: thread.getId() } });

    assert.equal(sheet.getRange(2, 6).getValue(), 'X');
    assert.ok(thread.hasLabel(gmail.labelNamed('rally/denied')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/email-requires-review')));
    assert.equal(propertiesService.store.get('denied_msg_' + thread.getId()), msgId);
    assert.equal(result.notification._text, 'Bonus Denied');
  });

  test('clears a previously-recorded approval when a message is denied', () => {
    const { context, gmail, ss, propertiesService } = setupFullEnv();
    buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/approved', 'rally/scored']);
    const msgId = thread.getMessages()[0].getId();
    propertiesService.store.set('approved_msg_' + thread.getId(), msgId);

    context.handleDeny({ parameters: { msgId, threadId: thread.getId() } });

    assert.equal(propertiesService.store.get('approved_msg_' + thread.getId()), undefined);
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/approved')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/scored')));
  });

  test('on sheet-write failure, cleans up opposing labels', () => {
    const { context, gmail, ss } = setupFullEnv();
    buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ZZZZ', from: 'jane@example.com' }], [
      'rally/approved', 'rally/scored', 'rally/email-requires-review',
    ]);
    const msgId = thread.getMessages()[0].getId();

    const result = context.handleDeny({ parameters: { msgId, threadId: thread.getId() } });

    assert.ok(thread.hasLabel(gmail.labelNamed('rally/denied')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/approved')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/scored')));
    assert.match(result.notification._text, /sheet update failed/);
  });
});

describe('handleRevertApproved / handleRevertDenied - regression for the first-message-vs-actioned-message bug', () => {
  test('reverting targets the message that was actually approved, not just the first message in the thread', () => {
    const { context, gmail, ss } = setupFullEnv();
    const sheet = buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    // Two attempts in one thread with DIFFERENT bonus codes.
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
      { subject: '42 WXYZ', from: 'jane@example.com' },
    ], ['rally/email-requires-review']);
    const [firstMsg, secondMsg] = thread.getMessages();

    // The scorer approves the SECOND message (WXYZ), not the first.
    context.handleApprove({ parameters: { msgId: secondMsg.getId(), threadId: thread.getId() } });
    assert.equal(sheet.getRange(3, 4).getValue(), 'X'); // WXYZ row

    // The Revert button always sends messages[0].getId() as firstMsgId - exactly
    // what buildScoringCard_ does - regardless of which message was approved.
    context.handleRevertApproved({ parameters: { firstMsgId: firstMsg.getId() } });

    // The actually-approved row (WXYZ) must be cleared...
    assert.equal(sheet.getRange(3, 4).getValue(), '', 'WXYZ (the message that was really approved) should be reverted');
    // ...and the never-approved row (ABCD) must remain untouched.
    assert.equal(sheet.getRange(2, 4).getValue(), '', 'ABCD was never approved and should stay empty');
    assert.ok(thread.hasLabel(gmail.labelNamed('rally/email-requires-review')));
    assert.ok(!thread.hasLabel(gmail.labelNamed('rally/approved')));
  });

  test('reverting targets the message that was actually denied, not just the first message in the thread', () => {
    const { context, gmail, ss } = setupFullEnv();
    const sheet = buildRiderSheet(ss, '42', ['ABCD', 'WXYZ']);
    const thread = gmail.createThread('t1', [
      { subject: '42 ABCD', from: 'jane@example.com' },
      { subject: '42 WXYZ', from: 'jane@example.com' },
    ], ['rally/email-requires-review']);
    const [firstMsg, secondMsg] = thread.getMessages();

    context.handleDeny({ parameters: { msgId: secondMsg.getId(), threadId: thread.getId() } });
    assert.equal(sheet.getRange(3, 6).getValue(), 'X'); // WXYZ row, Denied column

    context.handleRevertDenied({ parameters: { firstMsgId: firstMsg.getId() } });

    assert.equal(sheet.getRange(3, 6).getValue(), '', 'WXYZ (the message that was really denied) should be reverted');
    assert.equal(sheet.getRange(2, 6).getValue(), '', 'ABCD was never denied and should stay empty');
  });

  test('falls back to first-valid-message data when no actioned message was recorded (e.g. manually labeled thread)', () => {
    const { context, gmail, ss } = setupFullEnv();
    const sheet = buildRiderSheet(ss, '42', ['ABCD']);
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/approved']);
    sheet.getRange(2, 4).setValue('X'); // pretend it was approved by some other path

    const result = context.handleRevertApproved({ parameters: { firstMsgId: thread.getMessages()[0].getId() } });

    assert.equal(sheet.getRange(2, 4).getValue(), '');
    assert.equal(result.notification._text, 'Reverted to email-requires-review.');
  });
});

describe('loadThreadContext_', () => {
  test('resolves message/thread/data/config/ss/labels for a valid message id', () => {
    const { context, gmail, ss } = setupFullEnv();
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }]);
    const msgId = thread.getMessages()[0].getId();

    const ctx = context.loadThreadContext_(msgId);

    assert.equal(ctx.thread, thread);
    assert.equal(ctx.data['rider-number'], '42');
    assert.equal(ctx.data.bonus, 'ABCD');
    assert.equal(ctx.ss, ss);
    assert.ok(ctx.labels);
    assert.ok(ctx.config);
  });

  test('still resolves config/ss/labels even when the message id is bogus', () => {
    const { context } = setupFullEnv();
    const ctx = context.loadThreadContext_('does-not-exist');
    assert.equal(ctx.message, undefined);
    assert.equal(ctx.thread, undefined);
    assert.ok(ctx.config);
    assert.ok(ctx.labels);
  });
});

describe('buildAddOn / buildScoringCard_', () => {
  test('shows an info card when no message in the thread is a valid submission', () => {
    const { context, gmail } = setupFullEnv();
    const thread = gmail.createThread('t1', [{ subject: 'hello there', from: 'jane@example.com' }]);
    const msgId = thread.getMessages()[0].getId();

    const card = context.buildAddOn({ gmail: { messageId: msgId } });

    assert.equal(card.header._subtitle, 'Not a rally submission');
  });

  test('shows an Approve button for a pending message and a Revert section once scored', () => {
    const { context, gmail } = setupFullEnv();
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/email-requires-review']);
    const msgId = thread.getMessages()[0].getId();

    const pending = context.buildAddOn({ gmail: { messageId: msgId } });
    const messageSection = pending.sections[1]; // [0] is thread status
    const buttonTexts = messageSection._widgets.filter((w) => w._text !== undefined && w._action).map((w) => w._text);
    assert.ok(buttonTexts.includes('Approve this message'));
    assert.equal(pending.sections.some((s) => s._header === 'Revert'), false);

    thread.addLabel(gmail.labelNamed('rally/approved'));
    thread.addLabel(gmail.labelNamed('rally/scored'));
    thread.removeLabel(gmail.labelNamed('rally/email-requires-review'));

    const scored = context.buildAddOn({ gmail: { messageId: msgId } });
    assert.ok(scored.sections.some((s) => s._header === 'Revert'));
  });

  test('shows the Error section text for a processing-error thread', () => {
    const { context, gmail } = setupFullEnv();
    const thread = gmail.createThread('t1', [{ subject: '42 ABCD', from: 'jane@example.com' }], ['rally/processing-error']);
    const msgId = thread.getMessages()[0].getId();

    const card = context.buildAddOn({ gmail: { messageId: msgId } });

    const errorSection = card.sections.find((s) => s._header === 'Error');
    assert.ok(errorSection, 'expected an Error section');
    assert.match(errorSection._widgets[0]._text, /bonus code could not be matched/);
  });
});
