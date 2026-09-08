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

const LABEL_CONFIG = {
  label_parent: 'rally',
  label_unprocessed: 'rally/unprocessed',
  label_format_error: 'rally/subject-line-error',
  label_email_error: 'rally/email-error',
  label_processing_error: 'rally/processing-error',
  label_needs_review: 'rally/email-requires-review',
  label_approved: 'rally/approved',
  label_denied: 'rally/denied',
  label_scored: 'rally/scored',
  trigger_interval_min: '10',
};

function configRows(config) {
  return [['key', 'value', 'notes'], ...Object.entries(config)];
}

function buildBaseSpreadsheet(env, extraConfig) {
  const ss = new MockSpreadsheet('ss1');
  ss.addSheet('Config', configRows({ ...LABEL_CONFIG, ...extraConfig }));
  ss.addSheet('Rider Master', [['Rider Number', 'Name', 'Email'], ['1', 'Jane', 'jane@example.com']]);
  ss.addSheet('Bonus Master', [['Bonus ID'], ['ABCD']]);
  env.spreadsheetApp.setActive(ss);
  return ss;
}

describe('setup()', () => {
  test('a missing Bonus Master does not prevent labels or the trigger from being set up', () => {
    // Bonus Master is the foundational prerequisite: createMasterScoring_,
    // createRiderSheet_, and (transitively, via Master Scoring) Leader Board
    // all hard-require it. Labels and the trigger have no such dependency
    // and must still complete.
    const env = loadApp();
    const ss = new MockSpreadsheet('ss1');
    ss.addSheet('Config', configRows(LABEL_CONFIG));
    ss.addSheet('Rider Master', [['Rider Number', 'Name', 'Email'], ['1', 'Jane', 'jane@example.com']]);
    // Deliberately no Bonus Master sheet.
    env.spreadsheetApp.setActive(ss);

    env.context.setup();

    for (const name of Object.values(LABEL_CONFIG).filter((v) => typeof v === 'string' && v.startsWith('rally'))) {
      assert.ok(env.gmail.labelNamed(name), 'expected label ' + name + ' to exist');
    }
    assert.equal(env.scriptApp.triggers.length, 1);
    assert.equal(env.scriptApp.triggers[0].getHandlerFunction(), 'processEmails');
    assert.equal(ss.getSheetByName('Master Scoring'), null);
    assert.equal(ss.getSheetByName('Leader Board'), null);
    assert.equal(ss.getSheetByName('1'), null, 'rider sheet creation also requires Bonus Master');
    assert.ok(env.logger.logs.some((l) => l.includes('Master Scoring') && l.includes('Bonus Master')),
      'expected the Master Scoring error to be logged');
  });

  test('with Rider Master and Bonus Master present, setup() creates Master Scoring, Leader Board, and rider sheets', () => {
    const env = loadApp();
    const ss = buildBaseSpreadsheet(env, {});

    env.context.setup();

    const scoring = ss.getSheetByName('Master Scoring');
    assert.ok(scoring, 'Master Scoring should have been created');
    assert.equal(scoring.getRange(2, 3).getValue(), '1');
    assert.equal(scoring._rawCell(5, 1), "='Bonus Master'!A2");

    const lb = ss.getSheetByName('Leader Board');
    assert.ok(lb, 'Leader Board should have been created');
    assert.equal(lb._rawCell(2, 1), "='Rider Master'!A2");

    assert.ok(ss.getSheetByName('1'), 'rider sheet should have been created');
  });

  test('Leader Board ends up leftmost and Master Scoring 2nd-from-left, after rider sheets are created', () => {
    const env = loadApp();
    const ss = buildBaseSpreadsheet(env, {});

    env.context.setup();

    const names = ss.getSheets().map((s) => s.getName());
    assert.equal(names[0], 'Leader Board');
    assert.equal(names[1], 'Master Scoring');
    // Rider sheet '1' exists (created before Master Scoring/Leader Board, per
    // the log order below) but is positioned after the pre-existing master
    // sheets, not repositioned by the Leader Board/Master Scoring pinning.
    assert.ok(names.includes('1'));

    // Master Scoring's own creation happens after rider sheets (so its formulas
    // reference a rider sheet that already exists at write time) - confirmed by
    // log ordering, since final tab position alone wouldn't distinguish this
    // from Master Scoring having been created first and repositioned later.
    const riderSheetLog = env.logger.logs.findIndex((l) => l.includes('Created rider sheet: 1'));
    const masterScoringLog = env.logger.logs.findIndex((l) => l.startsWith('Master Scoring: '));
    assert.ok(riderSheetLog >= 0 && masterScoringLog >= 0 && riderSheetLog < masterScoringLog,
      'rider sheets must be created before Master Scoring');
  });

  test('re-running setup() does not reposition sheets a human has since moved by hand', () => {
    const env = loadApp();
    const ss = buildBaseSpreadsheet(env, {});
    env.context.setup();

    // Simulate a human dragging Master Scoring's tab away from position 2.
    const scoring = ss.getSheetByName('Master Scoring');
    ss.setActiveSheet(scoring);
    ss.moveActiveSheet(ss.getSheets().length); // move to the end

    env.context.setup();

    const names = ss.getSheets().map((s) => s.getName());
    assert.equal(names[names.length - 1], 'Master Scoring',
      'a manually-moved tab must not be silently repositioned by a later setup() run');
  });

  test('re-running setup() does not create a duplicate trigger or duplicate rows/columns for the same rider', () => {
    const env = loadApp();
    const ss = buildBaseSpreadsheet(env, {});

    env.context.setup();
    env.context.setup();

    assert.equal(env.scriptApp.triggers.length, 1, 'setup() must not accumulate duplicate triggers');
    const lb = ss.getSheetByName('Leader Board');
    assert.equal(lb.getLastRow(), 2, 'only one data row for the one rider, even after two setup() runs');
    const scoring = ss.getSheetByName('Master Scoring');
    assert.equal(scoring.getLastColumn(), 3, 'only one rider column, even after two setup() runs');
  });
});
