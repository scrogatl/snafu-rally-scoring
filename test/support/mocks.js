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

const vm = require('vm');
const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// Minimal in-memory stand-ins for the Apps Script services code.js/Sidebar.gs/
// utility.js call into (GmailApp, SpreadsheetApp, PropertiesService,
// LockService, ScriptApp, CardService, Logger). Nothing here talks to a real
// Google API.
// ---------------------------------------------------------------------------

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.toUpperCase().charCodeAt(0) - 64);
  return n;
}

// --- Sheets ------------------------------------------------------------------

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._resolveCell(this.row + r, this.col + c));
      out.push(rowArr);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) this.sheet._setCell(this.row + r, this.col + c, values[r][c]);
    }
    return this;
  }
  setValue(value) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) this.sheet._setCell(this.row + r, this.col + c, value);
    }
    return this;
  }
  setFormulas(formulas) {
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) this.sheet._setCell(this.row + r, this.col + c, formulas[r][c]);
    }
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) this.sheet._setCell(this.row + r, this.col + c, '');
    }
    return this;
  }
  setFontWeight() { return this; }
  setNumberFormat(format) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) this.sheet._setFormat(this.row + r, this.col + c, format);
    }
    return this;
  }
  // Matches the real API: returns just the top-left cell's format.
  getNumberFormat() { return this.sheet._rawFormat(this.row, this.col); }
}

class MockSheet {
  constructor(spreadsheet, name, initialRows) {
    this.spreadsheet = spreadsheet;
    this.name = name;
    this.grid = [];
    this.formats = [];
    this.frozenRows = 0;
    if (initialRows) initialRows.forEach((row, r) => row.forEach((val, c) => this._setCell(r + 1, c + 1, val)));
  }
  getName() { return this.name; }
  _ensure(r, c) {
    while (this.grid.length < r) this.grid.push([]);
    const rowArr = this.grid[r - 1];
    while (rowArr.length < c) rowArr.push('');
  }
  _setCell(r, c, v) { this._ensure(r, c); this.grid[r - 1][c - 1] = v === undefined ? '' : v; }
  _rawCell(r, c) {
    if (r - 1 >= this.grid.length) return '';
    const row = this.grid[r - 1];
    if (c - 1 >= row.length) return '';
    const v = row[c - 1];
    return v === undefined ? '' : v;
  }
  _setFormat(r, c, format) {
    while (this.formats.length < r) this.formats.push([]);
    const rowArr = this.formats[r - 1];
    while (rowArr.length < c) rowArr.push('General');
    rowArr[c - 1] = format;
  }
  _rawFormat(r, c) {
    if (r - 1 >= this.formats.length) return 'General';
    const row = this.formats[r - 1];
    if (c - 1 >= row.length) return 'General';
    return row[c - 1];
  }
  // Resolves the one formula pattern the app actually writes: ='Sheet Name'!A2
  _resolveCell(r, c) {
    const raw = this._rawCell(r, c);
    if (typeof raw === 'string' && raw.startsWith('=')) {
      const m = raw.match(/^='?([^'!]+)'?!([A-Za-z]+)(\d+)$/);
      if (m) {
        const target = this.spreadsheet.getSheetByName(m[1]);
        if (target) return target._resolveCell(parseInt(m[3], 10), colLetterToIndex(m[2]));
      }
      return raw;
    }
    return raw;
  }
  getLastRow() {
    for (let r = this.grid.length; r >= 1; r--) {
      if (this.grid[r - 1].some((v) => v !== '' && v !== undefined)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    for (const row of this.grid) {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1] !== '' && row[c - 1] !== undefined) { max = Math.max(max, c); break; }
      }
    }
    return max;
  }
  getRange(row, col, numRows, numCols) {
    numRows = numRows || 1;
    numCols = numCols || 1;
    this._ensure(row + numRows - 1, col + numCols - 1);
    return new MockRange(this, row, col, numRows, numCols);
  }
  getDataRange() {
    const lastRow = Math.max(this.getLastRow(), 1);
    const lastCol = Math.max(this.getLastColumn(), 1);
    return this.getRange(1, 1, lastRow, lastCol);
  }
  setFrozenRows(n) { this.frozenRows = n; }
  autoResizeColumn() {}
}

class MockSpreadsheet {
  constructor(id) {
    this.id = id;
    this.sheets = new Map();
    this.order = [];
  }
  getId() { return this.id; }
  getSheetByName(name) { return this.sheets.get(name) || null; }
  getSheets() { return this.order.map((n) => this.sheets.get(n)); }
  insertSheet(name, index) {
    if (this.sheets.has(name)) throw new Error('A sheet with the name "' + name + '" already exists.');
    const sheet = new MockSheet(this, name);
    this.sheets.set(name, sheet);
    if (typeof index === 'number' && index < this.order.length) this.order.splice(index, 0, name);
    else this.order.push(name);
    return sheet;
  }
  deleteSheet(sheet) {
    this.sheets.delete(sheet.getName());
    this.order = this.order.filter((n) => n !== sheet.getName());
  }
  // Real Apps Script has no direct "move this sheet" call - repositioning an
  // already-created sheet is always this two-step setActiveSheet()/
  // moveActiveSheet(pos) pattern (pos is 1-based).
  setActiveSheet(sheet) { this._activeSheetForMove = sheet; }
  moveActiveSheet(pos) {
    const sheet = this._activeSheetForMove;
    if (!sheet) return;
    const name = sheet.getName();
    this.order = this.order.filter((n) => n !== name);
    this.order.splice(pos - 1, 0, name);
  }
  // Test fixture helper - not part of the real Sheets API.
  addSheet(name, rows) {
    const sheet = new MockSheet(this, name, rows);
    this.sheets.set(name, sheet);
    this.order.push(name);
    return sheet;
  }
}

function createSpreadsheetApp() {
  const stores = new Map();
  let active = null;
  const api = {
    getActiveSpreadsheet() {
      if (!active) throw new Error('No active spreadsheet set for this test.');
      return active;
    },
    openById(id) {
      const ss = stores.get(id);
      if (!ss) throw new Error('No spreadsheet with id: ' + id);
      return ss;
    },
  };
  return {
    api,
    setActive(ss) { active = ss; stores.set(ss.getId(), ss); },
    register(ss) { stores.set(ss.getId(), ss); },
  };
}

// --- Gmail ---------------------------------------------------------------

class MockLabel {
  constructor(name) { this._name = name; }
  getName() { return this._name; }
}

class MockMessage {
  constructor({ id, subject, from, date }) {
    this._id = id;
    this._subject = subject;
    this._from = from;
    this._date = date || new Date(2026, 0, 1, 12, 0, 0);
    this._thread = null;
  }
  getId() { return this._id; }
  getSubject() { return this._subject; }
  getFrom() { return this._from; }
  getDate() { return this._date; }
  getThread() { return this._thread; }
}

class MockThread {
  constructor(id, messages) {
    this._id = id;
    this._messages = messages;
    messages.forEach((m) => { m._thread = this; });
    this._labels = new Set();
  }
  getId() { return this._id; }
  getMessages() { return this._messages; }
  addLabel(label) { if (label) this._labels.add(label); return this; }
  removeLabel(label) { if (label) this._labels.delete(label); return this; }
  getLabels() { return Array.from(this._labels); }
  hasLabel(label) { return label ? this._labels.has(label) : false; }
  refresh() {}
}

function matchesSearchQuery(thread, query) {
  const parts = query.split(/\s+/).filter(Boolean);
  for (const part of parts) {
    if (part.startsWith('-label:')) {
      const name = part.slice('-label:'.length);
      if (thread.getLabels().some((l) => l.getName() === name)) return false;
    } else if (part.startsWith('label:')) {
      const name = part.slice('label:'.length);
      if (!thread.getLabels().some((l) => l.getName() === name)) return false;
    }
  }
  return true;
}

function createGmailApp() {
  const labels = new Map();
  const messagesById = new Map();
  const threads = [];

  const api = {
    getUserLabelByName(name) { return labels.get(name) || null; },
    createLabel(name) {
      if (labels.has(name)) return labels.get(name);
      const l = new MockLabel(name);
      labels.set(name, l);
      return l;
    },
    getMessageById(id) {
      const m = messagesById.get(id);
      if (!m) throw new Error('Message not found: ' + id);
      return m;
    },
    search(query) { return threads.filter((t) => matchesSearchQuery(t, query)); },
  };

  return {
    api,
    labels,
    // Test fixture helpers - not part of the real Gmail API.
    createThread(id, messageSpecs, initialLabelNames) {
      const messages = messageSpecs.map((spec, i) => new MockMessage({
        id: spec.id || (id + '-m' + (i + 1)),
        subject: spec.subject,
        from: spec.from,
        date: spec.date,
      }));
      const thread = new MockThread(id, messages);
      messages.forEach((m) => messagesById.set(m.getId(), m));
      (initialLabelNames || []).forEach((name) => thread.addLabel(api.createLabel(name)));
      threads.push(thread);
      return thread;
    },
    labelNamed(name) { return labels.get(name) || null; },
  };
}

// --- PropertiesService / LockService / ScriptApp / Logger -----------------

function createPropertiesService() {
  const store = new Map();
  const scriptProperties = {
    getProperty(key) { return store.has(key) ? store.get(key) : null; },
    setProperty(key, value) { store.set(key, value); return scriptProperties; },
    deleteProperty(key) { store.delete(key); return scriptProperties; },
  };
  return { api: { getScriptProperties: () => scriptProperties }, store };
}

function createLockService() {
  let locked = false;
  let waitCount = 0;
  let releaseCount = 0;
  const lock = {
    waitLock() {
      if (locked) throw new Error('Could not acquire lock');
      locked = true;
      waitCount++;
      return true;
    },
    tryLock() {
      if (locked) return false;
      locked = true;
      waitCount++;
      return true;
    },
    releaseLock() { locked = false; releaseCount++; },
    hasLock() { return locked; },
  };
  return {
    api: { getScriptLock: () => lock },
    lock,
    get waitCount() { return waitCount; },
    get releaseCount() { return releaseCount; },
  };
}

function createScriptApp() {
  const triggers = [];
  let idCounter = 0;
  class MockTrigger {
    constructor(handlerFunction, intervalMinutes) {
      this._handler = handlerFunction;
      this._interval = intervalMinutes;
      this._id = 't' + (++idCounter);
    }
    getHandlerFunction() { return this._handler; }
  }
  class TriggerBuilder {
    constructor(handlerFunction) { this._handler = handlerFunction; this._interval = null; }
    timeBased() { return this; }
    everyMinutes(n) { this._interval = n; return this; }
    create() { const t = new MockTrigger(this._handler, this._interval); triggers.push(t); return t; }
  }
  const api = {
    getProjectTriggers() { return triggers.slice(); },
    deleteTrigger(t) { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
    newTrigger(handlerFunction) { return new TriggerBuilder(handlerFunction); },
  };
  return { api, triggers };
}

function createLogger() {
  const logs = [];
  return { api: { log: (...args) => logs.push(args.map(String).join(' ')) }, logs };
}

// --- CardService -----------------------------------------------------------
// Stateless builder stubs: enough structure for tests to assert on card
// contents (buttons, text, notifications) without a real Gmail UI.

class CardBuilder {
  constructor() { this._header = null; this._sections = []; }
  setHeader(h) { this._header = h; return this; }
  addSection(s) { this._sections.push(s); return this; }
  build() { return { header: this._header, sections: this._sections }; }
}
class CardHeader {
  constructor() { this._title = null; this._subtitle = null; }
  setTitle(t) { this._title = t; return this; }
  setSubtitle(s) { this._subtitle = s; return this; }
}
class CardSection {
  constructor() { this._header = null; this._widgets = []; }
  setHeader(h) { this._header = h; return this; }
  addWidget(w) { this._widgets.push(w); return this; }
}
class KeyValue {
  constructor() { this._topLabel = null; this._content = null; }
  setTopLabel(l) { this._topLabel = l; return this; }
  setContent(c) { this._content = c; return this; }
}
class TextParagraph {
  constructor() { this._text = null; }
  setText(t) { this._text = t; return this; }
}
class TextButton {
  constructor() { this._text = null; this._action = null; this._style = null; this._bg = null; }
  setText(t) { this._text = t; return this; }
  setOnClickAction(a) { this._action = a; return this; }
  setTextButtonStyle(s) { this._style = s; return this; }
  setBackgroundColor(c) { this._bg = c; return this; }
}
class CardAction {
  constructor() { this._fn = null; this._params = null; }
  setFunctionName(f) { this._fn = f; return this; }
  setParameters(p) { this._params = p; return this; }
}
class CardNotification {
  constructor() { this._text = null; }
  setText(t) { this._text = t; return this; }
}
class CardNavigation {
  constructor() { this._ops = []; }
  pushCard(c) { this._ops.push(['push', c]); return this; }
  popCard() { this._ops.push(['pop']); return this; }
  updateCard(c) { this._ops.push(['update', c]); return this; }
}
class ActionResponseBuilder {
  constructor() { this._notification = null; this._navigation = null; this._stateChanged = false; }
  setNotification(n) { this._notification = n; return this; }
  setNavigation(n) { this._navigation = n; return this; }
  setStateChanged(v) { this._stateChanged = v; return this; }
  build() { return { notification: this._notification, navigation: this._navigation, stateChanged: this._stateChanged }; }
}

const CardService = {
  newCardBuilder: () => new CardBuilder(),
  newCardHeader: () => new CardHeader(),
  newCardSection: () => new CardSection(),
  newKeyValue: () => new KeyValue(),
  newTextParagraph: () => new TextParagraph(),
  newTextButton: () => new TextButton(),
  newAction: () => new CardAction(),
  newActionResponseBuilder: () => new ActionResponseBuilder(),
  newNotification: () => new CardNotification(),
  newNavigation: () => new CardNavigation(),
  TextButtonStyle: { FILLED: 'FILLED', TEXT: 'TEXT' },
};

// --- Wiring it all together --------------------------------------------------

const CODE_JS_PATH = path.join(__dirname, '..', '..', 'code.js');
const SIDEBAR_GS_PATH = path.join(__dirname, '..', '..', 'Sidebar.gs');
const UTILITY_JS_PATH = path.join(__dirname, '..', '..', 'utility.js');

/**
 * Builds a fresh mock GAS environment and runs the real code.js/Sidebar.gs/
 * utility.js source into it unmodified. Returns the vm context (every
 * top-level function across all three files is callable as a property on
 * it) plus handles to each mock service for building fixtures and making
 * assertions.
 */
function loadApp() {
  const gmail = createGmailApp();
  const spreadsheetApp = createSpreadsheetApp();
  const propertiesService = createPropertiesService();
  const lockService = createLockService();
  const scriptApp = createScriptApp();
  const logger = createLogger();

  const sandbox = {
    // Pass through host intrinsics so values built inside the vm (new Date(),
    // thrown Errors, array literals) are instanceof-compatible with the
    // Node realm the tests run in - vm.createContext otherwise creates a
    // separate realm with its own Date/Array/Error/Object constructors.
    Date, Array, Object, Error, RegExp, Math, JSON,
    GmailApp: gmail.api,
    SpreadsheetApp: spreadsheetApp.api,
    PropertiesService: propertiesService.api,
    LockService: lockService.api,
    ScriptApp: scriptApp.api,
    CardService,
    Logger: logger.api,
    console,
  };
  const context = vm.createContext(sandbox);

  vm.runInContext(fs.readFileSync(CODE_JS_PATH, 'utf8'), context, { filename: 'code.js' });
  vm.runInContext(fs.readFileSync(SIDEBAR_GS_PATH, 'utf8'), context, { filename: 'Sidebar.gs' });
  vm.runInContext(fs.readFileSync(UTILITY_JS_PATH, 'utf8'), context, { filename: 'utility.js' });

  return { context, gmail, spreadsheetApp, propertiesService, lockService, scriptApp, logger };
}

/** Registers ss as both the "active" spreadsheet and the one openById(id) resolves to. */
function registerActiveSpreadsheet(env, ss) {
  env.spreadsheetApp.setActive(ss);
  env.propertiesService.store.set('SPREADSHEET_ID', ss.getId());
}

/**
 * Copies a vm-realm plain object's own enumerable properties onto a
 * host-realm object, so it can be compared with assert.deepStrictEqual
 * (which also checks prototype identity - objects returned by functions
 * loaded via vm.runInContext have a different Object.prototype than
 * literals written directly in a test file, even though they're
 * structurally identical).
 */
function toPlain(obj) { return Object.assign({}, obj); }

/**
 * Same cross-realm problem as toPlain, but for arrays: Array.prototype.map()/
 * filter() called inside code loaded via vm.runInContext builds a vm-realm
 * array, which isn't instanceof the host realm's Array - so it trips
 * assert.deepStrictEqual's prototype check even when its elements match.
 * Array.from(), called from the host realm, copies elements into a plain
 * host-realm array.
 */
function toPlainArray(arr) { return Array.from(arr); }

module.exports = {
  loadApp,
  registerActiveSpreadsheet,
  MockSpreadsheet,
  colLetterToIndex,
  toPlain,
  toPlainArray,
};
