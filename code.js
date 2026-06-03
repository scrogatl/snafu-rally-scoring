/**
 * Rally Scoring — code.js
 * Config is read at runtime from the "Config" sheet.
 * Run setup() once to create the Config sheet, Gmail labels, and time trigger.
 */

// ─── Config loader ────────────────────────────────────────────────────────────
function loadConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) throw new Error('Config sheet not found. Run setup() first.');
  const rows = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]).trim();
    const val = String(rows[i][1]).trim();
    if (key) config[key] = val;
  }
  return config;
}

// ─── One-time setup ───────────────────────────────────────────────────────────
function setup() {
  Logger.log('Running setup...');
  const ss = SpreadsheetApp.openById('1AJIR9y46MZ6ArUg4krRyTIIWYk3zLJL_7Z_-YdCSI_o');

  // Write Config sheet
  createConfigSheet_(ss);

  // Save spreadsheet ID to Script Properties (used by Sidebar.gs)
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // Load config and create Gmail labels
  const config = loadConfig(ss);
  const labelKeys = [
    'label_unprocessed','label_format_error','label_email_error',
    'label_processing_error','label_needs_review','label_approved','label_scored'
  ];
  let created = 0, skipped = 0;
  for (const key of labelKeys) {
    const name = config[key];
    if (!name) { Logger.log('Config key missing: ' + key); continue; }
    if (GmailApp.getUserLabelByName(name)) {
      Logger.log('Label already exists: ' + name); skipped++;
    } else {
      GmailApp.createLabel(name);
      Logger.log('Created label: ' + name); created++;
    }
  }
  Logger.log('Labels: created=' + created + ', skipped=' + skipped);

  // Set time-driven trigger (removes existing ones first)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('Setup complete. Trigger set for every 10 minutes.');
}

function createConfigSheet_(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) { sheet = ss.insertSheet('Config'); }
  const rows = [
    ['key', 'value', 'notes'],
    ['event_name',            'SNAFU 2026 JUNE 2 TEST',       'Display only'],
    ['organizer_email',       'scott.srogers+snafu@gmail.com',         'Display only'],
    ['spreadsheet_id',        '1AJIR9y46MZ6ArUg4krRyTIIWYk3zLJL_7Z_-YdCSI_o',          'Do not change after setup'],
    ['sheet_rider_master',    'Rider Master',      'Tab name of the rider roster'],
    ['master_col_rider_number','Rider Number',  'Column header in Rider Master'],
    ['master_col_email',      'Email',      'Column header in Rider Master'],
    ['header_row',            '1',        'Row number of headers in rider sheets'],
    ['col_bonus_id',          '1',       'Column index of Bonus ID (A=1)'],
    ['col_submitted',         '2',        'Column index for submitted X'],
    ['col_submitted_time',    '3',    'Column index for submitted timestamp'],
    ['col_approved',          '4',       'Column index for approved X'],
    ['col_approved_time',     '5',   'Column index for approved timestamp'],
    ['trigger_interval_min',  '10',         'Re-run setup() to change'],
    ['label_unprocessed',     'rally/unprocessed',     ''],
    ['label_format_error',    'rally/format-error',     ''],
    ['label_email_error',     'rally/email-error',      ''],
    ['label_processing_error','rally/processing-error',       ''],
    ['label_needs_review',    'rally/needs-review',     ''],
    ['label_approved',        'rally/approved',        ''],
    ['label_scored',          'rally/scored',          ''],
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  Logger.log('Config sheet written.');
}

// ─── Main processing loop ─────────────────────────────────────────────────────
function processEmails() {
  Logger.log('Starting email processing...');
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '1AJIR9y46MZ6ArUg4krRyTIIWYk3zLJL_7Z_-YdCSI_o'
  );
  let config;
  try { config = loadConfig(ss); }
  catch (e) { Logger.log('FATAL: ' + e.message); return; }

  const labels = loadLabels_(config);
  if (!labels) return;

  let threads = GmailApp.search('label:' + config['label_unprocessed']);
  Logger.log('Unprocessed: ' + threads.length);
  for (const thread of threads)
    handleUnprocessedEmail(ss, config, thread.getMessages()[0], labels);

  threads = GmailApp.search(
    'label:' + config['label_approved'] + ' -label:' + config['label_scored']
  );
  Logger.log('Approved/unscored: ' + threads.length);
  for (const thread of threads)
    addApprovedCheck(ss, config, thread.getMessages()[0], labels);

  Logger.log('Processing complete.');
}

// ─── Email handlers ───────────────────────────────────────────────────────────
function handleUnprocessedEmail(ss, config, message, labels) {
  const thread = message.getThread();
  const subject = message.getSubject();
  Logger.log('Processing: ' + subject);

  if (!isValidSubject(subject)) {
    Logger.log('-> Format error.');
    thread.addLabel(labels.formatError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
    return;
  }

  const data = extractEmailData(message);
  if (!validateEmailAddress(ss, config, data.sender, data['rider-number'])) {
    Logger.log('-> Email validation failed for Rider ' + data['rider-number']);
    thread.addLabel(labels.emailError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
    return;
  }

  try {
    updateSpreadsheet(ss, config, data, parseInt(config['col_submitted'], 10), true);
    thread.addLabel(labels.needsReview);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
    Logger.log('-> Tagged as needs-review.');
  } catch (e) {
    Logger.log('-> Processing error: ' + e.message);
    thread.addLabel(labels.processingError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
  }
}

function addApprovedCheck(ss, config, message, labels) {
  const thread = message.getThread();
  const subject = message.getSubject();
  Logger.log('Recording approval: ' + subject);
  try {
    updateSpreadsheet(ss, config, extractEmailData(message), parseInt(config['col_approved'], 10), false);
    thread.removeLabel(labels.needsReview);
    thread.addLabel(labels.scored);
    thread.refresh();
  } catch (e) {
    Logger.log('-> Error: ' + e.message);
    thread.addLabel(labels.processingError);
    thread.refresh();
  }
}

// ─── Spreadsheet update ───────────────────────────────────────────────────────
function updateSpreadsheet(ss, config, data, columnIndex, useEmailTime) {
  const sheet = ss.getSheetByName(data['rider-number']);
  if (!sheet) throw new Error('No sheet for rider: ' + data['rider-number']);
  const lastRow = sheet.getLastRow();
  const startRow = parseInt(config['header_row'], 10) + 1;
  if (lastRow < startRow) throw new Error('No data rows in sheet ' + data['rider-number']);
  const bonusCol = parseInt(config['col_bonus_id'], 10);
  const values = sheet.getRange(startRow, bonusCol, lastRow - startRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === data['bonus'].trim()) {
      const row = startRow + i;
      sheet.getRange(row, columnIndex).setValue('X');
      sheet.getRange(row, columnIndex + 1).setValue(useEmailTime ? data.date : new Date());
      return;
    }
  }
  throw new Error('Bonus ID "' + data['bonus'] + '" not found in sheet ' + data['rider-number']);
}

// ─── Validation ───────────────────────────────────────────────────────────────
function validateEmailAddress(ss, config, senderString, riderNumber) {
  const re = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/;
  const sm = senderString.match(re);
  if (!sm) { Logger.log('No valid email in: ' + senderString); return false; }
  const senderEmail = sm[0].toLowerCase();

  const masterSheet = ss.getSheetByName(config['sheet_rider_master'] || 'Rider Master');
  if (!masterSheet) { Logger.log('Rider Master sheet not found.'); return false; }
  const rows = masterSheet.getDataRange().getValues();
  if (rows.length < 2) return false;

  const headers = rows[0];
  const rci = headers.indexOf(config['master_col_rider_number'] || 'Rider Number');
  const eci = headers.indexOf(config['master_col_email'] || 'Email');
  if (rci === -1 || eci === -1) { Logger.log('Missing column headers in Rider Master.'); return false; }

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][rci]).trim() !== String(riderNumber).trim()) continue;
    const em = String(rows[i][eci]).trim().match(re);
    if (!em) { Logger.log('Invalid email in row ' + (i+1)); return false; }
    const reg = em[0].toLowerCase();
    if (reg === senderEmail) { Logger.log('Email validated for Rider ' + riderNumber); return true; }
    Logger.log('Email mismatch for Rider ' + riderNumber + ': expected ' + reg + ', got ' + senderEmail);
    return false;
  }
  Logger.log('Rider ' + riderNumber + ' not found.');
  return false;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function loadLabels_(config) {
  const defs = {
    unprocessed:'label_unprocessed', formatError:'label_format_error',
    emailError:'label_email_error', processingError:'label_processing_error',
    needsReview:'label_needs_review', approved:'label_approved', scored:'label_scored'
  };
  const labels = {}, missing = [];
  for (const [key, ck] of Object.entries(defs)) {
    const name = config[ck];
    if (!name) { missing.push(ck); continue; }
    const label = GmailApp.getUserLabelByName(name);
    if (!label) missing.push('Gmail label: ' + name);
    else labels[key] = label;
  }
  if (missing.length) {
    Logger.log('ERROR: Missing — ' + missing.join(', ') + '. Run setup() first.');
    return null;
  }
  return labels;
}

function isValidSubject(subject) {
  return /^\d+\s.+$/.test(subject.trim());
}

function extractEmailData(message) {
  const subject = message.getSubject().trim();
  const match = subject.match(/^(\d+)\s(.*)$/);
  return {
    date: message.getDate(), subject,
    sender: message.getFrom(),
    'rider-number': match ? match[1] : null,
    'bonus': match ? match[2] : null
  };
}
