/**
 * Rally Scoring - code.js
 * Config is read at runtime from the "Config" sheet.
 * Run setup() once to create the Config sheet, Gmail labels, and time trigger.
 */

// --- Config loader ------------------------------------------------------------
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

// --- One-time setup -----------------------------------------------------------
// Before running setup():
//   1. Import config.csv (from the setup wizard) into your spreadsheet as a sheet named "Config"
//   2. Make sure your Rider Master and Bonus Master sheets are populated
// Then run this function once to create Gmail labels, rider sheets, and the time trigger.

function setup() {
  Logger.log('Running setup...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Save spreadsheet ID to Script Properties for use by Sidebar.gs
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  Logger.log('Spreadsheet ID saved: ' + ss.getId());

  // Load config from the Config sheet (imported from config.csv)
  let config;
  try {
    config = loadConfig(ss);
  } catch (e) {
    Logger.log('FATAL: ' + e.message);
    Logger.log('Make sure you have imported config.csv as a sheet named "Config" before running setup().');
    return;
  }

  // Create the bare parent label so all submissions are visible in one click
  const parentLabel = config['label_parent'] || 'rally';
  if (!GmailApp.getUserLabelByName(parentLabel)) {
    GmailApp.createLabel(parentLabel);
    Logger.log('Created label: ' + parentLabel);
  } else {
    Logger.log('Label already exists: ' + parentLabel);
  }

  // Create all sub-labels
  const labelKeys = [
    'label_unprocessed', 'label_format_error', 'label_email_error',
    'label_processing_error', 'label_needs_review', 'label_approved',
    'label_denied', 'label_scored'
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

  // Create rider score sheets for all riders in Rider Master
  createAllRiderSheets_(ss, config);

  // Set time-driven trigger
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const intervalMin = parseInt(config['trigger_interval_min'], 10) || 10;
  ScriptApp.newTrigger('processEmails').timeBased().everyMinutes(intervalMin).create();

  Logger.log('Setup complete. Trigger set for every ' + intervalMin + ' minutes.');
}

/**
 * Loops through Rider Master and creates a score sheet for any
 * rider that doesn't already have one.
 */
function createAllRiderSheets_(ss, config) {
  const masterSheet = ss.getSheetByName(config['sheet_rider_master'] || 'Rider Master');
  if (!masterSheet) {
    Logger.log('createAllRiderSheets_: Rider Master sheet not found - skipping.');
    return;
  }

  const rows = masterSheet.getDataRange().getValues();
  if (rows.length < 2) {
    Logger.log('createAllRiderSheets_: No riders found in Rider Master.');
    return;
  }

  const headers = rows[0];
  const rci = headers.indexOf(config['master_col_rider_number'] || 'Rider Number');
  if (rci === -1) {
    Logger.log('createAllRiderSheets_: Rider Number column not found in Rider Master.');
    return;
  }

  let created = 0, skipped = 0;
  for (let i = 1; i < rows.length; i++) {
    const riderNumber = String(rows[i][rci]).trim();
    if (!riderNumber) continue;
    if (ss.getSheetByName(riderNumber)) {
      Logger.log('Rider sheet already exists: ' + riderNumber); skipped++;
    } else {
      try {
        createRiderSheet_(ss, config, riderNumber);
        created++;
      } catch (e) {
        Logger.log('Error creating sheet for Rider ' + riderNumber + ': ' + e.message);
      }
    }
  }
  Logger.log('Rider sheets: created=' + created + ', skipped=' + skipped);
}


// --- Main processing loop -----------------------------------------------------
function processEmails() {
  Logger.log('Starting email processing...');
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
  );
  let config;
  try { config = loadConfig(ss); }
  catch (e) { Logger.log('FATAL: ' + e.message); return; }

  const labels = loadLabels_(config);
  if (!labels) return;

  // Process new unprocessed submissions
  let threads = GmailApp.search('label:' + config['label_unprocessed']);
  Logger.log('Unprocessed: ' + threads.length);
  for (const thread of threads)
    handleUnprocessedEmail(ss, config, thread.getMessages()[0], labels);

  // Record approved submissions that haven't been scored yet
  threads = GmailApp.search(
    'label:' + config['label_approved'] + ' -label:' + config['label_scored']
  );
  Logger.log('Approved/unscored: ' + threads.length);
  for (const thread of threads)
    addApprovedCheck(ss, config, thread.getMessages()[0], labels);

  Logger.log('Processing complete.');
}

// --- Email handlers -----------------------------------------------------------
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
  Logger.log('Recording approval: ' + message.getSubject());
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

// --- Spreadsheet update -------------------------------------------------------

/**
 * Finds the rider's sheet (creating it from Bonus Master if missing),
 * locates the bonus row, and writes X + timestamp to the given column.
 */
function updateSpreadsheet(ss, config, data, columnIndex, useEmailTime) {
  const riderNumber = data['rider-number'];
  const bonusToFind = data['bonus'];

  let sheet = ss.getSheetByName(riderNumber);
  if (!sheet) {
    sheet = createRiderSheet_(ss, config, riderNumber);
  }

  const lastRow = sheet.getLastRow();
  const startRow = parseInt(config['header_row'], 10) + 1;
  if (lastRow < startRow) throw new Error('No data rows in sheet ' + riderNumber);

  const bonusCol = parseInt(config['col_bonus_id'], 10);
  const values = sheet.getRange(startRow, bonusCol, lastRow - startRow + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === bonusToFind.trim()) {
      const row = startRow + i;
      sheet.getRange(row, columnIndex).setValue('X');
      sheet.getRange(row, columnIndex + 1).setValue(useEmailTime ? data.date : new Date());
      return;
    }
  }
  throw new Error('Bonus ID "' + bonusToFind + '" not found in sheet ' + riderNumber);
}

/**
 * Creates a new rider sheet by copying bonus IDs from the Bonus Master sheet.
 * Headers are written based on the configured column layout.
 */
function createRiderSheet_(ss, config, riderNumber) {
  Logger.log('Creating rider sheet for: ' + riderNumber);

  const bonusMasterName = config['sheet_bonus_master'] || 'Bonus Master';
  const bonusMaster = ss.getSheetByName(bonusMasterName);
  if (!bonusMaster) {
    throw new Error(
      'Cannot create rider sheet - "' + bonusMasterName + '" sheet not found. ' +
      'Create a sheet named "' + bonusMasterName + '" with one bonus ID per row in column A.'
    );
  }

  const sheet = ss.insertSheet(riderNumber);

  // Write header row
  const headerRow = parseInt(config['header_row'], 10);
  const headers = ['Bonus ID', 'Submitted', 'Submit Time', 'Approved', 'Approve Time', 'Denied', 'Deny Time'];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(headerRow);

  // Copy bonus IDs from Bonus Master (skip header row)
  const bonusMasterLastRow = bonusMaster.getLastRow();
  if (bonusMasterLastRow > 1) {
    const bonusIds = bonusMaster.getRange(2, 1, bonusMasterLastRow - 1, 1).getValues();
    const startRow = headerRow + 1;
    sheet.getRange(startRow, 1, bonusIds.length, 1).setValues(bonusIds);
  }

  sheet.autoResizeColumn(1);
  Logger.log('Created rider sheet: ' + riderNumber);
  return sheet;
}

// --- Validation ---------------------------------------------------------------
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
    if (!em) { Logger.log('Invalid email in row ' + (i + 1)); return false; }
    const reg = em[0].toLowerCase();
    if (reg === senderEmail) { Logger.log('Email validated for Rider ' + riderNumber); return true; }
    Logger.log('Mismatch for Rider ' + riderNumber + ': expected ' + reg + ', got ' + senderEmail);
    return false;
  }
  Logger.log('Rider ' + riderNumber + ' not found.');
  return false;
}

// --- Helpers ------------------------------------------------------------------
function loadLabels_(config) {
  const defs = {
    unprocessed:     'label_unprocessed',
    formatError:     'label_format_error',
    emailError:      'label_email_error',
    processingError: 'label_processing_error',
    needsReview:     'label_needs_review',
    approved:        'label_approved',
    denied:          'label_denied',
    scored:          'label_scored',
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
    Logger.log('ERROR: Missing - ' + missing.join(', ') + '. Run setup() first.');
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
    date: message.getDate(),
    subject,
    sender: message.getFrom(),
    'rider-number': match ? match[1] : null,
    'bonus':        match ? match[2] : null,
  };
}
