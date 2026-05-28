// Gmail Label (Tag) Names
const LABEL_UNPROCESSED = 'rally/unprocessed';
const LABEL_FORMAT_ERROR = 'rally/format-error';
const LABEL_EMAIL_ERROR = 'rally/email-error';
const LABEL_PROCESSING_ERROR = 'rally/processing-error';
const LABEL_NEEDS_REVIEW = 'rally/needs-review';
const LABEL_APPROVED = 'rally/approved';
const LABEL_SCORED = 'rally/scored';

// Define the structure for the Rider Sheets
const BONUS_COLUMN_INDEX = 1;         // Column A (1-based index)
const SUBMITTED_COLUMN_INDEX = 2;
const SUBMITTED_TIME_COLUMN_INDEX = 3;
const APPROVED_COLUMN_INDEX  = 4;
const APPROVED_TIME_COLUMN_INDEX  = 5;

/**
 * Main function to be set up as a time-driven trigger.
 */
function processEmails() {
  Logger.log('Starting email processing script...');

  // 1. Fetch and bundle all Gmail Labels into a single object
  const labelNames = {
    unprocessed:     LABEL_UNPROCESSED,
    formatError:     LABEL_FORMAT_ERROR,
    emailError:      LABEL_EMAIL_ERROR,
    processingError: LABEL_PROCESSING_ERROR,
    needsReview:     LABEL_NEEDS_REVIEW,
    approved:        LABEL_APPROVED,
    scored:          LABEL_SCORED
  };

  const labels = {};
  const missingLabels = [];

  for (const [key, name] of Object.entries(labelNames)) {
    const label = GmailApp.getUserLabelByName(name);
    if (!label) {
      missingLabels.push(name);
    } else {
      labels[key] = label;
    }
  }

  if (missingLabels.length > 0) {
    Logger.log(`ERROR: The following required Gmail labels are missing: ${missingLabels.join(', ')}. Script stopped.`);
    return;
  }

  // Open spreadsheet once per execution to prevent throttling
  // const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Validate formatting and add X in Submitted column
  let threads = GmailApp.search(`label:${LABEL_UNPROCESSED}`);
  Logger.log(`Found ${threads.length} unprocessed email(s).`);

  for (let thread of threads) {
    const message = thread.getMessages()[0];
    handleUnprocessedEmail(ss, message, labels);
  }

  // Look for Approved, update spreadsheet and remove Needs Review label
  threads = GmailApp.search(`label:${LABEL_APPROVED} !label:${LABEL_SCORED}`);
  Logger.log(`Found ${threads.length} approved and unscored email(s).`);

  for (let thread of threads) {
    const message = thread.getMessages()[0];
    addApprovedCheck(ss, message, labels);
  }

  Logger.log('Email processing script finished.');
}

/**
 * Adds X to the approved column and applies the scored label.
 */
function addApprovedCheck(ss, message, labels) {
  const thread = message.getThread();
  const subject = message.getSubject();
  const data = extractEmailData(message);

  Logger.log(`Processing approved email: ${subject}`);

  try {
    updateSpreadsheet(ss, data, APPROVED_COLUMN_INDEX, false);
    thread.removeLabel(labels.needsReview);
    thread.addLabel(labels.scored);
    thread.refresh();
  } catch (e) {
    Logger.log(`-> Error approving ${subject}: ${e.message}`);
    thread.addLabel(labels.processingError);
    thread.refresh();
  }
}

/**
 * Handles the processing for a single unprocessed email message.
 */
function handleUnprocessedEmail(ss, message, labels) {
  const thread = message.getThread();
  const subject = message.getSubject();

  Logger.log(`Processing email: ${subject}`);

  // 1. Format Check
  if (!isValidSubject(subject)) {
    Logger.log(`-> Format error: Subject does not match "Number Space String" pattern.`);
    thread.addLabel(labels.formatError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
    return;
  }

  const data = extractEmailData(message);

  // 2. Validate Email Address and Rider Number
  if (!validateEmailAddress(ss, data.sender, data['rider-number'])) {
    Logger.log(`-> Email validation failed for Rider ${data['rider-number']}. Tagging as email error.`);
    thread.addLabel(labels.emailError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
    return;
  }

  try {
    // 3. Success Path
    updateSpreadsheet(ss, data, SUBMITTED_COLUMN_INDEX);

    Logger.log('-> Successfully processed. Tagging as "needs-review".');
    thread.addLabel(labels.needsReview);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();

  } catch (e) {
    Logger.log(`-> Processing error for ${subject}: ${e.message}`);
    thread.addLabel(labels.processingError);
    thread.removeLabel(labels.unprocessed);
    thread.refresh();
  }
}

/**
 * Validates that the sender's email matches the registered email for the given rider number.
 */
function validateEmailAddress(ss, senderString, riderNumber) {
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;

  const senderMatch = senderString.match(emailRegex);
  if (!senderMatch) {
    Logger.log(`Validation Error: No valid email address found in sender string: "${senderString}"`);
    return false;
  }
  const senderEmail = senderMatch[0].toLowerCase();

  const masterSheet = ss.getSheetByName("Rider Master");
  if (!masterSheet) {
    Logger.log("Validation Error: 'Rider Master' sheet not found.");
    return false;
  }

  const data = masterSheet.getDataRange().getValues();
  if (data.length < 2) return false;

  const headers = data[0];
  const riderNumColIdx = headers.indexOf("Rider Number");
  const emailColIdx = headers.indexOf("Email");

  if (riderNumColIdx === -1 || emailColIdx === -1) {
    Logger.log("Validation Error: Missing 'Rider Number' or 'Email' column headers.");
    return false;
  }

  for (let i = 1; i < data.length; i++) {
    const sheetRiderNum = String(data[i][riderNumColIdx]).trim();

    if (sheetRiderNum === String(riderNumber).trim()) {
      const rawSheetEmail = String(data[i][emailColIdx]).trim();
      const sheetEmailMatch = rawSheetEmail.match(emailRegex);

      if (!sheetEmailMatch) {
        Logger.log(`Validation Error: Row ${i + 1} contains an invalid email format: "${rawSheetEmail}"`);
        return false;
      }

      const registeredEmail = sheetEmailMatch[0].toLowerCase();
      if (registeredEmail === senderEmail) {
        Logger.log(`-> Email validated for Rider ${riderNumber}.`);
        return true;
      } else {
        Logger.log(`-> Email mismatch for Rider ${riderNumber}. Expected: ${registeredEmail}, Got: ${senderEmail}`);
        return false;
      }
    }
  }

  Logger.log(`-> Validation Error: Rider ${riderNumber} not found.`);
  return false;
}

/**
 * Finds the row matching the bonus string in Column A and updates target status columns.
 */
function updateSpreadsheet(ss, data, columnIndex, useEmailTime = true) {
  const sheetName = data['rider-number'];
  const bonusToFind = data['bonus'];
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error(`Sheet / Rider Number not found: ${sheetName}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error(`Sheet ${sheetName} has no data rows.`);

  const range = sheet.getRange(2, BONUS_COLUMN_INDEX, lastRow - 1, 1);
  const values = range.getValues();
  let rowToUpdate = -1;

  for (let i = 0; i < values.length; i++) {
    const cellValue = String(values[i][0]).trim();
    if (cellValue === bonusToFind.trim()) {
      rowToUpdate = i + 2;
      break;
    }
  }

  if (rowToUpdate !== -1) {
    sheet.getRange(rowToUpdate, columnIndex).setValue("X");

    if (useEmailTime) {
      sheet.getRange(rowToUpdate, columnIndex + 1).setValue(data.date);
    } else {
      sheet.getRange(rowToUpdate, columnIndex + 1).setValue(new Date());
    }
  } else {
    throw new Error(`Row identifier (Bonus ID: ${bonusToFind}) not found in sheet ${sheetName}.`);
  }
}

/**
 * Returns true if the subject matches the expected "Number Space String" format.
 */
function isValidSubject(subject) {
  return /^\d+\s.+$/.test(subject.trim());
}

/**
 * Extracts structured data from an email message using the subject line.
 */
function extractEmailData(message) {
  const subject = message.getSubject().trim();
  const match = subject.match(/^(\d+)\s(.*)$/);

  return {
    date: message.getDate(),
    subject: subject,
    sender: message.getFrom(),
    'rider-number': match ? match[1] : null,
    'bonus': match ? match[2] : null
  };
}
