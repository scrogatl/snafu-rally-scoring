// --- CONFIGURATION ---
// REPLACE WITH YOUR ACTUAL SPREADSHEET ID
const SPREADSHEET_ID = '1COjEQApvj1uXyCExI58jy8BWpUTb7juTZ8lBfVmXIOo'; // cite: 1

// Gmail Label (Tag) Names
const LABEL_UNPROCESSED = 'rally/unprocessed'; // cite: 2
const LABEL_FORMAT_ERROR = 'rally/format-error'; // cite: 2
const LABEL_EMAIL_ERROR = 'rally/email-error';
const LABEL_PROCESSING_ERROR = 'rally/processing-error'; // cite: 2
const LABEL_NEEDS_REVIEW = 'rally/needs-review'; // cite: 2
const LABEL_APPROVED = 'rally/approved'; // cite: 3
const LABEL_SCORED = 'rally/scored'; // cite: 3

// Define the structure for the Rider Sheets
const BONUS_COLUMN_INDEX = 1; // Column A (1-based index) // cite: 3, 4
const SUBMITTED_COLUMN_INDEX = 2; // cite: 4
const SUBMITTED_TIME_COLUMN_INDEX = 3; // cite: 5
const APPROVED_COLUMN_INDEX  = 4; // cite: 5
const APPROVED_TIME_COLUMN_INDEX  = 5; // cite: 6

/**
 * Main function to be set up as a time-driven trigger.
 */
function processEmails() { // cite: 9
  Logger.log('Starting email processing script...'); // cite: 9
  
  // 1. Fetch and bundle all Gmail Labels into a single object
  const labels = {
    unprocessed:     GmailApp.getUserLabelByName(LABEL_UNPROCESSED), // cite: 9
    formatError:     GmailApp.getUserLabelByName(LABEL_FORMAT_ERROR), // cite: 10
    emailError:      GmailApp.getUserLabelByName(LABEL_EMAIL_ERROR),
    processingError: GmailApp.getUserLabelByName(LABEL_PROCESSING_ERROR), // cite: 10
    needsReview:     GmailApp.getUserLabelByName(LABEL_NEEDS_REVIEW), // cite: 10
    approved:        GmailApp.getUserLabelByName(LABEL_APPROVED), // cite: 11
    scored:          GmailApp.getUserLabelByName(LABEL_SCORED) // cite: 11
  };

  if (!labels.unprocessed) { // cite: 12
    Logger.log(`Label not found: ${LABEL_UNPROCESSED}. Script stopped.`); // cite: 12
    return; // cite: 12
  }

  // Open spreadsheet once per execution to prevent throttling
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // cite: 24

  // Validate formatting and add X in Submitted column
  let threads = GmailApp.search(`label:${LABEL_UNPROCESSED}`); // cite: 12
  Logger.log(`Found ${threads.length} unprocessed emails(s).`); // cite: 13

  for (let thread of threads) { // cite: 13
    const message = thread.getMessages()[0]; // cite: 13
    handleUnprocessedEmail(ss, message, labels);
  }

  // Look for Approved, update spreadsheet and remove Needs Review label
  threads = GmailApp.search(`label:${LABEL_APPROVED} !label:${LABEL_SCORED}`); // cite: 14
  Logger.log(`Found ${threads.length} Approved and unscored emails(s).`); // cite: 14
  for (let thread of threads) { // cite: 15
    const message = thread.getMessages()[0]; // cite: 15
    addApprovedCheck(ss, message, labels);
  }
  
  Logger.log('Email processing script finished.'); // cite: 16
}

/**
 * Adds X to the approved column and applies the scored label
 */
function addApprovedCheck(ss, message, labels) {
  const thread = message.getThread(); // cite: 16
  const subject = message.getSubject(); // cite: 17
  const data = extractEmailData(message); // cite: 17
  
  Logger.log(`Processing email: ${subject}`); // cite: 17
  
  try {
    updateSpreadsheet(ss, data, APPROVED_COLUMN_INDEX, false); // cite: 18
    thread.removeLabel(labels.needsReview); // cite: 18
    thread.addLabel(labels.scored); // cite: 18
    thread.refresh(); // cite: 18
  } catch (e) {
    Logger.log(`-> Error approving ${subject}: ${e.message}`); // cite: 20
    thread.addLabel(labels.processingError); // cite: 21
  }
}

/**
 * Handles the processing for a single email message.
 */
function handleUnprocessedEmail(ss, message, labels) {
  const thread = message.getThread(); // cite: 16
  const subject = message.getSubject(); // cite: 17
  
  Logger.log(`Processing email: ${subject}`); // cite: 17

  // 1. Format Check 
  if (!isValidSubject(subject)) {
    Logger.log(`-> Format error: Subject does not match "Number Space String" pattern.`);
    thread.addLabel(labels.formatError); // cite: 19
    thread.removeLabel(labels.unprocessed); // cite: 19
    return; // cite: 19
  }
  
  const data = extractEmailData(message); // cite: 17
  
  // 2. Validate Email Address and Rider Number 
  if (!validateEmailAddress(ss, data.sender, data['rider-number'])) { // cite: 19
    Logger.log(`-> Email validation failed for Rider ${data['rider-number']}. Tagging as email error.`);
    thread.addLabel(labels.emailError);
    thread.removeLabel(labels.unprocessed); // cite: 19
    return; // cite: 19
  }

  try {
    // 3. Success Path
    updateSpreadsheet(ss, data, SUBMITTED_COLUMN_INDEX);

    Logger.log('-> Successfully processed. Tagging as "needs-review".');
    thread.addLabel(labels.needsReview);
    thread.removeLabel(labels.unprocessed);
    thread.refresh(); // cite: 18
    
  } catch (e) {
    Logger.log(`-> Processing error for ${subject}: ${e.message}`); // cite: 20
    thread.addLabel(labels.processingError); // cite: 21
    thread.removeLabel(labels.unprocessed); // cite: 21
  }
}

/**
 * Validates that the sender's email matches the registered email for the given rider number.
 */
function validateEmailAddress(ss, senderString, riderNumber) {
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/; // cite: 74
  
  const senderMatch = senderString.match(emailRegex); // cite: 75
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
  const riderNumColIdx = headers.indexOf("Rider Number"); // cite: 63
  const emailColIdx = headers.indexOf("Email"); // cite: 63

  if (riderNumColIdx === -1 || emailColIdx === -1) {
    Logger.log("Validation Error: Missing 'Rider Number' or 'Email' column headers.");
    return false;
  }

  for (let i = 1; i < data.length; i++) {
    const sheetRiderNum = String(data[i][riderNumColIdx]).trim();
    
    if (sheetRiderNum === String(riderNumber).trim()) {
      const rawSheetEmail = String(data[i][emailColIdx]).trim();
      const sheetEmailMatch = rawSheetEmail.match(emailRegex); // cite: 81
      
      if (!sheetEmailMatch) {
        Logger.log(`Validation Error: Row ${i + 1} contains an invalid email format: "${rawSheetEmail}"`);
        return false;
      }
      
      const registeredEmail = sheetEmailMatch[0].toLowerCase(); // cite: 81
      if (registeredEmail === senderEmail) { // cite: 81
        Logger.log(`-> Email validated for Rider ${riderNumber}.`);
        return true; 
      } else {
        Logger.log(`-> Email mismatch for Rider ${riderNumber}. Expected: ${registeredEmail}, Got: ${senderEmail}`); // cite: 81
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
  const sheetName = data['rider-number']; // cite: 25
  const bonusToFind = data['bonus']; // cite: 25
  const sheet = ss.getSheetByName(sheetName); // cite: 25
  
  if (!sheet) throw new Error(`Sheet / Rider Number not found: ${sheetName}`); // cite: 27

  const lastRow = sheet.getLastRow(); // cite: 28
  if (lastRow < 2) throw new Error(`Sheet ${sheetName} has no data rows.`);

  const range = sheet.getRange(2, BONUS_COLUMN_INDEX, lastRow - 1, 1); // cite: 28, 52
  const values = range.getValues(); // cite: 29
  let rowToUpdate = -1; // cite: 29
  
  for (let i = 0; i < values.length; i++) { // cite: 29
    const cellValue = String(values[i][0]).trim(); // cite: 29
    if (cellValue === bonusToFind.trim()) { // cite: 30
      rowToUpdate = i + 2; // cite: 30
      break; // cite: 31
    }
  }

  if (rowToUpdate !== -1) { // cite: 32
    sheet.getRange(rowToUpdate, columnIndex).setValue("X");
    
    if (useEmailTime) { 
      sheet.getRange(rowToUpdate, columnIndex + 1).setValue(data.date);
    } else {
      sheet.getRange(rowToUpdate, columnIndex + 1).setValue(new Date()); 
    }
  } else {
    throw new Error(`Row identifier (Bonus ID: ${bonusToFind}) not found in sheet ${sheetName}.`); // cite: 35
  }
}

function isValidSubject(subject) { // cite: 36
  return /^\d+\s.+$/.test(subject.trim()); // cite: 37
}

function extractEmailData(message) { // cite: 38
  const subject = message.getSubject().trim(); // cite: 38
  const match = subject.match(/^(\d+)\s(.*)$/); // cite: 38

  return {
    date: message.getDate(), // cite: 40
    subject: subject, // cite: 40
    sender: message.getFrom(), // cite: 40
    'rider-number': match ? match[1] : null, // cite: 38, 40
    'bonus': match ? match[2] : null // cite: 38, 40
  };
}