// ===========================================================================
// GMAIL ADD-ON - Rally Scoring Sidebar
// Paste this as a script file named "Sidebar" alongside code.js in Apps Script.
// ===========================================================================

// --- Add-on entry point -------------------------------------------------------
function buildAddOn(e) {
  const message = GmailApp.getMessageById(e.gmail.messageId);
  const thread  = message.getThread();
  const subject = message.getSubject();

  let config, ss;
  try {
    ss     = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return errorCard_('Config error', err.message + ' - make sure setup() has been run.');
  }

  if (!isValidSubject(subject)) {
    return infoCard_('Not a rally submission', subject,
      'Subject does not match the required format: Rider Number followed by Bonus ID.');
  }

  const data         = extractEmailData(message);
  const threadLabels = thread.getLabels().map(l => l.getName());
  const status       = getStatus_(threadLabels, config);

  return buildScoringCard_(data, status, e.gmail.messageId, config, threadLabels);
}

// --- Card builders ------------------------------------------------------------
function buildScoringCard_(data, status, messageId, config, threadLabels) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Rally Scoring')
      .setSubtitle('Bonus submission'));

  // Submission details
  const details = CardService.newCardSection().setHeader('Submission details');
  details.addWidget(CardService.newKeyValue().setTopLabel('Rider number').setContent(data['rider-number'] || '-'));
  details.addWidget(CardService.newKeyValue().setTopLabel('Bonus ID').setContent(data['bonus'] || '-'));
  details.addWidget(CardService.newKeyValue().setTopLabel('Submitted').setContent(data.date ? data.date.toLocaleString() : '-'));
  details.addWidget(CardService.newKeyValue().setTopLabel('From').setContent(data.sender || '-'));
  card.addSection(details);

  // Status
  const statusSec = CardService.newCardSection().setHeader('Status');
  statusSec.addWidget(CardService.newKeyValue().setTopLabel('Current status').setContent(status));
  card.addSection(statusSec);

  // Actions
  const actions  = CardService.newCardSection().setHeader('Actions');
  const isScored  = threadLabels.includes(config['label_scored']);
  const isApproved = threadLabels.includes(config['label_approved']);
  const isDenied  = threadLabels.includes(config['label_denied']);

  if (isScored) {
    actions.addWidget(CardService.newTextParagraph()
      .setText('This submission has already been scored. No further action needed.'));
  } else if (isDenied) {
    actions.addWidget(CardService.newTextParagraph()
      .setText('This submission has been denied.'));
  } else {
    if (!isApproved) {
      actions.addWidget(CardService.newTextButton()
        .setText('Approve submission')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleApprove')
          .setParameters({messageId}))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor('#1A56AA'));
    }

    actions.addWidget(CardService.newTextButton()
      .setText('Deny submission')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleDeny')
        .setParameters({messageId})));

    actions.addWidget(CardService.newTextButton()
      .setText('Flag for review')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('handleFlag')
        .setParameters({messageId})));
  }
  card.addSection(actions);
  return card.build();
}

function errorCard_(subtitle, msg) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Rally Scoring').setSubtitle(subtitle))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(msg)))
    .build();
}

function infoCard_(subtitle, subject, msg) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Rally Scoring').setSubtitle(subtitle))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newKeyValue().setTopLabel('Subject').setContent(subject))
      .addWidget(CardService.newTextParagraph().setText(msg)))
    .build();
}

// --- Action handlers ----------------------------------------------------------

function handleApprove(e) {
  const {thread, config, ss, labels, data} = loadContext_(e.parameters.messageId);
  if (!config) return notify_('Config error - run setup() first.');
  if (!labels) return notify_('Gmail labels missing - run setup() first.');
  try {
    // Write approval to spreadsheet immediately, then apply approved + scored labels
    updateSpreadsheet(ss, config, data, parseInt(config['col_approved'], 10), false);
    thread.addLabel(labels.approved);
    thread.addLabel(labels.scored);
    thread.removeLabel(labels.needsReview);
    thread.refresh();
    return notifyRefresh_('Approved: Rider ' + data['rider-number'] + ' - ' + data['bonus']);
  } catch (err) {
    // Sheet write failed - still apply approved label so trigger can pick it up
    try { thread.addLabel(labels.approved); thread.removeLabel(labels.needsReview); thread.refresh(); } catch(_) {}
    return notifyRefresh_('Approved (sheet update failed: ' + err.message + ')');
  }
}

function handleDeny(e) {
  const {thread, config, ss, labels, data} = loadContext_(e.parameters.messageId);
  if (!config) return notify_('Config error - run setup() first.');
  if (!labels) return notify_('Gmail labels missing - run setup() first.');
  try {
    // Write X + timestamp to the denied column immediately
    updateSpreadsheet(ss, config, data, parseInt(config['col_denied'], 10), false);

    thread.addLabel(labels.denied);
    thread.removeLabel(labels.needsReview);
    thread.removeLabel(labels.approved);
    thread.refresh();
    return notifyRefresh_('Denied: Rider ' + data['rider-number'] + ' - ' + data['bonus']);
  } catch (err) {
    // Still apply the label even if the sheet write fails
    try {
      thread.addLabel(labels.denied);
      thread.removeLabel(labels.needsReview);
      thread.refresh();
    } catch (_) {}
    return notifyRefresh_('Denied (sheet update failed: ' + err.message + ')');
  }
}

function handleFlag(e) {
  const {thread, config, labels, data} = loadContext_(e.parameters.messageId);
  if (!config) return notify_('Config error - run setup() first.');
  if (!labels) return notify_('Gmail labels missing - run setup() first.');
  try {
    thread.addLabel(labels.needsReview);
    thread.removeLabel(labels.approved);
    thread.refresh();
    return notifyRefresh_('Flagged for review: Rider ' + data['rider-number'] + ' - ' + data['bonus']);
  } catch (err) {
    return notify_('Error: ' + err.message);
  }
}

// --- Helpers ------------------------------------------------------------------

function loadContext_(messageId) {
  const message = GmailApp.getMessageById(messageId);
  const thread  = message.getThread();
  const data    = extractEmailData(message);
  let config = null, ss = null, labels = null;
  try {
    ss     = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
    labels = loadLabels_(config);
  } catch (_) {}
  return {message, thread, config, ss, labels, data};
}

function getSpreadsheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return id;
  throw new Error('Spreadsheet ID not set. Run setup() from code.js first.');
}

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
  if (missing.length) return null;
  return labels;
}

function getStatus_(threadLabels, config) {
  const has = n => threadLabels.includes(n);
  if (has(config['label_scored']))           return 'Scored';
  if (has(config['label_approved']))         return 'Approved - pending score';
  if (has(config['label_denied']))           return 'Denied';
  if (has(config['label_needs_review']))     return 'Needs review';
  if (has(config['label_processing_error'])) return 'Processing error';
  if (has(config['label_email_error']))      return 'Email error';
  if (has(config['label_format_error']))     return 'Format error';
  if (has(config['label_unprocessed']))      return 'Unprocessed';
  return 'Unknown';
}

function notify_(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text)).build();
}

function notifyRefresh_(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .setStateChanged(true).build();
}
