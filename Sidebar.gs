// ═══════════════════════════════════════════════════════════════════════════
// GMAIL ADD-ON — Rally Scoring Sidebar
// Paste this file as a new script file (e.g. Sidebar.gs) in your
// Apps Script project alongside code.js
// ═══════════════════════════════════════════════════════════════════════════

// ─── Add-on Entry Point ───────────────────────────────────────────────────────

/**
 * Called by Gmail when the add-on is opened on an email.
 * Returns a Card showing the submission details and action buttons.
 */
function buildAddOn(e) {
  const messageId = e.gmail.messageId;
  const message = GmailApp.getMessageById(messageId);
  const thread = message.getThread();
  const subject = message.getSubject();

  // Load config from the linked spreadsheet
  let config, ss;
  try {
    ss = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return errorCard_('Config error', `Could not load config: ${err.message}. Make sure setup() has been run.`);
  }

  // Parse the email subject
  if (!isValidSubject(subject)) {
    return infoCard_('Not a rally submission', subject, 'This email subject does not match the rally format (Rider Number + Bonus ID).', null);
  }

  const data = extractEmailData(message);
  const threadLabels = thread.getLabels().map(l => l.getName());
  const status = getStatus_(threadLabels, config);

  return buildScoringCard_(data, status, messageId, config, threadLabels);
}

// ─── Card Builders ────────────────────────────────────────────────────────────

function buildScoringCard_(data, status, messageId, config, threadLabels) {
  const card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader()
      .setTitle('Rally Scoring')
      .setSubtitle('Bonus submission')
  );

  // ── Submission details section ──
  const detailsSection = CardService.newCardSection().setHeader('Submission details');

  detailsSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Rider number')
      .setContent(data['rider-number'] || '—')
  );
  detailsSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Bonus ID')
      .setContent(data['bonus'] || '—')
  );
  detailsSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Submitted')
      .setContent(data.date ? data.date.toLocaleString() : '—')
  );
  detailsSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('From')
      .setContent(data.sender || '—')
  );

  card.addSection(detailsSection);

  // ── Current status section ──
  const statusSection = CardService.newCardSection().setHeader('Current status');
  statusSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Status')
      .setContent(status.label)
      .setIconUrl(status.iconUrl)
  );
  card.addSection(statusSection);

  // ── Actions section ──
  const actionsSection = CardService.newCardSection().setHeader('Actions');

  const isScored   = threadLabels.includes(config['label_scored']);
  const isApproved = threadLabels.includes(config['label_approved']);
  const isRejected = threadLabels.includes(config['label_processing_error']);

  if (isScored) {
    actionsSection.addWidget(
      CardService.newTextParagraph().setText('✅ This submission has already been scored. No further action needed.')
    );
  } else {
    // Approve button
    if (!isApproved) {
      const approveAction = CardService.newAction()
        .setFunctionName('handleApprove')
        .setParameters({ messageId });
      actionsSection.addWidget(
        CardService.newTextButton()
          .setText('✅  Approve submission')
          .setOnClickAction(approveAction)
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#1A56AA')
      );
    }

    // Reject button
    if (!isRejected) {
      const rejectAction = CardService.newAction()
        .setFunctionName('handleReject')
        .setParameters({ messageId });
      actionsSection.addWidget(
        CardService.newTextButton()
          .setText('❌  Reject submission')
          .setOnClickAction(rejectAction)
      );
    }

    // Flag for review button
    const flagAction = CardService.newAction()
      .setFunctionName('handleFlag')
      .setParameters({ messageId });
    actionsSection.addWidget(
      CardService.newTextButton()
        .setText('🚩  Flag for review')
        .setOnClickAction(flagAction)
    );
  }

  card.addSection(actionsSection);
  return card.build();
}

function errorCard_(title, message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Rally Scoring').setSubtitle(title))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(`⚠️ ${message}`))
    )
    .build();
}

function infoCard_(title, subject, message, messageId) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Rally Scoring').setSubtitle(title))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newKeyValue().setTopLabel('Subject').setContent(subject))
        .addWidget(CardService.newTextParagraph().setText(message))
    )
    .build();
}

// ─── Action Handlers ──────────────────────────────────────────────────────────

function handleApprove(e) {
  const messageId = e.parameters.messageId;
  const message = GmailApp.getMessageById(messageId);
  const thread = message.getThread();

  let config, ss;
  try {
    ss = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Config error: ${err.message}`))
      .build();
  }

  const labels = getLabels_(config);
  if (!labels) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Could not load Gmail labels. Run setup() first.'))
      .build();
  }

  try {
    const data = extractEmailData(message);

    // Only apply the approved label — processEmails() will detect it on its
    // next run, write the spreadsheet, and apply the scored label.
    thread.addLabel(labels.approved);
    thread.removeLabel(labels.needsReview);
    thread.refresh();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✅ Approved: Rider ${data['rider-number']} — ${data['bonus']}. Score will be recorded on the next processing run.`))
      .setStateChanged(true)
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Error: ${err.message}`))
      .build();
  }
}

function handleReject(e) {
  const messageId = e.parameters.messageId;
  const message = GmailApp.getMessageById(messageId);
  const thread = message.getThread();

  let config, ss;
  try {
    ss = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Config error: ${err.message}`))
      .build();
  }

  const labels = getLabels_(config);
  if (!labels) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Could not load Gmail labels.'))
      .build();
  }

  try {
    const data = extractEmailData(message);
    thread.addLabel(labels.processingError);
    thread.removeLabel(labels.needsReview);
    thread.removeLabel(labels.approved);
    thread.refresh();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`❌ Rejected: Rider ${data['rider-number']} — ${data['bonus']}`))
      .setStateChanged(true)
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Error: ${err.message}`))
      .build();
  }
}

function handleFlag(e) {
  const messageId = e.parameters.messageId;
  const message = GmailApp.getMessageById(messageId);
  const thread = message.getThread();

  let config;
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Config error: ${err.message}`))
      .build();
  }

  const labels = getLabels_(config);
  if (!labels) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Could not load Gmail labels.'))
      .build();
  }

  try {
    const data = extractEmailData(message);
    thread.addLabel(labels.needsReview);
    thread.removeLabel(labels.approved);
    thread.refresh();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`🚩 Flagged for review: Rider ${data['rider-number']} — ${data['bonus']}`))
      .setStateChanged(true)
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Error: ${err.message}`))
      .build();
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Reads the Spreadsheet ID from script properties.
 * Falls back to the constant in code.js if not set.
 */
function getSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return id;
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) return SPREADSHEET_ID;
  throw new Error('Spreadsheet ID not configured. Run setup() from code.js first.');
}

/**
 * Loads all label objects from Gmail using config keys.
 * Returns null if any label is missing.
 */
function getLabels_(config) {
  const keys = {
    unprocessed:     'label_unprocessed',
    formatError:     'label_format_error',
    emailError:      'label_email_error',
    processingError: 'label_processing_error',
    needsReview:     'label_needs_review',
    approved:        'label_approved',
    scored:          'label_scored',
  };
  const labels = {};
  for (const [key, configKey] of Object.entries(keys)) {
    const name = config[configKey];
    if (!name) return null;
    const label = GmailApp.getUserLabelByName(name);
    if (!label) return null;
    labels[key] = label;
  }
  return labels;
}

/**
 * Determines a human-readable status string from the thread's labels.
 */
function getStatus_(threadLabels, config) {
  const has = name => threadLabels.includes(name);
  if (has(config['label_scored']))           return { label: 'Scored ✅',          iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/check_circle/default/24px.svg' };
  if (has(config['label_approved']))         return { label: 'Approved — pending score', iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/thumb_up/default/24px.svg' };
  if (has(config['label_needs_review']))     return { label: 'Needs review 👀',    iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/visibility/default/24px.svg' };
  if (has(config['label_processing_error'])) return { label: 'Rejected ❌',        iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/cancel/default/24px.svg' };
  if (has(config['label_email_error']))      return { label: 'Email error ⚠️',     iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/warning/default/24px.svg' };
  if (has(config['label_format_error']))     return { label: 'Format error ⚠️',    iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/warning/default/24px.svg' };
  if (has(config['label_unprocessed']))      return { label: 'Unprocessed ⏳',     iconUrl: 'https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/schedule/default/24px.svg' };
  return { label: 'Unknown', iconUrl: '' };
}