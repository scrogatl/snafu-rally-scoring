/**
 * Rally Scoring - Sidebar.gs
 * Gmail Add-on sidebar. Shows all messages in a thread individually,
 * with Approve / Deny buttons per message and Revert controls for the thread.
 * Reads config from the Config sheet via loadConfig() defined in code.js.
 * Add this as a script file named "Sidebar" in Apps Script.
 */

// --- Add-on entry point -------------------------------------------------------
function buildAddOn(e) {
  const messageId = e.gmail.messageId;
  const message   = GmailApp.getMessageById(messageId);
  const thread    = message.getThread();
  const messages  = thread.getMessages();

  let config, ss;
  try {
    ss     = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return errorCard_('Config error', err.message + ' - make sure setup() has been run.');
  }

  // Check if at least one message in thread is a valid rally submission
  const hasValidMessage = messages.some(m => isValidSubject(m.getSubject()));
  if (!hasValidMessage) {
    return infoCard_('Not a rally submission',
      message.getSubject(),
      'No message in this thread matches the required format: Rider Number followed by Bonus ID.');
  }

  const threadLabels = thread.getLabels().map(l => l.getName());
  const threadStatus = getStatus_(threadLabels, config);
  const threadId     = thread.getId();

  return buildScoringCard_(messages, threadStatus, threadId, config, threadLabels, ss);
}

// --- Card builder -------------------------------------------------------------
function buildScoringCard_(messages, threadStatus, threadId, config, threadLabels, ss) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Rally Scoring')
      .setSubtitle('Thread: ' + messages.length + ' message(s)'));

  // Thread status section
  const statusSec = CardService.newCardSection().setHeader('Thread status');
  statusSec.addWidget(
    CardService.newKeyValue().setTopLabel('Status').setContent(threadStatus)
  );
  card.addSection(statusSec);

  // One section per message
  const isScored  = threadLabels.includes(config['label_scored']);
  const isDenied  = threadLabels.includes(config['label_denied']);
  const isApproved = threadLabels.includes(config['label_approved']);

  messages.forEach(function(message, idx) {
    const subject = message.getSubject();
    const valid   = isValidSubject(subject);
    const data    = valid ? extractEmailData(message) : null;
    const msgId   = message.getId();
    const label   = 'Message ' + (idx + 1) + (idx === messages.length - 1 ? ' (latest)' : '');

    const sec = CardService.newCardSection().setHeader(label);
    sec.addWidget(CardService.newKeyValue().setTopLabel('Subject').setContent(subject));
    sec.addWidget(CardService.newKeyValue().setTopLabel('From').setContent(message.getFrom()));
    sec.addWidget(CardService.newKeyValue().setTopLabel('Date').setContent(message.getDate().toLocaleString()));

    if (!valid) {
      sec.addWidget(CardService.newTextParagraph().setText('Invalid format - not a rally submission.'));
    } else if (isScored) {
      sec.addWidget(CardService.newTextParagraph().setText('Thread is scored. Use Revert below to undo.'));
    } else if (isDenied) {
      sec.addWidget(CardService.newTextParagraph().setText('Thread is denied. Use Revert below to undo.'));
    } else {
      // Approve button for this message
      if (!isApproved) {
        sec.addWidget(CardService.newTextButton()
          .setText('Approve this message')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('handleApprove')
            .setParameters({ msgId: msgId, threadId: threadId }))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor('#1A56AA'));
      }
      // Deny button for this message
      sec.addWidget(CardService.newTextButton()
        .setText('Deny this message')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleDeny')
          .setParameters({ msgId: msgId, threadId: threadId })));
    }

    card.addSection(sec);
  });

  // Revert section - always shown if thread has been actioned
  if (isApproved || isDenied || isScored) {
    const revertSec = CardService.newCardSection().setHeader('Revert');

    const firstMsgId = messages[0].getId();

    if (isApproved || isScored) {
      revertSec.addWidget(CardService.newTextButton()
        .setText('Remove approval / revert to needs-review')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleRevertApproved')
          .setParameters({ firstMsgId: firstMsgId })));
    }

    if (isDenied) {
      revertSec.addWidget(CardService.newTextButton()
        .setText('Remove denial / revert to needs-review')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleRevertDenied')
          .setParameters({ firstMsgId: firstMsgId })));
    }

    card.addSection(revertSec);
  }

  return card.build();
}

// --- Action handlers ----------------------------------------------------------

function handleApprove(e) {
  const msgId    = e.parameters.msgId;
  const threadId = e.parameters.threadId;
  const ctx      = loadThreadContext_(msgId);
  if (!ctx.config) return notify_('Config error - run setup() first.');
  if (!ctx.labels) return notify_('Gmail labels missing - run setup() first.');

  try {
    // Write approval to sheet using this specific message's data
    updateSpreadsheet(ctx.ss, ctx.config, ctx.data, parseInt(ctx.config['col_approved'], 10), false);
    ctx.thread.addLabel(ctx.labels.approved);
    ctx.thread.addLabel(ctx.labels.scored);
    ctx.thread.removeLabel(ctx.labels.needsReview);
    ctx.thread.removeLabel(ctx.labels.denied);
    ctx.thread.refresh();
    return notifyRefresh_('Approved: Rider ' + ctx.data['rider-number'] + ' - ' + ctx.data['bonus']);
  } catch (err) {
    try {
      ctx.thread.addLabel(ctx.labels.approved);
      ctx.thread.removeLabel(ctx.labels.needsReview);
      ctx.thread.refresh();
    } catch(_) {}
    return notifyRefresh_('Approved (sheet update failed: ' + err.message + ')');
  }
}

function handleDeny(e) {
  const msgId    = e.parameters.msgId;
  const threadId = e.parameters.threadId;
  const ctx      = loadThreadContext_(msgId);
  if (!ctx.config) return notify_('Config error - run setup() first.');
  if (!ctx.labels) return notify_('Gmail labels missing - run setup() first.');

  try {
    updateSpreadsheet(ctx.ss, ctx.config, ctx.data, parseInt(ctx.config['col_denied'], 10), false);
    ctx.thread.addLabel(ctx.labels.denied);
    ctx.thread.removeLabel(ctx.labels.needsReview);
    ctx.thread.removeLabel(ctx.labels.approved);
    ctx.thread.removeLabel(ctx.labels.scored);
    ctx.thread.refresh();
    return notifyRefresh_('Denied: Rider ' + ctx.data['rider-number'] + ' - ' + ctx.data['bonus']);
  } catch (err) {
    try {
      ctx.thread.addLabel(ctx.labels.denied);
      ctx.thread.removeLabel(ctx.labels.needsReview);
      ctx.thread.refresh();
    } catch(_) {}
    return notifyRefresh_('Denied (sheet update failed: ' + err.message + ')');
  }
}

function handleRevertApproved(e) {
  const ctx = loadThreadContext_(e.parameters.firstMsgId);
  if (!ctx.config) return notify_('Config error - run setup() first.');
  if (!ctx.labels) return notify_('Gmail labels missing - run setup() first.');

  try {
    // Clear approved column and timestamp in sheet
    const data = getFirstValidData_(ctx.thread) || ctx.data;
    if (data) updateSpreadsheet(ctx.ss, ctx.config, data, parseInt(ctx.config['col_approved'], 10), false, null);
    ctx.thread.removeLabel(ctx.labels.approved);
    ctx.thread.removeLabel(ctx.labels.scored);
    ctx.thread.addLabel(ctx.labels.needsReview);
    ctx.thread.refresh();
    return notifyRefresh_('Reverted to needs-review.');
  } catch (err) {
    return notify_('Error: ' + err.message);
  }
}

function handleRevertDenied(e) {
  const ctx = loadThreadContext_(e.parameters.firstMsgId);
  if (!ctx.config) return notify_('Config error - run setup() first.');
  if (!ctx.labels) return notify_('Gmail labels missing - run setup() first.');

  try {
    // Clear denied column and timestamp in sheet
    const data = getFirstValidData_(ctx.thread) || ctx.data;
    if (data) updateSpreadsheet(ctx.ss, ctx.config, data, parseInt(ctx.config['col_denied'], 10), false, null);
    ctx.thread.removeLabel(ctx.labels.denied);
    ctx.thread.addLabel(ctx.labels.needsReview);
    ctx.thread.refresh();
    return notifyRefresh_('Reverted to needs-review.');
  } catch (err) {
    return notify_('Error: ' + err.message);
  }
}

// --- Helpers ------------------------------------------------------------------

/**
 * Loads context from a message ID. Always pass a msgId -
 * for revert handlers, pass the first message ID in the thread.
 */
function loadThreadContext_(msgId) {
  let message, thread, data;
  try {
    message = GmailApp.getMessageById(msgId);
    thread  = message.getThread();
    data    = extractEmailData(message);
  } catch(_) {}

  let config = null, ss = null, labels = null;
  try {
    ss     = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
    labels = loadLabels_(config);
  } catch(_) {}

  return { message, thread, data, config, ss, labels };
}

/**
 * Returns extractEmailData for the first message in the thread
 * that has a valid rally subject line.
 */
function getFirstValidData_(thread) {
  if (!thread) return null;
  const messages = thread.getMessages();
  for (const msg of messages) {
    if (isValidSubject(msg.getSubject())) return extractEmailData(msg);
  }
  return null;
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
  if (has(config['label_scored']))            return 'Scored';
  if (has(config['label_approved']))          return 'Approved - pending score';
  if (has(config['label_denied']))            return 'Denied';
  if (has(config['label_needs_review']))      return 'Needs review';
  if (has(config['label_processing_error']))  return 'Processing error';
  if (has(config['label_email_error']))       return 'Email error';
  if (has(config['label_format_error']))      return 'Format error';
  if (has(config['label_unprocessed']))       return 'Unprocessed';
  return 'Unknown';
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

function notify_(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text)).build();
}

function notifyRefresh_(text) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(text))
    .setStateChanged(true).build();
}