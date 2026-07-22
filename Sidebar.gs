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
      'No message in this thread matches the required format: Rider Number followed by a 4-letter Bonus Code.');
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

  const isScored      = threadLabels.includes(config['label_scored']);
  const isDenied       = threadLabels.includes(config['label_denied']);
  const isApproved     = threadLabels.includes(config['label_approved']);
  const isFormatError  = threadLabels.includes(config['label_format_error']);
  const isEmailError    = threadLabels.includes(config['label_email_error']);
  const isProcError     = threadLabels.includes(config['label_processing_error']);
  const isUnprocessed   = threadLabels.includes(config['label_unprocessed']);
  const hasErrorLabel  = isFormatError || isEmailError || isProcError;

  // Error explanation section - shown once at the top if any error label is present
  if (hasErrorLabel) {
    const errSec = CardService.newCardSection().setHeader('Error');
    let errText = '';
    if (isFormatError) {
      errText = 'Subject line does not match the required format: Rider Number followed by a 4-letter Bonus Code (e.g. "42 ABCD"). Ask the rider to resend with the correct subject.';
    } else if (isEmailError) {
      errText = 'The sender email does not match the registered address for this rider number in Rider Master. Verify the rider\'s registered email or check for a typo.';
    } else if (isProcError) {
      errText = 'A bonus code error occurred while recording this submission - the bonus code could not be matched in the rider sheet. Check the Apps Script Executions log for details.';
    }
    errSec.addWidget(CardService.newTextParagraph().setText(errText));
    card.addSection(errSec);
  }

  // One section per message
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
    } else if (hasErrorLabel) {
      sec.addWidget(CardService.newTextParagraph().setText('Approve/Deny unavailable while an error label is present. See Error section above.'));
    } else if (isUnprocessed) {
      sec.addWidget(CardService.newTextParagraph().setText('Waiting for the next processing run to validate this submission. Approve/Deny will be available once it moves to Email Requires Review.'));
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
      // Deny button for this message - goes to confirmation card first
      sec.addWidget(CardService.newTextButton()
        .setText('Deny this message')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('showDenyConfirmation')
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
        .setText('Remove approval / revert to email-requires-review')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleRevertApproved')
          .setParameters({ firstMsgId: firstMsgId })));
    }

    if (isDenied) {
      revertSec.addWidget(CardService.newTextButton()
        .setText('Remove denial / revert to email-requires-review')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('handleRevertDenied')
          .setParameters({ firstMsgId: firstMsgId })));
    }

    card.addSection(revertSec);
  }

  return card.build();
}

// --- Deny confirmation card ----------------------------------------------------

/**
 * Shows a confirmation card before denying. Nothing is changed in Gmail
 * or the spreadsheet until the scorer clicks Confirm.
 */
function showDenyConfirmation(e) {
  const msgId    = e.parameters.msgId;
  const threadId = e.parameters.threadId;

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Confirm Bonus Denial')
      .setSubtitle('This will mark the bonus as denied'));

  const sec = CardService.newCardSection();
  sec.addWidget(CardService.newTextParagraph()
    .setText('Are you sure you want to deny this bonus submission? This will record a denial in the spreadsheet and apply the denied label.'));

  sec.addWidget(CardService.newTextButton()
    .setText('Confirm Bonus Denial')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleDeny')
      .setParameters({ msgId: msgId, threadId: threadId }))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor('#B3261E'));

  sec.addWidget(CardService.newTextButton()
    .setText('Cancel')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('cancelDeny')
      .setParameters({ msgId: msgId })));

  card.addSection(sec);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Cancel button on the deny confirmation card - returns to the main
 * scoring card without changing anything.
 */
function cancelDeny(e) {
  const msgId = e.parameters.msgId;
  const message = GmailApp.getMessageById(msgId);
  const thread  = message.getThread();
  const messages = thread.getMessages();

  let config, ss;
  try {
    ss     = SpreadsheetApp.openById(getSpreadsheetId_());
    config = loadConfig(ss);
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  }

  const threadLabels = thread.getLabels().map(l => l.getName());
  const threadStatus = getStatus_(threadLabels, config);
  const card = buildScoringCard_(messages, threadStatus, thread.getId(), config, threadLabels, ss);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard().updateCard(card))
    .build();
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
    updateSpreadsheet(ctx.ss, ctx.config, ctx.data,
      configInt_(ctx.config, 'col_approved', 4),
      configInt_(ctx.config, 'col_approved_time', 5),
      false);
    ctx.thread.addLabel(ctx.labels.approved);
    ctx.thread.addLabel(ctx.labels.scored);
    ctx.thread.removeLabel(ctx.labels.needsReview);
    ctx.thread.removeLabel(ctx.labels.denied);
    ctx.thread.refresh();

    // Remember exactly which message was approved so Revert targets the right row
    setActionedMessage_(threadId, 'approved', msgId);
    clearActionedMessage_(threadId, 'denied');

    // Rebuild and push the updated card so the sidebar reflects the new state immediately
    const messages      = ctx.thread.getMessages();
    const threadLabels  = ctx.thread.getLabels().map(l => l.getName());
    const threadStatus  = getStatus_(threadLabels, ctx.config);
    const card = buildScoringCard_(messages, threadStatus, ctx.thread.getId(), ctx.config, threadLabels, ctx.ss);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Bonus Accepted'))
      .setNavigation(CardService.newNavigation().updateCard(card))
      .build();
  } catch (err) {
    try {
      ctx.thread.addLabel(ctx.labels.approved);
      ctx.thread.removeLabel(ctx.labels.needsReview);
      ctx.thread.removeLabel(ctx.labels.denied);
      ctx.thread.removeLabel(ctx.labels.scored);
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
    updateSpreadsheet(ctx.ss, ctx.config, ctx.data,
      configInt_(ctx.config, 'col_denied', 6),
      configInt_(ctx.config, 'col_denied_time', 7),
      false);
    ctx.thread.addLabel(ctx.labels.denied);
    ctx.thread.removeLabel(ctx.labels.needsReview);
    ctx.thread.removeLabel(ctx.labels.approved);
    ctx.thread.removeLabel(ctx.labels.scored);
    ctx.thread.refresh();

    // Remember exactly which message was denied so Revert targets the right row
    setActionedMessage_(threadId, 'denied', msgId);
    clearActionedMessage_(threadId, 'approved');

    // Return to the main scoring card with an updated state and confirmation toast
    const messages = ctx.thread.getMessages();
    const threadLabels = ctx.thread.getLabels().map(l => l.getName());
    const threadStatus = getStatus_(threadLabels, ctx.config);
    const card = buildScoringCard_(messages, threadStatus, ctx.thread.getId(), ctx.config, threadLabels, ctx.ss);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Bonus Denied'))
      .setNavigation(CardService.newNavigation().popCard().updateCard(card))
      .build();
  } catch (err) {
    try {
      ctx.thread.addLabel(ctx.labels.denied);
      ctx.thread.removeLabel(ctx.labels.needsReview);
      ctx.thread.removeLabel(ctx.labels.approved);
      ctx.thread.removeLabel(ctx.labels.scored);
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
    // Clear approved column and timestamp in sheet - target the message that was
    // actually approved, not just the first message in the thread
    const threadId = ctx.thread.getId();
    const data = dataForMsgId_(getActionedMessage_(threadId, 'approved')) ||
      getFirstValidData_(ctx.thread) || ctx.data;
    if (data) updateSpreadsheet(ctx.ss, ctx.config, data,
      configInt_(ctx.config, 'col_approved', 4),
      configInt_(ctx.config, 'col_approved_time', 5),
      false, null);
    ctx.thread.removeLabel(ctx.labels.approved);
    ctx.thread.removeLabel(ctx.labels.scored);
    ctx.thread.addLabel(ctx.labels.needsReview);
    ctx.thread.refresh();
    clearActionedMessage_(threadId, 'approved');

    const messages      = ctx.thread.getMessages();
    const threadLabels  = ctx.thread.getLabels().map(l => l.getName());
    const threadStatus  = getStatus_(threadLabels, ctx.config);
    const card = buildScoringCard_(messages, threadStatus, ctx.thread.getId(), ctx.config, threadLabels, ctx.ss);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Reverted to email-requires-review.'))
      .setNavigation(CardService.newNavigation().updateCard(card))
      .build();
  } catch (err) {
    return notify_('Error: ' + err.message);
  }
}

function handleRevertDenied(e) {
  const ctx = loadThreadContext_(e.parameters.firstMsgId);
  if (!ctx.config) return notify_('Config error - run setup() first.');
  if (!ctx.labels) return notify_('Gmail labels missing - run setup() first.');

  try {
    // Clear denied column and timestamp in sheet - target the message that was
    // actually denied, not just the first message in the thread
    const threadId = ctx.thread.getId();
    const data = dataForMsgId_(getActionedMessage_(threadId, 'denied')) ||
      getFirstValidData_(ctx.thread) || ctx.data;
    if (data) updateSpreadsheet(ctx.ss, ctx.config, data,
      configInt_(ctx.config, 'col_denied', 6),
      configInt_(ctx.config, 'col_denied_time', 7),
      false, null);
    ctx.thread.removeLabel(ctx.labels.denied);
    ctx.thread.addLabel(ctx.labels.needsReview);
    ctx.thread.refresh();
    clearActionedMessage_(threadId, 'denied');

    const messages      = ctx.thread.getMessages();
    const threadLabels  = ctx.thread.getLabels().map(l => l.getName());
    const threadStatus  = getStatus_(threadLabels, ctx.config);
    const card = buildScoringCard_(messages, threadStatus, ctx.thread.getId(), ctx.config, threadLabels, ctx.ss);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Reverted to email-requires-review.'))
      .setNavigation(CardService.newNavigation().updateCard(card))
      .build();
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

/**
 * Remembers which specific message was approved/denied for a thread, so
 * Revert can target that exact message's data instead of guessing from
 * the first message in the thread.
 */
function setActionedMessage_(threadId, kind, msgId) {
  PropertiesService.getScriptProperties().setProperty(kind + '_msg_' + threadId, msgId);
}

function getActionedMessage_(threadId, kind) {
  return PropertiesService.getScriptProperties().getProperty(kind + '_msg_' + threadId);
}

function clearActionedMessage_(threadId, kind) {
  PropertiesService.getScriptProperties().deleteProperty(kind + '_msg_' + threadId);
}

/**
 * extractEmailData for a message ID, tolerating messages that no longer
 * exist or no longer have a valid subject.
 */
function dataForMsgId_(msgId) {
  if (!msgId) return null;
  try {
    const msg = GmailApp.getMessageById(msgId);
    return isValidSubject(msg.getSubject()) ? extractEmailData(msg) : null;
  } catch (e) {
    return null;
  }
}

function getSpreadsheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return id;
  throw new Error('Spreadsheet ID not set. Run setup() from code.js first.');
}

// loadLabels_ is defined once in code.js and shared via Apps Script's global scope.

function getStatus_(threadLabels, config) {
  const has = n => threadLabels.includes(n);
  if (has(config['label_scored']))            return 'Scored';
  if (has(config['label_approved']))          return 'Approved - pending score';
  if (has(config['label_denied']))            return 'Denied';
  if (has(config['label_needs_review']))      return 'Email Requires Review';
  if (has(config['label_processing_error']))  return 'Bonus Code Error';
  if (has(config['label_email_error']))       return 'Email error';
  if (has(config['label_format_error']))      return 'Subject Line Error';
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