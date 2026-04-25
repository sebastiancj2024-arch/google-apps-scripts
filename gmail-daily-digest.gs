// ============================================================
// GMAIL DAILY DIGEST
// Get a clean morning summary of your unread emails, every day.
// Works with any Gmail account. Set up in 2 minutes.
//
// HOW TO USE:
//   1. Go to script.google.com → New project
//   2. Paste this file
//   3. Edit CONFIG below (at minimum set SEND_HOUR)
//   4. Click Run → setupTrigger
//   5. Authorize Gmail access when prompted
//   6. Done — digest lands in your inbox every morning
// ============================================================

const CONFIG = {
  RECIPIENT_EMAIL: Session.getActiveUser().getEmail(), // Change to 'you@gmail.com' if needed
  SEND_HOUR: 7,       // What hour to send (24h clock — 7 = 7am, 8 = 8am)
  MAX_EMAILS: 20,     // Max emails to show per digest
  UNREAD_ONLY: true,  // true = unread only | false = all email from last 24h
};

function sendDailyDigest() {
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const afterDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const query = (CONFIG.UNREAD_ONLY ? 'is:unread ' : '') + 'after:' + afterDate;
  const threads = GmailApp.search(query);

  let html = '<h2 style="color:#1a73e8;font-family:sans-serif;">📬 Daily Email Digest</h2>';
  html += '<p style="font-family:sans-serif;color:#444"><b>Date:</b> ' + now.toDateString() + '</p>';
  html += '<p style="font-family:sans-serif;color:#444"><b>' +
    (CONFIG.UNREAD_ONLY ? 'Unread emails (last 24h):' : 'Emails (last 24h):') + '</b> ' +
    threads.length + '</p><hr>';

  if (threads.length === 0) {
    html += '<p style="font-family:sans-serif;">✅ <b>Inbox zero!</b> Nothing unread from the last 24 hours.</p>';
  } else {
    threads.slice(0, CONFIG.MAX_EMAILS).forEach(thread => {
      const msg = thread.getMessages()[0];
      const preview = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ').trim();
      html += '<div style="font-family:sans-serif;margin-bottom:12px;">';
      html += '<b>From:</b> ' + msg.getFrom() + '<br>';
      html += '<b>Subject:</b> ' + msg.getSubject() + '<br>';
      html += '<span style="color:#666">' + preview + '...</span>';
      html += '</div><hr>';
    });
    if (threads.length > CONFIG.MAX_EMAILS) {
      html += '<p style="font-family:sans-serif;color:#888"><i>+ ' +
        (threads.length - CONFIG.MAX_EMAILS) + ' more emails not shown.</i></p>';
    }
  }

  GmailApp.sendEmail(
    CONFIG.RECIPIENT_EMAIL,
    '📬 Daily Digest — ' + now.toDateString(),
    '',
    { htmlBody: html }
  );
  Logger.log('Digest sent to ' + CONFIG.RECIPIENT_EMAIL + '. Emails found: ' + threads.length);
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.SEND_HOUR)
    .create();
  Logger.log('✅ Daily digest scheduled for ' + CONFIG.SEND_HOUR + ':00. Sending test now...');
  sendDailyDigest();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('All triggers removed.');
}
