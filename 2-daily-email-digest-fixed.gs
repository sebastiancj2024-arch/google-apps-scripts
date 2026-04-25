// DAILY EMAIL DIGEST — fixed version
// Sends a morning email summary of your unread Gmail
// HOW TO USE:
//   1. Go to script.google.com → open your "Daily Email Digest" project
//   2. REPLACE all existing code with this
//   3. Run → setupDigestTrigger (do this ONCE to activate the 7am trigger)
//   4. Authorize Gmail access when prompted
//   5. You'll get a daily email at 7am — check Execution Log if it doesn't arrive

function dailyEmailDigest() {
  const today = new Date();
  const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);

  // Proper date format for Gmail search (YYYY/MM/DD)
  const afterDate = Utilities.formatDate(
    yesterday,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );

  const threads = GmailApp.search('after:' + afterDate + ' is:unread');

  let html = '<h2 style="color:#1a73e8;">📬 Daily Email Summary</h2>';
  html += '<p><b>Date:</b> ' + today.toDateString() + '</p>';
  html += '<p><b>Unread emails since yesterday:</b> ' + threads.length + '</p>';
  html += '<hr>';

  if (threads.length === 0) {
    html += '<p>✅ <b>Inbox zero!</b> No unread emails from yesterday.</p>';
  } else {
    threads.slice(0, 15).forEach(thread => {
      const msg = thread.getMessages()[0];
      const from = msg.getFrom();
      const subject = msg.getSubject();
      const preview = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ');
      html += '<p>';
      html += '<b>From:</b> ' + from + '<br>';
      html += '<b>Subject:</b> ' + subject + '<br>';
      html += '<b>Preview:</b> ' + preview + '...<br>';
      html += '</p><hr>';
    });

    if (threads.length > 15) {
      html += '<p><i>+ ' + (threads.length - 15) + ' more emails not shown.</i></p>';
    }
  }

  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    '📬 Daily Email Digest — ' + today.toDateString(),
    '',
    { htmlBody: html }
  );

  Logger.log('Digest sent. Unread count: ' + threads.length);
}

function setupDigestTrigger() {
  // Removes old triggers for this script only, then creates fresh one
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyEmailDigest')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  Logger.log('✅ Trigger set! Daily digest will run at 7am.');
  // Run it immediately so you can confirm it works right now
  dailyEmailDigest();
}
