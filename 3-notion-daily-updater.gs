// NOTION DAILY UPDATER — keeps your Home Dashboard live
// Runs every morning at 6am:
//   → Updates the date in your Home Dashboard
//   → Unchecks all completed to-do items (fresh start each day)
//   → On Mondays: resets the Weekly Focus section
// HOW TO USE:
//   1. Go to script.google.com → New project → name it "Notion Daily Updater"
//   2. Paste this entire file
//   3. Run → setupDailyTrigger (ONCE to activate)
//   4. Authorize when prompted

const NOTION_TOKEN = 'YOUR_NOTION_TOKEN_HERE'; // Get from notion.so/my-integrations
const HOME_DASHBOARD_ID = '34c52d56f77781e8b2c6f91e692e2d21';

function updateHomeDashboard() {
  const today = new Date();
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December'];
  const dateStr = days[today.getDay()] + ', ' + months[today.getMonth()] + ' ' +
                  today.getDate() + ', ' + today.getFullYear();
  const isMonday = today.getDay() === 1;

  const headers = {
    'Authorization': 'Bearer ' + NOTION_TOKEN,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28'
  };

  // Get all blocks in Home Dashboard
  const res = UrlFetchApp.fetch(
    'https://api.notion.com/v1/blocks/' + HOME_DASHBOARD_ID + '/children?page_size=100',
    { method: 'get', headers }
  );
  const blocks = JSON.parse(res.getContentText()).results;

  blocks.forEach(block => {
    // Update the bold date paragraph
    if (block.type === 'paragraph') {
      const texts = block.paragraph.rich_text;
      if (texts.length > 0 && texts[0].annotations && texts[0].annotations.bold) {
        const content = texts.map(t => t.plain_text).join('');
        // Only update if it looks like a date (contains a comma and year)
        if (content.includes(',') && content.match(/20\d\d/)) {
          UrlFetchApp.fetch('https://api.notion.com/v1/blocks/' + block.id, {
            method: 'patch',
            headers,
            payload: JSON.stringify({
              paragraph: {
                rich_text: [{
                  type: 'text',
                  text: { content: dateStr },
                  annotations: { bold: true }
                }]
              }
            })
          });
          Logger.log('Date updated to: ' + dateStr);
        }
      }
    }

    // Uncheck any completed to-do boxes (reset priorities each morning)
    if (block.type === 'to_do' && block.to_do.checked) {
      UrlFetchApp.fetch('https://api.notion.com/v1/blocks/' + block.id, {
        method: 'patch',
        headers,
        payload: JSON.stringify({ to_do: { checked: false } })
      });
    }
  });

  // On Mondays, log a reminder to update Weekly Focus
  if (isMonday) {
    Logger.log('📅 It\'s Monday — remember to update your Weekly Focus in Notion!');
    // Optional: send yourself a reminder email
    GmailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      '📅 Monday — Update your Notion Weekly Focus',
      'Open your Home Dashboard and set your main goal, one thing to improve, and something to be proud of this week.',
      {}
    );
  }

  Logger.log('✅ Home Dashboard updated for ' + dateStr);
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateHomeDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  Logger.log('✅ Daily updater set for 6am!');
  // Run once immediately to confirm it works
  updateHomeDashboard();
}
