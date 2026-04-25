// ============================================================
// NOTION DAILY RESET
// Keeps any Notion dashboard fresh automatically every morning:
//   ✅ Updates today's date on your dashboard
//   ✅ Unchecks completed to-do items (fresh start each day)
//   ✅ Sends a Monday email reminder to set your weekly goals
//
// HOW TO USE:
//   1. Create a Notion integration at notion.so/my-integrations
//      → New integration → Internal → copy the token
//   2. Share your Notion page with the integration
//      → Open page → ··· menu → Connections → add your integration
//   3. Paste this script into script.google.com → New project
//   4. Fill in CONFIG below
//   5. Run setupTrigger() once — authorize when prompted
// ============================================================

const CONFIG = {
  // From notion.so/my-integrations → your integration → Internal Integration Secret
  NOTION_TOKEN: 'YOUR_NOTION_TOKEN_HERE',

  // From your Notion page URL: notion.so/YOUR_PAGE_ID_HERE
  // Remove dashes if present
  PAGE_ID: 'YOUR_PAGE_ID_HERE',

  // Hour to run daily (24h clock — 6 = 6am)
  RUN_HOUR: 6,

  // Update the bold date line on your page each morning
  UPDATE_DATE: true,

  // Uncheck all completed to-do items each morning
  RESET_TODOS: true,

  // Send a Monday email reminder to set weekly goals
  WEEKLY_REMINDER_EMAIL: Session.getActiveUser().getEmail(), // or '' to disable
};

function runDailyReset() {
  const headers = {
    'Authorization': 'Bearer ' + CONFIG.NOTION_TOKEN,
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28',
  };

  const today = new Date();
  const DAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const dateStr = DAYS[today.getDay()] + ', ' + MONTHS[today.getMonth()] + ' ' + today.getDate() + ', ' + today.getFullYear();

  const res = UrlFetchApp.fetch(
    'https://api.notion.com/v1/blocks/' + CONFIG.PAGE_ID + '/children?page_size=100',
    { method: 'get', headers }
  );
  const blocks = JSON.parse(res.getContentText()).results;

  blocks.forEach(block => {
    // Update bold date paragraph
    if (CONFIG.UPDATE_DATE && block.type === 'paragraph') {
      const texts = block.paragraph.rich_text;
      if (texts.length > 0 && texts[0].annotations && texts[0].annotations.bold) {
        const content = texts.map(t => t.plain_text).join('');
        if (content.includes(',') && content.match(/20\d\d/)) {
          UrlFetchApp.fetch('https://api.notion.com/v1/blocks/' + block.id, {
            method: 'patch', headers,
            payload: JSON.stringify({
              paragraph: {
                rich_text: [{ type: 'text', text: { content: dateStr }, annotations: { bold: true } }]
              }
            })
          });
          Logger.log('Date updated to: ' + dateStr);
        }
      }
    }

    // Uncheck completed to-dos
    if (CONFIG.RESET_TODOS && block.type === 'to_do' && block.to_do.checked) {
      UrlFetchApp.fetch('https://api.notion.com/v1/blocks/' + block.id, {
        method: 'patch', headers,
        payload: JSON.stringify({ to_do: { checked: false } })
      });
    }
  });

  // Monday weekly reminder
  if (today.getDay() === 1 && CONFIG.WEEKLY_REMINDER_EMAIL) {
    GmailApp.sendEmail(
      CONFIG.WEEKLY_REMINDER_EMAIL,
      '📅 Monday — Set your weekly goals in Notion',
      'Open your Notion dashboard and set your top priorities for the week.',
      {}
    );
    Logger.log('Monday reminder sent to ' + CONFIG.WEEKLY_REMINDER_EMAIL);
  }

  Logger.log('✅ Notion reset complete — ' + dateStr);
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runDailyReset')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.RUN_HOUR)
    .create();
  Logger.log('✅ Daily reset scheduled for ' + CONFIG.RUN_HOUR + ':00. Running now to test...');
  runDailyReset();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('All triggers removed.');
}
