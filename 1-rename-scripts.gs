// RENAME SCRIPTS — run once, then delete this file
// Uses DriveApp (built into Apps Script) to rename your script projects by ID
// HOW TO USE:
//   1. Go to script.google.com → New project
//   2. Paste this entire file
//   3. Click Run → renameAllScripts
//   4. Authorize when prompted
//   5. Check Execution Log — should say "Done!"

function renameAllScripts() {
  const renames = [
    {
      id: '1pp8emClEHvFrdjXpn8POOFiJSHVpi0ngNsKfUf8xnJfAWeZzIwDqgMIX',
      name: 'Notion HQ Builder'
    },
    {
      id: '1NDX73wlkJ8E51C_1-fdhRjiU99Bfq1B3o_A027vQGwuSGUTzBY_EfxQ2',
      name: 'Daily Email Digest'
    },
    {
      id: '1Fz5tMNiVa1_6CAGoYITQqrwR6kIVvEhQ4tWKqxaybxlsv-aVrrQfROU-',
      name: 'School Drive Organizer'
    },
    {
      id: '1Ik3H24O4TwoaARnqGoXy1uw2yNXxyQtskSEoOJYX0diCWGfaS-UaDi0A',
      name: 'Deadline Alert System'
    },
    {
      id: '1oR2jpiU9XtAUIV7STkUiPJiw5CGOIuK9a_59Wut85pQNlInJtz_iTu5D',
      name: 'MLA Formatter (Empty — build or delete)'
    }
  ];

  renames.forEach(({ id, name }) => {
    try {
      DriveApp.getFileById(id).setName(name);
      Logger.log('✅ Renamed: ' + name);
    } catch (e) {
      Logger.log('❌ Error renaming "' + name + '": ' + e.message);
    }
  });

  Logger.log('Done! Check script.google.com to confirm names.');
}
