// ============================================================
// GOOGLE DRIVE AUTO-ORGANIZER
// Automatically sorts files in any Drive folder into subfolders
// based on keywords, file types, or file age.
// Set it once, run it on a schedule, never sort manually again.
//
// HOW TO USE:
//   1. Go to script.google.com → New project
//   2. Paste this file
//   3. Set SOURCE_FOLDER_ID below (from your Drive folder URL)
//   4. Customize SORT_RULES and EXTENSION_RULES as needed
//   5. Run organizeFiles() with DRY_RUN: true first to preview
//   6. Set DRY_RUN: false, then run setupWeeklyTrigger() to automate
// ============================================================

const CONFIG = {
  // Folder to organize — get the ID from the URL:
  // drive.google.com/drive/folders/THIS_PART_HERE
  SOURCE_FOLDER_ID: 'YOUR_FOLDER_ID_HERE',

  // Move files older than X days into an Archive subfolder (0 = off)
  ARCHIVE_AFTER_DAYS: 90,

  // Preview mode — logs what WOULD happen without moving anything
  DRY_RUN: true,

  // Keyword rules — checks filename for these words (case-insensitive)
  // First match wins. Add, remove, or rename rows as needed.
  SORT_RULES: [
    { keyword: 'invoice',    folder: 'Invoices' },
    { keyword: 'receipt',    folder: 'Receipts' },
    { keyword: 'resume',     folder: 'Resumes' },
    { keyword: 'contract',   folder: 'Contracts' },
    { keyword: 'report',     folder: 'Reports' },
    { keyword: 'photo',      folder: 'Photos' },
    { keyword: 'screenshot', folder: 'Screenshots' },
    { keyword: 'budget',     folder: 'Finance' },
    { keyword: 'tax',        folder: 'Finance' },
  ],

  // File extension rules — applied after keyword rules
  EXTENSION_RULES: [
    { ext: '.pdf',  folder: 'PDFs' },
    { ext: '.xlsx', folder: 'Spreadsheets' },
    { ext: '.csv',  folder: 'Spreadsheets' },
    { ext: '.docx', folder: 'Documents' },
    { ext: '.doc',  folder: 'Documents' },
    { ext: '.pptx', folder: 'Presentations' },
    { ext: '.mp4',  folder: 'Videos' },
    { ext: '.mp3',  folder: 'Audio' },
    { ext: '.zip',  folder: 'Archives' },
  ],
};

function organizeFiles() {
  const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
  const files = sourceFolder.getFiles();
  let moved = 0, skipped = 0;

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    const ageDays = (new Date() - file.getDateCreated()) / (1000 * 60 * 60 * 24);
    let targetFolderName = null;

    // Archive rule
    if (CONFIG.ARCHIVE_AFTER_DAYS > 0 && ageDays > CONFIG.ARCHIVE_AFTER_DAYS) {
      targetFolderName = 'Archive';
    }

    // Keyword rules
    if (!targetFolderName) {
      for (const rule of CONFIG.SORT_RULES) {
        if (name.includes(rule.keyword.toLowerCase())) {
          targetFolderName = rule.folder;
          break;
        }
      }
    }

    // Extension rules
    if (!targetFolderName) {
      for (const rule of CONFIG.EXTENSION_RULES) {
        if (name.endsWith(rule.ext.toLowerCase())) {
          targetFolderName = rule.folder;
          break;
        }
      }
    }

    if (targetFolderName) {
      if (CONFIG.DRY_RUN) {
        Logger.log('[PREVIEW] "' + file.getName() + '" → ' + targetFolderName);
      } else {
        const dest = getOrCreateSubfolder(sourceFolder, targetFolderName);
        file.moveTo(dest);
        Logger.log('Moved: "' + file.getName() + '" → ' + targetFolderName);
      }
      moved++;
    } else {
      skipped++;
    }
  }

  Logger.log('─────────────────────────────');
  Logger.log((CONFIG.DRY_RUN ? '[PREVIEW] ' : '') + moved + ' files processed, ' + skipped + ' skipped (no rule matched).');
  if (CONFIG.DRY_RUN) Logger.log('Set DRY_RUN: false in CONFIG to apply changes for real.');
}

function getOrCreateSubfolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  return existing.hasNext() ? existing.next() : parent.createFolder(name);
}

function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('organizeFiles')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();
  Logger.log('✅ Auto-organizer scheduled for Sundays at 2am. Running a preview now...');
  organizeFiles();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('All triggers removed.');
}
