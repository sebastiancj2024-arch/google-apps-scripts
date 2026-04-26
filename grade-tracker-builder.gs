// ============================================================
// GRADE TRACKER BUILDER
// Run createGradeTracker() once to generate a fully formatted
// Grade Tracker spreadsheet — formulas, colour coding, and
// category breakdowns all included.
//
// HOW TO USE:
//   1. script.google.com → New project
//   2. Paste this file
//   3. Edit TRACKER_CONFIG below (subjects, weights, target %)
//   4. Click Run → createGradeTracker
//   5. Authorize → find the new spreadsheet in your Drive
// ============================================================

const TRACKER_CONFIG = {
  studentName: 'Your Name',
  school:      'Your School',
  semester:    'Semester 1',
  year:        '2025–2026',

  // Rows available per category for entering individual grades
  rowsPerCategory: 8,

  // Add, remove, or rename subjects freely.
  // Weights inside each subject MUST add up to 100.
  subjects: [
    {
      name: 'Science',
      target: 80,
      categories: [
        { name: 'Tests & Exams',   weight: 40 },
        { name: 'Labs & Projects', weight: 30 },
        { name: 'Assignments',     weight: 20 },
        { name: 'Participation',   weight: 10 },
      ],
    },
    {
      name: 'English',
      target: 80,
      categories: [
        { name: 'Essays & Writing', weight: 40 },
        { name: 'Tests & Exams',    weight: 30 },
        { name: 'Assignments',      weight: 20 },
        { name: 'Participation',    weight: 10 },
      ],
    },
    {
      name: 'History',
      target: 80,
      categories: [
        { name: 'Tests & Exams', weight: 40 },
        { name: 'Essays',        weight: 30 },
        { name: 'Assignments',   weight: 20 },
        { name: 'Participation', weight: 10 },
      ],
    },
    {
      name: 'Business',
      target: 80,
      categories: [
        { name: 'Tests & Exams', weight: 35 },
        { name: 'Projects',      weight: 35 },
        { name: 'Assignments',   weight: 20 },
        { name: 'Participation', weight: 10 },
      ],
    },
  ],

  colors: {
    headerDark:  '#1a1a2e',
    headerMid:   '#0f3460',
    rowAlt:      '#f8f9fa',
    catAlt:      '#e8f0fe',
    white:       '#ffffff',
    green:       '#34a853',
    yellow:      '#fbbc04',
    red:         '#ea4335',
    textLight:   '#ffffff',
    textMuted:   '#888888',
  },
};

// ─────────────────────────────────────────────────────────────
// Entry point
// ─────────────────────────────────────────────────────────────

function createGradeTracker() {
  const cfg = TRACKER_CONFIG;
  const ss = SpreadsheetApp.create(cfg.studentName + ' — Grade Tracker ' + cfg.year);

  const dash = ss.getActiveSheet();
  dash.setName('📊 Dashboard');

  const subjectMeta = cfg.subjects.map(subject => {
    const sheet = ss.insertSheet(subject.name);
    return buildSubjectSheet(sheet, subject, cfg);
  });

  buildDashboard(dash, cfg, subjectMeta);

  const blank = ss.getSheetByName('Sheet1');
  if (blank) ss.deleteSheet(blank);

  ss.setActiveSheet(dash);

  Logger.log('✅ Grade Tracker created: ' + ss.getUrl());
  SpreadsheetApp.getUi().alert(
    '✅ Done!\n\n' +
    '1. Open each subject tab\n' +
    '2. Enter grades in the white rows (column B)\n' +
    '3. Dashboard updates automatically'
  );
}

// ─────────────────────────────────────────────────────────────
// Dashboard
// ─────────────────────────────────────────────────────────────

function buildDashboard(sheet, cfg, subjectMeta) {
  const C = cfg.colors;

  sheet.setColumnWidths(1, 6, 1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 190);

  // Title
  sheet.setRowHeight(1, 52);
  const titleCell = sheet.getRange('A1:E1');
  titleCell.merge()
    .setValue('📊  Grade Tracker — ' + cfg.semester + '  ' + cfg.year)
    .setBackground(C.headerDark)
    .setFontColor(C.textLight)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Student info
  sheet.setRowHeight(2, 28);
  sheet.getRange('A2:E2').merge()
    .setValue(cfg.studentName + '   ·   ' + cfg.school)
    .setBackground(C.headerMid)
    .setFontColor(C.catAlt)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(11);

  // Spacer
  sheet.setRowHeight(3, 14);

  // Column headers
  sheet.setRowHeight(4, 28);
  [['A4','Subject'], ['B4','Target'], ['C4','Current'], ['D4','Status'], ['E4','Notes']].forEach(([addr, label]) => {
    sheet.getRange(addr)
      .setValue(label)
      .setBackground(C.headerMid)
      .setFontColor(C.textLight)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  // Subject rows
  subjectMeta.forEach((meta, i) => {
    const row = 5 + i;
    sheet.setRowHeight(row, 26);
    const bg = i % 2 === 0 ? C.white : C.rowAlt;
    const ref = "'" + meta.name + "'!" + meta.weightedAvgCell;

    sheet.getRange(row, 1).setValue(meta.name).setFontWeight('bold').setBackground(bg);

    sheet.getRange(row, 2)
      .setValue(meta.target / 100)
      .setNumberFormat('0%')
      .setHorizontalAlignment('center')
      .setBackground(bg);

    sheet.getRange(row, 3)
      .setFormula('=IFERROR(' + ref + '/100,"—")')
      .setNumberFormat('0.0%')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(bg);

    sheet.getRange(row, 4)
      .setFormula(
        '=IFERROR(IF(' + ref + '>=' + meta.target + ',"✅ On Track",' +
        'IF(' + ref + '>=' + (meta.target - 5) + ',"⚠️ Almost","❌ Below Target")),"—")'
      )
      .setHorizontalAlignment('center')
      .setBackground(bg);

    sheet.getRange(row, 5).setBackground(bg).setFontColor(C.textMuted)
      .setValue('← open subject tab to enter grades');
  });

  // Overall average
  const avgRow = 5 + subjectMeta.length + 1;
  sheet.setRowHeight(avgRow, 32);
  const cCells = subjectMeta.map((_, i) => 'C' + (5 + i)).join(',');
  sheet.getRange(avgRow, 1, 1, 5).setBackground(C.headerMid).setFontColor(C.textLight);
  sheet.getRange(avgRow, 1).setValue('Overall Average').setFontWeight('bold').setBackground(C.headerMid).setFontColor(C.textLight);
  sheet.getRange(avgRow, 3)
    .setFormula('=IFERROR(AVERAGE(' + cCells + '),"—")')
    .setNumberFormat('0.0%')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(C.headerMid)
    .setFontColor(C.textLight)
    .setFontSize(13);

  // Conditional formatting — current grade column
  const cfRules = [];
  subjectMeta.forEach((meta, i) => {
    const cell = sheet.getRange(5 + i, 3);
    cfRules.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(meta.target / 100)
        .setBackground('#e6f4ea').setFontColor(C.green).setRanges([cell]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo((meta.target - 5) / 100)
        .setBackground('#fef9e7').setFontColor('#b5870a').setRanges([cell]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan((meta.target - 5) / 100)
        .setBackground('#fce8e6').setFontColor(C.red).setRanges([cell]).build()
    );
  });
  sheet.setConditionalFormatRules(cfRules);

  sheet.setFrozenRows(4);
  sheet.hideGridlines();
}

// ─────────────────────────────────────────────────────────────
// Subject sheet
// ─────────────────────────────────────────────────────────────

function buildSubjectSheet(sheet, subject, cfg) {
  const C = cfg.colors;
  const rpc = cfg.rowsPerCategory;

  sheet.setColumnWidth(1, 220); // Assessment description
  sheet.setColumnWidth(2, 90);  // Grade
  sheet.setColumnWidth(3, 120); // Notes
  sheet.setColumnWidth(4, 24);  // Spacer
  sheet.setColumnWidth(5, 160); // Summary label
  sheet.setColumnWidth(6, 100); // Summary value

  // Title
  sheet.setRowHeight(1, 44);
  sheet.getRange('A1:F1').merge()
    .setValue(subject.name + ' — Grade Breakdown')
    .setBackground(C.headerDark)
    .setFontColor(C.textLight)
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Column headers
  sheet.setRowHeight(2, 24);
  [['A2','Assessment'], ['B2','Grade (%)'], ['C2','Notes'], ['E2','Summary'], ['F2','']].forEach(([addr, label]) => {
    sheet.getRange(addr)
      .setValue(label)
      .setBackground(C.headerMid)
      .setFontColor(C.textLight)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
  sheet.getRange('D2').setBackground(C.headerMid);

  let row = 3;
  const catAvgCells = [];

  subject.categories.forEach((cat, ci) => {
    // Category header
    sheet.setRowHeight(row, 26);
    sheet.getRange(row, 1, 1, 3).merge()
      .setValue(cat.name + '  ·  Weight: ' + cat.weight + '%')
      .setBackground(C.headerMid)
      .setFontColor(C.textLight)
      .setFontWeight('bold');
    row++;

    const gradeStart = row;
    for (let r = 0; r < rpc; r++) {
      sheet.setRowHeight(row, 22);
      const bg = r % 2 === 0 ? C.white : C.rowAlt;
      sheet.getRange(row, 1).setBackground(bg);
      sheet.getRange(row, 2)
        .setBackground(bg)
        .setNumberFormat('0.0')
        .setHorizontalAlignment('center');
      sheet.getRange(row, 3)
        .setBackground(bg)
        .setFontColor(C.textMuted)
        .setFontSize(9);
      row++;
    }
    const gradeEnd = row - 1;

    // Category average
    sheet.setRowHeight(row, 24);
    sheet.getRange(row, 1).setValue(cat.name + ' Average').setFontWeight('bold').setBackground(C.catAlt);
    const avgCell = sheet.getRange(row, 2);
    avgCell.setFormula('=IFERROR(AVERAGE(B' + gradeStart + ':B' + gradeEnd + '),"—")')
      .setNumberFormat('0.0')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(C.catAlt);
    sheet.getRange(row, 3).setBackground(C.catAlt);

    catAvgCells.push({ addr: 'B' + row, weight: cat.weight });
    row += 2;
  });

  // ── Summary panel (column E-F) ───────────────────────────
  const totalWeight = subject.categories.reduce((s, c) => s + c.weight, 0);
  const weightedFormula = catAvgCells
    .map(a => 'IFERROR(' + a.addr + ',0)*' + a.weight)
    .join('+');

  const weightedAvgCell = 'F3';

  // Current average
  sheet.getRange('E3').setValue('Current Average').setFontWeight('bold');
  const wAvg = sheet.getRange('F3');
  wAvg.setFormula('=IFERROR((' + weightedFormula + ')/' + totalWeight + ',"—")')
    .setNumberFormat('0.0')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(C.catAlt)
    .setFontSize(14);

  // Target
  sheet.getRange('E4').setValue('Target').setFontWeight('bold');
  sheet.getRange('F4').setValue(subject.target).setNumberFormat('0.0').setHorizontalAlignment('center');

  // Status
  sheet.getRange('E5').setValue('Status').setFontWeight('bold');
  sheet.getRange('F5')
    .setFormula(
      '=IFERROR(IF(F3>=' + subject.target + ',"✅ On Track",' +
      'IF(F3>=' + (subject.target - 5) + ',"⚠️ Almost","❌ Below Target")),"—")'
    )
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // Category breakdown
  sheet.getRange('E7').setValue('Category Breakdown').setFontWeight('bold').setFontColor('#555555');
  catAvgCells.forEach((a, i) => {
    const cat = subject.categories[i];
    sheet.getRange(8 + i, 5).setValue(cat.name + ' (' + cat.weight + '%)').setFontSize(9);
    sheet.getRange(8 + i, 6)
      .setFormula('=' + a.addr)
      .setNumberFormat('0.0')
      .setHorizontalAlignment('center')
      .setFontSize(9);
  });

  sheet.setFrozenRows(2);
  sheet.hideGridlines();

  return { name: subject.name, target: subject.target, weightedAvgCell };
}
