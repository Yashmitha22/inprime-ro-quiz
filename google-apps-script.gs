// ==================================================================
// InPrime RO Quiz — Google Sheets backend (Apps Script)
// ------------------------------------------------------------------
// Tabs:
//   Responses – one row per RO per day: details, Q1–Q15 answers, total marks
//   Summary   – one row per RO: Day 1, Day 2, Day 3 marks and total out of 45
//
// SETUP: open the sheet > Extensions > Apps Script > paste this file >
// Save > run "testSetup" once > delete the TEST rows >
// Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
// (First time only: Deploy > New deployment > Web app, Execute as: Me,
//  Who has access: Anyone.)  The /exec URL stays the same after updates.
// ==================================================================

var QUESTIONS_PER_DAY = 15;
var DAYS = 3;
// If an RO takes the same day twice, which marks go into Summary:
// 'first' = first attempt (recommended), 'best' = highest marks.
var SUMMARY_RULE = 'first';

var GREEN = '#dcfce7', RED = '#fee2e2';

var TABS = {
  responses: {
    name: 'Responses',
    headers: (function () {
      var h = ['Timestamp', 'Employee ID', 'Full Name', 'Position', 'Area Office', 'Region', 'Day'];
      for (var i = 1; i <= QUESTIONS_PER_DAY; i++) h.push('Q' + i);
      h.push('Total Marks');
      return h;
    })()
  },
  summary: {
    name: 'Summary',
    headers: ['Employee ID', 'Full Name', 'Position', 'Area Office', 'Region',
              'Day 1', 'Day 2', 'Day 3', 'Total Marks', 'Out Of', 'Days Completed', 'Last Updated']
  }
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Only finished quizzes are recorded.
    if (data.type !== 'submission') return reply({ status: 'success' });

    // Ignore exact repeats of the same submission (browser retries).
    var cache = CacheService.getScriptCache();
    if (data.submissionId) {
      if (cache.get(data.submissionId)) return reply({ status: 'duplicate' });
      cache.put(data.submissionId, '1', 21600);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var now = new Date();
      var who = [clean(id(data.employeeId)), clean(data.name), clean(data.position),
                 clean(data.areaOffice), clean(data.region)];
      var day = Number(data.day);
      var score = Number(data.score) || 0;

      // Q1..Q15: the answer picked, e.g. "A. 15%"
      var byNo = {};
      (data.answers || []).forEach(function (a) { byNo[a.questionNo] = a; });
      var answers = [], colours = [];
      for (var q = 1; q <= QUESTIONS_PER_DAY; q++) {
        var a = byNo[q];
        if (!a || a.selected === 'Not answered') { answers.push('Not answered'); colours.push(RED); continue; }
        answers.push(clean(a.selected + (a.selectedText ? '. ' + a.selectedText : '')));
        colours.push(a.isCorrect ? GREEN : RED);
      }

      var s = sheet(ss, TABS.responses);
      var r = s.getLastRow() + 1;
      s.getRange(r, 1, 1, TABS.responses.headers.length)
        .setValues([[now].concat(who, [day], answers, [score])]);
      s.getRange(r, 8, 1, QUESTIONS_PER_DAY).setBackgrounds([colours]);

      updateSummary(ss, who, day, score, now);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return reply({ status: 'success' });

  } catch (err) {
    return reply({ status: 'error', message: String(err) });
  }
}

// One row per Employee ID: writes that day's marks and recalculates the total.
function updateSummary(ss, who, day, score, now) {
  if (!(day >= 1 && day <= DAYS)) return;
  var s = sheet(ss, TABS.summary);
  var key = String(who[0]).trim().toUpperCase();
  var last = s.getLastRow();
  var rowIndex = -1;
  if (last >= 2) {
    var ids = s.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim().toUpperCase() === key) { rowIndex = i + 2; break; }
    }
  }
  var width = TABS.summary.headers.length;
  var row = rowIndex > 0 ? s.getRange(rowIndex, 1, 1, width).getValues()[0]
                         : new Array(width).fill('');
  for (var k = 0; k < 5; k++) row[k] = who[k];        // keep details current
  var col = 4 + day;                                   // Day 1 -> column F
  var existing = row[col];
  var has = existing !== '' && existing !== null;
  if (!has || (SUMMARY_RULE === 'best' && score > Number(existing))) row[col] = score;
  var total = 0, done = 0;
  for (var d = 0; d < DAYS; d++) {
    var v = row[5 + d];
    if (v !== '' && v !== null) { total += Number(v); done++; }
  }
  row[8] = total;
  row[9] = DAYS * QUESTIONS_PER_DAY;
  row[10] = done;
  row[11] = now;
  if (rowIndex > 0) s.getRange(rowIndex, 1, 1, width).setValues([row]);
  else s.appendRow(row);
}

function doGet() {
  return reply({ status: 'ok', message: 'InPrime RO Quiz API is running' });
}

// Run once from the editor: builds the two tabs, removes the old test-only
// tabs from the earlier version, and adds a TEST row to check it works.
function testSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Results', 'Answer Details', 'Quiz Starts'].forEach(function (name) {
    var old = ss.getSheetByName(name);
    if (old && onlyTestRows(old)) ss.deleteSheet(old);
  });
  var oldSummary = ss.getSheetByName('Summary');
  if (oldSummary && onlyTestRows(oldSummary) && oldSummary.getLastRow() > 1) {
    oldSummary.deleteRows(2, oldSummary.getLastRow() - 1);
  }
  var answers = [];
  for (var i = 1; i <= QUESTIONS_PER_DAY; i++) {
    answers.push({ questionNo: i, selected: i % 4 ? 'A' : 'B', selectedText: 'Sample',
                   correct: 'A', isCorrect: i % 4 !== 0 });
  }
  doPost({ postData: { contents: JSON.stringify({
    type: 'submission', submissionId: 'test-' + Math.random(), employeeId: 'TEST',
    name: 'Setup test', position: 'RO', areaOffice: 'Test', region: 'South India',
    day: 1, score: 11, totalQuestions: QUESTIONS_PER_DAY, answers: answers }) } });
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  Logger.log('Done. Check the Responses and Summary tabs, then delete the TEST rows.');
}

// ---------- helpers ----------
function sheet(ss, tab) {
  var s = ss.getSheetByName(tab.name);
  if (!s) s = ss.insertSheet(tab.name);
  var n = tab.headers.length;
  if (s.getMaxColumns() < n) s.insertColumnsAfter(s.getMaxColumns(), n - s.getMaxColumns());
  var current = s.getLastRow() >= 1 ? s.getRange(1, 1, 1, n).getValues()[0].join('|') : '';
  if (current !== tab.headers.join('|')) {            // create or refresh header row
    s.getRange(1, 1, 1, n).setValues([tab.headers])
      .setFontWeight('bold').setBackground('#6b21a8').setFontColor('#ffffff');
    var extra = s.getMaxColumns() - n;                 // clear old headings past the end
    if (extra > 0) s.getRange(1, n + 1, 1, extra).clearContent().clearFormat();
    s.setFrozenRows(1);
    s.setFrozenColumns(3);
  }
  return s;
}

// True if every data row in the tab belongs to the TEST employee.
function onlyTestRows(s) {
  var last = s.getLastRow();
  if (last < 2) return true;
  var vals = s.getRange(2, 1, last - 1, Math.min(3, s.getMaxColumns())).getValues();
  return vals.every(function (r) { return r.join('|').indexOf('TEST') >= 0 || r.join('') === ''; });
}

// Employee IDs: trimmed and upper-case so the same person always matches.
function id(v) { return String(v === undefined || v === null ? '' : v).trim().toUpperCase(); }

// Stops typed text from being treated as a formula in the sheet.
function clean(v) {
  v = (v === undefined || v === null) ? '' : String(v).slice(0, 300);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
