// ==================================================================
// InPrime RO Quiz — Google Sheets backend (Apps Script)
// ------------------------------------------------------------------
// Tabs this script maintains:
//   Summary        – one row per RO: Day 1/2/3 scores and total out of 45
//   Results        – one row per RO per day (every attempt)
//   Answer Details – one row per question answered
//   Quiz Starts    – one row each time someone starts a day
//
// SETUP: open the sheet > Extensions > Apps Script > paste this file >
// Save > run "testSetup" once > Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone) > copy the /exec URL.
// After changing this code: Deploy > Manage deployments > Edit (pencil)
// > Version: New version > Deploy. The URL stays the same.
// ==================================================================

var QUESTIONS_PER_DAY = 15;
var DAYS = 3;
// Which attempt counts in Summary if an RO takes the same day twice:
// 'first' = first attempt counts (recommended), 'best' = highest score counts.
var SUMMARY_RULE = 'first';

var TABS = {
  summary: {
    name: 'Summary',
    headers: ['Employee ID', 'Name', 'Position', 'AM ID', 'AM Name', 'Area Office', 'Region',
              'Day 1', 'Day 2', 'Day 3', 'Total', 'Out Of', 'Days Completed', 'Last Updated']
  },
  results: {
    name: 'Results',
    headers: ['Timestamp', 'Employee ID', 'Name', 'Position', 'AM ID', 'AM Name', 'Area Office',
              'Region', 'Language', 'Day', 'Score', 'Out Of', 'Percentage', 'Time Taken (min)', 'Result']
  },
  answers: {
    name: 'Answer Details',
    headers: ['Timestamp', 'Employee ID', 'Name', 'AM ID', 'Day', 'Question No', 'Question',
              'Selected', 'Correct Answer', 'Is Correct']
  },
  starts: {
    name: 'Quiz Starts',
    headers: ['Timestamp', 'Employee ID', 'Name', 'Position', 'AM ID', 'AM Name', 'Area Office',
              'Region', 'Language', 'Day']
  }
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

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
                 clean(id(data.amId)), clean(data.amName), clean(data.areaOffice), clean(data.region)];

      if (data.type === 'registration') {
        sheet(ss, TABS.starts).appendRow([now].concat(who, [clean(data.language), data.day]));
      }

      if (data.type === 'submission') {
        var score = Number(data.score) || 0;
        var outOf = Number(data.totalQuestions) || QUESTIONS_PER_DAY;
        var band = score >= 12 ? 'Excellent' : (score >= 9 ? 'Good' : 'Needs improvement');

        sheet(ss, TABS.results).appendRow([now].concat(who, [
          clean(data.language), data.day, score, outOf, score / outOf, data.timeTaken, band
        ]));

        var rows = (data.answers || []).map(function (a) {
          return [now, who[0], who[1], who[3], data.day, a.questionNo, clean(a.question),
                  a.selected, a.correct, a.isCorrect ? 'Yes' : 'No'];
        });
        if (rows.length) {
          var s = sheet(ss, TABS.answers);
          s.getRange(s.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        }

        updateSummary(ss, who, Number(data.day), score, now);
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return reply({ status: 'success' });

  } catch (err) {
    return reply({ status: 'error', message: String(err) });
  }
}

// One row per Employee ID. Writes the day's score and recalculates the total.
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
  for (var k = 0; k < 7; k++) row[k] = who[k];          // keep details current
  var col = 6 + day;                                     // Day 1 -> index 7
  var existing = row[col];
  var has = existing !== '' && existing !== null;
  if (!has || (SUMMARY_RULE === 'best' && score > Number(existing))) row[col] = score;
  var total = 0, done = 0;
  for (var d = 0; d < DAYS; d++) {
    var v = row[7 + d];
    if (v !== '' && v !== null) { total += Number(v); done++; }
  }
  row[10] = total;
  row[11] = DAYS * QUESTIONS_PER_DAY;
  row[12] = done;
  row[13] = now;
  if (rowIndex > 0) s.getRange(rowIndex, 1, 1, width).setValues([row]);
  else s.appendRow(row);
}

function doGet() {
  return reply({ status: 'ok', message: 'InPrime RO Quiz API is running' });
}

// Run once from the editor: creates the tabs, checks permissions, adds TEST rows.
function testSetup() {
  var base = { employeeId: 'TEST', name: 'Setup test', position: 'RO', amId: 'AMTEST',
               amName: 'Test AM', areaOffice: 'Test', region: 'South India', language: 'en' };
  var send = function (extra) {
    var obj = {}; for (var k in base) obj[k] = base[k]; for (var j in extra) obj[j] = extra[j];
    obj.submissionId = 'test-' + Math.random();
    return doPost({ postData: { contents: JSON.stringify(obj) } });
  };
  send({ type: 'registration', day: 1 });
  send({ type: 'submission', day: 1, score: 12, totalQuestions: 15, timeTaken: 6.5,
         answers: [{ questionNo: 1, question: 'Test question', selected: 'A', correct: 'A', isCorrect: true }] });
  Logger.log('Done. Check the tabs for TEST rows, then delete those rows.');
}

// ---------- helpers ----------
function sheet(ss, tab) {
  var s = ss.getSheetByName(tab.name);
  if (!s) s = ss.insertSheet(tab.name);
  var n = tab.headers.length;
  var current = s.getLastRow() >= 1 ? s.getRange(1, 1, 1, n).getValues()[0].join('|') : '';
  if (current !== tab.headers.join('|')) {          // create or refresh header row
    s.getRange(1, 1, 1, n).setValues([tab.headers])
      .setFontWeight('bold').setBackground('#6b21a8').setFontColor('#ffffff');
    s.setFrozenRows(1);
    if (tab === TABS.results) s.getRange('M:M').setNumberFormat('0%');
  }
  return s;
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
