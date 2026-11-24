// ==================================================================
// InPrime RO Quiz — Google Sheets backend (Apps Script)
// ------------------------------------------------------------------
// SETUP (one time, about 3 minutes)
//  1. Create a new Google Sheet, e.g. "InPrime RO Quiz – Responses".
//  2. In that sheet: Extensions > Apps Script.
//  3. Delete the sample code, paste this whole file, click Save.
//  4. Run the function "testSetup" once (select it in the toolbar, click Run)
//     and allow the permissions it asks for. Check the sheet: you should
//     see three tabs with a TEST row in each.
//  5. Deploy > New deployment > gear icon > Web app
//       Execute as:      Me
//       Who has access:  Anyone
//     Click Deploy and copy the Web app URL (ends with /exec).
//  6. Paste that URL into SHEETS_API_URL at the top of index.html.
//
// If you change this code later: Deploy > Manage deployments > Edit
// (pencil) > Version: New version > Deploy. The URL stays the same.
// ==================================================================

var TABS = {
  starts: {
    name: 'Quiz Starts',
    headers: ['Timestamp', 'Employee ID', 'Name', 'Position', 'Area Office', 'Region', 'Language', 'Day']
  },
  results: {
    name: 'Results',
    headers: ['Timestamp', 'Employee ID', 'Name', 'Position', 'Area Office', 'Region', 'Language',
              'Day', 'Score', 'Out Of', 'Percentage', 'Time Taken (min)', 'Result']
  },
  answers: {
    name: 'Answer Details',
    headers: ['Timestamp', 'Employee ID', 'Name', 'Day', 'Question No', 'Selected', 'Correct', 'Is Correct']
  }
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Ignore exact repeats of the same submission (browser retries).
    var cache = CacheService.getScriptCache();
    if (data.submissionId) {
      if (cache.get(data.submissionId)) return reply({ status: 'duplicate' });
      cache.put(data.submissionId, '1', 21600); // remember for 6 hours
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var now = new Date();

      if (data.type === 'registration') {
        sheet(ss, TABS.starts).appendRow([
          now, clean(data.employeeId), clean(data.name), clean(data.position),
          clean(data.areaOffice), clean(data.region), clean(data.language), data.day
        ]);
      }

      if (data.type === 'submission') {
        var pct = Number(data.percentage) || 0;
        var score = Number(data.score) || 0;
        var band = score >= 12 ? 'Excellent' : (score >= 9 ? 'Good' : 'Needs improvement');
        sheet(ss, TABS.results).appendRow([
          now, clean(data.employeeId), clean(data.name), clean(data.position),
          clean(data.areaOffice), clean(data.region), clean(data.language),
          data.day, score, data.totalQuestions, pct / 100, data.timeTaken, band
        ]);

        var rows = (data.answers || []).map(function (a) {
          return [now, clean(data.employeeId), clean(data.name), data.day,
                  a.questionNo, a.selected, a.correct, a.isCorrect ? 'Yes' : 'No'];
        });
        if (rows.length) {
          var s = sheet(ss, TABS.answers);
          s.getRange(s.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        }
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

function doGet() {
  return reply({ status: 'ok', message: 'InPrime RO Quiz API is running' });
}

// Run this once from the editor to create the tabs and check permissions.
function testSetup() {
  var fake = function (obj) { return { postData: { contents: JSON.stringify(obj) } }; };
  doPost(fake({ type: 'registration', submissionId: 'test-' + Date.now(), employeeId: 'TEST',
                name: 'Setup test', position: 'RO', areaOffice: 'Test', region: 'South India',
                language: 'en', day: 1 }));
  doPost(fake({ type: 'submission', submissionId: 'test2-' + Date.now(), employeeId: 'TEST',
                name: 'Setup test', position: 'RO', areaOffice: 'Test', region: 'South India',
                language: 'en', day: 1, score: 12, totalQuestions: 15, percentage: 80,
                timeTaken: 6.5, answers: [{ questionNo: 1, selected: 'A', correct: 'A', isCorrect: true }] }));
  Logger.log('Done. Check the three tabs in your sheet for TEST rows, then delete them.');
}

// ---------- helpers ----------
function sheet(ss, tab) {
  var s = ss.getSheetByName(tab.name);
  if (!s) {
    s = ss.insertSheet(tab.name);
    s.appendRow(tab.headers);
    s.getRange(1, 1, 1, tab.headers.length)
      .setFontWeight('bold').setBackground('#6b21a8').setFontColor('#ffffff');
    s.setFrozenRows(1);
    if (tab === TABS.results) s.getRange('K:K').setNumberFormat('0%');
  }
  return s;
}

// Stops typed text from being treated as a formula in the sheet.
function clean(v) {
  v = (v === undefined || v === null) ? '' : String(v).slice(0, 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
