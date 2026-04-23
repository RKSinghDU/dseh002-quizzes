/**
 * ================================================================
 * DSEH002 — Full Syllabus Quiz Logger  v2.0
 * BOUND SCRIPT — paste via Extensions → Apps Script
 * ================================================================
 * Mirrors COMC004 Multi-Quiz Logger v4.0 exactly.
 * Functions: createAllTabs, createTab, authorizeMail,
 *            doPost, writeScoreToNamesTab, sendFeedbackEmail,
 *            lookupStudent, doGet
 *
 * SETUP:
 *  1. Open spreadsheet → Extensions → Apps Script
 *  2. Paste this file, replacing ALL existing content
 *  3. Run authorizeMail() once
 *  4. Run createAllTabs() once
 *  5. Deploy → New Deployment → Web App
 *       Execute as: Me  |  Who has access: Anyone
 *  6. Copy new URL → paste as LOGGER_URL in quiz HTML
 * ================================================================
 */

// ── Quiz routing table (mirrors COMC004 QUIZ_CONFIG) ────────────
var QUIZ_CONFIG = {
  'DSEH002_FullSyllabus': {
    tab:       'DSEH002_FullSyllabus',
    scoreCol:  2,          // col B in Student_Roster for score
    maxMarks:  30,
    questions: 15
  }
};

var ROSTER_TAB = 'Student_Roster';   // col A = Name, col C = Roll No
var NAMES_TAB  = 'Student_Roster';   // same tab — writeScoreToNamesTab writes score to col B

var HEADERS = [
  'Timestamp', 'Submitted Name', 'Submitted Email',
  'Matched Name', 'Roll No',
  'Score (/30)', 'Percentage (%)', 'Grade Band',
  'Time Taken (s)', 'Auto-Submit',
  'Q1 Chosen',  'Q1 Correct',
  'Q2 Chosen',  'Q2 Correct',
  'Q3 Chosen',  'Q3 Correct',
  'Q4 Chosen',  'Q4 Correct',
  'Q5 Chosen',  'Q5 Correct',
  'Q6 Chosen',  'Q6 Correct',
  'Q7 Chosen',  'Q7 Correct',
  'Q8 Chosen',  'Q8 Correct',
  'Q9 Chosen',  'Q9 Correct',
  'Q10 Chosen', 'Q10 Correct',
  'Q11 Chosen', 'Q11 Correct',
  'Q12 Chosen', 'Q12 Correct',
  'Q13 Chosen', 'Q13 Correct',
  'Q14 Chosen', 'Q14 Correct',
  'Q15 Chosen', 'Q15 Correct',
  'Email Sent', 'Match Status'
];

// ── Run once to create all quiz tabs ────────────────────────────
function createAllTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var quiz in QUIZ_CONFIG) {
    createTab(ss, QUIZ_CONFIG[quiz].tab);
  }
}

// ── Create a single results tab with headers ─────────────────────
function createTab(ss, tabName) {
  if (ss.getSheetByName(tabName)) {
    Logger.log(tabName + ' already exists.');
    return;
  }
  var sheet = ss.insertSheet(tabName);
  var hdr   = sheet.getRange(1, 1, 1, HEADERS.length);
  hdr.setValues([HEADERS]);
  hdr.setBackground('#1a1410');
  hdr.setFontColor('#faf7f2');
  hdr.setFontWeight('bold');
  hdr.setFontFamily('Arial');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 220);
  sheet.autoResizeColumns(4, HEADERS.length - 3);
  Logger.log(tabName + ' tab created.');
}

// ── Authorize mail — run once before deploying ──────────────────
function authorizeMail() {
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    'DSEH002 Logger — Mail Authorized',
    'Mail scope is now authorized. You can deploy as Web App.'
  );
  Logger.log('Mail authorized.');
}

// ── Health check ─────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status:  'live',
      script:  'DSEH002 Full Syllabus Quiz Logger v2.0',
      quizzes: Object.keys(QUIZ_CONFIG)
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Receives every quiz submission ───────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var quizKey = payload.quiz || 'DSEH002_FullSyllabus';
    var cfg     = QUIZ_CONFIG[quizKey] || QUIZ_CONFIG['DSEH002_FullSyllabus'];

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.tab);
    if (!sheet) { createTab(ss, cfg.tab); sheet = ss.getSheetByName(cfg.tab); }

    // Lookup student in roster
    var lookup = lookupStudent(ss, payload.name);

    var total = payload.score   || 0;
    var max   = payload.max     || cfg.maxMarks;
    var pct   = payload.pct     || parseFloat(((total / max) * 100).toFixed(1));
    var band  = pct >= 85 ? 'O — Outstanding'
              : pct >= 70 ? 'A — Excellent'
              : pct >= 55 ? 'B — Satisfactory'
              :              'C — Needs Development';

    // Send feedback email
    var emailSent = 'No';
    if (payload.email && payload.email.indexOf('@') > -1) {
      try {
        sendFeedbackEmail(payload.email, payload.name, quizKey, total, max, pct, band, payload);
        emailSent = 'Yes';
      } catch (mailErr) {
        Logger.log('Mail error: ' + mailErr.message);
        emailSent = 'Error: ' + mailErr.message;
      }
    }

    // Build row
    var row = [
      new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      payload.name       || '',
      payload.email      || '',
      lookup.name        || '',
      lookup.rollNo      || 'NOT FOUND',
      total,
      pct + '%',
      band,
      payload.timeTaken  || 0,
      payload.autoSubmit || 'No'
    ];
    for (var q = 1; q <= cfg.questions; q++) {
      row.push(payload['Q' + q + '_chosen']  || '');
      row.push(payload['Q' + q + '_correct'] || '');
    }
    row.push(emailSent, lookup.status);

    sheet.appendRow(row);

    // Colour score cell by band
    var lastRow   = sheet.getLastRow();
    var bandColor = pct >= 85 ? '#C8E6C9'
                  : pct >= 70 ? '#BBDEFB'
                  : pct >= 55 ? '#FFE0B2'
                  :              '#FFCDD2';
    sheet.getRange(lastRow, 6).setBackground(bandColor);
    if (lookup.status === 'NOT FOUND') {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground('#FFF3E0');
    }

    // Write score back to Student_Roster col B
    writeScoreToNamesTab(ss, lookup.name || payload.name, cfg.scoreCol, total);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', emailSent: emailSent, row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Write score back to Student_Roster ──────────────────────────
function writeScoreToNamesTab(ss, studentName, col, score) {
  try {
    var nSheet = ss.getSheetByName(NAMES_TAB);
    if (!nSheet) return;
    var data  = nSheet.getDataRange().getValues();
    var query = (studentName || '').trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var cell = (data[i][0] || '').toString().trim().toLowerCase();
      if (cell === query || (query && cell.startsWith(query.split(' ')[0]))) {
        nSheet.getRange(i + 1, col).setValue(score);
        break;
      }
    }
  } catch (e) {
    Logger.log('writeScoreToNamesTab error: ' + e.message);
  }
}

// ── Send personalised feedback email ────────────────────────────
function sendFeedbackEmail(email, name, quizKey, total, max, pct, band, payload) {
  var topicMap = {
    1:  'OD Foundations',
    2:  'Force Field Analysis (Lewin)',
    3:  'Schein — Three Models of Consultation',
    4:  'OCTAPACE Culture Framework',
    5:  'Survey Feedback (Likert, 1947)',
    6:  'Johari Window',
    7:  'Argyris — Three Conditions',
    8:  'Hackman–Oldham JCM',
    9:  'STS Theory',
    10: 'Appreciative Inquiry',
    11: 'Positive Deviance',
    12: 'Burke–Litwin Model',
    13: 'ADKAR Model',
    14: 'Cummings & Worley Taxonomy',
    15: 'Gandhian OD Principles'
  };

  var tierMessages = {
    'O — Outstanding':      'Exceptional performance across the full syllabus. Your responses demonstrate integrated, systems-level thinking and a strong command of both Western OD frameworks and Gandhian philosophical foundations.',
    'A — Excellent':        'Strong performance. Revisit causal sequencing in the Burke–Litwin model and distinctions within Cummings & Worley\'s four intervention categories.',
    'B — Satisfactory':     'Core concepts are in place. Move from recognition to application — diagnose scenarios using frameworks rather than recalling definitions. Review OCTAPACE, Argyris\'s three conditions, and AI\'s 4-D cycle.',
    'C — Needs Development':'Revisit foundational frameworks: Force Field Analysis, Schein\'s three models, JCM, and STS theory. Practise applying each to organisational scenarios.'
  };

  var qSummary = '';
  for (var i = 1; i <= 15; i++) {
    var chosen  = payload['Q' + i + '_chosen']  || '';
    var correct = payload['Q' + i + '_correct'] || '';
    var mark    = chosen === correct ? '✓' : '○';
    qSummary += mark + '  Q' + i + ' — ' + (topicMap[i] || '') + '\n';
  }

  var subject = 'DSEH002 Full Syllabus Quiz Result — ' + (name || 'Student');
  var body =
    'Dear ' + (name || 'Student') + ',\n\n' +
    'Your result for the DSEH002 Full Syllabus Assessment has been recorded.\n\n' +
    '════════════════════════════════════════\n' +
    'SCORE:       ' + parseFloat(total).toFixed(2) + ' / ' + max + '  (' + pct + '%)\n' +
    'GRADE BAND:  ' + band + '\n' +
    '════════════════════════════════════════\n\n' +
    'TOPIC-WISE SUMMARY  (✓ full marks  ○ partial / missed):\n' +
    qSummary + '\n' +
    'FEEDBACK:\n' + (tierMessages[band] || '') + '\n\n' +
    'Your response has been saved to the course gradebook.\n\n' +
    'Best regards,\n' +
    'Prof. R K Singh\n' +
    'Department of Commerce, University of Delhi\n' +
    'DSEH002 — Organizational Development and Change Management\n\n' +
    '─────────────────────────────────────────\n' +
    'PRIVATE CIRCULATION ONLY · AY 2025–26';

  MailApp.sendEmail(email, subject, body);
}

// ── Lookup student in Student_Roster ─────────────────────────────
function lookupStudent(ss, submittedName) {
  if (!submittedName)
    return { name: '', rollNo: '', status: 'NO NAME PROVIDED' };

  var rosterSheet = ss.getSheetByName(ROSTER_TAB);
  if (!rosterSheet)
    return { name: submittedName, rollNo: '', status: 'NO ROSTER TAB' };

  var data      = rosterSheet.getDataRange().getValues();
  var query     = submittedName.trim().toLowerCase();
  var firstName = query.split(' ')[0];
  var exact     = null;
  var firstOnly = null;

  // Student_Roster: col A = Name, col C = Roll No
  for (var i = 1; i < data.length; i++) {
    var cell = (data[i][0] || '').toString().trim().toLowerCase();
    if (!cell) continue;
    if (cell === query) {
      exact = { name: data[i][0], rollNo: data[i][2] || '' };
      break;
    }
    if (!firstOnly && cell.startsWith(firstName)) {
      firstOnly = { name: data[i][0], rollNo: data[i][2] || '' };
    }
  }

  if (exact)     return Object.assign(exact,     { status: 'EXACT MATCH' });
  if (firstOnly) return Object.assign(firstOnly, { status: 'FIRST-NAME MATCH — verify' });
  return { name: '', rollNo: '', status: 'NOT FOUND' };
}
