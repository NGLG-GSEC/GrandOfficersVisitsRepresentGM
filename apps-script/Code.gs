/**
 * Backend για τη Δήλωση Διάθεσης Εκπροσώπησης Μ.Δ.
 * Αποθηκεύει τις δηλώσεις σε Google Sheet και επιστρέφει την αναφορά
 * μόνο όταν δοθεί σωστός κωδικός (ADMIN_KEY στις Script Properties).
 *
 * Οδηγίες εγκατάστασης: δείτε README.md
 */

const HEADERS = ['Χρονοσήμανση', 'Αξιωματικός', 'Κωδικός Εργασίας', 'Α/Α Προβαδίσματος', 'Ακυρώθηκε'];
const COL = { time: 0, officer: 1, meeting: 2, rank: 3, cancelled: 4 };

// Μορφή κωδικού Εργασίας (π.χ. "atlas_118"). Οι Εργασίες ορίζονται μόνο
// στο index.html, ώστε η προσθήκη νέων να μη χρειάζεται νέα ανάπτυξη.
const MEETING_ID_PATTERN = /^[a-z0-9_]{1,40}$/;
const MAX_MEETINGS = 50;

const MAX_OFFICER_LENGTH = 200;
const MAX_RANK = 10000;

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  if (body.action === 'submit') return submit_(body);
  if (body.action === 'cancel') return cancel_(body);
  if (body.action === 'report') return report_(body);
  return json_({ ok: false, error: 'unknown_action' });
}

// Κοινός έλεγχος για υποβολή και ακύρωση
function parseRequest_(body) {
  const officer = String(body.officer || '').replace(/\s+/g, ' ').trim();
  const meetingIds = Array.isArray(body.meetings) ? body.meetings : [];
  const meetings = meetingIds
    .filter(id => typeof id === 'string' && MEETING_ID_PATTERN.test(id))
    .filter((id, i, arr) => arr.indexOf(id) === i)
    .slice(0, MAX_MEETINGS);

  if (!officer || officer.length > MAX_OFFICER_LENGTH) return { error: 'invalid_officer' };
  if (meetings.length === 0) return { error: 'invalid_meetings' };
  return { officer: officer, meetings: meetings };
}

function submit_(body) {
  const req = parseRequest_(body);
  if (req.error) return json_({ ok: false, error: req.error });
  const rank = parseInt(body.rank, 10);
  const rankCell = rank > 0 && rank <= MAX_RANK ? rank : '';

  return withLock_(() => {
    const sheet = getSheet_();
    const existing = new Set(
      readRows_(sheet).filter(r => !r.cancelled)
        .map(r => normalize_(r.officer) + '|' + r.meetingId)
    );
    const now = new Date();
    const toAdd = [];
    req.meetings.forEach(id => {
      const key = normalize_(req.officer) + '|' + id;
      if (!existing.has(key)) {
        existing.add(key);
        toAdd.push([now, req.officer, id, rankCell, '']);
      }
    });
    if (toAdd.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, HEADERS.length).setValues(toAdd);
    }
    return json_({ ok: true, added: toAdd.length, duplicates: req.meetings.length - toAdd.length });
  });
}

// Η ακύρωση δεν σβήνει τη γραμμή: σημειώνει την ώρα ακύρωσης στη στήλη «Ακυρώθηκε»
function cancel_(body) {
  const req = parseRequest_(body);
  if (req.error) return json_({ ok: false, error: req.error });

  return withLock_(() => {
    const sheet = getSheet_();
    const who = normalize_(req.officer);
    const now = new Date();
    let cancelled = 0;
    readRows_(sheet).forEach(r => {
      if (!r.cancelled && normalize_(r.officer) === who && req.meetings.indexOf(r.meetingId) !== -1) {
        sheet.getRange(r.row, COL.cancelled + 1).setValue(now);
        cancelled++;
      }
    });
    return json_({ ok: true, cancelled: cancelled });
  });
}

function report_(body) {
  const adminKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!adminKey || String(body.key || '') !== adminKey) {
    Utilities.sleep(1000); // επιβράδυνση δοκιμών κωδικού
    return json_({ ok: false, error: 'unauthorized' });
  }
  const byMeeting = {};
  readRows_(getSheet_()).filter(r => !r.cancelled).forEach(r => {
    (byMeeting[r.meetingId] = byMeeting[r.meetingId] || []).push({ officer: r.officer, rank: r.rank });
  });
  return json_({ ok: true, data: byMeeting });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  // Χρησιμοποιεί το πρώτο φύλλο του Sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  if (header.getValues()[0].join('|') !== HEADERS.join('|')) header.setValues([HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .map((r, i) => ({
      row: i + 2,
      officer: String(r[COL.officer]),
      meetingId: String(r[COL.meeting]),
      rank: r[COL.rank] === '' ? null : Number(r[COL.rank]),
      cancelled: r[COL.cancelled] !== ''
    }))
    .filter(r => r.officer && r.meetingId);
}

// Για τον εντοπισμό διπλοεγγραφών: χωρίς τόνους, πεζά, ενιαία κενά
function normalize_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
