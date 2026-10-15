/**
 * Backend για τη Δήλωση Διάθεσης Εκπροσώπησης Μ.Δ.
 * Αποθηκεύει τις δηλώσεις σε Google Sheet και επιστρέφει την αναφορά
 * μόνο όταν δοθεί σωστός κωδικός (ADMIN_KEY στις Script Properties).
 *
 * Οδηγίες εγκατάστασης: δείτε README.md
 */

const SHEET_NAME = 'Δηλώσεις';
const HEADERS = ['Χρονοσήμανση', 'Αξιωματικός', 'Κωδικός Εργασίας'];

// Πρέπει να συμφωνεί με τη λίστα MEETINGS στο index.html
const VALID_MEETING_IDS = [
  'akropolis_104', 'epimelites_74', 'isotis_97', 'triptolemos_19',
  'enosis_16', 'athina_78', 'lapaix_95', 'omiros_59',
  'makedon_106', 'nagos_111', 'empedoklis_15'
];

const MAX_OFFICER_LENGTH = 200;

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  if (body.action === 'submit') return submit_(body);
  if (body.action === 'report') return report_(body);
  return json_({ ok: false, error: 'unknown_action' });
}

function submit_(body) {
  const officer = String(body.officer || '').replace(/\s+/g, ' ').trim();
  const meetingIds = Array.isArray(body.meetings) ? body.meetings : [];
  const valid = meetingIds.filter(id => VALID_MEETING_IDS.indexOf(id) !== -1);

  if (!officer || officer.length > MAX_OFFICER_LENGTH) {
    return json_({ ok: false, error: 'invalid_officer' });
  }
  if (valid.length === 0) {
    return json_({ ok: false, error: 'invalid_meetings' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const existing = new Set(
      readRows_(sheet).map(r => normalize_(r.officer) + '|' + r.meetingId)
    );
    const now = new Date();
    const toAdd = [];
    valid.forEach(id => {
      const key = normalize_(officer) + '|' + id;
      if (!existing.has(key)) {
        existing.add(key);
        toAdd.push([now, officer, id]);
      }
    });
    if (toAdd.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, HEADERS.length).setValues(toAdd);
    }
    return json_({ ok: true, added: toAdd.length, duplicates: valid.length - toAdd.length });
  } finally {
    lock.releaseLock();
  }
}

function report_(body) {
  const adminKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!adminKey || String(body.key || '') !== adminKey) {
    Utilities.sleep(1000); // επιβράδυνση δοκιμών κωδικού
    return json_({ ok: false, error: 'unauthorized' });
  }
  const byMeeting = {};
  readRows_(getSheet_()).forEach(r => {
    (byMeeting[r.meetingId] = byMeeting[r.meetingId] || []).push(r.officer);
  });
  return json_({ ok: true, data: byMeeting });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r[1] && r[2])
    .map(r => ({ officer: String(r[1]), meetingId: String(r[2]) }));
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
