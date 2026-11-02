/**
 * «Наради» — бекенд на Google Apps Script.
 * Зберігає наради в Google Таблиці, аудіо й тексти — на Google Диску,
 * синхронізує зустрічі з Google Календарем.
 *
 * Перший запуск: виконати функцію setup() з редактора (Run).
 * Потім: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const APP_VERSION = '1.0.0';
const SHEETS = { meetings: 'Наради', people: 'Люди', glossary: 'Словник' };
const M_COLS = ['id', 'eventId', 'title', 'start', 'end', 'participants', 'notes', 'folderId',
  'audioFileId', 'audioName', 'durationSec', 'engines', 'primaryEngine', 'speakerMap',
  'docId', 'status', 'updatedAt'];
const P_COLS = ['id', 'name', 'role', 'aliases'];
const DESC_MARK = '\n\n— Наради —\n';

/* ---------- Перше налаштування ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();

  let root = null;
  const rootId = props.getProperty('ROOT_FOLDER_ID');
  if (rootId) { try { root = DriveApp.getFolderById(rootId); if (root.isTrashed()) root = null; } catch (e) { root = null; } }
  if (!root) {
    root = DriveApp.createFolder('Наради — записи');
    props.setProperty('ROOT_FOLDER_ID', root.getId());
  }

  let ss = null;
  const sid = props.getProperty('SHEET_ID');
  if (sid) { try { ss = SpreadsheetApp.openById(sid); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Наради — база');
    DriveApp.getFileById(ss.getId()).moveTo(root);
    props.setProperty('SHEET_ID', ss.getId());
  }
  ensureSheet_(ss, SHEETS.meetings, M_COLS);
  ensureSheet_(ss, SHEETS.people, P_COLS);
  ensureSheet_(ss, SHEETS.glossary, ['term']);
  ss.getSheets().forEach(function (sh) {
    const known = [SHEETS.meetings, SHEETS.people, SHEETS.glossary].indexOf(sh.getName()) >= 0;
    if (!known && sh.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(sh);
  });

  let cal = null;
  const calId = props.getProperty('CALENDAR_ID');
  if (calId) cal = CalendarApp.getCalendarById(calId);
  if (!cal) {
    const found = CalendarApp.getCalendarsByName('Наради');
    cal = found.length ? found[0] : CalendarApp.createCalendar('Наради');
    props.setProperty('CALENDAR_ID', cal.getId());
  }

  let secret = props.getProperty('SECRET');
  if (!secret) {
    secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    props.setProperty('SECRET', secret);
  }
  Logger.log('Готово. Секретний ключ для додатку:');
  Logger.log(secret);
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  return sh;
}

/* ---------- Вхідні запити ---------- */

function doGet() {
  return ContentService.createTextOutput('Наради — бекенд працює. Версія ' + APP_VERSION);
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) return json_({ ok: false, error: 'not_setup' });
  if (!req || req.token !== secret) return json_({ ok: false, error: 'unauthorized' });
  const fn = ACTIONS[req.action];
  if (!fn) return json_({ ok: false, error: 'unknown_action: ' + req.action });
  try {
    return json_({ ok: true, data: fn(req.args || {}) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  ping: ping_,
  getConfig: getConfig_,
  saveConfig: saveConfig_,
  listCalendars: listCalendars_,
  listPeople: function () { return readTable_(sheet_(SHEETS.people), P_COLS).map(cleanPerson_); },
  savePerson: savePerson_,
  deletePerson: deletePerson_,
  getGlossary: getGlossary_,
  saveGlossary: saveGlossary_,
  listMeetings: listMeetings_,
  getMeeting: getMeeting_,
  saveMeeting: saveMeeting_,
  deleteMeeting: deleteMeeting_,
  driveToken: driveToken_,
  attachAudio: attachAudio_,
  saveTranscript: saveTranscript_,
  getTranscripts: getTranscripts_,
  saveMeetingState: saveMeetingState_,
  saveAnalysis: saveAnalysis_,
  getAnalysis: getAnalysis_,
  buildDoc: buildDoc_,
  proxy: proxy_
};

/* ---------- Конфігурація ---------- */

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function setProp_(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); }
function propJson_(k, def) { try { const v = prop_(k); return v ? JSON.parse(v) : def; } catch (e) { return def; } }

function info_() {
  const cal = getCal_();
  return {
    version: APP_VERSION,
    email: Session.getEffectiveUser().getEmail(),
    timeZone: Session.getScriptTimeZone(),
    calendarId: cal.getId(),
    calendarName: cal.getName(),
    folderUrl: DriveApp.getFolderById(prop_('ROOT_FOLDER_ID')).getUrl(),
    sheetUrl: SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl()
  };
}

function ping_() { return info_(); }

function getConfig_() {
  return { keys: propJson_('KEYS', {}), settings: propJson_('SETTINGS', {}), info: info_() };
}

function saveConfig_(a) {
  if (a.keys) {
    const keys = propJson_('KEYS', {});
    Object.keys(a.keys).forEach(function (k) {
      const v = a.keys[k];
      if (v === '' || v === null) delete keys[k];
      else if (typeof v === 'string') keys[k] = v.trim();
    });
    setProp_('KEYS', JSON.stringify(keys));
  }
  if (a.settings) setProp_('SETTINGS', JSON.stringify(a.settings));
  if (a.calendarId) {
    const cal = CalendarApp.getCalendarById(a.calendarId);
    if (!cal) throw new Error('Календар не знайдено');
    setProp_('CALENDAR_ID', cal.getId());
  }
  return getConfig_();
}

function listCalendars_() {
  return CalendarApp.getAllOwnedCalendars().map(function (c) {
    return { id: c.getId(), name: c.getName(), primary: c.isMyPrimaryCalendar() };
  });
}

function getCal_() {
  const id = prop_('CALENDAR_ID');
  const cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) throw new Error('Календар не налаштовано — запусти setup()');
  return cal;
}

/* ---------- Таблиця ---------- */

function sheet_(name) {
  const sh = SpreadsheetApp.openById(prop_('SHEET_ID')).getSheetByName(name);
  if (!sh) throw new Error('Немає аркуша «' + name + '» — запусти setup()');
  return sh;
}

function readTable_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  const out = [];
  vals.forEach(function (r, i) {
    if (!r[0]) return;
    const o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = r[j]; });
    out.push(o);
  });
  return out;
}

function writeRow_(sh, cols, obj) {
  const vals = cols.map(function (c) {
    const v = obj[c];
    if (v === null || v === undefined) return '';
    return typeof v === 'object' ? JSON.stringify(v) : String(v);
  });
  let r = obj._row;
  if (!r) {
    r = sh.getLastRow() + 1;
    if (r > sh.getMaxRows()) {
      sh.insertRowsAfter(sh.getMaxRows(), 100);
      sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    }
    obj._row = r;
  }
  sh.getRange(r, 1, 1, cols.length).setValues([vals]);
  return obj;
}

function findById_(sh, id) {
  return readTable_(sh, sh.getName() === SHEETS.people ? P_COLS : M_COLS)
    .filter(function (r) { return r.id === id; })[0] || null;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function parseJson_(s, def) { if (!s) return def; try { return JSON.parse(s); } catch (e) { return def; } }
function newId_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

/* ---------- Люди і словник ---------- */

function cleanPerson_(p) { return { id: p.id, name: p.name, role: p.role, aliases: p.aliases }; }

function savePerson_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.people);
    const p = a.person || {};
    if (!p.name || !String(p.name).trim()) throw new Error("Вкажи ім'я");
    let row = p.id ? findById_(sh, p.id) : null;
    if (!row) row = { id: newId_('p_') };
    row.name = String(p.name).trim();
    row.role = p.role ? String(p.role).trim() : '';
    row.aliases = p.aliases ? String(p.aliases).trim() : '';
    writeRow_(sh, P_COLS, row);
    return cleanPerson_(row);
  });
}

function deletePerson_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.people);
    const row = findById_(sh, a.id);
    if (row) sh.deleteRow(row._row);
    return true;
  });
}

function getGlossary_() {
  return readTable_(sheet_(SHEETS.glossary), ['term']).map(function (r) { return r.term; });
}

function saveGlossary_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.glossary);
    const terms = (a.terms || []).map(function (t) { return String(t).trim(); }).filter(String);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
    if (terms.length) {
      if (terms.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), terms.length + 1 - sh.getMaxRows());
      sh.getRange(2, 1, terms.length, 1).setNumberFormat('@').setValues(terms.map(function (t) { return [t]; }));
    }
    return terms;
  });
}

function peopleById_() {
  const m = {};
  readTable_(sheet_(SHEETS.people), P_COLS).forEach(function (p) { m[p.id] = p; });
  return m;
}

/* ---------- Наради ---------- */

function sameMinute_(iso, date) {
  if (!iso) return false;
  return Math.abs(new Date(iso).getTime() - date.getTime()) < 60000;
}

function notesFromDesc_(desc) {
  if (!desc) return '';
  const i = desc.indexOf(DESC_MARK);
  return (i >= 0 ? desc.slice(0, i) : desc).trim();
}

function buildDesc_(notes, participantIds, docId) {
  const people = peopleById_();
  const names = (participantIds || []).map(function (id) { return people[id] ? people[id].name : null; }).filter(String);
  let block = '';
  if (names.length) block += 'Учасники: ' + names.join(', ') + '\n';
  if (docId) block += 'Підсумок: https://docs.google.com/document/d/' + docId + '/edit\n';
  const appUrl = (propJson_('SETTINGS', {}) || {}).appUrl;
  if (appUrl) block += 'Додаток: ' + appUrl + '\n';
  return (notes || '').trim() + (block ? DESC_MARK + block : '');
}

function toMeeting_(row, ev) {
  const m = {
    id: row ? row.id : null,
    eventId: ev ? ev.getId() : (row ? row.eventId : ''),
    title: ev ? ev.getTitle() : (row ? row.title : ''),
    start: ev ? ev.getStartTime().toISOString() : (row ? row.start : ''),
    end: ev ? ev.getEndTime().toISOString() : (row ? row.end : ''),
    recurring: ev ? ev.isRecurringEvent() : false,
    allDay: ev ? ev.isAllDayEvent() : false,
    location: ev ? ev.getLocation() : '',
    participants: row ? parseJson_(row.participants, []) : [],
    notes: row ? row.notes : (ev ? notesFromDesc_(ev.getDescription()) : ''),
    audioFileId: row ? row.audioFileId : '',
    audioName: row ? row.audioName : '',
    durationSec: row && row.durationSec ? Number(row.durationSec) : 0,
    engines: row ? parseJson_(row.engines, []) : [],
    primaryEngine: row ? row.primaryEngine : '',
    speakerMap: row ? parseJson_(row.speakerMap, {}) : {},
    docId: row ? row.docId : '',
    status: row ? row.status : '',
    eventMissing: !!(row && row.eventId && !ev)
  };
  if (row && ev && !m.recurring) { m.title = ev.getTitle(); }
  return m;
}

function listMeetings_(a) {
  const from = new Date(a.from), to = new Date(a.to);
  const events = a.onlyRecorded ? [] : getCal_().getEvents(from, to);
  const rows = readTable_(sheet_(SHEETS.meetings), M_COLS);
  const used = {};
  const out = [];
  events.forEach(function (ev) {
    if (ev.isAllDayEvent()) return;
    const evId = ev.getId(), st = ev.getStartTime();
    let row = rows.filter(function (r) { return r.eventId === evId && sameMinute_(r.start, st); })[0];
    if (!row && !ev.isRecurringEvent()) row = rows.filter(function (r) { return r.eventId === evId && !used[r.id]; })[0];
    if (row) used[row.id] = true;
    out.push(toMeeting_(row || null, ev));
  });
  rows.forEach(function (r) {
    if (used[r.id] || !r.start) return;
    const s = new Date(r.start);
    if (s >= from && s < to) {
      if (a.onlyRecorded && !r.audioFileId) return;
      out.push(toMeeting_(r, null));
    }
  });
  out.sort(function (x, y) { return new Date(x.start) - new Date(y.start); });
  return out;
}

function eventForRow_(row) {
  if (!row.eventId) return null;
  try {
    const cal = getCal_();
    const st = new Date(row.start);
    const evs = cal.getEvents(new Date(st.getTime() - 60000), new Date(st.getTime() + 60000))
      .filter(function (e) { return e.getId() === row.eventId; });
    if (evs.length) return evs[0];
    return cal.getEventById(row.eventId);
  } catch (e) { return null; }
}

function getMeeting_(a) {
  const row = findById_(sheet_(SHEETS.meetings), a.id);
  if (!row) throw new Error('Нараду не знайдено');
  const ev = eventForRow_(row);
  const m = toMeeting_(row, ev);
  if (ev && ev.isRecurringEvent()) { m.start = row.start; m.end = row.end; }
  return m;
}

function saveMeeting_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const m = a.meeting || {};
    if (!m.title || !m.start || !m.end) throw new Error('Потрібні назва, початок і кінець');
    const st = new Date(m.start), en = new Date(m.end);
    if (en <= st) throw new Error('Кінець має бути пізніше за початок');

    let row = m.id ? findById_(sh, m.id) : null;
    if (!row && m.eventId) {
      row = readTable_(sh, M_COLS).filter(function (r) { return r.eventId === m.eventId && sameMinute_(r.start, st); })[0] || null;
    }
    const isNew = !row;
    if (!row) row = { id: newId_('m_'), status: 'planned' };

    const participants = m.participants || [];
    const cal = getCal_();
    let eventId = m.eventId || row.eventId || '';
    if (eventId) {
      if (!m.recurring) {
        const ev = eventForRow_({ eventId: eventId, start: row.start || m.start });
        if (ev && !ev.isRecurringEvent()) {
          if (ev.getTitle() !== m.title) ev.setTitle(m.title);
          if (ev.getStartTime().getTime() !== st.getTime() || ev.getEndTime().getTime() !== en.getTime()) ev.setTime(st, en);
          ev.setDescription(buildDesc_(m.notes, participants, row.docId));
        }
      }
    } else if (m.addToCalendar !== false) {
      const ev = cal.createEvent(m.title, st, en, { description: buildDesc_(m.notes, participants, row.docId) });
      eventId = ev.getId();
    }

    row.eventId = eventId;
    row.title = m.title;
    row.start = st.toISOString();
    row.end = en.toISOString();
    row.participants = JSON.stringify(participants);
    row.notes = m.notes || '';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    if (!isNew && row.folderId) renameFolder_(row);
    return toMeeting_(row, null);
  });
}

function deleteMeeting_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = findById_(sh, a.id);
    if (!row) return true;
    if (a.withEvent && row.eventId) {
      const ev = eventForRow_(row);
      if (ev && !ev.isRecurringEvent()) ev.deleteEvent();
    }
    if (row.folderId) { try { DriveApp.getFolderById(row.folderId).setTrashed(true); } catch (e) { } }
    sh.deleteRow(row._row);
    return true;
  });
}

/* ---------- Файли наради ---------- */

function folderName_(row) {
  const tz = Session.getScriptTimeZone();
  const d = Utilities.formatDate(new Date(row.start), tz, 'yyyy-MM-dd HH.mm');
  return (d + ' ' + String(row.title || '').replace(/[\\/:*?"<>|]/g, ' ')).trim().slice(0, 120);
}

function ensureFolder_(sh, row) {
  if (row.folderId) {
    try { const f = DriveApp.getFolderById(row.folderId); if (!f.isTrashed()) return f; } catch (e) { }
  }
  const root = DriveApp.getFolderById(prop_('ROOT_FOLDER_ID'));
  const f = root.createFolder(folderName_(row));
  row.folderId = f.getId();
  writeRow_(sh, M_COLS, row);
  return f;
}

function renameFolder_(row) {
  try { DriveApp.getFolderById(row.folderId).setName(folderName_(row)); } catch (e) { }
}

function fileByName_(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) return f; }
  return null;
}

function writeJson_(folder, name, obj) {
  const s = JSON.stringify(obj);
  const f = fileByName_(folder, name);
  if (f) { f.setContent(s); return f; }
  return folder.createFile(name, s, MimeType.PLAIN_TEXT);
}

function readJson_(folder, name) {
  const f = fileByName_(folder, name);
  return f ? parseJson_(f.getBlob().getDataAsString('UTF-8'), null) : null;
}

function rowOrThrow_(sh, id) {
  const row = findById_(sh, id);
  if (!row) throw new Error('Нараду не знайдено');
  return row;
}

function driveToken_(a) {
  const sh = sheet_(SHEETS.meetings);
  const out = { token: ScriptApp.getOAuthToken() };
  if (a.meetingId) {
    out.folderId = withLock_(function () { return ensureFolder_(sh, rowOrThrow_(sh, a.meetingId)).getId(); });
  }
  return out;
}

function attachAudio_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const old = row.audioFileId;
    row.audioFileId = a.fileId;
    row.audioName = a.name || '';
    row.durationSec = a.durationSec ? String(Math.round(a.durationSec)) : '';
    row.status = 'recorded';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    if (old && old !== a.fileId) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) { } }
    return toMeeting_(row, null);
  });
}

function saveTranscript_(a) {
  if (!/^[a-z0-9_-]+$/.test(a.engine || '')) throw new Error('Невірна назва двигуна');
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const folder = ensureFolder_(sh, row);
    writeJson_(folder, 'transcript_' + a.engine + '.json', a.transcript);
    const engines = parseJson_(row.engines, []);
    if (engines.indexOf(a.engine) < 0) engines.push(a.engine);
    row.engines = JSON.stringify(engines);
    if (!row.primaryEngine) row.primaryEngine = a.engine;
    row.status = 'transcribed';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    return toMeeting_(row, null);
  });
}

function getTranscripts_(a) {
  const sh = sheet_(SHEETS.meetings);
  const row = rowOrThrow_(sh, a.meetingId);
  const out = {};
  if (!row.folderId) return out;
  const folder = DriveApp.getFolderById(row.folderId);
  parseJson_(row.engines, []).forEach(function (eng) {
    const t = readJson_(folder, 'transcript_' + eng + '.json');
    if (t) out[eng] = t;
  });
  return out;
}

function saveMeetingState_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    if (a.primaryEngine !== undefined) row.primaryEngine = a.primaryEngine;
    if (a.speakerMap !== undefined) row.speakerMap = JSON.stringify(a.speakerMap);
    if (a.removeEngine) {
      const engines = parseJson_(row.engines, []).filter(function (e) { return e !== a.removeEngine; });
      row.engines = JSON.stringify(engines);
      if (row.primaryEngine === a.removeEngine) row.primaryEngine = engines[0] || '';
      if (row.folderId) {
        const f = fileByName_(DriveApp.getFolderById(row.folderId), 'transcript_' + a.removeEngine + '.json');
        if (f) f.setTrashed(true);
      }
    }
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    return toMeeting_(row, null);
  });
}

function saveAnalysis_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const folder = ensureFolder_(sh, row);
    writeJson_(folder, 'analysis.json', a.analysis);
    row.status = 'analyzed';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    return toMeeting_(row, null);
  });
}

function getAnalysis_(a) {
  const row = rowOrThrow_(sheet_(SHEETS.meetings), a.meetingId);
  if (!row.folderId) return null;
  return readJson_(DriveApp.getFolderById(row.folderId), 'analysis.json');
}

/* ---------- Google Doc з підсумком ---------- */

function fmtTime_(sec) {
  sec = Math.max(0, Math.round(sec || 0));
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  const mm = (m < 10 ? '0' : '') + m, ss = (s < 10 ? '0' : '') + s;
  return h ? h + ':' + mm + ':' + ss : mm + ':' + ss;
}

function speakerName_(label, map, people) {
  const v = map[label];
  if (!v) return label || '';
  if (people[v]) return people[v].name;
  return v;
}

function buildDoc_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const folder = ensureFolder_(sh, row);
    const analysis = readJson_(folder, 'analysis.json');
    const engines = parseJson_(row.engines, []);
    const primary = row.primaryEngine || engines[0] || '';
    const tr = primary ? readJson_(folder, 'transcript_' + primary + '.json') : null;
    const map = (parseJson_(row.speakerMap, {}) || {})[primary] || {};
    const people = peopleById_();
    const tz = Session.getScriptTimeZone();
    const st = new Date(row.start), en = new Date(row.end);
    const title = Utilities.formatDate(st, tz, 'dd.MM.yyyy') + ' — ' + row.title;

    let doc = null;
    if (row.docId) { try { doc = DocumentApp.openById(row.docId); } catch (e) { doc = null; } }
    if (!doc) {
      doc = DocumentApp.create(title);
      DriveApp.getFileById(doc.getId()).moveTo(folder);
      row.docId = doc.getId();
      writeRow_(sh, M_COLS, row);
    } else {
      doc.setName(title);
    }

    const body = doc.getBody();
    body.clear();
    body.getChild(0).asParagraph().setText(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(Utilities.formatDate(st, tz, 'dd.MM.yyyy, HH:mm') + '–' + Utilities.formatDate(en, tz, 'HH:mm'));
    const names = parseJson_(row.participants, []).map(function (id) { return people[id] ? people[id].name + (people[id].role ? ' (' + people[id].role + ')' : '') : null; }).filter(String);
    if (names.length) body.appendParagraph('Учасники: ' + names.join(', '));

    const H = DocumentApp.ParagraphHeading;
    if (analysis) {
      body.appendParagraph('Підсумок').setHeading(H.HEADING1);
      String(analysis.summary || '').split(/\n+/).forEach(function (p) { if (p.trim()) body.appendParagraph(p.trim()); });

      const list = function (head, items) {
        if (!items || !items.length) return;
        body.appendParagraph(head).setHeading(H.HEADING1);
        items.forEach(function (it) { body.appendListItem(String(it)).setGlyphType(DocumentApp.GlyphType.BULLET); });
      };
      list('Рішення', analysis.decisions);

      if (analysis.tasks && analysis.tasks.length) {
        body.appendParagraph('Задачі').setHeading(H.HEADING1);
        const rows = [['Хто', 'Що', 'Термін']].concat(analysis.tasks.map(function (t) {
          return [String(t.owner || '—'), String(t.task || ''), String(t.due || '—')];
        }));
        const table = body.appendTable(rows);
        const hr = table.getRow(0);
        for (let i = 0; i < hr.getNumCells(); i++) hr.getCell(i).editAsText().setBold(true);
      }
      list('Відкриті питання', analysis.openQuestions);
      list('Цифри й факти', analysis.facts);
    }

    if (tr && tr.segments && tr.segments.length) {
      body.appendParagraph('Повний текст' + (tr.engineName ? ' (' + tr.engineName + ')' : '')).setHeading(H.HEADING1);
      tr.segments.forEach(function (s) {
        const t = '[' + fmtTime_(s.start) + '] ';
        const who = s.speaker ? speakerName_(s.speaker, map, people) + ': ' : '';
        const p = body.appendParagraph(t + who + (s.text || ''));
        const txt = p.editAsText();
        txt.setForegroundColor(0, t.length - 1, '#888888');
        if (who) txt.setBold(t.length, t.length + who.length - 2, true);
      });
    } else if (tr && tr.text) {
      body.appendParagraph('Повний текст').setHeading(H.HEADING1);
      body.appendParagraph(tr.text);
    }
    doc.saveAndClose();

    const ev = eventForRow_(row);
    if (ev && !ev.isRecurringEvent()) {
      try { ev.setDescription(buildDesc_(row.notes, parseJson_(row.participants, []), row.docId)); } catch (e) { }
    }
    return { docId: row.docId, url: 'https://docs.google.com/document/d/' + row.docId + '/edit' };
  });
}

/* ---------- Запасний шлях до сервісів розпізнавання ---------- */

function proxy_(a) {
  if (!/^https:\/\/(api\.openai\.com|api\.groq\.com|api\.elevenlabs\.io|generativelanguage\.googleapis\.com|api\.anthropic\.com)\//.test(a.url || '')) {
    throw new Error('Адреса не дозволена');
  }
  const opt = { method: a.method || 'post', muteHttpExceptions: true, headers: a.headers || {} };
  if (a.json !== undefined && a.json !== null) {
    opt.contentType = 'application/json';
    opt.payload = JSON.stringify(a.json);
  } else if (a.form) {
    const mp = multipart_(a.form);
    opt.contentType = 'multipart/form-data; boundary=' + mp.boundary;
    opt.payload = mp.bytes;
  }
  const res = UrlFetchApp.fetch(a.url, opt);
  return { status: res.getResponseCode(), text: res.getContentText('UTF-8') };
}

function multipart_(form) {
  const boundary = '----nar' + Utilities.getUuid().replace(/-/g, '');
  let bytes = [];
  const add = function (s) { bytes = bytes.concat(Utilities.newBlob(s).getBytes()); };
  form.forEach(function (f) {
    let head = '--' + boundary + '\r\nContent-Disposition: form-data; name="' + f.name + '"';
    if (f.base64 !== undefined) {
      head += '; filename="' + (f.filename || 'audio.mp3') + '"\r\nContent-Type: ' + (f.mimeType || 'application/octet-stream') + '\r\n\r\n';
      add(head);
      bytes = bytes.concat(Utilities.base64Decode(f.base64));
      add('\r\n');
    } else {
      add(head + '\r\n\r\n' + String(f.value) + '\r\n');
    }
  });
  add('--' + boundary + '--\r\n');
  return { boundary: boundary, bytes: bytes };
}
