/**
 * «Наради» — бекенд на Google Apps Script.
 * Зберігає наради в Google Таблиці, аудіо й тексти — на Google Диску,
 * синхронізує зустрічі з Google Календарем.
 *
 * Перший запуск: виконати функцію setup() з редактора (Run).
 * Потім: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const APP_VERSION = '1.1.0';
const SHEETS = { meetings: 'Наради', people: 'Люди', glossary: 'Словник', series: 'Серії', tasks: 'Задачі' };
const M_COLS = ['id', 'eventId', 'title', 'start', 'end', 'participants', 'notes', 'folderId',
  'audioFileId', 'audioName', 'durationSec', 'engines', 'primaryEngine', 'speakerMap',
  'docId', 'status', 'updatedAt'];
const P_COLS = ['id', 'name', 'role', 'aliases'];
const S_COLS = ['eventId', 'title', 'participants', 'notes', 'freq', 'until'];
const T_COLS = ['id', 'meetingId', 'seriesKey', 'meetingTitle', 'meetingStart', 'owner', 'task', 'due', 'dueDate', 'status', 'doneAt', 'source', 'edited', 'createdAt', 'updatedAt'];
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
  ensureSheet_(ss, SHEETS.series, S_COLS);
  ensureSheet_(ss, SHEETS.tasks, T_COLS);
  ss.getSheets().forEach(function (sh) {
    const known = [SHEETS.meetings, SHEETS.people, SHEETS.glossary, SHEETS.series, SHEETS.tasks].indexOf(sh.getName()) >= 0;
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
  listTasks: listTasks_,
  carriedTasks: carriedTasks_,
  saveTask: saveTask_,
  deleteTask: deleteTask_,
  searchCorpus: searchCorpus_,
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

/* ---------- Серії повторюваних нарад ---------- */

function seriesSheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  return ss.getSheetByName(SHEETS.series) || ensureSheet_(ss, SHEETS.series, S_COLS);
}

function seriesMap_() {
  const m = {};
  readTable_(seriesSheet_(), S_COLS).forEach(function (r) { m[r.eventId] = r; });
  return m;
}

function saveSeries_(eventId, fields) {
  const sh = seriesSheet_();
  const row = readTable_(sh, S_COLS).filter(function (r) { return r.eventId === eventId; })[0] || { eventId: eventId };
  Object.keys(fields).forEach(function (k) { row[k] = fields[k]; });
  writeRow_(sh, S_COLS, row);
}

function buildRecurrence_(rep) {
  if (!rep.until) throw new Error('Не вказано, до якої дати повторювати');
  const until = new Date(rep.until);
  const W = CalendarApp.Weekday;
  const rec = CalendarApp.newRecurrence();
  let rule;
  if (rep.freq === 'weekdays') rule = rec.addWeeklyRule().onlyOnWeekdays([W.MONDAY, W.TUESDAY, W.WEDNESDAY, W.THURSDAY, W.FRIDAY]);
  else if (rep.freq === 'daily') rule = rec.addDailyRule();
  else if (rep.freq === 'weekly') rule = rec.addWeeklyRule();
  else if (rep.freq === 'biweekly') rule = rec.addWeeklyRule().interval(2);
  else if (rep.freq === 'monthly') rule = rec.addMonthlyRule();
  else throw new Error('Невідома періодичність: ' + rep.freq);
  rule.until(until);
  return rec;
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
  const series = events.some(function (ev) { return ev.isRecurringEvent(); }) ? seriesMap_() : {};
  events.forEach(function (ev) {
    if (ev.isAllDayEvent()) return;
    const evId = ev.getId(), st = ev.getStartTime();
    let row = rows.filter(function (r) { return r.eventId === evId && sameMinute_(r.start, st); })[0];
    if (!row && !ev.isRecurringEvent()) row = rows.filter(function (r) { return r.eventId === evId && !used[r.id]; })[0];
    if (row) used[row.id] = true;
    const mt = toMeeting_(row || null, ev);
    const ser = series[evId];
    if (!row && ser) {
      mt.participants = parseJson_(ser.participants, []);
      if (ser.notes) mt.notes = ser.notes;
    }
    out.push(mt);
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
    if (m.repeat && m.repeat.freq && !m.recurring) {
      // Нова повторювана нарада або перетворення одиночної на серію.
      if (eventId) {
        const old = eventForRow_({ eventId: eventId, start: row.start || m.start });
        if (old && !old.isRecurringEvent()) old.deleteEvent();
      }
      const ser = cal.createEventSeries(m.title, st, en, buildRecurrence_(m.repeat), { description: buildDesc_(m.notes, participants, '') });
      eventId = ser.getId();
      saveSeries_(eventId, { title: m.title, participants: JSON.stringify(participants), notes: m.notes || '', freq: m.repeat.freq, until: m.repeat.until });
    } else if (eventId) {
      if (m.recurring && m.applyToSeries) {
        saveSeries_(eventId, { title: m.title, participants: JSON.stringify(participants), notes: m.notes || '' });
      }
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
    syncTasks_(row, a.analysis);
    return toMeeting_(row, null);
  });
}

function getAnalysis_(a) {
  const row = rowOrThrow_(sheet_(SHEETS.meetings), a.meetingId);
  if (!row.folderId) return null;
  return readJson_(DriveApp.getFolderById(row.folderId), 'analysis.json');
}

/* ---------- Задачі з нарад ---------- */

function tasksSheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  return ss.getSheetByName(SHEETS.tasks) || ensureSheet_(ss, SHEETS.tasks, T_COLS);
}

function normTitle_(t) { return String(t || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function normText_(t) { return String(t || '').toLowerCase().replace(/[^0-9a-zа-яіїєґ']+/gi, ' ').trim(); }

// Ключ «серії»: повторювана подія календаря або нарада з тією самою назвою.
function seriesKeyFor_(row) {
  const ev = eventForRow_(row);
  return ev && ev.isRecurringEvent() ? 'ev:' + row.eventId : 't:' + normTitle_(row.title);
}

function cleanTask_(t) {
  return { id: t.id, meetingId: t.meetingId, seriesKey: t.seriesKey, meetingTitle: t.meetingTitle, meetingStart: t.meetingStart,
    owner: t.owner, task: t.task, due: t.due, dueDate: t.dueDate, status: t.status || 'open', doneAt: t.doneAt, source: t.source, edited: t.edited === '1' };
}

// Після кожного підсумку оновлюємо задачі наради. Виконані й відредаговані вручну не чіпаємо.
function syncTasks_(row, analysis) {
  const sh = tasksSheet_();
  const mine = readTable_(sh, T_COLS).filter(function (t) { return t.meetingId === row.id; });
  const drop = mine.filter(function (t) { return t.source === 'auto' && t.status !== 'done' && t.edited !== '1'; });
  const keep = mine.filter(function (t) { return drop.indexOf(t) < 0; });
  drop.sort(function (x, y) { return y._row - x._row; }).forEach(function (t) { sh.deleteRow(t._row); });
  const seen = {};
  keep.forEach(function (t) { seen[normText_(t.task)] = true; });
  const key = seriesKeyFor_(row);
  const now = new Date().toISOString();
  ((analysis && analysis.tasks) || []).forEach(function (t) {
    const txt = String(t.task || '').trim();
    if (!txt || seen[normText_(txt)]) return;
    seen[normText_(txt)] = true;
    writeRow_(sh, T_COLS, { id: newId_('t_'), meetingId: row.id, seriesKey: key, meetingTitle: row.title, meetingStart: row.start,
      owner: String(t.owner || ''), task: txt, due: String(t.due || ''), dueDate: /^\d{4}-\d{2}-\d{2}$/.test(t.dueDate || '') ? t.dueDate : '',
      status: 'open', doneAt: '', source: 'auto', edited: '', createdAt: now, updatedAt: now });
  });
}

// Один раз переносимо задачі з підсумків, зроблених до появи списку задач.
function ensureTasksMigrated_() {
  if (prop_('TASKS_V')) return;
  withLock_(function () {
    if (prop_('TASKS_V')) return;
    const existing = {};
    readTable_(tasksSheet_(), T_COLS).forEach(function (t) { existing[t.meetingId] = true; });
    readTable_(sheet_(SHEETS.meetings), M_COLS).forEach(function (row) {
      if (existing[row.id] || !row.folderId) return;
      try {
        const an = readJson_(DriveApp.getFolderById(row.folderId), 'analysis.json');
        if (an && an.tasks && an.tasks.length) syncTasks_(row, an);
      } catch (e) { }
    });
    setProp_('TASKS_V', '1');
  });
}

function listTasks_(a) {
  ensureTasksMigrated_();
  return readTable_(tasksSheet_(), T_COLS)
    .filter(function (t) { return !a.meetingId || t.meetingId === a.meetingId; })
    .map(cleanTask_);
}

// Відкриті задачі з попередніх нарад тієї ж серії (або з тією самою назвою).
function carriedTasks_(a) {
  ensureTasksMigrated_();
  let evKey = '', title = '', start = '', selfId = a.meetingId || '';
  if (a.meetingId) {
    const row = rowOrThrow_(sheet_(SHEETS.meetings), a.meetingId);
    const k = seriesKeyFor_(row);
    evKey = k.indexOf('ev:') === 0 ? k : '';
    title = row.title; start = row.start;
  } else {
    evKey = a.recurring && a.eventId ? 'ev:' + a.eventId : '';
    title = a.title; start = a.start;
  }
  const nt = normTitle_(title), st = new Date(start).getTime();
  return readTable_(tasksSheet_(), T_COLS).filter(function (t) {
    if (t.status === 'done' || t.meetingId === selfId) return false;
    if (!(new Date(t.meetingStart).getTime() < st)) return false;
    return (evKey && t.seriesKey === evKey) || (nt && normTitle_(t.meetingTitle) === nt);
  }).map(cleanTask_);
}

function saveTask_(a) {
  return withLock_(function () {
    const sh = tasksSheet_();
    const t = a.task || {};
    const now = new Date().toISOString();
    let row = t.id ? readTable_(sh, T_COLS).filter(function (x) { return x.id === t.id; })[0] : null;
    if (!row) {
      if (!t.task || !String(t.task).trim()) throw new Error('Опиши задачу');
      row = { id: newId_('t_'), meetingId: '', seriesKey: '', meetingTitle: '', meetingStart: '', status: 'open', source: 'manual', createdAt: now };
    }
    const content = ['task', 'owner', 'dueDate', 'due'].some(function (k) { return t[k] !== undefined && String(t[k]) !== String(row[k] || ''); });
    ['task', 'owner', 'due'].forEach(function (k) { if (t[k] !== undefined) row[k] = String(t[k]).trim(); });
    if (t.dueDate !== undefined) row.dueDate = /^\d{4}-\d{2}-\d{2}$/.test(t.dueDate) ? t.dueDate : '';
    if (t.status !== undefined && t.status !== row.status) { row.status = t.status === 'done' ? 'done' : 'open'; row.doneAt = row.status === 'done' ? now : ''; }
    if (content && row.source === 'auto') row.edited = '1';
    row.updatedAt = now;
    writeRow_(sh, T_COLS, row);
    return cleanTask_(row);
  });
}

function deleteTask_(a) {
  return withLock_(function () {
    const sh = tasksSheet_();
    const row = readTable_(sh, T_COLS).filter(function (x) { return x.id === a.id; })[0];
    if (row) sh.deleteRow(row._row);
    return true;
  });
}

/* ---------- Пошук: тексти нарад для індексу на пристрої ---------- */

function searchCorpus_(a) {
  const rows = readTable_(sheet_(SHEETS.meetings), M_COLS).filter(function (r) { return r.folderId && parseJson_(r.engines, []).length; });
  const since = a.since || '';
  const changed = rows.filter(function (r) { return !since || (r.updatedAt || '') > since; })
    .sort(function (x, y) { return (x.updatedAt || '') < (y.updatedAt || '') ? -1 : 1; });
  const items = [];
  let size = 0, cursor = since;
  for (let i = 0; i < changed.length; i++) {
    const r = changed[i];
    const eng = r.primaryEngine || parseJson_(r.engines, [])[0];
    let segs = [], summary = '';
    try {
      const folder = DriveApp.getFolderById(r.folderId);
      const tr = readJson_(folder, 'transcript_' + eng + '.json');
      const an = readJson_(folder, 'analysis.json');
      if (tr && tr.segments) segs = tr.segments.map(function (s) { return [Math.round(s.start || 0), s.speaker || '', s.text || '']; });
      else if (tr && tr.text) segs = [[0, '', tr.text]];
      if (an) summary = an.summary || '';
    } catch (e) { }
    const item = { id: r.id, title: r.title, start: r.start, engine: eng, speakerMap: (parseJson_(r.speakerMap, {}) || {})[eng] || {}, updatedAt: r.updatedAt, summary: summary, segs: segs };
    items.push(item);
    cursor = r.updatedAt || cursor;
    size += JSON.stringify(item).length;
    if (size > 3000000) break;
  }
  return { items: items, ids: rows.map(function (r) { return r.id; }), cursor: cursor, more: items.length < changed.length };
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
    const titlePara = body.getChild(0).asParagraph();
    titlePara.setText(title);
    titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
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
          const due = t.dueDate ? t.dueDate.split('-').reverse().join('.') : (t.due || '—');
          return [String(t.owner || '—'), String(t.task || ''), String(due)];
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
