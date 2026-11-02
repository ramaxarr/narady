/**
 * «Наради» — бекенд на Google Apps Script.
 * Зберігає наради в Google Таблиці, аудіо й тексти — на Google Диску,
 * синхронізує зустрічі з Google Календарем.
 *
 * Перший запуск: виконати функцію setup() з редактора (Run).
 * Потім: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 */

const APP_VERSION = '1.2.0';
const SHEETS = { meetings: 'Наради', people: 'Люди', glossary: 'Словник', series: 'Серії', tasks: 'Задачі' };
const M_COLS = ['id', 'eventId', 'title', 'start', 'end', 'participants', 'notes', 'folderId',
  'audioFileId', 'audioName', 'durationSec', 'engines', 'primaryEngine', 'speakerMap',
  'docId', 'status', 'updatedAt', 'audioPieces'];
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
  // Перевірка дозволу на тригери — без нього фонова обробка не продовжиться після закриття додатку.
  ScriptApp.getProjectTriggers();
  Logger.log('Готово. Версія ' + APP_VERSION + ', фонова обробка дозволена. Секретний ключ для додатку:');
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
  proxy: proxy_,
  queueJob: queueJob_,
  jobStatus: jobStatus_,
  cancelJob: cancelJob_,
  runQueue: function () { return runQueue_(); }
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
    const terms = (a.terms || []).map(function (t) { return String(t).trim(); }).filter(Boolean);
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
  const names = (participantIds || []).map(function (id) { return people[id] ? people[id].name : null; }).filter(Boolean);
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
    audioPieces: row ? parseJson_(row.audioPieces, null) : null,
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
    ensureHeader_(sh, M_COLS);
    const row = rowOrThrow_(sh, a.meetingId);
    const oldIds = [row.audioFileId].concat((parseJson_(row.audioPieces, []) || []).map(function (p) { return p.id; })).filter(Boolean);
    let keep = {};
    if (a.pieces && a.pieces.length) {
      // Запис частинами по ~5 хв: так їх може обробити скрипт у фоні.
      const pieces = a.pieces.map(function (p) { return { id: String(p.id), o: Number(p.o) || 0, d: Number(p.d) || 0 }; });
      row.audioPieces = JSON.stringify(pieces);
      row.audioFileId = pieces[0].id;
      pieces.forEach(function (p) { keep[p.id] = true; });
      const folder = ensureFolder_(sh, row);
      writeJson_(folder, 'audio.json', { pieces: pieces, duration: Number(a.durationSec) || 0 });
    } else {
      row.audioFileId = a.fileId;
      row.audioPieces = '';
      keep[a.fileId] = true;
    }
    row.audioName = a.name || '';
    row.durationSec = a.durationSec ? String(Math.round(a.durationSec)) : '';
    row.status = 'recorded';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    oldIds.forEach(function (id) { if (!keep[id]) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { } } });
    return toMeeting_(row, null);
  });
}

function ensureHeader_(sh, cols) {
  const c = sh.getRange(1, cols.length);
  if (!c.getValue()) c.setValue(cols[cols.length - 1]).setFontWeight('bold');
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
    if (row.status !== 'processing') row.status = 'transcribed';
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
    if (row.status !== 'processing') row.status = 'analyzed';
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
    const names = parseJson_(row.participants, []).map(function (id) { return people[id] ? people[id].name + (people[id].role ? ' (' + people[id].role + ')' : '') : null; }).filter(Boolean);
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
    if (f.base64 !== undefined || f.bytes !== undefined) {
      head += '; filename="' + (f.filename || 'audio.mp3') + '"\r\nContent-Type: ' + (f.mimeType || 'application/octet-stream') + '\r\n\r\n';
      add(head);
      bytes = bytes.concat(f.bytes !== undefined ? f.bytes : Utilities.base64Decode(f.base64));
      add('\r\n');
    } else {
      add(head + '\r\n\r\n' + String(f.value) + '\r\n');
    }
  });
  add('--' + boundary + '--\r\n');
  return { boundary: boundary, bytes: bytes };
}

/* =====================================================================
 * Фонова обробка: розпізнавання й підсумок на боці Google.
 * Телефон завантажує запис частинами по ~5 хв і ставить нараду в чергу.
 * Далі скрипт обробляє частину за частиною (кожен запит до сервісу ≤ ~60 с),
 * а якщо часу одного запуску не вистачає — продовжує за тригером.
 * ===================================================================== */

const ENGINE_NAMES = { gemini: 'Gemini', groq: 'Groq Whisper', openai: 'OpenAI (ChatGPT)', elevenlabs: 'ElevenLabs Scribe' };
const ENGINE_SPEAKERS = { gemini: true, groq: false, openai: true, elevenlabs: true };
const GROUP_SIZE = { gemini: 1, openai: 1, groq: 3, elevenlabs: 3 }; // скільки 5-хв частин в одному запиті
const ANALYZER_NAMES = { gemini: 'Gemini', openai: 'OpenAI', anthropic: 'Claude' };
const DEFAULT_SETTINGS_GS = {
  language: 'uk', autoAnalyze: true, autoDoc: true, analysisEngine: 'gemini',
  models: { gemini: 'gemini-2.5-flash', groq: 'whisper-large-v3', openai: 'gpt-4o-transcribe-diarize', elevenlabs: 'scribe_v2' },
  analysisModels: { gemini: 'gemini-2.5-flash', openai: 'gpt-5-mini', anthropic: 'claude-sonnet-5-5' }
};
const RUN_BUDGET_MS = 240000;   // після 4 хв нових кроків не починаємо
const RETRY_WAITS = [15000, 30000, 60000];
const WEEKDAYS_UK = ['неділя', 'понеділок', 'вівторок', 'середа', 'четвер', 'пʼятниця', 'субота'];

function settingsGs_() {
  const s = propJson_('SETTINGS', {}) || {};
  return {
    language: s.language || DEFAULT_SETTINGS_GS.language,
    autoDoc: s.autoDoc !== undefined ? !!s.autoDoc : true,
    analysisEngine: s.analysisEngine || DEFAULT_SETTINGS_GS.analysisEngine,
    models: Object.assign({}, DEFAULT_SETTINGS_GS.models, s.models || {}),
    analysisModels: Object.assign({}, DEFAULT_SETTINGS_GS.analysisModels, s.analysisModels || {})
  };
}

function ctxFor_(row) {
  const people = peopleById_();
  return {
    keys: propJson_('KEYS', {}) || {},
    settings: settingsGs_(),
    participants: parseJson_(row.participants, []).map(function (id) { return people[id]; }).filter(Boolean),
    glossary: getGlossary_(),
    people: people
  };
}

/* ---------- Черга ---------- */

function queueJob_(a) {
  const job = withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const folder = ensureFolder_(sh, row);
    const audio = readJson_(folder, 'audio.json');
    const keys = propJson_('KEYS', {}) || {};
    const engines = (a.engines || []).filter(function (e) { return ENGINE_NAMES[e]; });
    if (engines.length && !(audio && audio.pieces && audio.pieces.length)) throw new Error('Для фонової обробки запис має бути завантажений частинами');
    engines.forEach(function (e) { if (!keys[e]) throw new Error(ENGINE_NAMES[e] + ': не додано ключ API'); });
    const analyze = !!(a.analyzeAfter || a.analyzeOnly);
    if (!engines.length && !analyze) throw new Error('Не вибрано, що робити');
    const old = readJson_(folder, 'job.json');
    if (old && (old.status === 'queued' || old.status === 'running')) throw new Error('Ця нарада вже обробляється');
    const j = {
      id: newId_('j_'), meetingId: row.id, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
      status: 'queued', primaryBefore: row.primaryEngine || '', analysisEngine: a.analysisEngine || '',
      engines: engines.map(function (id) {
        return { id: id, status: 'queued', done: 0, total: groupsFor_(id, audio).length, segs: [], warnings: [], attempts: 0, nextAt: 0, error: '', detail: '' };
      }),
      analysis: { status: analyze ? 'queued' : 'skip', attempts: 0, nextAt: 0, error: '', detail: '' },
      doc: { status: analyze ? 'queued' : 'skip', error: '' }
    };
    writeJson_(folder, 'job.json', j);
    const q = propJson_('QUEUE', []);
    if (q.indexOf(row.id) < 0) q.push(row.id);
    setProp_('QUEUE', JSON.stringify(q));
    row.status = 'processing';
    row.updatedAt = new Date().toISOString();
    writeRow_(sh, M_COLS, row);
    return j;
  });
  ensureTrigger_(60);
  return jobView_(job);
}

function jobStatus_(a) {
  const row = rowOrThrow_(sheet_(SHEETS.meetings), a.meetingId);
  if (!row.folderId) return null;
  const job = readJson_(DriveApp.getFolderById(row.folderId), 'job.json');
  return job ? jobView_(job) : null;
}

function cancelJob_(a) {
  return withLock_(function () {
    const sh = sheet_(SHEETS.meetings);
    const row = rowOrThrow_(sh, a.meetingId);
    const folder = ensureFolder_(sh, row);
    const job = readJson_(folder, 'job.json');
    if (job && (job.status === 'queued' || job.status === 'running')) {
      job.status = 'cancelled';
      job.updatedAt = new Date().toISOString();
      writeJson_(folder, 'job.json', job);
    }
    dequeue_(row.id);
    row.status = finalStatus_(row, folder);
    writeRow_(sh, M_COLS, row);
    return toMeeting_(row, null);
  });
}

function jobView_(j) {
  return {
    id: j.id, status: j.status, createdAt: j.createdAt, updatedAt: j.updatedAt,
    engines: j.engines.map(function (e) {
      return { id: e.id, name: ENGINE_NAMES[e.id], status: e.status, done: e.done, total: e.total, error: e.error, detail: e.detail, words: e.words || 0, startedAt: e.startedAt || '', endedAt: e.endedAt || '' };
    }),
    analysis: { status: j.analysis.status, error: j.analysis.error, detail: j.analysis.detail, tasks: j.analysis.tasks || 0 },
    doc: { status: j.doc.status, error: j.doc.error }
  };
}

function dequeue_(meetingId) {
  const q = propJson_('QUEUE', []).filter(function (x) { return x !== meetingId; });
  setProp_('QUEUE', JSON.stringify(q));
}

function finalStatus_(row, folder) {
  if (readJson_(folder, 'analysis.json')) return 'analyzed';
  if (parseJson_(row.engines, []).length) return 'transcribed';
  return row.audioFileId ? 'recorded' : 'planned';
}

/* ---------- Запуск обробки (тригер або запит з додатку) ---------- */

function processQueue() {
  clearTriggers_();
  runQueue_();
}

function acquireLease_() {
  return withLock_(function () {
    const l = Number(prop_('QUEUE_LEASE') || 0);
    if (l > Date.now()) return false;
    setProp_('QUEUE_LEASE', String(Date.now() + 330000));
    return true;
  });
}
function releaseLease_() {
  withLock_(function () { PropertiesService.getScriptProperties().deleteProperty('QUEUE_LEASE'); });
}

function runQueue_() {
  // Хтось уже обробляє чергу. Ставимо запасний запуск: якщо той процес обірветься, черга не застрягне.
  if (!acquireLease_()) { ensureTrigger_(120); return { busy: true }; }
  const t0 = Date.now();
  let steps = 0;
  try {
    while (Date.now() - t0 < RUN_BUDGET_MS) {
      const q = propJson_('QUEUE', []);
      if (!q.length) break;
      let did = false, wait = Infinity;
      for (let i = 0; i < q.length; i++) {
        const r = stepJob_(q[i]);
        if (r === true) { did = true; steps++; break; }
        if (typeof r === 'number') wait = Math.min(wait, r);
      }
      if (did) continue;
      // Усі чекають повтору: якщо недовго — чекаємо тут, інакше продовжить тригер.
      if (wait < 35000 && Date.now() - t0 + wait < RUN_BUDGET_MS) { Utilities.sleep(Math.max(1000, wait)); continue; }
      break;
    }
  } finally {
    releaseLease_();
  }
  const left = propJson_('QUEUE', []).length;
  if (left) ensureTrigger_(steps ? 10 : 45); else clearTriggers_();
  return { steps: steps, left: left };
}

function ensureTrigger_(sec) {
  try {
    clearTriggers_();
    ScriptApp.newTrigger('processQueue').timeBased().after(sec * 1000).create();
  } catch (e) { }
}
function clearTriggers_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'processQueue') ScriptApp.deleteTrigger(t); });
  } catch (e) { }
}

/** Один крок обробки наради. true — щось зроблено; число — скільки мс чекати; false — нічого. */
function stepJob_(meetingId) {
  const sh = sheet_(SHEETS.meetings);
  const row = findById_(sh, meetingId);
  if (!row || !row.folderId) { dequeue_(meetingId); return true; }
  const folder = DriveApp.getFolderById(row.folderId);
  const job = readJson_(folder, 'job.json');
  if (!job || (job.status !== 'queued' && job.status !== 'running')) { dequeue_(meetingId); return true; }
  const now = Date.now();
  job.status = 'running';

  const active = job.engines.filter(function (e) { return e.status === 'queued' || e.status === 'running'; });
  if (active.length) {
    const ready = active.filter(function (e) { return !(e.nextAt > now); })[0];
    if (!ready) return Math.min.apply(null, active.map(function (e) { return e.nextAt - now; }));
    engineStep_(job, ready, row, folder);
    saveJob_(folder, job);
    return true;
  }

  if (job.analysis.status === 'queued' || job.analysis.status === 'running') {
    if (job.analysis.nextAt > now) return job.analysis.nextAt - now;
    analysisStep_(job, row, folder);
    saveJob_(folder, job);
    return true;
  }

  if (job.doc.status === 'queued') {
    if (job.analysis.status === 'done' && settingsGs_().autoDoc) {
      try { buildDoc_({ meetingId: row.id }); job.doc.status = 'done'; }
      catch (e) { job.doc.status = 'error'; job.doc.error = String(e.message || e); }
    } else job.doc.status = 'skip';
  }

  const anyOk = job.engines.some(function (e) { return e.status === 'done'; }) || job.analysis.status === 'done';
  job.status = anyOk ? 'done' : 'error';
  job.finishedAt = new Date().toISOString();
  saveJob_(folder, job);
  withLock_(function () {
    const r = findById_(sh, meetingId);
    if (r) { r.status = finalStatus_(r, folder); r.updatedAt = new Date().toISOString(); writeRow_(sh, M_COLS, r); }
  });
  dequeue_(meetingId);
  return true;
}

function saveJob_(folder, job) {
  job.updatedAt = new Date().toISOString();
  writeJson_(folder, 'job.json', job);
}

function groupsFor_(id, audio) {
  const n = GROUP_SIZE[id] || 1, out = [];
  const pieces = (audio && audio.pieces) || [];
  for (let i = 0; i < pieces.length; i += n) out.push(pieces.slice(i, i + n));
  return out;
}

function engineStep_(job, eng, row, folder) {
  const audio = readJson_(folder, 'audio.json');
  const groups = groupsFor_(eng.id, audio);
  eng.total = groups.length;
  if (eng.status === 'queued') { eng.status = 'running'; eng.startedAt = new Date().toISOString(); }
  try {
    const ctx = ctxFor_(row);
    const g = groups[eng.done];
    if (!g) throw new Error('Аудіо змінилося під час обробки — запусти розпізнавання ще раз');
    const segs = ENGINE_STEP[eng.id](g, eng.done, groups.length, ctx, eng);
    eng.segs = eng.segs.concat(segs);
    eng.done++;
    eng.attempts = 0; eng.nextAt = 0;
    eng.detail = groups.length > 1 ? 'частина ' + eng.done + ' з ' + groups.length : '';
    if (eng.done >= groups.length) {
      const tr = finalizeTranscript_(eng, audio, ctx);
      saveTranscript_({ meetingId: row.id, engine: eng.id, transcript: tr });
      eng.status = 'done';
      eng.words = tr.text.split(/\s+/).filter(Boolean).length;
      eng.endedAt = new Date().toISOString();
      eng.detail = eng.words + ' слів' + (tr.warnings.length ? ' · ' + tr.warnings[0] : '');
      eng.segs = [];
    }
  } catch (e) {
    eng.attempts++;
    if (e.retryable && eng.attempts <= RETRY_WAITS.length) {
      eng.nextAt = Date.now() + RETRY_WAITS[eng.attempts - 1];
      eng.detail = String(e.message || e).slice(0, 160) + ' — повтор ' + (eng.attempts + 1) + '…';
    } else {
      eng.status = 'error';
      eng.error = String(e.message || e).slice(0, 400);
      eng.segs = [];
      eng.endedAt = new Date().toISOString();
    }
  }
}

/* ---------- HTTP до сервісів ---------- */

function fetchJson_(name, url, opt) {
  let res;
  try { res = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true }, opt)); }
  catch (e) {
    const err = new Error(name + ': сервіс не відповів вчасно');
    err.retryable = true;
    throw err;
  }
  const code = res.getResponseCode(), text = res.getContentText('UTF-8');
  if (code !== 200) throw httpErr_(name, code, text);
  return JSON.parse(text);
}

function httpErr_(name, code, text) {
  let msg = text || '';
  try {
    const j = JSON.parse(text);
    msg = (j.error && (j.error.message || j.error.type || JSON.stringify(j.error))) || (j.detail && (j.detail.message || j.detail)) || j.message || msg;
    if (typeof msg !== 'string') msg = JSON.stringify(msg);
  } catch (e) { }
  const hint = code === 401 || code === 403 ? ' Перевір ключ API в Налаштуваннях.' : code === 429 ? ' Перевищено ліміт.' : '';
  const err = new Error(name + ': помилка ' + code + '. ' + String(msg).slice(0, 300) + hint);
  err.status = code; err.body = msg;
  err.retryable = [429, 500, 502, 503, 504].indexOf(code) >= 0;
  return err;
}

function postForm_(name, url, headers, form) {
  // Запит multipart/form-data з аудіо (байти) і полями.
  const mp = multipart_(form);
  return { opt: { method: 'post', headers: headers, contentType: 'multipart/form-data; boundary=' + mp.boundary, payload: mp.bytes } };
}

function groupBytes_(group) {
  let bytes = [];
  group.forEach(function (p) { bytes = bytes.concat(DriveApp.getFileById(p.id).getBlob().getBytes()); });
  return bytes;
}

/* ---------- Допоміжне (так само, як у додатку) ---------- */

function parseTsGs_(s) {
  if (typeof s === 'number') return s;
  const parts = String(s || '0').trim().replace(',', '.').split(':').map(Number);
  if (parts.some(isNaN)) return 0;
  return parts.reduce(function (acc, p) { return acc * 60 + p; }, 0);
}

function parseJsonLooseGs_(text) {
  const t = String(text || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch (e) { }
  const a = t.indexOf('{');
  if (a >= 0) {
    const b = t.lastIndexOf('}');
    try { return JSON.parse(t.slice(a, b + 1)); } catch (e) { }
    let pos = t.length;
    for (let k = 0; k < 400; k++) {
      pos = t.lastIndexOf('}', pos - 1);
      if (pos <= a) break;
      const tails = [']}', ']}]}', '}'];
      for (let i = 0; i < tails.length; i++) {
        try { const j = JSON.parse(t.slice(a, pos + 1) + tails[i]); j.__truncated = true; return j; } catch (e) { }
      }
    }
  }
  throw new Error('Не вдалося розібрати відповідь моделі');
}

function hintTermsGs_(ctx, max) {
  const out = [];
  (ctx.participants || []).forEach(function (p) {
    if (p.name) out.push(p.name);
    String(p.aliases || '').split(/[,;]/).forEach(function (x) { if (x.trim()) out.push(x.trim()); });
  });
  (ctx.glossary || []).forEach(function (t) { if (t) out.push(t); });
  const seen = {}, res = [];
  out.forEach(function (t) { if (!seen[t]) { seen[t] = true; res.push(t); } });
  return res.slice(0, max || 100);
}

function geminiPromptGs_(ctx, idx, total, prevTail) {
  const people = (ctx.participants || []).map(function (p) { return '- ' + p.name + (p.role ? ' — ' + p.role : ''); }).join('\n') || '- (не вказано)';
  const lang = ctx.settings.language === 'uk' ? 'Основна мова — українська.' : '';
  let s = 'Зроби дослівну розшифровку аудіозапису робочої наради на виробництві.\n' +
    lang + ' Пиши мовою оригіналу: не перекладай, не виправляй стиль; якщо хтось говорить російською чи суржиком — записуй як є.\n' +
    'Розбий текст на репліки за спікерами. Однаковий голос — завжди та сама мітка.\n' +
    'Учасники наради:\n' + people + '\n' +
    "Якщо з розмови зрозуміло, хто говорить (звертаються на ім'я, згадують посаду), підписуй репліку іменем учасника. Якщо ні — «Спікер 1», «Спікер 2» і т. д.\n" +
    'Для кожної репліки вкажи час початку у форматі MM:SS від початку ЦЬОГО аудіо.\n' +
    'Нічого не скорочуй, не узагальнюй і не пропускай. Нерозбірливе місце познач [нерозбірливо].';
  const terms = hintTermsGs_(ctx, 150);
  if (terms.length) s += '\nСлова й назви, які можуть звучати: ' + terms.join(', ') + '.';
  if (total > 1) {
    s += '\nЦе фрагмент ' + (idx + 1) + ' з ' + total + ' одного запису.';
    if (prevTail) s += ' Кінець попереднього фрагмента (для узгодження міток спікерів):\n' + prevTail;
  }
  return s;
}

const SEG_SCHEMA_GS = {
  type: 'OBJECT',
  properties: { segments: { type: 'ARRAY', items: { type: 'OBJECT',
    properties: { start: { type: 'STRING', description: 'час початку репліки MM:SS від початку цього аудіо' }, speaker: { type: 'STRING' }, text: { type: 'STRING' } },
    required: ['start', 'speaker', 'text'], propertyOrdering: ['start', 'speaker', 'text'] } } },
  required: ['segments']
};

/* ---------- Двигуни: один крок = одна група частин ---------- */

const ENGINE_STEP = {
  gemini: function (group, idx, total, ctx, eng) {
    const p = group[0];
    const model = ctx.settings.models.gemini;
    const body = {
      contents: [{ role: 'user', parts: [
        { inlineData: { mimeType: 'audio/mp3', data: Utilities.base64Encode(DriveApp.getFileById(p.id).getBlob().getBytes()) } },
        { text: geminiPromptGs_(ctx, idx, total, eng.prevTail || '') }] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: SEG_SCHEMA_GS, maxOutputTokens: 65536 }
    };
    if (/2\.5-flash/.test(model)) body.generationConfig.thinkingConfig = { thinkingBudget: 0 };
    const data = fetchJson_('Gemini', 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent',
      { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': ctx.keys.gemini }, payload: JSON.stringify(body) });
    const cand = (data.candidates || [])[0];
    if (!cand) throw new Error('Gemini не повернув відповіді');
    const text = ((cand.content && cand.content.parts) || []).map(function (x) { return x.text || ''; }).join('');
    const j = parseJsonLooseGs_(text);
    if (cand.finishReason && cand.finishReason !== 'STOP') eng.warnings.push('Частина ' + (idx + 1) + ': відповідь обірвалась (' + cand.finishReason + ')');
    const segs = (j.segments || []).filter(function (s) { return s && s.text; }).map(function (s) {
      return { start: p.o + Math.min(parseTsGs_(s.start), p.d), speaker: String(s.speaker || '').trim(), text: String(s.text).trim() };
    });
    segs.sort(function (a, b) { return a.start - b.start; });
    segs.forEach(function (s, k) { s.end = k + 1 < segs.length ? segs[k + 1].start : p.o + p.d; });
    eng.prevTail = segs.slice(-4).map(function (s) { return s.speaker + ': ' + s.text.slice(0, 160); }).join('\n');
    return segs;
  },

  groq: function (group, idx, total, ctx, eng) {
    const o = group[0].o;
    const form = [{ name: 'file', bytes: groupBytes_(group), filename: 'audio.mp3', mimeType: 'audio/mpeg' },
      { name: 'model', value: ctx.settings.models.groq }, { name: 'response_format', value: 'verbose_json' }, { name: 'temperature', value: '0' }];
    if (ctx.settings.language === 'uk') form.push({ name: 'language', value: 'uk' });
    const prompt = hintTermsGs_(ctx, 60).join(', ').slice(0, 600);
    if (prompt) form.push({ name: 'prompt', value: prompt });
    const req = postForm_('Groq', '', { Authorization: 'Bearer ' + ctx.keys.groq }, form);
    const j = fetchJson_('Groq', 'https://api.groq.com/openai/v1/audio/transcriptions', req.opt);
    const out = [];
    (j.segments || []).forEach(function (s) { const t = String(s.text || '').trim(); if (t) out.push({ start: o + s.start, end: o + s.end, speaker: '', text: t }); });
    if (!out.length && j.text) out.push({ start: o, end: o + group.reduce(function (a, p) { return a + p.d; }, 0), speaker: '', text: String(j.text).trim() });
    return out;
  },

  openai: function (group, idx, total, ctx, eng) {
    const p = group[0];
    const model = ctx.settings.models.openai;
    const diarize = /diarize/.test(model);
    const bytes = groupBytes_(group);
    let form = [{ name: 'file', bytes: bytes, filename: 'audio.mp3', mimeType: 'audio/mpeg' }, { name: 'model', value: model }];
    if (diarize) form = form.concat([{ name: 'response_format', value: 'diarized_json' }, { name: 'chunking_strategy', value: 'auto' }]);
    else {
      form.push({ name: 'response_format', value: 'json' });
      const kw = hintTermsGs_(ctx, 60);
      if (kw.length) form.push({ name: 'prompt', value: 'Робоча нарада на виробництві. Назви й імена: ' + kw.join(', ') });
    }
    const withLang = ctx.settings.language === 'uk' ? form.concat([{ name: diarize ? 'language' : 'languages[]', value: 'uk' }]) : form;
    let j;
    try {
      j = fetchJson_('OpenAI', 'https://api.openai.com/v1/audio/transcriptions', postForm_('OpenAI', '', { Authorization: 'Bearer ' + ctx.keys.openai }, withLang).opt);
    } catch (e) {
      if (e.status === 400 && withLang !== form && /language/i.test(String(e.body))) {
        j = fetchJson_('OpenAI', 'https://api.openai.com/v1/audio/transcriptions', postForm_('OpenAI', '', { Authorization: 'Bearer ' + ctx.keys.openai }, form).opt);
      } else throw e;
    }
    const out = [];
    if (diarize && j.segments) {
      j.segments.forEach(function (s) {
        const t = String(s.text || '').trim(); if (!t) return;
        const spk = String(s.speaker === undefined || s.speaker === null ? '' : s.speaker);
        out.push({ start: p.o + (s.start || 0), end: p.o + (s.end || 0), speaker: total > 1 ? (idx + 1) + '·' + spk : spk, text: t });
      });
    } else if (j.text) out.push({ start: p.o, end: p.o + p.d, speaker: '', text: String(j.text).trim() });
    if (diarize && total > 1 && idx === 0) eng.warnings.push('Запис оброблено частинами: мітки спікерів у різних частинах незалежні (1·A, 2·A…).');
    return out;
  },

  elevenlabs: function (group, idx, total, ctx, eng) {
    const o = group[0].o;
    const base = [{ name: 'file', bytes: groupBytes_(group), filename: 'audio.mp3', mimeType: 'audio/mpeg' },
      { name: 'model_id', value: ctx.settings.models.elevenlabs }, { name: 'diarize', value: 'true' },
      { name: 'timestamps_granularity', value: 'word' }, { name: 'tag_audio_events', value: 'false' }];
    if (ctx.settings.language === 'uk') base.push({ name: 'language_code', value: 'uk' });
    const terms = hintTermsGs_(ctx, 100).filter(function (t) { return t.length <= 50; });
    const url = 'https://api.elevenlabs.io/v1/speech-to-text';
    let j;
    try {
      j = fetchJson_('ElevenLabs', url, postForm_('ElevenLabs', '', { 'xi-api-key': ctx.keys.elevenlabs }, base.concat(terms.map(function (t) { return { name: 'keyterms', value: t }; }))).opt);
    } catch (e) {
      if (terms.length && (e.status === 400 || e.status === 422)) {
        if (eng.warnings.indexOf('Словник не прийнято сервісом, розпізнано без нього.') < 0) eng.warnings.push('Словник не прийнято сервісом, розпізнано без нього.');
        j = fetchJson_('ElevenLabs', url, postForm_('ElevenLabs', '', { 'xi-api-key': ctx.keys.elevenlabs }, base).opt);
      } else throw e;
    }
    const out = [];
    let cur = null;
    (j.words || []).forEach(function (w) {
      if (w.type === 'audio_event') return;
      const spk = (total > 1 ? (idx + 1) + '·' : '') + String(w.speaker_id || '');
      if (w.type === 'spacing') { if (cur) cur.text += w.text; return; }
      const st = o + (w.start || 0), en = o + (w.end || w.start || 0);
      const brk = !cur || cur.speaker !== spk || st - cur.end > 2.5 || (cur.text.length > 600 && /[.!?…]\s*$/.test(cur.text));
      if (brk) { if (cur) out.push(cur); cur = { start: st, end: en, speaker: spk, text: w.text }; }
      else { cur.text += w.text; cur.end = en; }
    });
    if (cur) out.push(cur);
    if (!(j.words || []).length && j.text) out.push({ start: o, end: o, speaker: '', text: j.text });
    out.forEach(function (s) { s.text = String(s.text).replace(/\s+/g, ' ').trim(); });
    if (total > 1 && idx === 0) eng.warnings.push('Запис оброблено частинами по ~15 хв: мітки спікерів у різних частинах незалежні.');
    return out.filter(function (s) { return s.text; });
  }
};

function finalizeTranscript_(eng, audio, ctx) {
  let segs = eng.segs.slice().sort(function (a, b) { return a.start - b.start; });
  if (eng.id === 'groq') {
    const out = [];
    segs.forEach(function (s) {
      const last = out[out.length - 1];
      if (last && s.start - last.start < 45 && s.start - last.end < 2.5) { last.text = (last.text + ' ' + s.text).trim(); last.end = s.end; }
      else out.push({ start: s.start, end: s.end, speaker: '', text: s.text });
    });
    segs = out;
  } else if (eng.id === 'openai') {
    const out = [];
    segs.forEach(function (s) {
      const last = out[out.length - 1];
      if (last && last.speaker && last.speaker === s.speaker && s.start - last.end < 1.5 && (last.end - last.start) < 60) { last.text += ' ' + s.text; last.end = s.end; }
      else out.push({ start: s.start, end: s.end, speaker: s.speaker, text: s.text });
    });
    segs = out;
  }
  segs = segs.filter(function (s) { return s.text && String(s.text).trim(); });
  if (!segs.length) throw new Error(ENGINE_NAMES[eng.id] + ': сервіс повернув порожній текст. Перевір, чи в записі чути мову.');
  const speakers = [];
  segs.forEach(function (s) { if (s.speaker && speakers.indexOf(s.speaker) < 0) speakers.push(s.speaker); });
  return {
    engine: eng.id, engineName: ENGINE_NAMES[eng.id], model: ctx.settings.models[eng.id], language: ctx.settings.language,
    createdAt: new Date().toISOString(), durationSec: audio.duration, pieces: eng.total,
    elapsedSec: Math.round((Date.now() - new Date(eng.startedAt).getTime()) / 1000),
    segments: segs.map(function (s) { return { start: Math.round(s.start * 100) / 100, end: Math.round((s.end || s.start) * 100) / 100, speaker: s.speaker || '', text: s.text }; }),
    text: segs.map(function (s) { return s.text; }).join(' '),
    speakers: speakers, warnings: eng.warnings || [], via: 'background'
  };
}

/* ---------- Підсумок у фоні ---------- */

const ANALYSIS_SCHEMA_GS = {
  type: 'OBJECT',
  properties: {
    summary: { type: 'STRING' },
    decisions: { type: 'ARRAY', items: { type: 'STRING' } },
    tasks: { type: 'ARRAY', items: { type: 'OBJECT', properties: { owner: { type: 'STRING' }, task: { type: 'STRING' }, due: { type: 'STRING' }, dueDate: { type: 'STRING', description: 'РРРР-ММ-ДД або порожньо' } }, required: ['owner', 'task', 'due', 'dueDate'] } },
    openQuestions: { type: 'ARRAY', items: { type: 'STRING' } },
    facts: { type: 'ARRAY', items: { type: 'STRING' } },
    reportedDone: { type: 'ARRAY', items: { type: 'STRING' } }
  },
  required: ['summary', 'decisions', 'tasks', 'openQuestions', 'facts']
};

function fmtTsGs_(sec) {
  sec = Math.round(sec || 0);
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  const p = function (n) { return (n < 10 ? '0' : '') + n; };
  return h ? h + ':' + p(m) + ':' + p(s) : p(m) + ':' + p(s);
}

function analysisPromptGs_(meta, transcriptText, ctx) {
  const people = (ctx.participants || []).map(function (p) { return p.name + (p.role ? ' (' + p.role + ')' : ''); }).join(', ') || 'не вказано';
  const terms = (ctx.glossary || []).join(', ');
  const carry = (ctx.carryTasks || []).map(function (t, i) { return '[k' + (i + 1) + '] ' + t.task + ' — ' + (t.owner || '—') + (t.meetingStart ? ' (з наради ' + String(t.meetingStart).slice(0, 10) + ')' : ''); }).join('\n');
  return 'Ти — помічник керівника виробництва. Нижче автоматична розшифровка робочої наради (можуть бути помилки розпізнавання).\n' +
    'Склади короткий структурований підсумок українською мовою.\n\n' +
    'Нарада: «' + meta.title + '», ' + meta.dateText + '. Дата наради: ' + meta.dateISO + ' (' + meta.weekday + ').\n' +
    'Учасники: ' + people + '.\n' +
    (terms ? 'Терміни й назви, які можуть зустрічатися: ' + terms + '.\n' : '') + (carry ? 'Відкриті задачі з попередніх нарад:\n' + carry + '\n' : '') + '\n' +
    'Поверни ТІЛЬКИ JSON такого вигляду:\n' +
    '{"summary": "3–6 речень: про що нарада і до чого дійшли",\n "decisions": ["прийняте рішення"],\n' +
    ' "tasks": [{"owner": "хто", "task": "що зробити", "due": "термін словами або порожній рядок", "dueDate": "РРРР-ММ-ДД або порожній рядок"}],\n' +
    ' "openQuestions": ["питання, що лишилося відкритим"],\n "facts": ["важливі цифри, дати, обсяги, показники"],\n "reportedDone": []}\n\n' +
    'Правила:\n- Тільки те, що є в розмові. Нічого не вигадуй.\n' +
    '- Задача — це конкретна дія, яку хтось взяв на себе або доручили. Власника пиши іменем учасника; якщо незрозуміло — «—».\n' +
    "- Якщо термін не названо — due і dueDate порожні. Якщо термін можна перевести в конкретну дату відносно дати наради («до п'ятниці», «завтра», «до 15-го»), запиши її в dueDate у форматі РРРР-ММ-ДД.\n" +
    '- reportedDone — мітки (k1, k2…) відкритих задач з попередніх нарад, про які на цій нараді прямо сказали, що їх виконано. Якщо таких немає або список задач не дано — [].\n' +
    '- Не повторюй у tasks задачі з попередніх нарад, якщо по них не дали нового доручення чи нового терміну.\n' +
    '- Очевидні помилки розпізнавання в іменах і термінах виправляй.\n' +
    '- Пиши коротко й по суті. Порожні списки — [].\n\n' +
    'Розшифровка:\n' + transcriptText;
}

function analysisStep_(job, row, folder) {
  const an = job.analysis;
  an.status = 'running';
  try {
    const sh = sheet_(SHEETS.meetings);
    // Основний текст: якщо раніше не було — перший обраний двигун, що розрізняє спікерів.
    let cur = findById_(sh, row.id);
    const done = job.engines.filter(function (e) { return e.status === 'done'; }).map(function (e) { return e.id; });
    if (!job.primaryBefore && done.length) {
      const best = done.filter(function (id) { return ENGINE_SPEAKERS[id]; })[0] || done[0];
      if (cur.primaryEngine !== best) {
        saveMeetingState_({ meetingId: row.id, primaryEngine: best });
        cur = findById_(sh, row.id);
      }
    }
    const engines = parseJson_(cur.engines, []);
    const primary = cur.primaryEngine && engines.indexOf(cur.primaryEngine) >= 0 ? cur.primaryEngine : engines[0];
    if (!primary) throw new Error('Немає тексту для підсумку');
    const tr = readJson_(folder, 'transcript_' + primary + '.json');
    const ctx = ctxFor_(cur);
    const engine = job.analysisEngine || ctx.settings.analysisEngine || 'gemini';
    const key = ctx.keys[engine];
    if (!key) throw new Error(ANALYZER_NAMES[engine] + ': не додано ключ API для підсумку');
    const model = ctx.settings.analysisModels[engine];
    const map = (parseJson_(cur.speakerMap, {}) || {})[primary] || {};
    const nameOf = function (label) { const v = map[label]; if (!v) return label; return ctx.people[v] ? ctx.people[v].name : v; };
    const text = tr.segments && tr.segments.length
      ? tr.segments.map(function (s) { return '[' + fmtTsGs_(s.start) + '] ' + (s.speaker ? nameOf(s.speaker) + ': ' : '') + s.text; }).join('\n')
      : (tr.text || '');
    try { ctx.carryTasks = carriedTasks_({ meetingId: row.id }); } catch (e) { ctx.carryTasks = []; }
    const tz = Session.getScriptTimeZone(), st = new Date(cur.start);
    const meta = { title: cur.title, dateText: Utilities.formatDate(st, tz, 'dd.MM.yyyy, HH:mm'), dateISO: Utilities.formatDate(st, tz, 'yyyy-MM-dd'),
      weekday: WEEKDAYS_UK[Number(Utilities.formatDate(st, tz, 'u')) % 7] };
    const prompt = analysisPromptGs_(meta, text, ctx);
    let out;
    if (engine === 'gemini') {
      const d = fetchJson_('Gemini', 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent',
        { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
          payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, responseMimeType: 'application/json', responseSchema: ANALYSIS_SCHEMA_GS, maxOutputTokens: 16384 } }) });
      out = (((d.candidates || [])[0] || {}).content || { parts: [] }).parts.map(function (x) { return x.text || ''; }).join('');
    } else if (engine === 'openai') {
      const d = fetchJson_('OpenAI', 'https://api.openai.com/v1/chat/completions', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({ model: model, messages: [{ role: 'user', content: prompt }], response_format: { type: 'json_object' } }) });
      out = ((d.choices || [])[0] || { message: {} }).message.content || '';
    } else {
      const d = fetchJson_('Claude', 'https://api.anthropic.com/v1/messages', { method: 'post', contentType: 'application/json',
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: model, max_tokens: 8000, messages: [{ role: 'user', content: prompt }] }) });
      out = (d.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    }
    const j = parseJsonLooseGs_(out);
    const arr = function (x) { return Array.isArray(x) ? x.map(function (v) { return typeof v === 'string' ? v : JSON.stringify(v); }).filter(Boolean) : []; };
    const analysis = {
      summary: String(j.summary || ''), decisions: arr(j.decisions),
      tasks: Array.isArray(j.tasks) ? j.tasks.map(function (t) {
        return { owner: String(t.owner || '—'), task: String(t.task || ''), due: String(t.due || ''), dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(t.dueDate || '')) ? String(t.dueDate) : '' };
      }).filter(function (t) { return t.task; }) : [],
      openQuestions: arr(j.openQuestions), facts: arr(j.facts),
      reportedDone: arr(j.reportedDone).map(function (k) { const m = /k?(\d+)/i.exec(k); const t = m && ctx.carryTasks[Number(m[1]) - 1]; return t ? t.id : null; }).filter(Boolean),
      engine: engine, engineName: ANALYZER_NAMES[engine], model: model, basedOn: primary, createdAt: new Date().toISOString(), via: 'background'
    };
    saveAnalysis_({ meetingId: row.id, analysis: analysis });
    an.status = 'done';
    an.tasks = analysis.tasks.length;
    an.detail = analysis.tasks.length + ' задач, ' + analysis.decisions.length + ' рішень' + (analysis.reportedDone.length ? ', ' + analysis.reportedDone.length + ' з минулих — схоже, виконано' : '');
  } catch (e) {
    an.attempts = (an.attempts || 0) + 1;
    if (e.retryable && an.attempts <= RETRY_WAITS.length) {
      an.status = 'queued';
      an.nextAt = Date.now() + RETRY_WAITS[an.attempts - 1];
      an.detail = String(e.message || e).slice(0, 160) + ' — повтор…';
    } else {
      an.status = 'error';
      an.error = String(e.message || e).slice(0, 400);
    }
  }
}
