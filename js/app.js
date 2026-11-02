// «Наради» — головний модуль інтерфейсу.

import {
  h, put, $, $$, clear, icon, toast, modal, confirmDialog, fmtDuration, fmtHuman, fmtTime, fmtDay, fmtDateShort,
  fmtDateFull, fmtWeekday, fmtBytes, ymd, hm, startOfWeek, addDays, sameDay, initials, colorFor, copyText, spinner
} from './ui.js';
import * as api from './api.js';
import { ENGINES, ANALYZERS, transcriptAsText } from './engines.js';
import { processAudio, runTranscription, runAnalysis, onJobs, jobFor, clearJob, speakerNamer, anyJobRunning } from './pipeline.js';
import { Recorder, pendingRecordings, assembleRecording, deleteRecording } from './recorder.js';
import { idb, getPrepared } from './store.js';
import { downloadFile } from './drive.js';
import { taskCard, editTask, sortTasks, isOverdue, weeklyReport, tasksText as tasksListText } from './tasks.js';
import { syncIndex, searchAll, queryWords, snippet, highlight } from './search.js';

const APP_VERSION = '1.1.1';
const view = $('#view');
const S = {
  week: startOfWeek(new Date()),
  evCache: new Map(),
  meetings: new Map(),
  recorder: null,
  recMeetingId: null,
  installEvt: null,
  player: null
};

/* =================== Маршрути =================== */

const routes = [
  [/^#\/calendar$/, viewCalendar],
  [/^#\/recordings$/, viewRecordings],
  [/^#\/people$/, viewPeople],
  [/^#\/settings$/, viewSettings],
  [/^#\/connect$/, viewConnect],
  [/^#\/new(?:\?(.*))?$/, (q) => viewEdit(null, new URLSearchParams(q || ''))],
  [/^#\/edit\/([^?]+)$/, (id) => viewEdit(decodeURIComponent(id))],
  [/^#\/m\/([^?]+)(?:\?(.*))?$/, (id, q) => viewMeeting(decodeURIComponent(id), new URLSearchParams(q || ''))],
  [/^#\/tasks$/, viewTasks],
  [/^#\/ev\/([^?]+)$/, (key) => viewMeetingEv(decodeURIComponent(key))],
  [/^#\/rec\/([^?]+)$/, (id) => viewRecord(decodeURIComponent(id))],
  [/^#\/import$/, viewImport]
];

let seq = 0;
const alive = (t) => t === seq;

function go(hash, replace = false) {
  if (replace) { history.replaceState(null, '', hash); render(); }
  else if (location.hash === hash) render();
  else location.hash = hash;
}

async function render() {
  const t = ++seq;
  let hash = location.hash || '#/calendar';
  if (hash === '#' || hash === '#/') hash = '#/calendar';
  if (!api.getConn() && !hash.startsWith('#/connect')) { go('#/connect', true); return; }
  setNav(hash);
  S.player = null;
  clear(view);
  window.scrollTo(0, 0);
  for (const [re, fn] of routes) {
    const mm = hash.match(re);
    if (mm) {
      try { await fn(...mm.slice(1)); }
      catch (e) { if (alive(t)) showError(e, () => render()); }
      return;
    }
  }
  go('#/calendar', true);
}

function setNav(hash) {
  const sec = hash.startsWith('#/recordings') ? 'recordings' : hash.startsWith('#/tasks') ? 'tasks' : hash.startsWith('#/people') ? 'people'
    : hash.startsWith('#/settings') || hash.startsWith('#/connect') ? 'settings' : 'calendar';
  $$('.nav a').forEach(a => a.classList.toggle('active', a.dataset.sec === sec));
  updateRecPill();
}

function showError(e, retry) {
  console.error(e);
  view.appendChild(h('div', { class: 'error-box' },
    h('p', null, e.message || String(e)),
    e.code === 'no_conn' || e.code === 'unauthorized' ? h('a', { class: 'btn', href: '#/connect' }, 'До підключення') : null,
    retry ? h('button', { class: 'btn', onclick: retry }, 'Спробувати ще') : null));
}

function topbar(title, { back, actions } = {}) {
  return h('header', { class: 'topbar' },
    back ? h('button', { class: 'icon-btn', 'aria-label': 'Назад', onclick: () => (history.length > 1 ? history.back() : go(back)) }, icon('back')) : null,
    h('h1', null, title),
    h('div', { class: 'topbar-actions' }, actions || []));
}

/* =================== Календар =================== */

function weekLabel(ws) {
  const we = addDays(ws, 6);
  const sameMonth = ws.getMonth() === we.getMonth();
  const f = (d, o) => new Intl.DateTimeFormat('uk-UA', o).format(d);
  return sameMonth ? `${ws.getDate()}–${f(we, { day: 'numeric', month: 'long' })}` : `${f(ws, { day: 'numeric', month: 'short' })} – ${f(we, { day: 'numeric', month: 'short' })}`;
}

function evKey(m) {
  const key = (m.eventId || 'x') + '|' + m.start;
  S.evCache.set(key, m);
  return encodeURIComponent(key);
}

function meetingHref(m) { return m.id ? '#/m/' + encodeURIComponent(m.id) : '#/ev/' + evKey(m); }

function badges(m) {
  const b = [];
  if (m.recurring) b.push(h('span', { class: 'badge muted', title: 'Повторювана' }, '↻'));
  if (m.audioFileId) b.push(h('span', { class: 'badge rec' }, icon('wave'), m.durationSec ? fmtHuman(m.durationSec) : 'запис'));
  if (m.engines && m.engines.length) b.push(h('span', { class: 'badge' }, 'текст' + (m.engines.length > 1 ? ' ×' + m.engines.length : '')));
  if (m.status === 'analyzed' || m.docId) b.push(h('span', { class: 'badge ok' }, 'підсумок'));
  const j = m.id && jobFor(m.id);
  if (j && j.running) b.push(h('span', { class: 'badge run' }, 'обробка…'));
  return b;
}

function personChips(ids, max = 4) {
  const people = api.peopleSync();
  const ps = (ids || []).map(id => people.find(p => p.id === id)).filter(Boolean);
  if (!ps.length) return null;
  return h('div', { class: 'faces' },
    ps.slice(0, max).map(p => h('span', { class: 'face', style: { background: colorFor(p.id) }, title: p.name }, initials(p.name))),
    ps.length > max ? h('span', { class: 'face more' }, '+' + (ps.length - max)) : null);
}

function meetingCard(m, { showDate = false } = {}) {
  return h('a', { class: 'card mcard', href: meetingHref(m) },
    h('div', { class: 'mc-time' },
      showDate ? h('b', null, fmtDateShort(m.start)) : null,
      h('span', null, fmtTime(m.start)), h('small', null, fmtTime(m.end))),
    h('div', { class: 'mc-main' },
      h('div', { class: 'mc-title' }, m.title || 'Без назви'),
      h('div', { class: 'mc-meta' }, personChips(m.participants), badges(m))));
}

async function viewCalendar() {
  const t = seq;
  const ws = S.week;
  const today = new Date();
  view.appendChild(topbar('Календар', {
    actions: [h('button', { class: 'btn ghost small', onclick: () => { S.week = startOfWeek(new Date()); render(); } }, 'Сьогодні')]
  }));

  view.appendChild(h('div', { class: 'weeknav' },
    h('button', { class: 'icon-btn', 'aria-label': 'Попередній тиждень', onclick: () => { S.week = addDays(S.week, -7); render(); } }, icon('left')),
    h('div', { class: 'weeklabel' }, weekLabel(ws)),
    h('button', { class: 'icon-btn', 'aria-label': 'Наступний тиждень', onclick: () => { S.week = addDays(S.week, 7); render(); } }, icon('right'))));

  const strip = h('div', { class: 'daystrip' });
  for (let i = 0; i < 7; i++) {
    const d = addDays(ws, i);
    strip.appendChild(h('button', {
      class: 'day' + (sameDay(d, today) ? ' today' : ''), dataset: { day: ymd(d) },
      onclick: () => { const sec = $(`#day-${ymd(d)}`); if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' }); else go('#/new?d=' + ymd(d)); }
    }, h('small', null, fmtWeekday(d)), h('b', null, d.getDate()), h('i', { class: 'dot' })));
  }
  view.appendChild(strip);

  view.appendChild(h('div', { class: 'quick' },
    h('a', { class: 'btn primary', href: '#/new?d=' + ymd(sameDay(ws, startOfWeek(today)) ? today : ws) }, icon('plus'), 'Нарада'),
    h('button', { class: 'btn rec', onclick: quickRecord }, icon('mic'), 'Записати зараз'),
    h('label', { class: 'btn' }, icon('upload'), 'Файл з диктофона',
      h('input', { type: 'file', accept: 'audio/*,video/*,.m4a,.mp3,.wav,.ogg,.webm,.aac,.amr,.3gp', hidden: true, onchange: (e) => quickUpload(e.target.files[0]) }))));

  const pend = h('div');
  view.appendChild(pend);
  showPending(pend);

  const list = h('div', { class: 'daylist' }, spinner());
  view.appendChild(list);
  const [meetings] = await Promise.all([api.call('listMeetings', { from: ws.toISOString(), to: addDays(ws, 7).toISOString() }), api.getPeople().catch(() => [])]);
  if (!alive(t)) return;
  clear(list);
  meetings.forEach(m => { if (m.id) S.meetings.set(m.id, m); });
  const byDay = {};
  for (const m of meetings) (byDay[ymd(m.start)] = byDay[ymd(m.start)] || []).push(m);
  Object.keys(byDay).forEach(d => { const b = strip.querySelector(`[data-day="${d}"]`); if (b) b.classList.add('has'); });
  if (!meetings.length) {
    list.appendChild(h('div', { class: 'empty' }, h('p', null, 'На цьому тижні нарад немає.'), h('p', { class: 'hint' }, 'Додай нараду кнопкою «Нарада» — вона зʼявиться і в Google Календарі.')));
    return;
  }
  for (let i = 0; i < 7; i++) {
    const d = addDays(ws, i), key = ymd(d);
    if (!byDay[key]) continue;
    list.appendChild(h('section', { class: 'dayblock', id: 'day-' + key },
      h('h2', { class: sameDay(d, today) ? 'today' : '' }, fmtDay(d), sameDay(d, today) ? h('span', { class: 'tag' }, 'сьогодні') : null),
      byDay[key].map(m => meetingCard(m))));
  }
  const todaySec = $(`#day-${ymd(today)}`);
  if (todaySec && sameDay(startOfWeek(today), ws)) todaySec.scrollIntoView({ block: 'start' });
}

async function quickRecord() {
  if (S.recorder) { go('#/rec/' + encodeURIComponent(S.recMeetingId)); return; }
  const now = new Date(); now.setSeconds(0, 0);
  try {
    const m = await api.call('saveMeeting', { meeting: { title: 'Нарада ' + hm(now), start: now.toISOString(), end: new Date(now.getTime() + 3600e3).toISOString(), participants: [], notes: '' } });
    S.meetings.set(m.id, m);
    go('#/rec/' + encodeURIComponent(m.id));
  } catch (e) { toast(e.message, 'err'); }
}

async function quickUpload(file) {
  if (!file) return;
  await idb.put('shared', { id: 's_' + Date.now(), file, name: file.name, type: file.type, lastModified: file.lastModified, size: file.size });
  go('#/import');
}

async function showPending(box) {
  const list = (await pendingRecordings()).filter(r => !(S.recorder && r.recId === S.recorder.recId));
  clear(box);
  for (const r of list) {
    box.appendChild(h('div', { class: 'banner warn' },
      h('div', null, h('b', null, 'Незбережений запис'), h('div', { class: 'hint' }, `${fmtDateShort(r.startedAt)}, ${fmtTime(r.startedAt)} · ${fmtHuman(r.elapsed || 0)}`)),
      h('div', { class: 'row' },
        h('button', { class: 'btn small primary', onclick: () => recoverRecording(r) }, 'Обробити'),
        h('button', { class: 'btn small ghost', onclick: async () => { if (await confirmDialog('Видалити цей запис з телефона назавжди?', 'Видалити', true)) { await deleteRecording(r.recId); showPending(box); } } }, 'Видалити'))));
  }
}

async function recoverRecording(r) {
  const { blob } = await assembleRecording(r.recId);
  if (!blob.size) { toast('Запис порожній', 'err'); await deleteRecording(r.recId); render(); return; }
  let m = null;
  if (r.meetingId) { try { m = await api.call('getMeeting', { id: r.meetingId }); } catch { m = null; } }
  if (!m) {
    const st = new Date(r.startedAt);
    m = await api.call('saveMeeting', { meeting: { title: 'Нарада ' + hm(st), start: st.toISOString(), end: new Date(st.getTime() + Math.max(600, r.elapsed || 0) * 1000).toISOString(), participants: [] } });
  }
  startProcessing(m, blob, { fileName: 'запис.webm', onUploaded: () => deleteRecording(r.recId) });
}

/* =================== Записи (архів) =================== */

async function viewRecordings() {
  const t = seq;
  view.appendChild(topbar('Записи'));
  const full = backendAtLeast('1.1.0');
  const search = h('input', { type: 'search', class: 'input', placeholder: full ? 'Пошук у назвах і текстах нарад…' : 'Пошук за назвою…', value: S.searchQ || '' });
  const status = h('p', { class: 'hint' });
  view.appendChild(h('div', { class: 'pad' }, search, status));
  const list = h('div', { class: 'stack' }, spinner());
  view.appendChild(list);
  const now = new Date();
  const [items] = await Promise.all([
    api.call('listMeetings', { from: addDays(now, -365).toISOString(), to: addDays(now, 1).toISOString(), onlyRecorded: true }),
    api.getPeople().catch(() => [])
  ]);
  if (!alive(t)) return;
  items.reverse();
  items.forEach(m => S.meetings.set(m.id, m));
  let ready = false, drawSeq = 0;
  const draw = async () => {
    const my = ++drawSeq;
    const raw = search.value.trim();
    S.searchQ = raw;
    const q = raw.toLowerCase();
    if (full && ready && queryWords(raw).length) {
      const res = await searchAll(raw, api.peopleSync());
      if (my !== drawSeq || !alive(t)) return;
      clear(list);
      if (!res.length) { put(list, h('div', { class: 'empty' }, 'Нічого не знайдено')); return; }
      const words = queryWords(raw);
      const hitsTotal = res.reduce((n, r) => n + r.hits.length, 0);
      put(list, h('p', { class: 'hint' }, `Нарад: ${res.length}, збігів у текстах: ${hitsTotal}`));
      res.forEach(r => put(list, searchResultCard(r, words, raw)));
      return;
    }
    clear(list);
    const shown = items.filter(m => !q || (m.title || '').toLowerCase().includes(q));
    if (!shown.length) { put(list, h('div', { class: 'empty' }, q ? 'Нічого не знайдено' : 'Записів ще немає. Запиши нараду або завантаж файл з диктофона.')); return; }
    shown.forEach(m => list.appendChild(meetingCard(m, { showDate: true })));
  };
  let timer = null;
  search.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(draw, 200); });
  draw();
  if (full) {
    status.textContent = 'Готую пошук по текстах…';
    syncIndex((n) => { if (alive(t) && n) status.textContent = `Завантажую тексти для пошуку: ${n}…`; })
      .then(() => { if (!alive(t)) return; ready = true; status.textContent = ''; if (search.value.trim()) draw(); })
      .catch((e) => { if (alive(t)) status.textContent = 'Пошук по текстах зараз недоступний: ' + e.message; });
  } else {
    status.textContent = 'Пошук по текстах зʼявиться після оновлення скрипта Google до версії 1.1.0.';
  }
}

function searchResultCard(r, words, raw) {
  const href = (tt) => `#/m/${encodeURIComponent(r.id)}?t=${tt}&q=${encodeURIComponent(raw)}`;
  return h('div', { class: 'card sres' },
    h('a', { class: 'sres-head', href: '#/m/' + encodeURIComponent(r.id) },
      h('b', null, highlight(r.title || 'Без назви', words)), h('span', { class: 'hint' }, `${fmtDateShort(r.start)}, ${fmtTime(r.start)}`)),
    r.sumHit ? h('p', { class: 'hint sres-sum' }, highlight(snippet(r.summary, words), words)) : null,
    r.hits.slice(0, 5).map(x => h('a', { class: 'sres-hit', href: href(x.t) },
      h('span', { class: 'ts' }, fmtDuration(x.t)),
      h('span', null, x.speaker ? h('b', null, x.speaker + ': ') : null, highlight(snippet(x.text, words), words)))),
    r.hits.length > 5 ? h('a', { class: 'linkish small', href: href(r.hits[5].t) }, `ще ${r.hits.length - 5} збігів у цій нараді`) : null);
}

function backendAtLeast(v) {
  const c = api.cfg();
  return !!(c && c.info && verGte(c.info.version, v));
}

/* =================== Картка наради =================== */

async function viewMeetingEv(key) {
  const m = S.evCache.get(key);
  if (!m) { go('#/calendar', true); return; }
  renderMeeting(m);
}

async function viewMeeting(id, params) {
  const t = seq;
  const focus = params && params.get('t') !== null ? { t: Number(params.get('t')) || 0, q: params.get('q') || '' } : null;
  const cached = S.meetings.get(id);
  if (cached) renderMeeting(cached, focus);
  else view.appendChild(spinner());
  const [m] = await Promise.all([api.call('getMeeting', { id }), api.getPeople().catch(() => [])]);
  if (!alive(t)) return;
  S.meetings.set(m.id, m);
  // Нарада не змінилась — картка вже показана й тексти вантажаться, перемальовувати не треба.
  if (cached && JSON.stringify(cached) === JSON.stringify(m) && view.querySelector('.meeting')) return;
  clear(view);
  renderMeeting(m, focus);
}

async function ensureRow(m) {
  if (m.id) return m;
  const saved = await api.call('saveMeeting', { meeting: { ...m, participants: m.participants || [] } });
  S.meetings.set(saved.id, saved);
  history.replaceState(null, '', '#/m/' + encodeURIComponent(saved.id));
  return saved;
}

let MV = null; // стан відкритої картки

function renderMeeting(m, focus) {
  const t = seq;
  MV = { m, transcripts: null, analysis: null, tab: focus ? 'text' : (MV && MV.m && MV.m.id === m.id ? MV.tab : null), engine: null, t, focus: focus || null };
  const st = new Date(m.start), en = new Date(m.end);
  const people = api.peopleSync();
  const parts = (m.participants || []).map(id => people.find(p => p.id === id)).filter(Boolean);

  const root = h('div', { class: 'meeting' });
  root.appendChild(topbar('', {
    back: '#/calendar',
    actions: [
      h('button', { class: 'icon-btn', 'aria-label': 'Редагувати', title: 'Редагувати', onclick: async () => { const r = await ensureRow(m); go('#/edit/' + encodeURIComponent(r.id)); } }, icon('edit')),
      m.id ? h('button', { class: 'icon-btn', 'aria-label': 'Видалити', title: 'Видалити', onclick: () => deleteMeetingFlow(m) }, icon('trash')) : null
    ]
  }));
  root.appendChild(h('div', { class: 'mhead' },
    h('h1', null, m.title || 'Без назви'),
    h('p', { class: 'when' }, `${fmtDay(st)} · ${fmtTime(st)}–${fmtTime(en)}`, m.recurring ? h('span', { class: 'badge muted' }, '↻ повторювана') : null,
      m.eventMissing ? h('span', { class: 'badge warn' }, 'подію видалено з календаря') : null),
    parts.length
      ? h('div', { class: 'chips' }, parts.map(p => h('span', { class: 'chip' }, h('i', { style: { background: colorFor(p.id) } }, initials(p.name)), p.name, p.role ? h('small', null, p.role) : null)))
      : h('button', { class: 'linkish', onclick: async () => { const r = await ensureRow(m); go('#/edit/' + encodeURIComponent(r.id)); } }, '+ додати учасників'),
    m.notes ? h('p', { class: 'notes' }, m.notes) : null));

  if (m.status !== 'analyzed' && backendAtLeast('1.1.0')) root.appendChild(agendaBox(m));
  const jobBox = m.id ? jobPanel(m.id, (j) => { if (alive(t) && j.meeting) { S.meetings.set(j.meeting.id, j.meeting); clear(view); renderMeeting(j.meeting); } }) : h('div');
  root.appendChild(jobBox);
  root.appendChild(audioPanel(m));
  const tabsHost = h('div', { class: 'tabs-host' });
  root.appendChild(tabsHost);
  view.appendChild(root);
  if (m.id && (m.engines.length || m.status === 'analyzed')) loadTabs(m, tabsHost);
}

async function loadTabs(m, host) {
  const t = seq;
  clear(host).appendChild(spinner('Завантажую текст…'));
  try {
    const [transcripts, analysis] = await Promise.all([api.call('getTranscripts', { meetingId: m.id }), api.call('getAnalysis', { meetingId: m.id })]);
    if (!alive(t)) return;
    MV.transcripts = transcripts || {};
    MV.analysis = analysis;
    MV.engine = (m.primaryEngine && MV.transcripts[m.primaryEngine]) ? m.primaryEngine : Object.keys(MV.transcripts)[0] || null;
    drawTabs(m, host);
  } catch (e) {
    if (!alive(t)) return;
    clear(host).appendChild(h('div', { class: 'error-box' }, e.message, h('button', { class: 'btn small', onclick: () => loadTabs(m, host) }, 'Ще раз')));
  }
}

function drawTabs(m, host) {
  clear(host);
  const engines = Object.keys(MV.transcripts);
  const tabs = [['summary', 'Підсумок'], ['tasks', 'Задачі' + (MV.analysis && MV.analysis.tasks.length ? ` (${MV.analysis.tasks.length})` : '')], ['text', 'Текст']];
  if (engines.length > 1) tabs.push(['compare', 'Порівняння']);
  if (!MV.tab || !tabs.some(x => x[0] === MV.tab)) MV.tab = MV.analysis ? 'summary' : 'text';
  const bar = h('div', { class: 'tabs', role: 'tablist' }, tabs.map(([id, label]) =>
    h('button', { role: 'tab', class: 'tab' + (MV.tab === id ? ' active' : ''), 'aria-selected': String(MV.tab === id), onclick: () => { MV.tab = id; drawTabs(m, host); } }, label)));
  const body = h('div', { class: 'tabbody' });
  put(host, bar, body);
  if (MV.tab === 'summary') tabSummary(m, body);
  else if (MV.tab === 'tasks') tabTasks(m, body);
  else if (MV.tab === 'text') tabText(m, body);
  else tabCompare(m, body);
}

function tabSummary(m, el) {
  const a = MV.analysis;
  if (!a) {
    put(el, h('div', { class: 'empty' }, h('p', null, 'Підсумку ще немає.'),
      Object.keys(MV.transcripts).length ? h('button', { class: 'btn primary', onclick: () => analyzeFlow(m) }, 'Зробити підсумок') : null));
    return;
  }
  const list = (title, items) => items && items.length ? h('section', { class: 'sum-sec' }, h('h3', null, title), h('ul', null, items.map(x => h('li', null, x)))) : null;
  put(el, 
    h('section', { class: 'sum-sec' }, String(a.summary || '').split(/\n+/).filter(Boolean).map(p => h('p', null, p))),
    list('Рішення', a.decisions),
    list('Відкриті питання', a.openQuestions),
    list('Цифри й факти', a.facts),
    h('p', { class: 'hint' }, `Підсумок: ${a.engineName} (${a.model}) за текстом ${ENGINES[a.basedOn] ? ENGINES[a.basedOn].name : a.basedOn} · ${fmtDateShort(a.createdAt)}, ${fmtTime(a.createdAt)}`),
    h('div', { class: 'row wrap' },
      m.docId ? h('a', { class: 'btn primary', href: `https://docs.google.com/document/d/${m.docId}/edit`, target: '_blank', rel: 'noopener' }, icon('doc'), 'Google Doc') : null,
      h('button', { class: 'btn', onclick: () => buildDocFlow(m) }, icon('doc'), m.docId ? 'Оновити документ' : 'Створити Google Doc'),
      h('button', { class: 'btn ghost', onclick: () => analyzeFlow(m) }, 'Перезробити підсумок')));
}

async function tabTasks(m, el) {
  const a = MV.analysis;
  if (!backendAtLeast('1.1.0')) {
    // Старий скрипт: показуємо задачі з підсумку без статусів.
    if (!a) { put(el, h('div', { class: 'empty' }, 'Задачі зʼявляться після підсумку.')); return; }
    if (!a.tasks.length) { put(el, h('div', { class: 'empty' }, 'На нараді задач не зафіксовано.')); return; }
    put(el, h('div', { class: 'stack' }, a.tasks.map(x => h('div', { class: 'card task' }, h('div', { class: 'grow' }, h('div', { class: 'task-text' }, x.task),
      h('div', { class: 'task-meta' }, h('span', { class: 'chip small' }, x.owner || '—'), h('span', { class: 'due muted' }, x.due || 'без терміну')))))),
      h('div', { class: 'row' }, h('button', { class: 'btn', onclick: () => copyText(tasksListText(a.tasks)) }, icon('copy'), 'Скопіювати задачі')));
    return;
  }
  const t = seq;
  put(el, spinner());
  let mine, carried;
  try { [mine, carried] = await Promise.all([api.call('listTasks', { meetingId: m.id }), api.call('carriedTasks', { meetingId: m.id })]); }
  catch (e) { if (alive(t)) put(clear(el), h('div', { class: 'error-box' }, e.message)); return; }
  if (!alive(t) || !el.isConnected) return;
  clear(el);
  const sug = new Set((a && a.reportedDone) || []);
  const redraw = () => { clear(el); tabTasks(m, el); };
  const onChange = (x, kind) => { if (kind !== 'status') redraw(); };
  if (!mine.length && !carried.length) {
    put(el, h('div', { class: 'empty' }, a ? 'На нараді задач не зафіксовано.' : 'Задачі зʼявляться після підсумку.'));
  }
  if (mine.length) {
    put(el, h('h3', { class: 'sub-h' }, 'З цієї наради'), h('div', { class: 'stack' }, sortTasks(mine).map(x => taskCard(x, { onChange, showMeeting: false }))));
  }
  if (carried.length) {
    put(el, h('h3', { class: 'sub-h' }, `З минулих нарад — ще відкриті (${carried.length})`),
      h('p', { class: 'hint' }, sug.size ? 'Позначені «схоже, виконано» — за розмовою на цій нараді. Перевір і постав галочку.' : 'Постав галочку на тих, що вже виконано.'),
      h('div', { class: 'stack' }, sortTasks(carried).map(x => taskCard(x, { onChange, suggestedDone: sug.has(x.id) }))));
  }
  const all = [...mine, ...carried.filter(x => x.status !== 'done')];
  if (all.length) put(el, h('div', { class: 'row wrap' },
    h('button', { class: 'btn', onclick: () => copyText(tasksListText(sortTasks(all))) }, icon('copy'), 'Скопіювати задачі'),
    h('a', { class: 'btn ghost', href: '#/tasks' }, 'Усі задачі')));
}

// Порядок денний: відкриті задачі з минулих нарад серії (або з тією ж назвою).
function agendaBox(m) {
  const box = h('div', { class: 'card agenda', hidden: true });
  const args = m.id ? { meetingId: m.id } : { eventId: m.eventId, recurring: !!m.recurring, title: m.title, start: m.start };
  api.call('carriedTasks', args).then((list) => {
    if (!list.length || !box.isConnected) return;
    const draw = () => {
      const open = list.filter(x => x.status !== 'done');
      clear(box);
      box.hidden = !list.length;
      put(box, h('div', { class: 'row between' }, h('b', null, `До порядку денного: відкриті задачі (${open.length})`),
        h('button', { class: 'btn small ghost', onclick: () => copyText('Відкриті задачі з минулих нарад:\n' + tasksListText(sortTasks(open))) }, icon('copy'), 'Скопіювати')),
        h('p', { class: 'hint' }, 'З минулих нарад цієї серії. Відміть виконані — решта перейде на наступну нараду.'),
        h('div', { class: 'stack' }, sortTasks(list).map(x => taskCard(x, { onChange: (y, kind) => { if (kind === 'deleted') list.splice(list.indexOf(x), 1); draw(); } }))));
    };
    draw();
  }).catch(() => { });
  return box;
}

/* =================== Задачі (усі) =================== */

async function viewTasks() {
  const t = seq;
  put(view, topbar('Задачі', { actions: [h('button', { class: 'btn primary small', onclick: () => editTask(null, () => render()) }, icon('plus'), 'Задача')] }));
  if (!backendAtLeast('1.1.0')) {
    await api.loadConfig().catch(() => null);
    if (!alive(t)) return;
    if (!backendAtLeast('1.1.0')) { put(view, h('div', { class: 'empty' }, 'Щоб бачити всі задачі з нарад, онови скрипт Google до версії 1.1.0.')); return; }
  }
  const f = S.taskFilter || (S.taskFilter = { status: 'open', owner: '', q: '' });
  const chips = h('div', { class: 'chips' });
  const owner = h('select', { class: 'input' });
  const q = h('input', { type: 'search', class: 'input', placeholder: 'Пошук у задачах…', value: f.q });
  const list = h('div', { class: 'stack' }, spinner());
  let all = [];
  const reportBtn = h('button', { class: 'btn small', onclick: () => copyText(weeklyReport(all)) }, icon('copy'), 'Звіт за тиждень');
  put(view, chips, h('div', { class: 'grid2' }, owner, q), h('div', { class: 'row wrap' }, reportBtn,
    h('span', { class: 'hint' }, 'виконане за тиждень, прострочене й у роботі — текстом для звіту')), list);

  try { [all] = await Promise.all([api.call('listTasks', {}), api.getPeople().catch(() => [])]); }
  catch (e) { if (alive(t)) put(clear(list), h('div', { class: 'error-box' }, e.message)); return; }
  if (!alive(t)) return;

  const draw = () => {
    const open = all.filter(x => x.status !== 'done'), late = all.filter(isOverdue), done = all.filter(x => x.status === 'done');
    clear(chips);
    [['open', `Відкриті (${open.length})`], ['late', `Прострочені (${late.length})`], ['done', `Виконані (${done.length})`], ['all', 'Усі']].forEach(([id, label]) =>
      put(chips, h('button', { class: 'chip toggle' + (f.status === id ? ' on' : '') + (id === 'late' && late.length ? ' warn' : ''), onclick: () => { f.status = id; draw(); } }, label)));
    const owners = [...new Set(all.map(x => (x.owner || '—').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'uk'));
    clear(owner);
    put(owner, h('option', { value: '' }, 'Усі відповідальні'), owners.map(o => h('option', { value: o }, o)));
    owner.value = owners.includes(f.owner) ? f.owner : '';
    const qq = f.q.trim().toLowerCase();
    let shown = f.status === 'open' ? open : f.status === 'late' ? late : f.status === 'done' ? done : all;
    if (owner.value) shown = shown.filter(x => (x.owner || '—').trim() === owner.value);
    if (qq) shown = shown.filter(x => (x.task + ' ' + x.owner + ' ' + (x.meetingTitle || '')).toLowerCase().includes(qq));
    clear(list);
    if (!shown.length) { put(list, h('div', { class: 'empty' }, all.length ? 'Нічого не знайдено' : 'Задач ще немає. Вони зʼявляються після підсумку наради або додаються кнопкою «Задача».')); return; }
    sortTasks(shown).forEach(x => put(list, taskCard(x, { onChange: (y, kind) => {
      if (kind === 'deleted') all = all.filter(z => z.id !== x.id);
      else if (kind === 'created') all.push(y);
      draw();
    } })));
  };
  owner.addEventListener('change', () => { f.owner = owner.value; draw(); });
  q.addEventListener('input', () => { f.q = q.value; draw(); });
  draw();
}

function engineChips(selected, onPick) {
  return h('div', { class: 'chips' }, Object.keys(MV.transcripts).map(id =>
    h('button', { class: 'chip toggle' + (selected === id ? ' on' : ''), onclick: () => onPick(id) },
      ENGINES[id] ? ENGINES[id].name : id, MV.m.primaryEngine === id ? h('small', null, 'основний') : null)));
}

function tabText(m, el) {
  const people = api.peopleSync();
  const id = MV.engine;
  const tr = MV.transcripts[id];
  if (!tr) { put(el, h('div', { class: 'empty' }, 'Тексту ще немає.')); return; }
  const nameOf = speakerNamer(m, id, people);
  put(el, engineChips(id, (x) => { MV.engine = x; clear(el); tabText(m, el); }));
  put(el, h('p', { class: 'hint' }, `${tr.engineName} · ${tr.model} · ${tr.segments.length} реплік · оброблено за ${tr.elapsedSec} с`));
  (tr.warnings || []).forEach(w => put(el, h('div', { class: 'banner warn small' }, w)));

  const actions = h('div', { class: 'row wrap' },
    m.primaryEngine !== id ? h('button', { class: 'btn small', onclick: async () => {
      const r = await api.call('saveMeetingState', { meetingId: m.id, primaryEngine: id });
      MV.m = r; S.meetings.set(r.id, r); m.primaryEngine = id; toast('Тепер це основний текст. Онови підсумок, щоб він спирався на нього.', 'ok'); clear(el); tabText(m, el);
    } }, 'Зробити основним') : null,
    h('button', { class: 'btn small', onclick: () => copyText(transcriptAsText(tr, nameOf)) }, icon('copy'), 'Скопіювати'),
    h('button', { class: 'btn small ghost', onclick: async () => {
      if (!await confirmDialog(`Видалити текст від ${tr.engineName}?`, 'Видалити', true)) return;
      const r = await api.call('saveMeetingState', { meetingId: m.id, removeEngine: id });
      S.meetings.set(r.id, r); clear(view); renderMeeting(r);
    } }, 'Видалити варіант'));
  put(el, actions);

  if (tr.speakers && tr.speakers.length) put(el, speakerPanel(m, id, tr, () => { clear(el); tabText(m, el); }));

  const q = h('input', { type: 'search', class: 'input', placeholder: 'Пошук у тексті…' });
  const list = h('div', { class: 'segs' });
  const focus = MV.focus;
  const words = focus && focus.q ? queryWords(focus.q) : [];
  const draw = () => {
    clear(list);
    const s = q.value.trim().toLowerCase();
    const segs = tr.segments.filter(x => !s || x.text.toLowerCase().includes(s) || nameOf(x.speaker).toLowerCase().includes(s));
    if (!segs.length) { put(list, h('div', { class: 'empty' }, 'Нічого не знайдено')); return; }
    segs.forEach(x => put(list, h('div', { class: 'seg', dataset: { t: Math.round(x.start) } },
      h('button', { class: 'ts', onclick: () => seek(x.start) }, fmtDuration(x.start)),
      h('div', null,
        x.speaker ? h('b', { class: 'spk', style: { color: colorFor(nameOf(x.speaker)) } }, nameOf(x.speaker)) : null,
        h('span', null, words.length ? highlight(x.text, words) : x.text)))));
  };
  q.addEventListener('input', draw);
  put(el, q, list);
  draw();
  if (focus) {
    MV.focus = null;
    const nodes = $$('.seg', list);
    let best = null;
    for (const n of nodes) { if (Number(n.dataset.t) <= focus.t + 1) best = n; else break; }
    best = best || nodes[0];
    if (best) { best.classList.add('flash'); setTimeout(() => best.scrollIntoView({ block: 'center', behavior: 'smooth' }), 80); }
  }
}

function speakerPanel(m, engine, tr, redraw) {
  const people = api.peopleSync();
  const map = { ...((m.speakerMap || {})[engine] || {}) };
  const partIds = m.participants || [];
  const ordered = [...people.filter(p => partIds.includes(p.id)), ...people.filter(p => !partIds.includes(p.id))];
  const rows = tr.speakers.map(label => {
    const segs = tr.segments.filter(s => s.speaker === label);
    const sample = (segs.find(s => s.text.length > 40) || segs[0] || { text: '' }).text.slice(0, 110);
    const sel = h('select', { class: 'input', onchange: (e) => { if (e.target.value) map[label] = e.target.value; else delete map[label]; } },
      h('option', { value: '' }, '— як є —'),
      ordered.map(p => h('option', { value: p.id }, p.name + (partIds.includes(p.id) ? '' : ' (не в учасниках)'))));
    sel.value = map[label] || '';
    return h('div', { class: 'spkrow' },
      h('div', null, h('b', null, label), h('small', null, ` · ${segs.length} реплік`), h('div', { class: 'hint' }, `«${sample}…»`)),
      sel);
  });
  return h('details', { class: 'card spkpanel' },
    h('summary', null, 'Хто є хто — підписати спікерів'),
    people.length ? null : h('p', { class: 'hint' }, 'Спершу додай людей у розділі «Люди».'),
    rows,
    h('div', { class: 'row' }, h('button', { class: 'btn small primary', onclick: async () => {
      const all = { ...(m.speakerMap || {}), [engine]: map };
      const r = await api.call('saveMeetingState', { meetingId: m.id, speakerMap: all });
      m.speakerMap = r.speakerMap; S.meetings.set(r.id, r);
      toast('Імена збережено. Щоб вони потрапили в підсумок — перезроби його.', 'ok', 5000);
      redraw();
    } }, 'Зберегти імена')));
}

function tabCompare(m, el) {
  const ids = Object.keys(MV.transcripts);
  MV.cmp = (MV.cmp || ids.slice(0, 3)).filter(x => ids.includes(x));
  if (MV.cmp.length < 2) MV.cmp = ids.slice(0, 2);
  const people = api.peopleSync();
  put(el, h('p', { class: 'hint' }, 'Обери 2–3 варіанти. Текст вирівняно по хвилинах запису.'));
  put(el, h('div', { class: 'chips' }, ids.map(id => h('button', {
    class: 'chip toggle' + (MV.cmp.includes(id) ? ' on' : ''),
    onclick: () => {
      if (MV.cmp.includes(id)) { if (MV.cmp.length > 2) MV.cmp = MV.cmp.filter(x => x !== id); }
      else if (MV.cmp.length < 3) MV.cmp = [...MV.cmp, id];
      else MV.cmp = [...MV.cmp.slice(1), id];
      clear(el); tabCompare(m, el);
    }
  }, ENGINES[id] ? ENGINES[id].name : id))));

  const cols = MV.cmp.map(id => MV.transcripts[id]);
  put(el, h('div', { class: 'cmp-stats', style: { gridTemplateColumns: `repeat(${cols.length}, minmax(0,1fr))` } }, cols.map(tr => {
    const words = tr.text.split(/\s+/).filter(Boolean).length;
    const unk = (tr.text.match(/нерозбірливо/gi) || []).length;
    return h('div', { class: 'card' }, h('b', null, tr.engineName),
      h('div', { class: 'hint' }, `${words} слів · ${tr.speakers.length || '—'} спікерів · ${tr.elapsedSec} с`),
      unk ? h('div', { class: 'hint' }, `нерозбірливо: ${unk}`) : null);
  })));

  const dur = Math.max(...cols.map(tr => tr.durationSec || 0), 1);
  const WIN = 60;
  const table = h('div', { class: 'cmp' });
  for (let w = 0; w < dur; w += WIN) {
    const cells = cols.map(tr => {
      const nameOf = speakerNamer(m, tr.engine, people);
      const segs = tr.segments.filter(s => s.start >= w && s.start < w + WIN);
      if (!segs.length) return h('div', { class: 'cmp-cell empty-cell' }, h('div', { class: 'cmp-eng' }, tr.engineName), h('p', { class: 'hint' }, '—'));
      return h('div', { class: 'cmp-cell' }, h('div', { class: 'cmp-eng' }, tr.engineName),
        segs.map(s => h('p', null, s.speaker ? h('b', null, nameOf(s.speaker) + ': ') : null, s.text)));
    });
    if (cells.every(c => c.classList.contains('empty-cell'))) continue;
    put(table, h('div', { class: 'cmp-row' },
      h('button', { class: 'ts', onclick: () => seek(w) }, fmtDuration(w)),
      h('div', { class: 'cmp-cells', style: { gridTemplateColumns: `repeat(${cols.length}, minmax(220px,1fr))` } }, cells)));
  }
  put(el, table);
}

/* ---------- аудіо ---------- */

function audioPanel(m) {
  const box = h('div', { class: 'card audio' });
  if (!m.audioFileId) {
    const fileIn = h('input', { type: 'file', hidden: true, accept: 'audio/*,video/*,.m4a,.mp3,.wav,.ogg,.webm,.aac,.amr,.3gp', onchange: (e) => onFile(e.target.files[0]) });
    const onFile = async (file) => { if (!file) return; const r = await ensureRow(m); startProcessing(r, file, { fileName: file.name }); };
    box.classList.add('drop');
    put(box, 
      h('p', { class: 'drop-title' }, 'Запису ще немає'),
      h('div', { class: 'row wrap center' },
        h('button', { class: 'btn rec', onclick: async () => { const r = await ensureRow(m); go('#/rec/' + encodeURIComponent(r.id)); } }, icon('mic'), 'Записати'),
        h('button', { class: 'btn', onclick: () => fileIn.click() }, icon('upload'), 'Завантажити файл')),
      h('p', { class: 'hint center' }, 'Або перетягни сюди файл з диктофона (m4a, mp3, wav…)'),
      fileIn);
    box.addEventListener('dragover', (e) => { e.preventDefault(); box.classList.add('over'); });
    box.addEventListener('dragleave', () => box.classList.remove('over'));
    box.addEventListener('drop', (e) => { e.preventDefault(); box.classList.remove('over'); onFile(e.dataTransfer.files[0]); });
    return box;
  }
  const playerHost = h('div', { class: 'player-host' });
  const fileIn = h('input', { type: 'file', hidden: true, accept: 'audio/*,video/*,.m4a,.mp3,.wav,.ogg,.webm,.aac,.amr,.3gp', onchange: async (e) => {
    const file = e.target.files[0]; if (!file) return;
    if (!await confirmDialog('Замінити аудіо цієї наради? Старий файл піде в кошик Google Диска, а наявні тексти лишаться.', 'Замінити')) return;
    startProcessing(m, file, { fileName: file.name });
  } });
  put(box, 
    h('div', { class: 'row wrap' },
      h('button', { class: 'btn', onclick: () => loadPlayer(m, playerHost) }, icon('play'), 'Слухати'),
      h('div', { class: 'grow' }, h('b', null, fmtDuration(m.durationSec)), h('div', { class: 'hint ellipsis' }, m.audioName || 'аудіо')),
      h('button', { class: 'btn primary', onclick: () => chooseEnginesFlow(m) }, 'Розпізнати…'),
      h('button', { class: 'btn ghost small', onclick: () => fileIn.click() }, 'Замінити')),
    playerHost, fileIn);
  return box;
}

async function loadPlayer(m, host) {
  if (S.player && host.contains(S.player)) { S.player.play(); return; }
  put(clear(host), spinner('Завантажую аудіо…'));
  try {
    let blob = null;
    const p = await getPrepared(m.id);
    if (p && p.fileId === m.audioFileId) blob = p.full;
    if (!blob) blob = await downloadFile(m.audioFileId);
    const a = h('audio', { controls: true, preload: 'auto', src: URL.createObjectURL(blob) });
    put(clear(host), a);
    S.player = a;
    a.play().catch(() => { });
  } catch (e) { put(clear(host), h('div', { class: 'error-box' }, e.message)); }
}

function seek(sec) {
  if (!S.player) {
    const host = view.querySelector('.player-host');
    if (host && MV && MV.m.audioFileId) loadPlayer(MV.m, host).then(() => { if (S.player) { S.player.currentTime = sec; } });
    else toast('Аудіо немає');
    return;
  }
  S.player.currentTime = sec;
  S.player.play().catch(() => { });
  S.player.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

/* ---------- обробка ---------- */

function jobPanel(meetingId, onFinished) {
  const box = h('div', { class: 'job card' });
  let wasRunning = !!(jobFor(meetingId) && jobFor(meetingId).running);
  const icons = { ok: '✓', err: '!', run: '', wait: '·' };
  const draw = () => {
    const j = jobFor(meetingId);
    clear(box);
    box.hidden = !j;
    if (!j) return;
    put(box, 
      h('div', { class: 'job-head' }, j.running ? h('span', { class: 'spin' }) : null,
        h('b', null, j.running ? 'Обробка…' : j.error ? 'Обробку зупинено' : 'Готово'),
        !j.running ? h('button', { class: 'btn small ghost', onclick: () => clearJob(meetingId) }, 'Сховати') : null),
      h('ul', { class: 'steps' }, j.steps.map(s => h('li', { class: 'step ' + s.status },
        h('span', { class: 'st-ic' }, icons[s.status]),
        h('div', { class: 'grow' }, h('div', null, s.label), s.detail ? h('div', { class: 'hint' }, s.detail) : null,
          s.progress != null ? h('div', { class: 'bar' }, h('i', { style: { width: Math.round(s.progress * 100) + '%' } })) : null)))),
      j.running ? h('p', { class: 'hint' }, 'Не закривай додаток до завершення — екран лишатиметься увімкненим.') : null);
  };
  const off = onJobs((id) => {
    if (id !== meetingId) return;
    if (!box.isConnected) { off(); return; }
    draw();
    const j = jobFor(meetingId);
    if (wasRunning && j && !j.running && onFinished) onFinished(j);
    wasRunning = !!(j && j.running);
  });
  draw();
  return box;
}

async function startProcessing(m, blob, { fileName, onUploaded } = {}) {
  const choice = await chooseEngines(m, { title: 'Що робити з записом?', okLabel: 'Почати', allowNone: true });
  if (!choice) return;
  if (location.hash !== '#/m/' + encodeURIComponent(m.id)) go('#/m/' + encodeURIComponent(m.id));
  else { clear(view); renderMeeting(m); }
  processAudio(m, blob, { fileName, engines: choice.engines, analyzeAfter: choice.analyzeAfter, onUploaded })
    .then((r) => { S.meetings.set(r.id, r); toast('Нараду оброблено', 'ok'); })
    .catch((e) => toast(e.message, 'err', 7000));
}

async function chooseEnginesFlow(m) {
  const choice = await chooseEngines(m, { title: 'Розпізнати запис', okLabel: 'Розпізнати' });
  if (!choice || !choice.engines.length) return;
  runTranscription(m, choice.engines, { analyzeAfter: choice.analyzeAfter })
    .then((r) => { S.meetings.set(r.id, r); toast('Готово', 'ok'); })
    .catch((e) => toast(e.message, 'err', 7000));
  clear(view); renderMeeting(m);
}

async function chooseEngines(m, { title, okLabel, allowNone = false }) {
  let c;
  try { c = await api.loadConfig(); } catch (e) { toast(e.message, 'err'); return null; }
  const done = new Set(m.engines || []);
  const defaults = new Set(c.settings.defaultEngines || []);
  const boxes = {};
  const list = h('div', { class: 'stack' }, Object.values(ENGINES).map(e => {
    const hasKey = !!c.keys[e.keyName];
    const cb = h('input', { type: 'checkbox', disabled: !hasKey, checked: hasKey && defaults.has(e.id) && !done.has(e.id) });
    boxes[e.id] = cb;
    return h('label', { class: 'card choice' + (hasKey ? '' : ' disabled') }, cb,
      h('div', null,
        h('b', null, e.name), done.has(e.id) ? h('span', { class: 'badge ok' }, 'вже є') : null,
        h('div', { class: 'tags' },
          h('span', { class: 'badge ' + (e.freeTier ? 'ok' : 'muted') }, e.freeTier ? 'є безкоштовно' : 'платно'),
          h('span', { class: 'badge muted' }, e.speakers ? 'спікери ✓' : 'без спікерів')),
        h('div', { class: 'hint' }, hasKey ? e.note : 'Немає ключа API — додай у Налаштуваннях.')));
  }));
  const an = h('input', { type: 'checkbox', checked: c.settings.autoAnalyze });
  const anName = (ANALYZERS[c.settings.analysisEngine] || ANALYZERS.gemini).name;
  return new Promise((resolve) => {
    let resolved = false;
    modal({
      title,
      body: h('div', null,
        h('p', { class: 'hint' }, 'Можна обрати кілька — тексти стануть поруч для порівняння.'),
        list,
        h('label', { class: 'check' }, an, ` Після розпізнавання зробити підсумок і задачі (${anName})`),
        allowNone ? h('p', { class: 'hint' }, 'Якщо нічого не обрати — запис просто збережеться на Диск.') : null),
      actions: [
        { label: 'Скасувати', onClick: () => { resolved = true; resolve(null); } },
        {
          label: okLabel, kind: 'primary', onClick: () => {
            const engines = Object.keys(boxes).filter(id => boxes[id].checked);
            if (!engines.length && !allowNone) { toast('Обери хоча б один двигун'); return false; }
            resolved = true; resolve({ engines, analyzeAfter: an.checked });
          }
        }
      ],
      onClose: () => { if (!resolved) resolve(null); }
    });
  });
}

async function analyzeFlow(m) {
  const c = await api.loadConfig();
  const sel = h('select', { class: 'input' }, Object.values(ANALYZERS).map(a => h('option', { value: a.id, disabled: !c.keys[a.keyName] }, a.name + (c.keys[a.keyName] ? '' : ' — немає ключа'))));
  sel.value = c.settings.analysisEngine;
  const primary = m.primaryEngine && ENGINES[m.primaryEngine] ? ENGINES[m.primaryEngine].name : '';
  modal({
    title: 'Підсумок наради',
    body: h('div', null, h('label', { class: 'field' }, 'Хто робить підсумок', sel),
      primary ? h('p', { class: 'hint' }, `За основу береться текст від ${primary}. Змінити основний можна на вкладці «Текст».`) : null),
    actions: [
      { label: 'Скасувати' },
      {
        label: 'Зробити', kind: 'primary', onClick: () => {
          runAnalysis(m, { engine: sel.value }).then((r) => { S.meetings.set(r.id, r); toast('Підсумок готовий', 'ok'); }).catch((e) => toast(e.message, 'err', 7000));
          clear(view); renderMeeting(m);
        }
      }
    ]
  });
}

async function buildDocFlow(m) {
  toast('Створюю документ…', 'info', 2500);
  try {
    const r = await api.call('buildDoc', { meetingId: m.id });
    m.docId = r.docId;
    const fresh = await api.call('getMeeting', { id: m.id });
    S.meetings.set(fresh.id, fresh);
    toast('Документ готовий', 'ok');
    window.open(r.url, '_blank', 'noopener');
    clear(view); renderMeeting(fresh);
  } catch (e) { toast(e.message, 'err', 7000); }
}

async function deleteMeetingFlow(m) {
  const withEvent = h('input', { type: 'checkbox', checked: !m.recurring, disabled: m.recurring });
  modal({
    title: 'Видалити нараду?',
    body: h('div', null,
      h('p', null, 'Аудіо, тексти й документ підуть у кошик Google Диска (звідти їх можна відновити 30 днів).'),
      h('label', { class: 'check' }, withEvent, ' Також видалити подію з Google Календаря'),
      m.recurring ? h('p', { class: 'hint' }, 'Повторювану подію видаляй у самому Календарі.') : null),
    actions: [
      { label: 'Скасувати' },
      {
        label: 'Видалити', kind: 'danger', onClick: async () => {
          try { await api.call('deleteMeeting', { id: m.id, withEvent: withEvent.checked }); S.meetings.delete(m.id); toast('Видалено', 'ok'); go('#/calendar'); }
          catch (e) { toast(e.message, 'err'); return false; }
        }
      }
    ]
  });
}

/* =================== Створення й редагування =================== */

function participantPicker(selected, people) {
  let sel = [...selected];
  const chips = h('div', { class: 'chips' });
  const input = h('input', { class: 'input', placeholder: "Почни вводити ім'я…", autocomplete: 'off' });
  const drop = h('div', { class: 'dropdown', hidden: true });
  const draw = () => {
    clear(chips);
    sel.forEach(id => {
      const p = people.find(x => x.id === id);
      if (p) put(chips, h('span', { class: 'chip' }, h('i', { style: { background: colorFor(p.id) } }, initials(p.name)), p.name,
        h('button', { class: 'chip-x', type: 'button', 'aria-label': 'Прибрати', onclick: () => { sel = sel.filter(x => x !== id); draw(); } }, '×')));
    });
  };
  const suggest = () => {
    const q = input.value.trim().toLowerCase();
    clear(drop);
    const matches = people.filter(p => !sel.includes(p.id) && (!q || p.name.toLowerCase().includes(q) || (p.aliases || '').toLowerCase().includes(q) || (p.role || '').toLowerCase().includes(q))).slice(0, 8);
    matches.forEach(p => put(drop, h('button', { type: 'button', class: 'dd-item', onmousedown: (e) => { e.preventDefault(); sel.push(p.id); input.value = ''; draw(); suggest(); } },
      p.name, p.role ? h('small', null, ' · ' + p.role) : null)));
    if (q && !people.some(p => p.name.toLowerCase() === q)) {
      put(drop, h('button', { type: 'button', class: 'dd-item add', onmousedown: async (e) => {
        e.preventDefault();
        try {
          const p = await api.call('savePerson', { person: { name: input.value.trim() } });
          people.push(p); api.setPeopleCache(people); sel.push(p.id); input.value = ''; draw(); suggest();
        } catch (err) { toast(err.message, 'err'); }
      } }, `+ Додати «${input.value.trim()}»`));
    }
    drop.hidden = !drop.childNodes.length;
  };
  input.addEventListener('focus', suggest);
  input.addEventListener('input', suggest);
  input.addEventListener('blur', () => setTimeout(() => { drop.hidden = true; }, 150));
  draw();
  return { el: h('div', { class: 'picker' }, chips, h('div', { class: 'dd-wrap' }, input, drop)), get: () => sel };
}

async function viewEdit(id, params) {
  const t = seq;
  view.appendChild(spinner());
  const people = await api.getPeople();
  let m = null;
  if (id) m = S.meetings.get(id) || await api.call('getMeeting', { id });
  if (!alive(t)) return;
  clear(view);
  const now = new Date();
  let start, end;
  if (m) { start = new Date(m.start); end = new Date(m.end); }
  else {
    const dParam = params && params.get('d');
    const day = dParam ? new Date(dParam + 'T00:00') : new Date(now);
    if (sameDay(day, now)) { start = new Date(now); start.setMinutes(start.getMinutes() < 30 ? 30 : 60, 0, 0); }
    else { start = new Date(day); start.setHours(9, 0, 0, 0); }
    end = new Date(start.getTime() + 3600e3);
  }
  const title = h('input', { class: 'input', value: m ? m.title : '', placeholder: 'Напр.: Планірка, нарада з дільничними', required: true });
  const date = h('input', { class: 'input', type: 'date', value: ymd(start) });
  const t1 = h('input', { class: 'input', type: 'time', value: hm(start) });
  const t2 = h('input', { class: 'input', type: 'time', value: hm(end) });
  const lockTime = m && m.recurring;
  if (lockTime) [date, t1, t2].forEach(x => { x.disabled = true; });
  const notes = h('textarea', { class: 'input', rows: 3, placeholder: 'Порядок денний, нотатки…' }, m ? m.notes : '');
  const picker = participantPicker(m ? m.participants : [], people);
  const addCal = h('input', { type: 'checkbox', checked: true });

  // Повторення: для нової наради або одиночної, яку хочемо зробити регулярною.
  const c = await api.loadConfig().catch(() => null);
  if (!alive(t)) return;
  const seriesOk = !!(c && c.info && verGte(c.info.version, '1.0.2'));
  const canRepeat = !(m && m.recurring);
  const freq = h('select', { class: 'input', disabled: !seriesOk });
  const term = h('select', { class: 'input' }, [[1, '1 місяць'], [2, '2 місяці'], [3, '3 місяці'], [6, 'пів року'], [12, 'рік']].map(([v, l]) => h('option', { value: v }, l)));
  const repInfo = h('p', { class: 'hint' });
  const termBox = h('label', { class: 'field', hidden: true }, 'На який термін створити', term);
  const applySeries = h('input', { type: 'checkbox' });
  const untilDate = () => { const u = new Date(`${date.value}T00:00`); u.setMonth(u.getMonth() + Number(term.value)); u.setHours(23, 59, 59, 0); return u; };
  const drawRepeat = () => {
    const d = new Date(`${date.value}T00:00`);
    const wd = new Intl.DateTimeFormat('uk-UA', { weekday: 'long' }).format(d);
    const cur = freq.value;
    clear(freq);
    put(freq, [['', 'Не повторювати'], ['weekdays', 'Щодня в робочі дні (пн–пт)'], ['daily', 'Щодня'], ['weekly', `Щотижня (${wd})`], ['biweekly', `Раз на 2 тижні (${wd})`], ['monthly', `Щомісяця (${d.getDate()} числа)`]]
      .map(([v, l]) => h('option', { value: v }, l)));
    freq.value = cur;
    const on = !!freq.value;
    termBox.hidden = !on;
    addCal.checked = on || addCal.checked; addCal.disabled = on;
    repInfo.textContent = !seriesOk ? 'Щоб створювати повторювані наради, онови скрипт Google до версії 1.0.2.'
      : on ? `У Google Календарі зʼявиться серія: ${countOccurrences(freq.value, d, untilDate())} нарад до ${fmtDateFull(untilDate())} включно. Час усієї серії потім змінюють у Google Календарі.` : '';
  };
  [date, freq, term].forEach(x => x.addEventListener('change', drawRepeat));
  drawRepeat();

  t1.addEventListener('change', () => {
    const [a, b] = t1.value.split(':').map(Number);
    const [c, d] = t2.value.split(':').map(Number);
    if (c * 60 + d <= a * 60 + b) { const e = a * 60 + b + 60; t2.value = `${String(Math.floor(e / 60) % 24).padStart(2, '0')}:${String(e % 60).padStart(2, '0')}`; }
  });

  const save = async (e) => {
    e.preventDefault();
    if (!title.value.trim()) { title.focus(); return; }
    const s = new Date(`${date.value}T${t1.value}`);
    let en = new Date(`${date.value}T${t2.value}`);
    // Нарада через північ (напр. 23:00–00:30): кінець — наступного дня.
    if (en <= s) { const next = addDays(en, 1); if (next - s <= 12 * 3600e3) en = next; }
    if (en <= s) { toast('Кінець має бути пізніше за початок', 'err'); return; }
    btn.disabled = true;
    try {
      const payload = { ...(m || {}), title: title.value.trim(), start: s.toISOString(), end: en.toISOString(), notes: notes.value, participants: picker.get() };
      if (lockTime) { payload.start = m.start; payload.end = m.end; }
      if (!m) payload.addToCalendar = addCal.checked;
      if (canRepeat && freq.value) payload.repeat = { freq: freq.value, until: untilDate().toISOString() };
      if (m && m.recurring && applySeries.checked) payload.applyToSeries = true;
      const r = await api.call('saveMeeting', { meeting: payload });
      S.meetings.set(r.id, r);
      toast('Збережено', 'ok');
      go('#/m/' + encodeURIComponent(r.id), true);
    } catch (err) { toast(err.message, 'err'); btn.disabled = false; }
  };
  const btn = h('button', { class: 'btn primary', type: 'submit' }, 'Зберегти');
  put(view, 
    topbar(m ? 'Редагування' : 'Нова нарада', { back: '#/calendar' }),
    h('form', { class: 'form', onsubmit: save },
      h('label', { class: 'field' }, 'Назва', title),
      h('div', { class: 'grid3' }, h('label', { class: 'field' }, 'Дата', date), h('label', { class: 'field' }, 'Початок', t1), h('label', { class: 'field' }, 'Кінець', t2)),
      lockTime ? h('p', { class: 'hint' }, 'Це повторювана подія — час змінюй у Google Календарі.') : null,
      canRepeat ? h('label', { class: 'field' }, 'Повторювати', freq) : null,
      canRepeat ? termBox : null,
      canRepeat ? repInfo : null,
      h('div', { class: 'field' }, 'Учасники', picker.el),
      h('label', { class: 'field' }, 'Нотатки', notes),
      m && m.recurring ? h('label', { class: 'check' }, applySeries, ' Застосувати учасників і нотатки до всієї серії') : null,
      m && m.recurring ? h('p', { class: 'hint' }, 'Наради серії, які ти вже записував чи змінював окремо, лишаться як є.') : null,
      !m ? h('label', { class: 'check' }, addCal, ' Додати в Google Календар') : null,
      h('div', { class: 'row' }, btn, h('button', { type: 'button', class: 'btn ghost', onclick: () => history.back() }, 'Скасувати'))));
}

function verGte(v, min) {
  const a = String(v || '0').split('.').map(Number), b = String(min).split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((a[i] || 0) !== (b[i] || 0)) return (a[i] || 0) > (b[i] || 0); }
  return true;
}

// Скільки нарад вийде в серії — для підказки у формі (так само рахує Google Календар).
function countOccurrences(freq, start, until) {
  const s0 = new Date(start);
  let n = 0;
  if (freq === 'monthly') {
    // Як у Google Календарі: місяці, де немає такого числа (напр. 31-го), пропускаються.
    for (let k = 0; k < 400; k++) {
      const d = new Date(s0.getFullYear(), s0.getMonth() + k, s0.getDate());
      if (new Date(s0.getFullYear(), s0.getMonth() + k, 1) > until) break;
      if (d.getDate() === s0.getDate() && d <= until) n++;
    }
    return n;
  }
  const step = freq === 'weekly' ? 7 : freq === 'biweekly' ? 14 : 1;
  for (let d = s0, g = 0; d <= until && g < 800; d = addDays(d, step), g++) {
    const wd = d.getDay();
    if (freq !== 'weekdays' || (wd >= 1 && wd <= 5)) n++;
  }
  return n;
}

/* =================== Запис =================== */

async function viewRecord(id) {
  const t = seq;
  let m = S.meetings.get(id);
  if (!m) { view.appendChild(spinner()); m = await api.call('getMeeting', { id }); if (!alive(t)) return; clear(view); S.meetings.set(m.id, m); }
  if (S.recorder && S.recMeetingId !== id) {
    put(view, topbar('Запис', { back: '#/calendar' }), h('div', { class: 'empty' }, 'Зараз уже йде запис іншої наради.',
      h('a', { class: 'btn', href: '#/rec/' + encodeURIComponent(S.recMeetingId) }, 'Перейти до нього')));
    return;
  }
  const c = await api.loadConfig().catch(() => null);
  const timer = h('div', { class: 'rec-timer' }, '00:00');
  const meter = h('div', { class: 'meter' }, h('i'));
  const stateTxt = h('div', { class: 'rec-state' }, 'Готовий до запису');
  const mainBtn = h('button', { class: 'rec-btn', 'aria-label': 'Почати запис' }, icon('mic'));
  const pauseBtn = h('button', { class: 'btn', hidden: true }, icon('pause'), 'Пауза');
  const stopBtn = h('button', { class: 'btn primary', hidden: true }, icon('stop'), 'Завершити');
  const cancelBtn = h('button', { class: 'linkish danger', hidden: true }, 'Скасувати запис');
  const dimBtn = h('button', { class: 'btn ghost', hidden: true, onclick: dimScreen }, '🌙 Затемнити екран');

  const sync = () => {
    const r = S.recorder;
    const st = r ? r.state : 'idle';
    mainBtn.hidden = st !== 'idle';
    pauseBtn.hidden = !(st === 'recording' || st === 'paused');
    stopBtn.hidden = pauseBtn.hidden;
    cancelBtn.hidden = pauseBtn.hidden;
    dimBtn.hidden = pauseBtn.hidden;
    put(clear(pauseBtn), icon(st === 'paused' ? 'mic' : 'pause'), st === 'paused' ? 'Продовжити' : 'Пауза');
    stateTxt.textContent = st === 'recording' ? 'Іде запис' : st === 'paused' ? 'Пауза' : 'Готовий до запису';
    view.querySelector('.recview')?.classList.toggle('live', st === 'recording');
  };
  const tick = setInterval(() => {
    if (!view.contains(timer)) { clearInterval(tick); return; }
    if (S.recorder) timer.textContent = fmtDuration(S.recorder.elapsed);
  }, 500);
  const level = (v) => { const i = meter.firstChild; if (i) i.style.width = Math.round(v * 100) + '%'; };
  if (S.recorder) S.recorder.onLevel = level;

  mainBtn.onclick = async () => {
    try {
      const r = new Recorder({ meetingId: m.id, onLevel: level, onState: () => { sync(); updateRecPill(); }, onInterrupted: () => finishRecording('interrupted'), onWarn: (msg, kind) => toast(msg, kind || 'err', 9000) });
      S.recorder = r; S.recMeetingId = m.id;
      await r.start({ noiseSuppression: !!(c && c.settings.noiseSuppression) });
      sync();
    } catch (e) {
      S.recorder = null; S.recMeetingId = null;
      const msg = e.name === 'NotAllowedError' ? 'Немає дозволу на мікрофон. Дозволь його в налаштуваннях браузера для цього сайту.' : e.message;
      toast(msg, 'err', 7000); sync();
    }
  };
  pauseBtn.onclick = () => { const r = S.recorder; if (!r) return; r.state === 'paused' ? r.resume() : r.pause(); sync(); };
  stopBtn.onclick = () => finishRecording('stop');
  cancelBtn.onclick = async () => {
    if (!await confirmDialog('Скасувати запис? Записане буде видалено.', 'Видалити запис', true)) return;
    const r = S.recorder; if (r) await r.discard();
    S.recorder = null; S.recMeetingId = null; updateRecPill(); sync(); timer.textContent = '00:00';
  };

  put(view, 
    topbar('Запис', { back: '#/m/' + encodeURIComponent(m.id) }),
    h('div', { class: 'recview' },
      h('h2', null, m.title),
      h('p', { class: 'hint' }, `${fmtDay(m.start)} · ${fmtTime(m.start)}–${fmtTime(m.end)}`),
      timer, meter, stateTxt, mainBtn,
      h('div', { class: 'row center' }, pauseBtn, stopBtn),
      dimBtn,
      cancelBtn,
      h('p', { class: 'hint center narrow' }, 'Поклади телефон посередині столу. Не вимикай екран кнопкою — натисни «Затемнити екран»: він стане чорним і майже не витрачатиме заряд, а запис ітиме далі. Записане зберігається кожні 5 секунд.')));
  sync();
  if (S.recorder) timer.textContent = fmtDuration(S.recorder.elapsed);
}

// Завершення запису: кнопкою «Завершити» або коли мікрофон несподівано відключився.
async function finishRecording(reason) {
  const r = S.recorder;
  if (!r || r.state === 'stopped') return;
  const m = S.meetings.get(S.recMeetingId);
  const blob = await r.stop();
  S.recorder = null; S.recMeetingId = null; updateRecPill();
  $$('.dim').forEach(x => x.click());
  if (reason === 'interrupted') toast('Мікрофон відключився, тому запис зупинено. Записане збережено.', 'err', 9000);
  if (!blob || !blob.size) { toast('Запис порожній', 'err'); return; }
  if (!m) { toast('Запис збережено на телефоні — обробиш його з календаря', 'info', 7000); return; }
  startProcessing(m, blob, { fileName: 'запис з додатку', onUploaded: () => deleteRecording(r.recId) });
}

// Чорний екран під час запису: на OLED-екрані майже не витрачає заряд, але запис не переривається.
function dimScreen() {
  if (!S.recorder) return;
  const clock = h('div', { class: 'dim-clock' });
  const ov = h('div', { class: 'dim', role: 'button', 'aria-label': 'Повернути екран' }, clock, h('div', { class: 'dim-hint' }, 'Іде запис · торкнись, щоб повернути екран'));
  const draw = () => { if (!S.recorder) { close(); return; } clock.textContent = fmtDuration(S.recorder.elapsed); };
  const t1 = setInterval(draw, 1000);
  const t2 = setInterval(() => { ov.style.paddingTop = Math.round(15 + Math.random() * 55) + 'vh'; }, 30000);
  const close = () => {
    clearInterval(t1); clearInterval(t2); ov.remove();
    try { if (document.fullscreenElement) document.exitFullscreen(); } catch { }
  };
  ov.addEventListener('click', close);
  document.body.appendChild(ov);
  try { if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => { }); } catch { }
  draw();
}

function updateRecPill() {
  let pill = $('#recpill');
  const on = !!S.recorder && !location.hash.startsWith('#/rec/');
  if (!on) { if (pill) pill.hidden = true; return; }
  if (!pill) {
    pill = h('a', { id: 'recpill', class: 'recpill' });
    document.body.appendChild(pill);
    setInterval(() => { if (S.recorder && !pill.hidden) pill.lastChild.textContent = ' ' + fmtDuration(S.recorder.elapsed); }, 1000);
  }
  pill.href = '#/rec/' + encodeURIComponent(S.recMeetingId);
  put(clear(pill), h('i', { class: 'dot' }), document.createTextNode(' ' + fmtDuration(S.recorder.elapsed)));
  pill.hidden = false;
}

/* =================== Імпорт файлу («Поділитися») =================== */

async function viewImport() {
  const t = seq;
  const items = (await idb.all('shared')) || [];
  put(view, topbar('Новий запис', { back: '#/calendar' }));
  if (!items.length) { put(view, h('div', { class: 'empty' }, 'Файлів для імпорту немає.')); return; }
  const it = items[items.length - 1];
  const fileTime = new Date(it.lastModified || Date.now());
  put(view, h('div', { class: 'card' }, h('b', null, it.name || 'аудіо'), h('div', { class: 'hint' }, `${fmtBytes(it.size || 0)} · змінено ${fmtDateShort(fileTime)}, ${fmtTime(fileTime)}`)));
  const list = h('div', { class: 'stack' }, spinner());
  put(view, h('h2', { class: 'sub' }, 'До якої наради додати?'), list);
  const meetings = await api.call('listMeetings', { from: addDays(fileTime, -1).toISOString(), to: addDays(fileTime, 1).toISOString() });
  await api.getPeople().catch(() => []);
  if (!alive(t)) return;
  clear(list);
  let choice = 'new';
  const radios = [];
  const opt = (value, label, sub) => {
    const r = h('input', { type: 'radio', name: 'tgt', value, checked: value === 'new', onchange: () => { choice = value; newBox.hidden = value !== 'new'; } });
    radios.push(r);
    return h('label', { class: 'card choice' }, r, h('div', null, h('b', null, label), sub ? h('div', { class: 'hint' }, sub) : null));
  };
  const free = meetings.filter(m => !m.audioFileId).reverse();
  free.forEach((m, i) => put(list, opt('m' + i, m.title, `${fmtDateShort(m.start)}, ${fmtTime(m.start)}–${fmtTime(m.end)}`)));
  put(list, opt('new', 'Нова нарада'));
  const title = h('input', { class: 'input', placeholder: 'Назва наради', value: 'Нарада ' + hm(fileTime) });
  const date = h('input', { class: 'input', type: 'date', value: ymd(fileTime) });
  const st = new Date(fileTime.getTime() - 3600e3);
  const t1 = h('input', { class: 'input', type: 'time', value: hm(st) });
  const t2 = h('input', { class: 'input', type: 'time', value: hm(fileTime) });
  const newBox = h('div', { class: 'form' }, h('label', { class: 'field' }, 'Назва', title),
    h('div', { class: 'grid3' }, h('label', { class: 'field' }, 'Дата', date), h('label', { class: 'field' }, 'Початок', t1), h('label', { class: 'field' }, 'Кінець', t2)));
  put(list, newBox);
  put(view, h('div', { class: 'row pad' },
    h('button', { class: 'btn primary', onclick: async (e) => {
      e.target.disabled = true;
      try {
        let m;
        if (choice === 'new') {
          const s = new Date(`${date.value}T${t1.value}`); let en = new Date(`${date.value}T${t2.value}`);
          if (en <= s) en = new Date(s.getTime() + 3600e3);
          m = await api.call('saveMeeting', { meeting: { title: title.value.trim() || 'Нарада', start: s.toISOString(), end: en.toISOString(), participants: [] } });
        } else {
          m = await ensureRow(free[Number(choice.slice(1))]);
        }
        S.meetings.set(m.id, m);
        const file = it.file instanceof Blob ? it.file : new Blob([it.file]);
        startProcessing(m, file, { fileName: it.name, onUploaded: async () => { for (const x of items) await idb.del('shared', x.id); } });
      } catch (err) { toast(err.message, 'err'); e.target.disabled = false; }
    } }, 'Далі'),
    h('button', { class: 'btn ghost', onclick: async () => { for (const x of items) await idb.del('shared', x.id); go('#/calendar'); } }, 'Скасувати')));
}

/* =================== Люди і словник =================== */

async function viewPeople() {
  const t = seq;
  put(view, topbar('Люди', { actions: [h('button', { class: 'btn primary small', onclick: () => editPerson(null) }, icon('plus'), 'Додати')] }));
  const list = h('div', { class: 'stack' }, spinner());
  put(view, h('p', { class: 'hint pad' }, 'Довідник учасників нарад. Імена й «як ще називають» підказують розпізнаванню, хто є хто.'), list);
  const glossBox = h('div');
  put(view, h('h2', { class: 'sub' }, 'Словник'), h('p', { class: 'hint pad' }, 'Кожен термін з нового рядка: назви дільниць, обладнання, матеріалів, клієнтів, продуктів. Двигуни отримують їх як підказку.'), glossBox);
  const [people, gloss] = await Promise.all([api.getPeople(true), api.getGlossary(true)]);
  if (!alive(t)) return;
  clear(list);
  if (!people.length) put(list, h('div', { class: 'empty' }, 'Поки нікого. Додай колег, з якими проводиш наради.'));
  people.slice().sort((a, b) => a.name.localeCompare(b.name, 'uk')).forEach(p => put(list, h('button', { class: 'card person', onclick: () => editPerson(p) },
    h('span', { class: 'face big', style: { background: colorFor(p.id) } }, initials(p.name)),
    h('div', { class: 'grow' }, h('b', null, p.name), p.role ? h('div', { class: 'hint' }, p.role) : null, p.aliases ? h('div', { class: 'hint' }, 'також: ' + p.aliases) : null))));
  const ta = h('textarea', { class: 'input', rows: 8, placeholder: 'IdeaCore\nлазерна порізка\nтиснення\n…' }, gloss.join('\n'));
  put(glossBox, h('div', { class: 'form' }, ta, h('div', { class: 'row' }, h('button', { class: 'btn', onclick: async (e) => {
    e.target.disabled = true;
    try { const g = await api.call('saveGlossary', { terms: ta.value.split('\n') }); api.setGlossaryCache(g); toast('Словник збережено', 'ok'); }
    catch (err) { toast(err.message, 'err'); }
    e.target.disabled = false;
  } }, 'Зберегти словник'))));
}

function editPerson(p) {
  const name = h('input', { class: 'input', value: p ? p.name : '', placeholder: "Ім'я та прізвище" });
  const role = h('input', { class: 'input', value: p ? p.role : '', placeholder: 'Посада або дільниця' });
  const aliases = h('input', { class: 'input', value: p ? p.aliases : '', placeholder: 'Напр.: Саша, Олександр Петрович' });
  const actions = [];
  if (p) actions.push({ label: 'Видалити', kind: 'danger ghost', onClick: async () => {
    if (!await confirmDialog(`Видалити ${p.name} з довідника?`, 'Видалити', true)) return false;
    await api.call('deletePerson', { id: p.id }); await api.getPeople(true); render();
  } });
  actions.push({ label: 'Скасувати' });
  actions.push({ label: 'Зберегти', kind: 'primary', onClick: async () => {
    if (!name.value.trim()) { name.focus(); return false; }
    try { await api.call('savePerson', { person: { id: p && p.id, name: name.value, role: role.value, aliases: aliases.value } }); await api.getPeople(true); render(); }
    catch (e) { toast(e.message, 'err'); return false; }
  } });
  modal({
    title: p ? 'Людина' : 'Нова людина',
    body: h('div', { class: 'form' }, h('label', { class: 'field' }, "Ім'я", name), h('label', { class: 'field' }, 'Посада', role), h('label', { class: 'field' }, 'Як ще називають (через кому)', aliases)),
    actions
  });
}

/* =================== Налаштування =================== */

async function viewSettings() {
  const t = seq;
  put(view, topbar('Налаштування'));
  put(view, spinner());
  let c, cals = [];
  try { [c, cals] = await Promise.all([api.loadConfig(true), api.call('listCalendars').catch(() => [])]); }
  catch (e) { if (!alive(t)) return; clear(view); put(view, topbar('Налаштування')); showError(e); connSection(); return; }
  if (!alive(t)) return;
  clear(view);
  put(view, topbar('Налаштування'));
  const s = JSON.parse(JSON.stringify(c.settings));

  // Ключі
  const keyInputs = {};
  const keyRows = [
    ['gemini', 'Gemini (Google AI Studio)', 'https://aistudio.google.com/apikey'],
    ['groq', 'Groq', 'https://console.groq.com/keys'],
    ['openai', 'OpenAI (ChatGPT)', 'https://platform.openai.com/api-keys'],
    ['elevenlabs', 'ElevenLabs', 'https://elevenlabs.io/app/settings/api-keys'],
    ['anthropic', 'Claude (Anthropic) — лише для підсумків', 'https://console.anthropic.com/settings/keys']
  ].map(([k, label, url]) => {
    const has = c.keys[k];
    const inp = h('input', { class: 'input', type: 'password', autocomplete: 'off', placeholder: has ? 'збережено ••••' + String(has).slice(-4) : 'встав ключ' });
    keyInputs[k] = inp;
    return h('div', { class: 'field' },
      h('div', { class: 'row between' }, h('span', null, label, has ? h('span', { class: 'badge ok' }, 'є') : null),
        h('a', { href: url, target: '_blank', rel: 'noopener', class: 'linkish small' }, 'де взяти ключ ↗')),
      h('div', { class: 'row' }, inp, has ? h('button', { class: 'btn small ghost', type: 'button', onclick: async () => {
        if (!await confirmDialog(`Видалити ключ ${label}?`, 'Видалити', true)) return;
        await api.saveConfig({ keys: { [k]: '' } }); render();
      } }, 'Прибрати') : null));
  });

  // Розпізнавання
  const lang = h('select', { class: 'input' }, h('option', { value: 'uk' }, 'Українська'), h('option', { value: 'auto' }, 'Визначати автоматично (змішана мова)'));
  lang.value = s.language;
  const defBoxes = Object.values(ENGINES).map(e => {
    const cb = h('input', { type: 'checkbox', checked: (s.defaultEngines || []).includes(e.id) });
    return { id: e.id, el: h('label', { class: 'check' }, cb, ' ' + e.name), cb };
  });
  const ns = h('input', { type: 'checkbox', checked: !!s.noiseSuppression });

  // Аналіз
  const anSel = h('select', { class: 'input' }, Object.values(ANALYZERS).map(a => h('option', { value: a.id }, a.name)));
  anSel.value = s.analysisEngine;
  const auto = h('input', { type: 'checkbox', checked: !!s.autoAnalyze });
  const autoDoc = h('input', { type: 'checkbox', checked: !!s.autoDoc });

  // Моделі
  const modelInputs = {};
  const modelRows = [
    ...Object.values(ENGINES).map(e => ['models', e.id, 'Розпізнавання: ' + e.name]),
    ...Object.values(ANALYZERS).map(a => ['analysisModels', a.id, 'Підсумок: ' + a.name])
  ].map(([group, id, label]) => {
    const inp = h('input', { class: 'input', value: (s[group] || {})[id] || '' });
    modelInputs[group + '.' + id] = inp;
    return h('label', { class: 'field' }, label, inp);
  });

  // Календар
  const calSel = h('select', { class: 'input' }, cals.map(x => h('option', { value: x.id }, x.name + (x.primary ? ' (основний)' : ''))));
  if (c.info) calSel.value = c.info.calendarId;

  const saveAll = async (e) => {
    e.target.disabled = true;
    try {
      const keys = {};
      for (const [k, inp] of Object.entries(keyInputs)) if (inp.value.trim()) keys[k] = inp.value.trim();
      s.language = lang.value;
      s.defaultEngines = defBoxes.filter(b => b.cb.checked).map(b => b.id);
      s.noiseSuppression = ns.checked;
      s.analysisEngine = anSel.value;
      s.autoAnalyze = auto.checked;
      s.autoDoc = autoDoc.checked;
      for (const [k, inp] of Object.entries(modelInputs)) { const [g, id] = k.split('.'); s[g] = s[g] || {}; if (inp.value.trim()) s[g][id] = inp.value.trim(); }
      const part = { keys, settings: s };
      if (cals.length && c.info && calSel.value !== c.info.calendarId) part.calendarId = calSel.value;
      await api.saveConfig(part);
      toast('Збережено', 'ok');
      render();
    } catch (err) { toast(err.message, 'err'); e.target.disabled = false; }
  };

  put(view, h('div', { class: 'settings' },
    h('section', { class: 'card' }, h('h2', null, 'Ключі API'),
      h('p', { class: 'hint' }, 'Ключі зберігаються у твоєму скрипті Google і підтягуються на всі пристрої. Досить одного — Gemini; решту додай, щоб порівнювати.'),
      keyRows),
    h('section', { class: 'card' }, h('h2', null, 'Розпізнавання'),
      h('label', { class: 'field' }, 'Мова нарад', lang),
      h('div', { class: 'field' }, 'Двигуни за замовчуванням', h('div', { class: 'stack tight' }, defBoxes.map(b => b.el))),
      h('label', { class: 'check' }, ns, ' Шумозаглушення під час запису (вимкни, якщо голоси звучать «зажовано»)')),
    h('section', { class: 'card' }, h('h2', null, 'Підсумок'),
      h('label', { class: 'field' }, 'Хто робить підсумок', anSel),
      h('label', { class: 'check' }, auto, ' Робити підсумок одразу після розпізнавання'),
      h('label', { class: 'check' }, autoDoc, ' Створювати Google Doc з підсумком і текстом')),
    h('section', { class: 'card' }, h('h2', null, 'Календар'),
      cals.length ? h('label', { class: 'field' }, 'Куди записувати наради', calSel) : h('p', { class: 'hint' }, 'Не вдалося отримати список календарів.'),
      h('p', { class: 'hint' }, 'За замовчуванням — окремий календар «Наради». Його видно в Google Календарі на телефоні поруч з основним.')),
    h('details', { class: 'card' }, h('summary', null, 'Моделі (для досвідчених)'),
      h('p', { class: 'hint' }, 'Назви моделей у сервісів змінюються. Якщо двигун пише, що модель не знайдено, — вкажи тут актуальну.'),
      modelRows),
    h('div', { class: 'row pad' }, h('button', { class: 'btn primary', onclick: saveAll }, 'Зберегти налаштування'))));
  connSection(c);
}

function connSection(c) {
  const info = c && c.info;
  put(view, h('section', { class: 'card settings' }, h('h2', null, 'Підключення'),
    info ? h('div', { class: 'stack tight' },
      h('div', null, 'Акаунт: ', h('b', null, info.email || '—')),
      h('div', null, 'Календар: ', h('b', null, info.calendarName)),
      h('div', null, 'Скрипт Google: ', h('b', null, 'версія ' + (info.version || '—'))),
      h('div', { class: 'row wrap' },
        h('a', { class: 'btn small', href: info.folderUrl, target: '_blank', rel: 'noopener' }, icon('open'), 'Папка на Диску'),
        h('a', { class: 'btn small', href: info.sheetUrl, target: '_blank', rel: 'noopener' }, icon('open'), 'Таблиця нарад'))) : null,
    h('div', { class: 'row wrap' },
      S.installEvt ? h('button', { class: 'btn small primary', onclick: async () => { S.installEvt.prompt(); S.installEvt = null; } }, 'Встановити як додаток') : null,
      h('button', { class: 'btn small ghost', onclick: async () => {
        if (!await confirmDialog('Відключити цей пристрій? Дані на Диску лишаться, потім просто підключишся знову.', 'Відключити')) return;
        api.setConn(null); go('#/connect');
      } }, 'Відключити цей пристрій')),
    h('p', { class: 'hint' }, 'Версія ' + APP_VERSION)));
}

function viewConnect() {
  const conn = api.getConn() || {};
  const url = h('input', { class: 'input', placeholder: 'https://script.google.com/macros/s/…/exec', value: conn.url || '' });
  const token = h('input', { class: 'input', type: 'password', placeholder: 'Секретний ключ зі скрипта', value: conn.token || '' });
  const btn = h('button', { class: 'btn primary', type: 'submit' }, 'Підключити');
  put(view, 
    topbar('Підключення'),
    h('form', { class: 'form card', onsubmit: async (e) => {
      e.preventDefault();
      const c = { url: url.value.trim(), token: token.value.trim() };
      if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(c.url)) { toast('Адреса має починатися з https://script.google.com/', 'err'); return; }
      btn.disabled = true;
      try {
        const info = await api.call('ping', {}, c);
        api.setConn(c);
        toast('Підключено: ' + (info.email || ''), 'ok');
        await api.loadConfig(true).catch(() => { });
        go('#/calendar');
      } catch (err) { toast(err.message, 'err', 7000); btn.disabled = false; }
    } },
      h('p', null, 'Встав адресу бекенду і секретний ключ — їх ти отримаєш, коли налаштуєш скрипт Google.'),
      h('label', { class: 'field' }, 'Адреса бекенду (Web app URL)', url),
      h('label', { class: 'field' }, 'Секретний ключ', token),
      h('div', { class: 'row' }, btn)));
}

/* =================== Запуск =================== */

function buildShell() {
  const nav = h('nav', { class: 'nav' },
    h('div', { class: 'brand' }, h('span', { class: 'logo' }, icon('wave')), 'Наради'),
    [['calendar', '#/calendar', 'Календар', 'calendar'], ['recordings', '#/recordings', 'Записи', 'list'], ['tasks', '#/tasks', 'Задачі', 'check'], ['people', '#/people', 'Люди', 'people'], ['settings', '#/settings', 'Налаштування', 'settings']]
      .map(([sec, href, label, ic]) => h('a', { href, dataset: { sec } }, icon(ic), h('span', null, label))));
  document.body.insertBefore(nav, document.body.firstChild);
}

window.addEventListener('hashchange', render);
window.addEventListener('beforeunload', (e) => { if (S.recorder || anyJobRunning()) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); S.installEvt = e; });
window.addEventListener('unhandledrejection', (e) => { const m = e.reason && e.reason.message; if (m) toast(m, 'err', 6000); });

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('./sw.js').catch(() => { });
}

buildShell();
(async () => {
  // Спершу підтягуємо налаштування й версію скрипта, щоб показати лише те, що він уміє.
  if (api.getConn()) { put(view, spinner()); try { await api.loadConfig(); } catch { } }
  render();
})();
