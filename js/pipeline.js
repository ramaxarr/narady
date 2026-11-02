// Обробка наради: стиснення → Диск → розпізнавання → підсумок → Google Doc.
// Прогрес кожної наради видно в її картці навіть якщо перейти на інший екран.

import { prepareAudio, regroup, PIECE_SEC } from './audio.js';
import { uploadBlob, uploadPieces, downloadFile, driveToken } from './drive.js';
import { call, loadConfig, getPeople, getGlossary } from './api.js';
import { transcribe, analyze, ENGINES } from './engines.js';
import { savePrepared, getPrepared, trimPrepared } from './store.js';
import { fmtDateFull, fmtTime, ymd } from './ui.js';

const jobs = new Map();
const listeners = new Set();

export function onJobs(fn) { listeners.add(fn); return () => listeners.delete(fn); }
export function jobFor(meetingId) { return jobs.get(meetingId) || null; }
export function anyJobRunning() { for (const j of jobs.values()) if (j.running) return true; return false; }
function emit(id) { for (const fn of listeners) { try { fn(id, jobs.get(id)); } catch { } } }

let wake = null;
async function holdWake() { try { if (!wake && 'wakeLock' in navigator) { wake = await navigator.wakeLock.request('screen'); wake.addEventListener('release', () => { wake = null; }); } } catch { } }
function releaseWake() { if (!anyJobRunning() && wake) { try { wake.release(); } catch { } wake = null; } }
let hiddenAt = 0;
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') { hiddenAt = anyJobRunning() ? Date.now() : 0; return; }
  if (anyJobRunning()) {
    holdWake();
    if (hiddenAt && Date.now() - hiddenAt > 20000) {
      for (const fn of listeners) { try { fn('__resumed', null); } catch { } }
    }
  }
  hiddenAt = 0;
});

class Job {
  constructor(meetingId) {
    this.meetingId = meetingId;
    this.steps = [];
    this.running = true;
    this.error = null;
    this.meeting = null;
    jobs.set(meetingId, this);
    holdWake();
    emit(meetingId);
  }
  add(id, label) { this.steps.push({ id, label, status: 'wait', detail: '', progress: null }); emit(this.meetingId); }
  set(id, patch) {
    const s = this.steps.find(x => x.id === id);
    if (!s) return;
    if (patch.status === 'run' && !s.startedAt) s.startedAt = Date.now();
    if ((patch.status === 'ok' || patch.status === 'err') && s.startedAt && !s.endedAt) s.endedAt = Date.now();
    Object.assign(s, patch);
    emit(this.meetingId);
  }
  // Повільний двигун можна пропустити — решта й підсумок не чекатимуть на нього.
  skip(id) { const c = this.controls && this.controls[id]; if (c) c.abort(); }
  canSkip(id) { return !!(this.controls && this.controls[id] && !this.controls[id].signal.aborted); }
  finish(err) {
    this.running = false; this.error = err || null;
    emit(this.meetingId);
    releaseWake();
  }
}

export function clearJob(meetingId) { const j = jobs.get(meetingId); if (j && !j.running) { jobs.delete(meetingId); emit(meetingId); } }

export async function buildCtx(meeting) {
  const c = await loadConfig();
  const people = await getPeople();
  const glossary = await getGlossary();
  const participants = (meeting.participants || []).map(id => people.find(p => p.id === id)).filter(Boolean);
  return { keys: c.keys, settings: c.settings, participants, glossary, people };
}

export function speakerNamer(meeting, engine, people) {
  const map = ((meeting.speakerMap || {})[engine]) || {};
  return (label) => {
    const v = map[label];
    if (!v) return label;
    const p = people.find(x => x.id === v);
    return p ? p.name : v;
  };
}

// Для обробки на пристрої 5-хв частини склеюємо у фрагменти до ~20 хв — менше запитів.
function forDevice(p) { return p.pieces ? { ...p, chunks: regroup(p.pieces) } : p; }

function baseName(meeting) {
  const d = new Date(meeting.start);
  return `${ymd(d)} ${fmtTime(d).replace(':', '.')} ${meeting.title || 'Нарада'}`.replace(/[\\/:*?"<>|]/g, ' ');
}

/** Аудіо наради одним файлом (для програвача): з пам'яті пристрою або з Диска. */
export async function audioBlob(meeting, onProgress) {
  const p = await getPrepared(meeting.id);
  if (p && p.fileId === meeting.audioFileId && p.full) return p.full;
  const pieces = meeting.audioPieces || [];
  if (pieces.length > 1) {
    const auth = await driveToken();
    const parts = [];
    for (let i = 0; i < pieces.length; i++) parts.push(await downloadFile(pieces[i].id, (f) => onProgress && onProgress((i + f) / pieces.length), auth));
    return new Blob(parts, { type: 'audio/mpeg' });
  }
  return downloadFile(meeting.audioFileId, onProgress);
}

async function loadPrepared(meeting, job) {
  let prepared = await getPrepared(meeting.id);
  if (prepared && prepared.fileId === meeting.audioFileId) return forDevice(prepared);
  if (!meeting.audioFileId) throw new Error('У наради ще немає аудіо');
  const list = meeting.audioPieces || [];
  if (list.length) {
    // Запис лежить на Диску частинами — просто склеюємо, перекодовувати не треба.
    job.add('dl', 'Аудіо з Google Диска');
    job.set('dl', { status: 'run' });
    const auth = await driveToken();
    const pieces = [];
    for (let i = 0; i < list.length; i++) {
      const blob = await downloadFile(list[i].id, (f) => job.set('dl', { progress: (i + f) / list.length }), auth);
      pieces.push({ blob, offset: list[i].o, duration: list[i].d });
    }
    const last = list[list.length - 1];
    prepared = { full: new Blob(pieces.map(p => p.blob), { type: 'audio/mpeg' }), pieces, duration: last.o + last.d, fileId: meeting.audioFileId };
    job.set('dl', { status: 'ok', progress: null });
    await savePrepared(meeting.id, prepared);
    return forDevice(prepared);
  }
  job.add('dl', 'Аудіо з Google Диска');
  job.set('dl', { status: 'run' });
  const blob = await downloadFile(meeting.audioFileId, (f) => job.set('dl', { progress: f }));
  job.set('dl', { status: 'ok', progress: null });
  job.add('prep', 'Підготовка аудіо');
  job.set('prep', { status: 'run' });
  prepared = await prepareAudio(blob, (stage, f) => job.set('prep', { progress: stage === 'encode' ? f : null }), { isArchive: true });
  prepared.fileId = meeting.audioFileId;
  job.set('prep', { status: 'ok', progress: null });
  await savePrepared(meeting.id, prepared);
  return prepared;
}

/**
 * Новий запис або файл з диктофона: стиснути, зберегти на Диск, за бажанням одразу розпізнати.
 * pieces — зберігати на Диск частинами по ~5 хв (скрипт версії 1.2.0+), щоб їх міг обробити скрипт у фоні.
 * background — розпізнавання й підсумок робить скрипт Google; додаток після завантаження можна закрити.
 */
export async function processAudio(meeting, blob, { fileName, engines = [], analyzeAfter = true, onDone, onUploaded, pieces = false, background = false } = {}) {
  const job = new Job(meeting.id);
  const bg = !!(background && pieces && engines.length);
  job.bg = bg;
  try {
    job.add('prep', 'Стиснення аудіо');
    job.add('up', 'Збереження на Google Диск');
    job.set('prep', { status: 'run', detail: 'Читаю файл…' });
    const prepared = await prepareAudio(blob, (stage, f) => {
      job.set('prep', { detail: stage === 'decode' ? 'Читаю файл…' : 'Стискаю…', progress: stage === 'encode' ? f : null });
    }, pieces ? { chunkSec: PIECE_SEC, searchSec: 15 } : {});
    job.set('prep', { status: 'ok', progress: null, detail: `${Math.round(prepared.duration / 60)} хв, ${(prepared.full.size / 1048576).toFixed(1)} МБ` });

    job.set('up', { status: 'run' });
    const name = baseName(meeting);
    let m, local = prepared;
    if (pieces) {
      const list = await uploadPieces(prepared.chunks, name, meeting.id, (f, i, n) => job.set('up', { progress: f, detail: n > 1 ? `частина ${i + 1} з ${n}` : '' }));
      local = { full: prepared.full, pieces: prepared.chunks, duration: prepared.duration, fileId: list[0].id };
      await savePrepared(meeting.id, local);
      trimPrepared();
      m = await call('attachAudio', { meetingId: meeting.id, pieces: list, name: fileName || name + '.mp3', durationSec: prepared.duration });
      local = forDevice(local);
    } else {
      const file = await uploadBlob(prepared.full, name + '.mp3', meeting.id, (f) => job.set('up', { progress: f }));
      prepared.fileId = file.id;
      await savePrepared(meeting.id, prepared);
      trimPrepared();
      m = await call('attachAudio', { meetingId: meeting.id, fileId: file.id, name: fileName || name + '.mp3', durationSec: prepared.duration });
    }
    job.meeting = m;
    job.set('up', { status: 'ok', progress: null, detail: '' });
    if (onUploaded) { try { await onUploaded(m); } catch { } }

    if (bg) m = await queueStep(m, { engines, analyzeAfter }, job);
    else if (engines.length) m = await runEngines(m, local, engines, job, analyzeAfter);
    job.finish();
    if (onDone) onDone(m);
    return m;
  } catch (e) {
    markFailed(job, e);
    job.finish(e);
    throw e;
  }
}

/* ---------- Фонова обробка (скрипт Google) ---------- */

let kicking = false, kickPauseUntil = 0;
/** Просимо скрипт обробляти чергу зараз, не чекаючи тригера. Поки додаток відкритий — так швидше. */
export function kickQueue() {
  if (kicking || Date.now() < kickPauseUntil) return;
  kicking = true;
  call('runQueue')
    .then((r) => { if (r && r.busy) kickPauseUntil = Date.now() + 30000; })
    .catch(() => { kickPauseUntil = Date.now() + 30000; })
    .finally(() => { kicking = false; });
}

async function queueStep(meeting, opts, job) {
  job.add('queue', 'Передача у фонову обробку');
  job.set('queue', { status: 'run' });
  await call('queueJob', { meetingId: meeting.id, ...opts });
  kickQueue();
  job.set('queue', { status: 'ok', detail: 'Далі все робить скрипт Google — додаток можна закрити.' });
  job.background = true;
  const m = { ...meeting, status: 'processing' };
  job.meeting = m;
  return m;
}

/** Старий запис (одним файлом) ріжемо на 5-хв частини, щоб його міг обробити скрипт. */
async function toPieces(meeting, job) {
  job.add('dl', 'Аудіо з Google Диска');
  job.set('dl', { status: 'run' });
  const blob = await downloadFile(meeting.audioFileId, (f) => job.set('dl', { progress: f }));
  job.set('dl', { status: 'ok', progress: null });
  job.add('prep', 'Нарізка на частини по 5 хв');
  job.set('prep', { status: 'run' });
  const p = await prepareAudio(blob, (stage, f) => job.set('prep', { progress: stage === 'encode' ? f : null }), { isArchive: true, chunkSec: PIECE_SEC, searchSec: 15 });
  job.set('prep', { status: 'ok', progress: null, detail: p.chunks.length + ' част.' });
  let list;
  if (p.chunks.length === 1) list = [{ id: meeting.audioFileId, o: 0, d: Math.round(p.duration * 100) / 100 }];
  else {
    job.add('up', 'Збереження частин на Google Диск');
    job.set('up', { status: 'run' });
    list = await uploadPieces(p.chunks, baseName(meeting), meeting.id, (f, i, n) => job.set('up', { progress: f, detail: `частина ${i + 1} з ${n}` }));
    job.set('up', { status: 'ok', progress: null, detail: '' });
  }
  const m = await call('attachAudio', { meetingId: meeting.id, pieces: list, name: meeting.audioName || '', durationSec: p.duration });
  await savePrepared(meeting.id, { full: blob, pieces: p.chunks, duration: p.duration, fileId: list[0].id });
  return m;
}

/** Розпізнати вже збережений запис обраними двигунами. */
export async function runTranscription(meeting, engines, { analyzeAfter = true, onDone, background = false } = {}) {
  const job = new Job(meeting.id);
  job.bg = background;
  try {
    let m;
    if (background) {
      const src = meeting.audioPieces && meeting.audioPieces.length ? meeting : await toPieces(meeting, job);
      m = await queueStep(src, { engines, analyzeAfter }, job);
    } else {
      const prepared = await loadPrepared(meeting, job);
      m = await runEngines(meeting, prepared, engines, job, analyzeAfter);
    }
    job.finish();
    if (onDone) onDone(m);
    return m;
  } catch (e) {
    markFailed(job, e);
    job.finish(e);
    throw e;
  }
}

/** Лише підсумок (і документ) на основі вже готового тексту. */
export async function runAnalysis(meeting, { engine, onDone, background = false } = {}) {
  const job = new Job(meeting.id);
  job.bg = background;
  try {
    let m;
    if (background) m = await queueStep(meeting, { engines: [], analyzeOnly: true, analysisEngine: engine }, job);
    else {
      const transcripts = await call('getTranscripts', { meetingId: meeting.id });
      m = await analyzeStep(meeting, transcripts, job, engine);
    }
    job.finish();
    if (onDone) onDone(m);
    return m;
  } catch (e) {
    markFailed(job, e);
    job.finish(e);
    throw e;
  }
}

function markFailed(job, e) {
  const s = job.steps.find(x => x.status === 'run');
  if (s) job.set(s.id, { status: 'err', detail: e.message, progress: null });
}

async function runEngines(meeting, prepared, engines, job, analyzeAfter) {
  const ctx = await buildCtx(meeting);
  for (const id of engines) job.add('eng_' + id, 'Розпізнавання: ' + ENGINES[id].name);
  let latest = meeting;
  const results = {};
  job.controls = {};
  await Promise.all(engines.map(async (id) => {
    const sid = 'eng_' + id;
    const ctrl = new AbortController();
    job.controls[sid] = ctrl;
    job.set(sid, { status: 'run' });
    try {
      const ectx = { ...ctx, signal: ctrl.signal, onStatus: (t) => job.set(sid, { detail: t }) };
      const tr = await transcribe(id, prepared, ectx, (i, n) => job.set(sid, { detail: n > 1 ? `частина ${i + 1} з ${n}` : 'обробка…' }));
      latest = await call('saveTranscript', { meetingId: meeting.id, engine: id, transcript: tr });
      results[id] = tr;
      const words = tr.text.split(/\s+/).filter(Boolean).length;
      job.set(sid, { status: 'ok', detail: `${words} слів, ${tr.elapsedSec} с${tr.warnings.length ? ' · ' + tr.warnings[0] : ''}` });
    } catch (e) {
      job.set(sid, { status: 'err', detail: e.skipped ? 'пропущено — можна запустити пізніше через «Розпізнати…»' : e.message });
    } finally {
      delete job.controls[sid];
    }
  }));
  if (!Object.keys(results).length) throw new Error('Жоден двигун не впорався');
  if (!meeting.primaryEngine) {
    // Основним робимо перший обраний двигун, що розрізняє спікерів, а не той, що відповів найшвидше.
    const ok = engines.filter(id => results[id]);
    const best = ok.find(id => ENGINES[id].speakers) || ok[0];
    if (best && latest.primaryEngine !== best) latest = await call('saveMeetingState', { meetingId: meeting.id, primaryEngine: best });
  }
  job.meeting = latest;
  if (analyzeAfter) {
    const all = await call('getTranscripts', { meetingId: meeting.id });
    latest = await analyzeStep(latest, all, job);
  }
  return latest;
}

function pickPrimary(meeting, transcripts) {
  if (meeting.primaryEngine && transcripts[meeting.primaryEngine]) return meeting.primaryEngine;
  const ids = Object.keys(transcripts);
  return ids.find(id => ENGINES[id] && ENGINES[id].speakers) || ids[0];
}

async function analyzeStep(meeting, transcripts, job, engineOverride) {
  const primary = pickPrimary(meeting, transcripts);
  if (!primary) throw new Error('Немає тексту для підсумку');
  const ctx = await buildCtx(meeting);
  if (engineOverride) ctx.settings = { ...ctx.settings, analysisEngine: engineOverride };
  ctx.nameOf = speakerNamer(meeting, primary, ctx.people);
  job.add('an', 'Підсумок і задачі');
  job.set('an', { status: 'run' });
  ctx.onStatus = (t) => job.set('an', { detail: t });
  try { ctx.carryTasks = await call('carriedTasks', { meetingId: meeting.id }); } catch { ctx.carryTasks = []; }
  const d = new Date(meeting.start);
  const meta = { title: meeting.title, dateText: `${fmtDateFull(d)}, ${fmtTime(d)}`, dateISO: ymd(d), weekday: new Intl.DateTimeFormat('uk-UA', { weekday: 'long' }).format(d) };
  const analysis = await analyze(transcripts[primary], meta, ctx);
  let m = await call('saveAnalysis', { meetingId: meeting.id, analysis });
  job.set('an', { status: 'ok', detail: `${analysis.tasks.length} задач, ${analysis.decisions.length} рішень${analysis.reportedDone.length ? `, ${analysis.reportedDone.length} з минулих — схоже, виконано` : ''}` });
  if (ctx.settings.autoDoc) {
    job.add('doc', 'Google Doc');
    job.set('doc', { status: 'run' });
    try {
      const r = await call('buildDoc', { meetingId: meeting.id });
      m = { ...m, docId: r.docId };
      job.set('doc', { status: 'ok' });
    } catch (e) { job.set('doc', { status: 'err', detail: e.message }); }
  }
  job.meeting = m;
  return m;
}
