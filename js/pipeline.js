// Обробка наради: стиснення → Диск → розпізнавання → підсумок → Google Doc.
// Прогрес кожної наради видно в її картці навіть якщо перейти на інший екран.

import { prepareAudio } from './audio.js';
import { uploadBlob, downloadFile } from './drive.js';
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
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && anyJobRunning()) holdWake(); });

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
  set(id, patch) { const s = this.steps.find(x => x.id === id); if (s) Object.assign(s, patch); emit(this.meetingId); }
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

async function loadPrepared(meeting, job) {
  let prepared = await getPrepared(meeting.id);
  if (prepared && prepared.fileId === meeting.audioFileId) return prepared;
  if (!meeting.audioFileId) throw new Error('У наради ще немає аудіо');
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
 */
export async function processAudio(meeting, blob, { fileName, engines = [], analyzeAfter = true, onDone, onUploaded } = {}) {
  const job = new Job(meeting.id);
  try {
    job.add('prep', 'Стиснення аудіо');
    job.add('up', 'Збереження на Google Диск');
    job.set('prep', { status: 'run', detail: 'Читаю файл…' });
    const prepared = await prepareAudio(blob, (stage, f) => {
      job.set('prep', { detail: stage === 'decode' ? 'Читаю файл…' : 'Стискаю…', progress: stage === 'encode' ? f : null });
    });
    job.set('prep', { status: 'ok', progress: null, detail: `${Math.round(prepared.duration / 60)} хв, ${(prepared.full.size / 1048576).toFixed(1)} МБ` });

    job.set('up', { status: 'run' });
    const d = new Date(meeting.start);
    const name = `${d.toISOString().slice(0, 10)} ${fmtTime(d).replace(':', '.')} ${meeting.title}.mp3`.replace(/[\\/:*?"<>|]/g, ' ');
    const file = await uploadBlob(prepared.full, name, meeting.id, (f) => job.set('up', { progress: f }));
    prepared.fileId = file.id;
    await savePrepared(meeting.id, prepared);
    trimPrepared();
    let m = await call('attachAudio', { meetingId: meeting.id, fileId: file.id, name: fileName || name, durationSec: prepared.duration });
    job.meeting = m;
    job.set('up', { status: 'ok', progress: null });
    if (onUploaded) { try { await onUploaded(m); } catch { } }

    if (engines.length) m = await runEngines(m, prepared, engines, job, analyzeAfter);
    job.finish();
    if (onDone) onDone(m);
    return m;
  } catch (e) {
    markFailed(job, e);
    job.finish(e);
    throw e;
  }
}

/** Розпізнати вже збережений запис обраними двигунами. */
export async function runTranscription(meeting, engines, { analyzeAfter = true, onDone } = {}) {
  const job = new Job(meeting.id);
  try {
    const prepared = await loadPrepared(meeting, job);
    const m = await runEngines(meeting, prepared, engines, job, analyzeAfter);
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
export async function runAnalysis(meeting, { engine, onDone } = {}) {
  const job = new Job(meeting.id);
  try {
    const transcripts = await call('getTranscripts', { meetingId: meeting.id });
    const m = await analyzeStep(meeting, transcripts, job, engine);
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
  await Promise.all(engines.map(async (id) => {
    const sid = 'eng_' + id;
    job.set(sid, { status: 'run' });
    try {
      const ectx = { ...ctx, onStatus: (t) => job.set(sid, { detail: t }) };
      const tr = await transcribe(id, prepared, ectx, (i, n) => job.set(sid, { detail: n > 1 ? `частина ${i + 1} з ${n}` : 'обробка…' }));
      latest = await call('saveTranscript', { meetingId: meeting.id, engine: id, transcript: tr });
      results[id] = tr;
      const words = tr.text.split(/\s+/).filter(Boolean).length;
      job.set(sid, { status: 'ok', detail: `${words} слів, ${tr.elapsedSec} с${tr.warnings.length ? ' · ' + tr.warnings[0] : ''}` });
    } catch (e) {
      job.set(sid, { status: 'err', detail: e.message });
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
