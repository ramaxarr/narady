// Двигуни розпізнавання й аналізу. Кожен двигун повертає текст у спільному форматі:
// { segments: [{start, end, speaker, text}], text, speakers, ... }
// Запити йдуть напряму з браузера; якщо сервіс не пускає браузер (CORS), — через бекенд.

import { call } from './api.js';
import { blobToBase64 } from './audio.js';

export const ENGINES = {
  gemini: {
    id: 'gemini', name: 'Gemini', keyName: 'gemini', speakers: true, freeTier: true,
    maxBytes: 40e6, maxSec: 1300,
    note: 'Є безкоштовний рівень (Google може використовувати ці дані для покращення моделей). Розділяє спікерів і підписує їх іменами з контексту.',
    keyUrl: 'https://aistudio.google.com/apikey'
  },
  groq: {
    id: 'groq', name: 'Groq Whisper', keyName: 'groq', speakers: false, freeTier: true,
    maxBytes: 24e6, maxSec: Infinity,
    note: 'Whisper large-v3, дуже швидкий. Є безкоштовний рівень з лімітами. Спікерів не розділяє.',
    keyUrl: 'https://console.groq.com/keys'
  },
  openai: {
    id: 'openai', name: 'OpenAI (ChatGPT)', keyName: 'openai', speakers: true, freeTier: false,
    maxBytes: 24e6, maxSec: 1380,
    note: 'Платний. Модель з розділенням спікерів (gpt-4o-transcribe-diarize) OpenAI вимкне 26.02.2027.',
    keyUrl: 'https://platform.openai.com/api-keys'
  },
  elevenlabs: {
    id: 'elevenlabs', name: 'ElevenLabs Scribe', keyName: 'elevenlabs', speakers: true, freeTier: false,
    maxBytes: 900e6, maxSec: 9 * 3600,
    note: 'Розділяє спікерів по всьому запису одразу. Є стартові безкоштовні хвилини.',
    keyUrl: 'https://elevenlabs.io/app/settings/api-keys'
  }
};

export const ANALYZERS = {
  gemini: { id: 'gemini', name: 'Gemini', keyName: 'gemini', keyUrl: 'https://aistudio.google.com/apikey' },
  openai: { id: 'openai', name: 'OpenAI (ChatGPT)', keyName: 'openai', keyUrl: 'https://platform.openai.com/api-keys' },
  anthropic: { id: 'anthropic', name: 'Claude', keyName: 'anthropic', keyUrl: 'https://console.anthropic.com/settings/keys' }
};

/* ---------- HTTP з запасним шляхом через бекенд ---------- */

async function httpPost(url, { headers = {}, json, form, timeoutMs = 20 * 60 * 1000 }) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    let body, h = { ...headers };
    if (json !== undefined) { body = JSON.stringify(json); h['Content-Type'] = 'application/json'; }
    else {
      body = new FormData();
      for (const f of form) {
        if (f.blob) body.append(f.name, f.blob, f.filename || 'audio.mp3');
        else body.append(f.name, f.value);
      }
    }
    const res = await fetch(url, { method: 'POST', headers: h, body, signal: ctrl.signal });
    return { status: res.status, text: await res.text(), via: 'direct' };
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Сервіс не відповів за відведений час');
    if (!(e instanceof TypeError) || !navigator.onLine) throw e;
    // Браузер не пустив запит (найчастіше — CORS). Пробуємо через бекенд.
    let pform;
    if (form) {
      pform = [];
      for (const f of form) {
        if (f.blob) pform.push({ name: f.name, filename: f.filename || 'audio.mp3', mimeType: f.blob.type || 'audio/mpeg', base64: await blobToBase64(f.blob) });
        else pform.push({ name: f.name, value: f.value });
      }
    }
    const r = await call('proxy', { url, headers, json, form: pform });
    return { ...r, via: 'proxy' };
  } finally {
    clearTimeout(timer);
  }
}

// Тимчасові відмови (перевантаження, ліміт за хвилину) пробуємо ще тричі з паузами.
const RETRY_STATUS = [429, 500, 502, 503, 504];
export const RETRY_DELAYS = [10, 20, 40];
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function post(ctx, name, url, opts) {
  for (let attempt = 0; ; attempt++) {
    const r = await httpPost(url, opts);
    if (!RETRY_STATUS.includes(r.status) || attempt >= RETRY_DELAYS.length) return r;
    const total = RETRY_DELAYS.length + 1;
    for (let s = RETRY_DELAYS[attempt]; s > 0; s--) {
      if (ctx.onStatus) ctx.onStatus(`${name} перевантажений (${r.status}), повтор через ${s} с — спроба ${attempt + 2} з ${total}`);
      await sleep(1000);
    }
    if (ctx.onStatus) ctx.onStatus(`${name}: спроба ${attempt + 2} з ${total}…`);
  }
}

function apiError(name, r) {
  let msg = r.text || '';
  try {
    const j = JSON.parse(r.text);
    msg = (j.error && (j.error.message || j.error.type || JSON.stringify(j.error))) || j.detail?.message || (typeof j.detail === 'string' ? j.detail : '') || j.message || msg;
    if (typeof msg !== 'string') msg = JSON.stringify(msg);
  } catch { }
  const busy = [500, 502, 503, 504].includes(r.status);
  const hint = r.status === 401 || r.status === 403 ? ' Перевір ключ API в Налаштуваннях.'
    : r.status === 429 ? ' Перевищено ліміт — спробуй трохи пізніше.'
    : busy ? ` Сервіс тимчасово перевантажений і не відповів після ${RETRY_DELAYS.length + 1} спроб. Натисни «Розпізнати…» ще раз за кілька хвилин або обери інший двигун.` : '';
  const e = new Error(`${name}: помилка ${r.status}. ${String(msg).slice(0, 300)}${hint}`);
  e.status = r.status; e.body = msg;
  return e;
}

/* ---------- Допоміжне ---------- */

export function parseTs(s) {
  if (typeof s === 'number') return s;
  const parts = String(s || '0').trim().replace(',', '.').split(':').map(Number);
  if (parts.some(isNaN)) return 0;
  return parts.reduce((acc, p) => acc * 60 + p, 0);
}

function parseJsonLoose(text) {
  const t = String(text || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch { }
  const a = t.indexOf('{');
  if (a >= 0) {
    const b = t.lastIndexOf('}');
    try { return JSON.parse(t.slice(a, b + 1)); } catch { }
    // Обрізана відповідь: відкидаємо незавершений хвіст масиву.
    let pos = t.length;
    for (let k = 0; k < 400; k++) {
      pos = t.lastIndexOf('}', pos - 1);
      if (pos <= a) break;
      for (const tail of [']}', ']}]}', '}']) {
        try { const j = JSON.parse(t.slice(a, pos + 1) + tail); j.__truncated = true; return j; } catch { }
      }
    }
  }
  throw new Error('Не вдалося розібрати відповідь моделі');
}

function hintTerms(ctx, max = 100) {
  const out = [];
  for (const p of ctx.participants || []) {
    if (p.name) out.push(p.name);
    for (const a of String(p.aliases || '').split(/[,;]/)) if (a.trim()) out.push(a.trim());
  }
  for (const t of ctx.glossary || []) if (t) out.push(t);
  return [...new Set(out)].slice(0, max);
}

function groupPlain(segs, maxLen = 45) {
  // Для двигунів без спікерів: склеюємо короткі шматки в абзаци до ~45 секунд.
  const out = [];
  for (const s of segs) {
    const last = out[out.length - 1];
    if (last && s.start - last.start < maxLen && s.start - last.end < 2.5) {
      last.text = (last.text + ' ' + s.text).trim(); last.end = s.end;
    } else out.push({ ...s, speaker: '' });
  }
  return out;
}

function plan(engine, audio) {
  if (audio.full.size <= engine.maxBytes && audio.duration <= engine.maxSec) {
    return [{ blob: audio.full, offset: 0, duration: audio.duration }];
  }
  return audio.chunks;
}

/* ---------- Gemini ---------- */

const SEG_SCHEMA = {
  type: 'OBJECT',
  properties: {
    segments: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          start: { type: 'STRING', description: 'час початку репліки MM:SS від початку цього аудіо' },
          speaker: { type: 'STRING' },
          text: { type: 'STRING' }
        },
        required: ['start', 'speaker', 'text'],
        propertyOrdering: ['start', 'speaker', 'text']
      }
    }
  },
  required: ['segments']
};

function geminiPrompt(ctx, idx, total, prevTail) {
  const people = (ctx.participants || []).map(p => `- ${p.name}${p.role ? ' — ' + p.role : ''}`).join('\n') || '- (не вказано)';
  const lang = ctx.settings.language === 'uk' ? 'Основна мова — українська.' : '';
  let s = `Зроби дослівну розшифровку аудіозапису робочої наради на виробництві.
${lang} Пиши мовою оригіналу: не перекладай, не виправляй стиль; якщо хтось говорить російською чи суржиком — записуй як є.
Розбий текст на репліки за спікерами. Однаковий голос — завжди та сама мітка.
Учасники наради:
${people}
Якщо з розмови зрозуміло, хто говорить (звертаються на ім'я, згадують посаду), підписуй репліку іменем учасника. Якщо ні — «Спікер 1», «Спікер 2» і т. д.
Для кожної репліки вкажи час початку у форматі MM:SS від початку ЦЬОГО аудіо.
Нічого не скорочуй, не узагальнюй і не пропускай. Нерозбірливе місце познач [нерозбірливо].`;
  const terms = hintTerms(ctx, 150);
  if (terms.length) s += `\nСлова й назви, які можуть звучати: ${terms.join(', ')}.`;
  if (total > 1) {
    s += `\nЦе фрагмент ${idx + 1} з ${total} одного запису.`;
    if (prevTail) s += ` Кінець попереднього фрагмента (для узгодження міток спікерів):\n${prevTail}`;
  }
  return s;
}

async function geminiTranscribe(audio, ctx, onProgress) {
  const key = ctx.keys.gemini;
  const model = ctx.settings.models.gemini || 'gemini-2.5-flash';
  const pieces = plan(ENGINES.gemini, audio);
  const segments = [], warnings = [];
  let prevTail = '';
  for (let i = 0; i < pieces.length; i++) {
    onProgress(i, pieces.length);
    const p = pieces[i];
    const body = {
      contents: [{ role: 'user', parts: [
        { inlineData: { mimeType: 'audio/mp3', data: await blobToBase64(p.blob) } },
        { text: geminiPrompt(ctx, i, pieces.length, prevTail) }
      ] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: SEG_SCHEMA, maxOutputTokens: 65536 }
    };
    if (/2\.5-flash/.test(model)) body.generationConfig.thinkingConfig = { thinkingBudget: 0 };
    const r = await post(ctx, 'Gemini', `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
      { headers: { 'x-goog-api-key': key }, json: body });
    if (r.status !== 200) throw apiError('Gemini', r);
    const data = JSON.parse(r.text);
    const cand = (data.candidates || [])[0];
    if (!cand) throw new Error('Gemini не повернув відповіді' + (data.promptFeedback ? ': ' + JSON.stringify(data.promptFeedback) : ''));
    const text = ((cand.content && cand.content.parts) || []).map(x => x.text || '').join('');
    const j = parseJsonLoose(text);
    if (cand.finishReason && cand.finishReason !== 'STOP') warnings.push(`Фрагмент ${i + 1}: відповідь обірвалась (${cand.finishReason})`);
    else if (j.__truncated) warnings.push(`Фрагмент ${i + 1}: відповідь обірвалась`);
    const segs = (j.segments || []).filter(s => s && s.text).map(s => ({
      start: p.offset + Math.min(parseTs(s.start), p.duration), speaker: String(s.speaker || '').trim(), text: String(s.text).trim()
    }));
    segs.sort((a, b) => a.start - b.start);
    segs.forEach((s, k) => { s.end = k + 1 < segs.length ? segs[k + 1].start : p.offset + p.duration; });
    segments.push(...segs);
    prevTail = segs.slice(-4).map(s => `${s.speaker}: ${s.text.slice(0, 160)}`).join('\n');
  }
  return { model, pieces: pieces.length, segments, warnings };
}

/* ---------- Groq Whisper ---------- */

async function groqTranscribe(audio, ctx, onProgress) {
  const key = ctx.keys.groq;
  const model = ctx.settings.models.groq || 'whisper-large-v3';
  const pieces = plan(ENGINES.groq, audio);
  const raw = [];
  const prompt = hintTerms(ctx, 60).join(', ').slice(0, 600);
  for (let i = 0; i < pieces.length; i++) {
    onProgress(i, pieces.length);
    const p = pieces[i];
    const form = [
      { name: 'file', blob: p.blob, filename: 'audio.mp3' },
      { name: 'model', value: model },
      { name: 'response_format', value: 'verbose_json' },
      { name: 'temperature', value: '0' }
    ];
    if (ctx.settings.language === 'uk') form.push({ name: 'language', value: 'uk' });
    if (prompt) form.push({ name: 'prompt', value: prompt });
    const r = await post(ctx, 'Groq', 'https://api.groq.com/openai/v1/audio/transcriptions', { headers: { Authorization: 'Bearer ' + key }, form });
    if (r.status !== 200) throw apiError('Groq', r);
    const j = JSON.parse(r.text);
    for (const s of j.segments || []) {
      const text = String(s.text || '').trim();
      if (text) raw.push({ start: p.offset + s.start, end: p.offset + s.end, text });
    }
    if (!(j.segments || []).length && j.text) raw.push({ start: p.offset, end: p.offset + p.duration, text: String(j.text).trim() });
  }
  return { model, pieces: pieces.length, segments: groupPlain(raw), warnings: [] };
}

/* ---------- OpenAI ---------- */

async function openaiTranscribe(audio, ctx, onProgress) {
  const key = ctx.keys.openai;
  const model = ctx.settings.models.openai || 'gpt-4o-transcribe-diarize';
  const diarize = /diarize/.test(model);
  const pieces = plan(ENGINES.openai, audio);
  const segments = [], warnings = [];
  for (let i = 0; i < pieces.length; i++) {
    onProgress(i, pieces.length);
    const p = pieces[i];
    const base = [{ name: 'file', blob: p.blob, filename: 'audio.mp3' }, { name: 'model', value: model }];
    let form;
    if (diarize) {
      form = [...base, { name: 'response_format', value: 'diarized_json' }, { name: 'chunking_strategy', value: 'auto' }];
    } else {
      form = [...base, { name: 'response_format', value: 'json' }];
      const kw = hintTerms(ctx, 60);
      if (kw.length) form.push({ name: 'prompt', value: 'Робоча нарада на виробництві. Назви й імена: ' + kw.join(', ') });
    }
    const withLang = ctx.settings.language === 'uk' ? [...form, { name: diarize ? 'language' : 'languages[]', value: 'uk' }] : form;
    let r = await post(ctx, 'OpenAI', 'https://api.openai.com/v1/audio/transcriptions', { headers: { Authorization: 'Bearer ' + key }, form: withLang });
    if (r.status === 400 && withLang !== form && /language/i.test(r.text)) {
      r = await post(ctx, 'OpenAI', 'https://api.openai.com/v1/audio/transcriptions', { headers: { Authorization: 'Bearer ' + key }, form });
    }
    if (r.status !== 200) throw apiError('OpenAI', r);
    const j = JSON.parse(r.text);
    if (diarize && Array.isArray(j.segments)) {
      for (const s of j.segments) {
        const text = String(s.text || '').trim();
        if (!text) continue;
        const spk = String(s.speaker ?? '');
        segments.push({ start: p.offset + (s.start || 0), end: p.offset + (s.end || 0), speaker: pieces.length > 1 ? `${i + 1}·${spk}` : spk, text });
      }
    } else if (j.text) {
      segments.push({ start: p.offset, end: p.offset + p.duration, speaker: '', text: String(j.text).trim() });
    }
  }
  if (diarize && pieces.length > 1) warnings.push('Запис оброблено частинами: мітки спікерів у різних частинах незалежні (1·A, 2·A…).');
  return { model, pieces: pieces.length, segments: mergeSameSpeaker(segments), warnings };
}

function mergeSameSpeaker(segs) {
  const out = [];
  for (const s of segs) {
    const last = out[out.length - 1];
    if (last && last.speaker && last.speaker === s.speaker && s.start - last.end < 1.5 && (last.end - last.start) < 60) {
      last.text += ' ' + s.text; last.end = s.end;
    } else out.push({ ...s });
  }
  return out;
}

/* ---------- ElevenLabs Scribe ---------- */

async function elevenTranscribe(audio, ctx, onProgress) {
  const key = ctx.keys.elevenlabs;
  const model = ctx.settings.models.elevenlabs || 'scribe_v2';
  const pieces = plan(ENGINES.elevenlabs, audio);
  const segments = [], warnings = [];
  for (let i = 0; i < pieces.length; i++) {
    onProgress(i, pieces.length);
    const p = pieces[i];
    const form = [
      { name: 'file', blob: p.blob, filename: 'audio.mp3' },
      { name: 'model_id', value: model },
      { name: 'diarize', value: 'true' },
      { name: 'timestamps_granularity', value: 'word' },
      { name: 'tag_audio_events', value: 'false' }
    ];
    if (ctx.settings.language === 'uk') form.push({ name: 'language_code', value: 'uk' });
    const terms = hintTerms(ctx, 100).filter(t => t.length <= 50);
    const withTerms = [...form, ...terms.map(t => ({ name: 'keyterms', value: t }))];
    let r = await post(ctx, 'ElevenLabs', 'https://api.elevenlabs.io/v1/speech-to-text', { headers: { 'xi-api-key': key }, form: terms.length ? withTerms : form });
    if (terms.length && (r.status === 400 || r.status === 422)) {
      warnings.push('Словник не прийнято сервісом, розпізнано без нього.');
      r = await post(ctx, 'ElevenLabs', 'https://api.elevenlabs.io/v1/speech-to-text', { headers: { 'xi-api-key': key }, form });
    }
    if (r.status !== 200) throw apiError('ElevenLabs', r);
    const j = JSON.parse(r.text);
    let cur = null;
    for (const w of j.words || []) {
      if (w.type === 'audio_event') continue;
      const spk = (pieces.length > 1 ? `${i + 1}·` : '') + String(w.speaker_id || '');
      if (w.type === 'spacing') { if (cur) cur.text += w.text; continue; }
      const st = p.offset + (w.start || 0), en = p.offset + (w.end || w.start || 0);
      const brk = !cur || cur.speaker !== spk || st - cur.end > 2.5 || (cur.text.length > 600 && /[.!?…]\s*$/.test(cur.text));
      if (brk) { if (cur) segments.push(cur); cur = { start: st, end: en, speaker: spk, text: w.text }; }
      else { cur.text += w.text; cur.end = en; }
    }
    if (cur) segments.push(cur);
    if (!(j.words || []).length && j.text) segments.push({ start: p.offset, end: p.offset + p.duration, speaker: '', text: j.text });
  }
  segments.forEach(s => { s.text = s.text.replace(/\s+/g, ' ').trim(); });
  return { model, pieces: pieces.length, segments: segments.filter(s => s.text), warnings };
}

const RUNNERS = { gemini: geminiTranscribe, groq: groqTranscribe, openai: openaiTranscribe, elevenlabs: elevenTranscribe };

export async function transcribe(engineId, audio, ctx, onProgress = () => { }) {
  const eng = ENGINES[engineId];
  if (!eng) throw new Error('Невідомий двигун: ' + engineId);
  if (!ctx.keys[eng.keyName]) throw new Error(`${eng.name}: не додано ключ API`);
  const t0 = Date.now();
  const r = await RUNNERS[engineId](audio, ctx, onProgress);
  if (!r.segments.some(s => s.text && s.text.trim())) throw new Error(`${eng.name}: сервіс повернув порожній текст. Перевір, чи в записі чути мову.`);
  const speakers = [...new Set(r.segments.map(s => s.speaker).filter(Boolean))];
  return {
    engine: engineId,
    engineName: eng.name,
    model: r.model,
    language: ctx.settings.language,
    createdAt: new Date().toISOString(),
    durationSec: audio.duration,
    pieces: r.pieces,
    elapsedSec: Math.round((Date.now() - t0) / 1000),
    segments: r.segments.map(s => ({ start: +s.start.toFixed(2), end: +(s.end || s.start).toFixed(2), speaker: s.speaker || '', text: s.text })),
    text: r.segments.map(s => s.text).join(' '),
    speakers,
    warnings: r.warnings || []
  };
}

/* ---------- Аналіз наради ---------- */

const ANALYSIS_SCHEMA = {
  type: 'OBJECT',
  properties: {
    summary: { type: 'STRING' },
    decisions: { type: 'ARRAY', items: { type: 'STRING' } },
    tasks: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: { owner: { type: 'STRING' }, task: { type: 'STRING' }, due: { type: 'STRING' }, dueDate: { type: 'STRING', description: 'РРРР-ММ-ДД або порожньо' } },
        required: ['owner', 'task', 'due', 'dueDate']
      }
    },
    openQuestions: { type: 'ARRAY', items: { type: 'STRING' } },
    facts: { type: 'ARRAY', items: { type: 'STRING' } },
    reportedDone: { type: 'ARRAY', items: { type: 'STRING' } }
  },
  required: ['summary', 'decisions', 'tasks', 'openQuestions', 'facts']
};

function fmtTs(sec) {
  sec = Math.round(sec || 0);
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  const p = (n) => String(n).padStart(2, '0');
  return h ? `${h}:${p(m)}:${p(s)}` : `${p(m)}:${p(s)}`;
}

export function transcriptAsText(tr, nameOf = (x) => x) {
  if (tr.segments && tr.segments.length) {
    return tr.segments.map(s => `[${fmtTs(s.start)}] ${s.speaker ? nameOf(s.speaker) + ': ' : ''}${s.text}`).join('\n');
  }
  return tr.text || '';
}

function analysisPrompt(meta, transcriptText, ctx) {
  const people = (ctx.participants || []).map(p => `${p.name}${p.role ? ' (' + p.role + ')' : ''}`).join(', ') || 'не вказано';
  const terms = (ctx.glossary || []).join(', ');
  const carry = (ctx.carryTasks || []).map((t, i) => `[k${i + 1}] ${t.task} — ${t.owner || '—'}${t.meetingStart ? ' (з наради ' + t.meetingStart.slice(0, 10) + ')' : ''}`).join('\n');
  return `Ти — помічник керівника виробництва. Нижче автоматична розшифровка робочої наради (можуть бути помилки розпізнавання).
Склади короткий структурований підсумок українською мовою.

Нарада: «${meta.title}», ${meta.dateText}. Дата наради: ${meta.dateISO} (${meta.weekday}).
Учасники: ${people}.
${terms ? 'Терміни й назви, які можуть зустрічатися: ' + terms + '.\n' : ''}${carry ? 'Відкриті задачі з попередніх нарад:\n' + carry + '\n' : ''}
Поверни ТІЛЬКИ JSON такого вигляду:
{"summary": "3–6 речень: про що нарада і до чого дійшли",
 "decisions": ["прийняте рішення"],
 "tasks": [{"owner": "хто", "task": "що зробити", "due": "термін словами або порожній рядок", "dueDate": "РРРР-ММ-ДД або порожній рядок"}],
 "openQuestions": ["питання, що лишилося відкритим"],
 "facts": ["важливі цифри, дати, обсяги, показники"],
 "reportedDone": []}

Правила:
- Тільки те, що є в розмові. Нічого не вигадуй.
- Задача — це конкретна дія, яку хтось взяв на себе або доручили. Власника пиши іменем учасника; якщо незрозуміло — «—».
- Якщо термін не названо — due і dueDate порожні. Якщо термін можна перевести в конкретну дату відносно дати наради («до п'ятниці», «завтра», «до 15-го»), запиши її в dueDate у форматі РРРР-ММ-ДД.
- reportedDone — мітки (k1, k2…) відкритих задач з попередніх нарад, про які на цій нараді прямо сказали, що їх виконано. Якщо таких немає або список задач не дано — [].
- Не повторюй у tasks задачі з попередніх нарад, якщо по них не дали нового доручення чи нового терміну.
- Очевидні помилки розпізнавання в іменах і термінах виправляй.
- Пиши коротко й по суті. Порожні списки — [].

Розшифровка:
${transcriptText}`;
}

export async function analyze(tr, meta, ctx) {
  const engine = ctx.settings.analysisEngine || 'gemini';
  const an = ANALYZERS[engine];
  const key = ctx.keys[an.keyName];
  if (!key) throw new Error(`${an.name}: не додано ключ API для аналізу`);
  const model = (ctx.settings.analysisModels || {})[engine];
  const prompt = analysisPrompt(meta, transcriptAsText(tr, ctx.nameOf), ctx);
  let text;
  if (engine === 'gemini') {
    const body = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, responseMimeType: 'application/json', responseSchema: ANALYSIS_SCHEMA, maxOutputTokens: 16384 }
    };
    const r = await post(ctx, 'Gemini', `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`,
      { headers: { 'x-goog-api-key': key }, json: body });
    if (r.status !== 200) throw apiError('Gemini', r);
    const d = JSON.parse(r.text);
    text = (((d.candidates || [])[0] || {}).content?.parts || []).map(x => x.text || '').join('');
  } else if (engine === 'openai') {
    const r = await post(ctx, 'OpenAI', 'https://api.openai.com/v1/chat/completions', {
      headers: { Authorization: 'Bearer ' + key },
      json: { model, messages: [{ role: 'user', content: prompt }], response_format: { type: 'json_object' } }
    });
    if (r.status !== 200) throw apiError('OpenAI', r);
    text = JSON.parse(r.text).choices?.[0]?.message?.content || '';
  } else if (engine === 'anthropic') {
    const r = await post(ctx, 'Claude', 'https://api.anthropic.com/v1/messages', {
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      json: { model, max_tokens: 8000, messages: [{ role: 'user', content: prompt }] }
    });
    if (r.status !== 200) throw apiError('Claude', r);
    text = (JSON.parse(r.text).content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  }
  const j = parseJsonLoose(text);
  const arr = (x) => Array.isArray(x) ? x.map(v => typeof v === 'string' ? v : JSON.stringify(v)).filter(Boolean) : [];
  return {
    summary: String(j.summary || ''),
    decisions: arr(j.decisions),
    tasks: Array.isArray(j.tasks) ? j.tasks.map(t => ({
      owner: String(t.owner || '—'), task: String(t.task || ''), due: String(t.due || ''),
      dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(t.dueDate || '')) ? String(t.dueDate) : ''
    })).filter(t => t.task) : [],
    openQuestions: arr(j.openQuestions),
    facts: arr(j.facts),
    reportedDone: arr(j.reportedDone).map(k => { const m = /k?(\d+)/i.exec(k); const t = m && (ctx.carryTasks || [])[Number(m[1]) - 1]; return t ? t.id : null; }).filter(Boolean),
    engine, engineName: an.name, model,
    basedOn: tr.engine,
    createdAt: new Date().toISOString()
  };
}
