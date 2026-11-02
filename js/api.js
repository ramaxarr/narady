// Зв'язок із бекендом на Google Apps Script і спільні налаштування.

const LS_CONN = 'narady.conn';

export const DEFAULT_SETTINGS = {
  language: 'uk',
  defaultEngines: ['gemini'],
  autoAnalyze: true,
  autoDoc: true,
  analysisEngine: 'gemini',
  models: {
    gemini: 'gemini-2.5-flash',
    groq: 'whisper-large-v3',
    openai: 'gpt-4o-transcribe-diarize',
    elevenlabs: 'scribe_v2'
  },
  analysisModels: {
    gemini: 'gemini-2.5-flash',
    openai: 'gpt-5-mini',
    anthropic: 'claude-sonnet-5-5'
  },
  noiseSuppression: false,
  appUrl: ''
};

export class ApiError extends Error {
  constructor(code, message) { super(message || code); this.code = code; }
}

export function getConn() {
  try { return JSON.parse(localStorage.getItem(LS_CONN)) || null; } catch { return null; }
}
export function setConn(c) {
  if (c) localStorage.setItem(LS_CONN, JSON.stringify(c)); else localStorage.removeItem(LS_CONN);
  config = null;
}

const ERRORS = {
  unauthorized: 'Невірний секретний ключ. Перевір його в Налаштуваннях.',
  not_setup: 'Бекенд не налаштовано: запусти функцію setup() у скрипті.',
  bad_json: 'Бекенд не зрозумів запит.'
};

export async function call(action, args = {}, conn = getConn()) {
  if (!conn || !conn.url) throw new ApiError('no_conn', 'Додаток не підключено до бекенду');
  let res;
  try {
    res = await fetch(conn.url, { method: 'POST', body: JSON.stringify({ token: conn.token, action, args }), redirect: 'follow' });
  } catch (e) {
    throw new ApiError('network', navigator.onLine ? 'Не вдалося звʼязатися з бекендом. Перевір адресу скрипта.' : 'Немає інтернету');
  }
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch {
    if (/<html|<!DOCTYPE/i.test(text)) throw new ApiError('bad_url', 'Бекенд повернув сторінку замість даних. Перевір, що скрипт розгорнуто з доступом «Anyone».');
    throw new ApiError('bad_response', 'Незрозуміла відповідь бекенду');
  }
  if (!data.ok) throw new ApiError(data.error, ERRORS[data.error] || data.error);
  // Лише getMeeting і listMeetings справді звіряють нараду з календарем;
  // у відповідях на збереження прапорець «подію видалено» хибний, тож прибираємо його.
  const d = data.data;
  if (d && typeof d === 'object' && !Array.isArray(d) && 'eventMissing' in d && action !== 'getMeeting') d.eventMissing = false;
  return d;
}

let config = null;

function merge(base, over) {
  const out = Array.isArray(base) ? [...base] : { ...base };
  for (const [k, v] of Object.entries(over || {})) {
    if (v && typeof v === 'object' && !Array.isArray(v) && base[k] && typeof base[k] === 'object') out[k] = merge(base[k], v);
    else if (v !== undefined) out[k] = v;
  }
  return out;
}

export async function loadConfig(force = false) {
  if (config && !force) return config;
  const c = await call('getConfig');
  config = { keys: c.keys || {}, settings: merge(DEFAULT_SETTINGS, c.settings || {}), info: c.info };
  const here = location.origin + location.pathname;
  if (config.settings.appUrl !== here) {
    config.settings.appUrl = here;
    call('saveConfig', { settings: config.settings }).catch(() => { });
  }
  return config;
}

export function cfg() { return config; }

export async function saveConfig(part) {
  const c = await call('saveConfig', part);
  config = { keys: c.keys || {}, settings: merge(DEFAULT_SETTINGS, c.settings || {}), info: c.info };
  return config;
}

// Люди і словник — кеш у пам'яті.
let people = null, glossary = null;
export async function getPeople(force = false) {
  if (!people || force) people = await call('listPeople');
  return people;
}
export function peopleSync() { return people || []; }
export async function getGlossary(force = false) {
  if (!glossary || force) glossary = await call('getGlossary');
  return glossary;
}
export function setGlossaryCache(g) { glossary = g; }
export function setPeopleCache(p) { people = p; }
