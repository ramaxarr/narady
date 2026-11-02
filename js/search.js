// Пошук по всіх розшифровках. Тексти зберігаються в пам'яті пристрою і дозавантажуються
// лише змінені наради, тож пошук миттєвий навіть без інтернету.

import { h } from './ui.js';
import { call } from './api.js';
import { idb } from './store.js';

let syncing = null;

export function syncIndex(onProgress) {
  if (syncing) return syncing;
  syncing = (async () => {
    let cursor = (await idb.get('kv', 'searchCursor')) || '';
    let ids = null, total = 0;
    for (let guard = 0; guard < 60; guard++) {
      const r = await call('searchCorpus', { since: cursor });
      if (!ids) ids = new Set(r.ids);
      for (const it of r.items) await idb.put('search', it);
      total += r.items.length;
      cursor = r.cursor || cursor;
      await idb.put('kv', cursor, 'searchCursor');
      if (onProgress) onProgress(total);
      if (!r.more) break;
    }
    const all = (await idb.all('search')) || [];
    for (const it of all) if (ids && !ids.has(it.id)) await idb.del('search', it.id);
    return total;
  })().finally(() => { syncing = null; });
  return syncing;
}

export const norm = (s) => String(s || '').toLowerCase().replace(/[’ʼ`']/g, "'").replace(/ё/g, 'е');
export const queryWords = (q) => norm(q).split(/\s+/).filter(w => w.length >= 2);

function speakerName(label, map, people) {
  if (!label) return '';
  const v = map && map[label];
  if (!v) return label;
  const p = people.find(x => x.id === v);
  return p ? p.name : v;
}

export async function searchAll(query, people) {
  const words = queryWords(query);
  if (!words.length) return [];
  const items = (await idb.all('search')) || [];
  const has = (s) => { const n = norm(s); return words.every(w => n.includes(w)); };
  const out = [];
  for (const it of items) {
    const hits = [];
    for (const [t, spk, text] of it.segs || []) {
      const name = speakerName(spk, it.speakerMap, people);
      if (has(name + ' ' + text)) hits.push({ t, speaker: name, text });
    }
    const titleHit = has(it.title);
    const sumHit = !!it.summary && has(it.summary);
    if (hits.length || titleHit || sumHit) out.push({ id: it.id, title: it.title, start: it.start, summary: it.summary, hits, titleHit, sumHit });
  }
  out.sort((a, b) => (b.hits.length - a.hits.length) || String(b.start).localeCompare(String(a.start)));
  return out;
}

/** Уривок навколо першого збігу. */
export function snippet(text, words, len = 220) {
  const n = norm(text);
  let idx = -1;
  for (const w of words) { const i = n.indexOf(w); if (i >= 0 && (idx < 0 || i < idx)) idx = i; }
  if (text.length <= len || idx < 0) return text.length <= len ? text : text.slice(0, len) + '…';
  const start = Math.max(0, idx - 70);
  return (start > 0 ? '…' : '') + text.slice(start, start + len) + (start + len < text.length ? '…' : '');
}

/** Текст із підсвіченими словами пошуку. */
export function highlight(text, words) {
  if (!words || !words.length) return [text];
  const n = norm(text);
  const marks = [];
  for (const w of words) {
    let i = n.indexOf(w);
    while (i >= 0) { marks.push([i, i + w.length]); i = n.indexOf(w, i + w.length); }
  }
  if (!marks.length) return [text];
  marks.sort((a, b) => a[0] - b[0]);
  const merged = [];
  for (const m of marks) { const last = merged[merged.length - 1]; if (last && m[0] <= last[1]) last[1] = Math.max(last[1], m[1]); else merged.push([...m]); }
  const out = [];
  let pos = 0;
  for (const [a, b] of merged) {
    if (a > pos) out.push(text.slice(pos, a));
    out.push(h('mark', null, text.slice(a, b)));
    pos = b;
  }
  if (pos < text.length) out.push(text.slice(pos));
  return out;
}
