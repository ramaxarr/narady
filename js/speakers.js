// Автоматичне підписування спікерів: мітки одного двигуна (speaker_0, 1·A…) отримують імена
// з тексту іншого двигуна, який уже знає імена (зазвичай Gemini). Порівнюються слова реплік
// поблизу за часом: та сама фраза в обох текстах → та сама людина.
// Блок між позначками <shared> скопійовано в backend/Code.gs без змін — правити в обох місцях однаково.

// <shared>
const SPK_STOP = { 'так': 1, 'вже': 1, 'але': 1, 'щоб': 1, 'якщо': 1, 'тому': 1, 'треба': 1, 'буде': 1, 'можна': 1, 'його': 1, 'вони': 1, 'воно': 1,
  'цього': 1, 'того': 1, 'там': 1, 'тут': 1, 'теж': 1, 'ось': 1, 'отже': 1, 'значить': 1, 'окей': 1, 'добре': 1, 'давайте': 1, 'тобто': 1, 'який': 1,
  'яка': 1, 'яке': 1, 'які': 1, 'для': 1, 'від': 1, 'при': 1, 'без': 1, 'про': 1, 'або': 1, 'ще': 1, 'нас': 1, 'вас': 1, 'нам': 1, 'вам': 1, 'мене': 1,
  'тебе': 1, 'зараз': 1, 'потім': 1, 'просто': 1, 'типу': 1, 'кажу': 1, 'ага': 1, 'угу': 1, 'еее': 1, 'ммм': 1 };

function spkIsRealName(n) {
  n = String(n || '').trim();
  if (!n || n.length < 2) return false;
  if (/^(спікер|спикер|speaker|мовець|учасник|невідом|unknown|голос|чоловік|жінка|ведуч)/i.test(n)) return false;
  if (/^\d+·/.test(n) || /^speaker_\d+$/i.test(n) || /^[A-Z]$/.test(n) || /^\d+$/.test(n)) return false;
  return true;
}

function spkStems(text) {
  const out = {};
  let n = 0;
  String(text || '').toLowerCase().replace(/[ʼ’'`]/g, '').split(/[^a-zа-яіїєґё0-9]+/i).forEach(function (w) {
    if (w.length < 3 || SPK_STOP[w]) return;
    const s = w.slice(0, 5);
    if (!out[s]) { out[s] = 1; n++; }
  });
  return { set: out, size: n };
}

/** Мітки цільового тексту → імена з тексту-зразка. Повертає { map: {мітка: імʼя}, words: {мітка: слів} }. */
function spkMatch(target, reference, refNameOf) {
  const refs = (reference.segments || []).map(function (s, i, a) {
    const next = a[i + 1];
    let end = s.end > s.start ? s.end : (next ? next.start : s.start + 15);
    if (end - s.start > 90) end = s.start + 90;
    return { start: s.start, end: end, name: refNameOf(s.speaker), st: spkStems(s.text) };
  });
  const WIN = 60;
  const votes = {}, words = {};
  (target.segments || []).forEach(function (t) {
    if (!t.speaker) return;
    const ts = spkStems(t.text);
    words[t.speaker] = (words[t.speaker] || 0) + ts.size;
    if (ts.size < 3) return;
    const tEnd = t.end > t.start ? t.end : t.start + 5;
    const mid = (t.start + tEnd) / 2;
    // Найсхожа репліка зразка поблизу; голос віддаємо, лише якщо вона помітно краща за схожу репліку іншої людини.
    let best = null, second = 0;
    refs.forEach(function (r) {
      if (r.start > tEnd + WIN || r.end < t.start - WIN || !r.st.size) return;
      let common = 0;
      for (const k in ts.set) if (r.st.set[k]) common++;
      if (common < 2) return;
      const score = common / Math.min(ts.size, r.st.size) - Math.abs((r.start + r.end) / 2 - mid) / 600;
      if (!best || score > best.score) { if (best && best.name !== r.name) second = Math.max(second, best.score); best = { score: score, name: r.name, common: common }; }
      else if (r.name !== best.name) second = Math.max(second, score);
    });
    if (!best || best.score < 0.35 || !spkIsRealName(best.name) || second > best.score * 0.8) return;
    const lv = votes[t.speaker] || (votes[t.speaker] = { _all: 0 });
    lv[best.name] = (lv[best.name] || 0) + ts.size;
    lv._all += ts.size;
  });
  const map = {};
  for (const label in votes) {
    const lv = votes[label];
    let best = null;
    for (const k in lv) if (k !== '_all' && (!best || lv[k] > lv[best])) best = k;
    if (best && lv[best] >= 12 && lv[best] / lv._all >= 0.55 && lv[best] / Math.max(1, words[label]) >= 0.25) map[label] = best;
  }
  return { map: map, words: words };
}

/** Імʼя з тексту → id людини з довідника, якщо вдається однозначно; інакше саме імʼя. */
function spkResolvePerson(name, people, participantIds) {
  const n = String(name || '').trim().toLowerCase();
  if (!n) return name;
  const list = people || [];
  let hit = list.filter(function (p) { return String(p.name || '').trim().toLowerCase() === n; });
  if (!hit.length) hit = list.filter(function (p) { return String(p.aliases || '').toLowerCase().split(/\s*,\s*/).indexOf(n) >= 0; });
  if (!hit.length) {
    // «Богдан» → «Богдан Смеречинський», якщо серед учасників такий один
    const pool = list.filter(function (p) { return (participantIds || []).indexOf(p.id) >= 0; });
    const first = n.split(/\s+/)[0];
    hit = pool.filter(function (p) { const pn = String(p.name || '').toLowerCase().split(/\s+/); return pn[0] === first && (n.split(/\s+/).length === 1 || pn.join(' ').indexOf(n) === 0); });
  }
  return hit.length === 1 ? hit[0].id : name;
}

/**
 * Для всіх текстів зі спікерами підписує ще не підписані мітки іменами з найкращого тексту-зразка.
 * transcripts: {двигун: текст}; speakerMap: {двигун: {мітка: id людини або імʼя}}; people: [{id, name, aliases}].
 * Повертає { speakerMap (новий), filled: {двигун: скільки міток підписано}, coverage: {двигун: частка слів з іменами}, reference }.
 */
function spkAutoMapAll(transcripts, speakerMap, people, participantIds) {
  const byId = {};
  (people || []).forEach(function (p) { byId[p.id] = p; });
  const sm = {};
  for (const e in (speakerMap || {})) { sm[e] = {}; for (const l in speakerMap[e]) sm[e][l] = speakerMap[e][l]; }
  const nameOf = function (engine) {
    const m = sm[engine] || {};
    return function (label) { const v = m[label]; if (!v) return label; return byId[v] ? byId[v].name : v; };
  };
  const named = function (engine) {
    const tr = transcripts[engine], f = nameOf(engine);
    let all = 0, real = 0;
    (tr.segments || []).forEach(function (s) { const w = String(s.text || '').length; all += w; if (s.speaker && spkIsRealName(f(s.speaker))) real += w; });
    return all ? real / all : 0;
  };
  const ids = Object.keys(transcripts || {}).filter(function (e) { const t = transcripts[e]; return t && t.segments && t.segments.some(function (s) { return s.speaker; }); });
  let reference = null, refScore = 0;
  ids.forEach(function (e) { const sc = named(e) + (e === 'gemini' ? 0.05 : 0); if (sc > refScore) { refScore = sc; reference = e; } });
  const filled = {}, coverage = {};
  if (reference && refScore >= 0.5) {
    const refNameOf = nameOf(reference);
    ids.forEach(function (e) {
      if (e === reference) return;
      const r = spkMatch(transcripts[e], transcripts[reference], refNameOf);
      const m = sm[e] || {};
      let n = 0;
      for (const label in r.map) if (!m[label]) { m[label] = spkResolvePerson(r.map[label], people, participantIds); n++; }
      if (n) sm[e] = m;
      filled[e] = n;
    });
  }
  ids.forEach(function (e) { coverage[e] = Math.round(named(e) * 100) / 100; });
  return { speakerMap: sm, filled: filled, coverage: coverage, reference: refScore >= 0.5 ? reference : null };
}
// </shared>

export { spkAutoMapAll, spkIsRealName, spkMatch, spkResolvePerson };
