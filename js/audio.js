// Підготовка аудіо в браузері: декодування, моно 16 кГц, вирівнювання гучності,
// нарізка по паузах і стиснення в MP3 32 кбіт/с (≈14 МБ на годину).

export const SR = 16000;
export const KBPS = 32;
export const CHUNK_SEC = 1200; // фрагменти ~20 хв: вміщаються в ліміти всіх двигунів

const AC = () => window.AudioContext || window.webkitAudioContext;

export async function decodeMono16k(blob) {
  const ab = await blob.arrayBuffer();
  let buf = null;
  try {
    const ctx = new (AC())({ sampleRate: SR });
    try { buf = await ctx.decodeAudioData(ab.slice(0)); } finally { try { ctx.close(); } catch { } }
  } catch (e) {
    buf = null;
  }
  if (!buf || buf.sampleRate !== SR) {
    let native = buf;
    if (!native) {
      const ctx = new (AC())();
      try { native = await ctx.decodeAudioData(ab.slice(0)); }
      catch { throw new Error('Браузер не зміг прочитати цей файл. Підходять m4a, mp3, wav, ogg, webm.'); }
      finally { try { ctx.close(); } catch { } }
    }
    const off = new OfflineAudioContext(1, Math.max(1, Math.ceil(native.duration * SR)), SR);
    const src = off.createBufferSource();
    src.buffer = native; src.connect(off.destination); src.start();
    buf = await off.startRendering();
  }
  const ch = buf.numberOfChannels, n = buf.length;
  const mono = buf.getChannelData(0);
  if (ch > 1) {
    const others = [];
    for (let c = 1; c < ch; c++) others.push(buf.getChannelData(c));
    for (let i = 0; i < n; i++) {
      let s = mono[i];
      for (let k = 0; k < others.length; k++) s += others[k][i];
      mono[i] = s / ch;
    }
  }
  return { pcm: mono, sampleRate: SR, duration: n / SR };
}

export function computeGain(pcm) {
  // Тихі записи (телефон далеко від людей) підсилюємо, гучні не чіпаємо.
  const hist = new Uint32Array(1001);
  let cnt = 0;
  for (let i = 0; i < pcm.length; i += 8) { const a = Math.min(1, Math.abs(pcm[i])); hist[Math.round(a * 1000)]++; cnt++; }
  let acc = 0, p999 = 1;
  for (let k = 1000; k >= 0; k--) { acc += hist[k]; if (acc > cnt * 0.001) { p999 = k / 1000; break; } }
  if (p999 <= 0.001) return 1;
  return p999 < 0.6 ? Math.min(0.9 / p999, 8) : 1;
}

export function findCuts(pcm, sr, chunkSec = CHUNK_SEC, searchSec = 30) {
  const total = pcm.length;
  const cuts = [0];
  const win = Math.round(sr * 0.4);
  const step = Math.round(win / 2);
  let target = chunkSec * sr;
  while (target < total - sr * 60) {
    const prev = cuts[cuts.length - 1];
    const lo = Math.max(prev + sr * 60, target - searchSec * sr);
    const hi = Math.min(total - sr * 30, target + searchSec * sr);
    let best = target, bestE = Infinity;
    for (let s = lo; s + win < hi; s += step) {
      let e = 0;
      for (let j = s; j < s + win; j += 4) e += pcm[j] * pcm[j];
      if (e < bestE) { bestE = e; best = s + step; }
    }
    cuts.push(best);
    target = best + chunkSec * sr;
  }
  cuts.push(total);
  return cuts;
}

let worker = null, seq = 0;
const pending = new Map();

function getWorker() {
  if (worker) return worker;
  worker = new Worker(new URL('./mp3worker.js', import.meta.url));
  worker.onmessage = (e) => {
    const d = e.data, p = pending.get(d.id);
    if (!p) return;
    if (d.progress !== undefined) { if (p.onProgress) p.onProgress(d.progress); return; }
    pending.delete(d.id);
    if (d.error) p.reject(new Error('Помилка стиснення: ' + d.error));
    else p.resolve(d.blob);
  };
  worker.onerror = (e) => {
    for (const p of pending.values()) p.reject(new Error('Кодувальник MP3 не запустився: ' + (e.message || '')));
    pending.clear();
    worker = null;
  };
  return worker;
}

export function encodeMp3(pcmSlice, sr, gain, onProgress) {
  return new Promise((resolve, reject) => {
    const id = ++seq;
    pending.set(id, { resolve, reject, onProgress });
    const copy = pcmSlice.slice();
    getWorker().postMessage({ id, pcm: copy, sampleRate: sr, kbps: KBPS, gain }, [copy.buffer]);
  });
}

/**
 * Готує запис до розпізнавання.
 * Повертає { full: Blob (mp3 усього запису), chunks: [{blob, offset, duration}], duration }.
 * Якщо передано вже стиснений архів (з Диска), він лишається як є, а нарізка робиться лише за потреби.
 */
export async function prepareAudio(blob, onProgress = () => { }, { isArchive = false } = {}) {
  onProgress('decode', 0);
  const { pcm, sampleRate, duration } = await decodeMono16k(blob);
  if (duration < 1) throw new Error('Запис порожній або надто короткий');
  onProgress('decode', 1);

  if (isArchive && duration <= CHUNK_SEC + 90) {
    return { full: blob, chunks: [{ blob, offset: 0, duration }], duration };
  }
  const gain = isArchive ? 1 : computeGain(pcm);
  const cuts = findCuts(pcm, sampleRate);
  const chunks = [];
  for (let i = 0; i < cuts.length - 1; i++) {
    const a = cuts[i], b = cuts[i + 1];
    const mp3 = await encodeMp3(pcm.subarray(a, b), sampleRate, gain,
      (f) => onProgress('encode', (a + f * (b - a)) / pcm.length));
    chunks.push({ blob: mp3, offset: a / sampleRate, duration: (b - a) / sampleRate });
  }
  onProgress('encode', 1);
  const full = isArchive ? blob : new Blob(chunks.map(c => c.blob), { type: 'audio/mpeg' });
  return { full, chunks, duration };
}

export function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => { const s = String(r.result); resolve(s.slice(s.indexOf(',') + 1)); };
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}
