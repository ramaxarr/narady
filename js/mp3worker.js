// Фоновий кодувальник MP3 (моно, 16 кГц). Використовує бібліотеку LAME (lamejs, LGPL).
/* global lamejs */
importScripts('../vendor/lame.min.js');

self.onmessage = (e) => {
  const { id, pcm, sampleRate, kbps, gain } = e.data;
  try {
    const enc = new lamejs.Mp3Encoder(1, sampleRate, kbps);
    const BLOCK = 1152 * 16;
    const out = [];
    const total = pcm.length;
    const buf = new Int16Array(BLOCK);
    let lastReport = 0;
    for (let i = 0; i < total; i += BLOCK) {
      const n = Math.min(BLOCK, total - i);
      const b = n === BLOCK ? buf : new Int16Array(n);
      for (let j = 0; j < n; j++) {
        let s = pcm[i + j] * gain;
        if (s > 1) s = 1; else if (s < -1) s = -1;
        b[j] = s < 0 ? s * 0x8000 : s * 0x7fff;
      }
      const mp3 = enc.encodeBuffer(b);
      if (mp3.length) out.push(new Uint8Array(mp3));
      if (i - lastReport > sampleRate * 30) { lastReport = i; self.postMessage({ id, progress: i / total }); }
    }
    const end = enc.flush();
    if (end.length) out.push(new Uint8Array(end));
    self.postMessage({ id, done: true, blob: new Blob(out, { type: 'audio/mpeg' }) });
  } catch (err) {
    self.postMessage({ id, error: String((err && err.message) || err) });
  }
};
