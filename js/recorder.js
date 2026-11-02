// Запис з мікрофона. Кожні 5 секунд шматок запису зберігається в пам'ять браузера,
// тож навіть якщо вкладка закриється, записане не пропаде.

import { idb } from './store.js';

function pickMime() {
  const opts = ['audio/webm;codecs=opus', 'audio/ogg;codecs=opus', 'audio/mp4;codecs=mp4a.40.2', 'audio/mp4', 'audio/webm'];
  for (const m of opts) { if (window.MediaRecorder && MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(m)) return m; }
  return '';
}

export class Recorder {
  constructor({ meetingId, onLevel, onState, onInterrupted, onWarn }) {
    this.meetingId = meetingId;
    this.onLevel = onLevel || (() => { });
    this.onState = onState || (() => { });
    this.onInterrupted = onInterrupted || null;
    this.onWarn = onWarn || null;
    this.recId = 'r_' + Date.now();
    this.seq = 0;
    this.chunks = [];
    this.pending = [];     // шматки, які ще не вдалося зберегти в пам'ять пристрою
    this.saving = false;
    this.warned = false;
    this.elapsedBefore = 0;
    this.startedAt = 0;
    this.state = 'idle';
    this.wakeLock = null;
    this._onVis = () => { if (document.visibilityState === 'visible' && this.state === 'recording') this._lock(); };
  }

  get elapsed() {
    return this.elapsedBefore + (this.state === 'recording' ? (Date.now() - this.startedAt) / 1000 : 0);
  }

  async start({ noiseSuppression = false } = {}) {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('Цей браузер не дає доступ до мікрофона');
    this.stream = await navigator.mediaDevices.getUserMedia({
      audio: { echoCancellation: false, noiseSuppression, autoGainControl: true, channelCount: 1 }
    });
    this.mime = pickMime();
    try { this.mr = new MediaRecorder(this.stream, this.mime ? { mimeType: this.mime, audioBitsPerSecond: 64000 } : undefined); }
    catch { this.mr = new MediaRecorder(this.stream); }
    this.mime = this.mr.mimeType || this.mime || 'audio/webm';
    // Якщо систему чи інший додаток забрав мікрофон — завершуємо запис, щоб записане пішло в обробку.
    const interrupted = () => { if (this.state !== 'stopped' && this.onInterrupted) this.onInterrupted(); };
    this.stream.getAudioTracks().forEach(t => t.addEventListener('ended', interrupted));
    this.mr.onerror = interrupted;
    this.mr.ondataavailable = (e) => {
      if (!e.data || !e.data.size) return;
      this.chunks.push(e.data);
      this.pending.push({ seq: this.seq++, blob: e.data });
      this._flush();
    };
    // Запис стартує одразу; резервна копія на пристрій пишеться у фоні й ніколи не затримує запис.
    this.mr.start(5000);
    this.firstStart = Date.now();
    this._flush();
    this.startedAt = Date.now();
    this.state = 'recording';
    this._meter();
    this._lock();
    document.addEventListener('visibilitychange', this._onVis);
    this.onState(this.state);
  }

  async _flush() {
    if (this.saving) return;
    this.saving = true;
    try {
      await idb.put('recMeta', { recId: this.recId, meetingId: this.meetingId, mime: this.mime, startedAt: this.firstStart || Date.now(), elapsed: this.elapsed });
      while (this.pending.length) {
        const c = this.pending[0];
        await idb.put('recChunks', { recId: this.recId, seq: c.seq, blob: c.blob });
        this.pending.shift();
      }
      if (this.warned && this.onWarn) this.onWarn('Резервна копія запису знову зберігається на пристрій.', 'ok');
      this.warned = false;
    } catch (err) {
      console.warn('Резервна копія запису не зберігається', err);
      if (!this.warned && this.onWarn) this.onWarn('Запис іде, але резервна копія на пристрій не зберігається. Закрий інші вікна й вкладки з «Нарадами».', 'err');
      this.warned = true;
    } finally {
      this.saving = false;
    }
  }

  async _lock() {
    try { if ('wakeLock' in navigator) this.wakeLock = await navigator.wakeLock.request('screen'); } catch { this.wakeLock = null; }
  }

  _meter() {
    try {
      const AC = window.AudioContext || window.webkitAudioContext;
      this.ac = new AC();
      if (this.ac.state === 'suspended' && this.ac.resume) this.ac.resume().catch(() => { });
      const src = this.ac.createMediaStreamSource(this.stream);
      const an = this.ac.createAnalyser();
      an.fftSize = 1024;
      src.connect(an);
      const data = new Float32Array(an.fftSize);
      const loop = () => {
        if (this.state === 'stopped') return;
        an.getFloatTimeDomainData(data);
        let s = 0; for (let i = 0; i < data.length; i++) s += data[i] * data[i];
        const rms = Math.sqrt(s / data.length);
        this.onLevel(this.state === 'recording' ? Math.min(1, rms * 6) : 0);
        this.raf = requestAnimationFrame(loop);
      };
      loop();
    } catch { }
  }

  pause() {
    if (this.state !== 'recording') return;
    this.mr.pause();
    this.elapsedBefore += (Date.now() - this.startedAt) / 1000;
    this.state = 'paused';
    this.onState(this.state);
  }

  resume() {
    if (this.state !== 'paused') return;
    this.mr.resume();
    this.startedAt = Date.now();
    this.state = 'recording';
    this.onState(this.state);
  }

  stop() {
    return new Promise((resolve) => {
      if (!this.mr || this.state === 'stopped') { resolve(null); return; }
      if (this.state === 'recording') this.elapsedBefore += (Date.now() - this.startedAt) / 1000;
      this.mr.addEventListener('stop', () => {
        setTimeout(() => resolve(new Blob(this.chunks, { type: this.mime })), 50);
      }, { once: true });
      this.state = 'stopped';
      try { this.mr.stop(); } catch { resolve(new Blob(this.chunks, { type: this.mime })); }
      this._release();
      this.onState(this.state);
    });
  }

  _release() {
    try { this.stream.getTracks().forEach(t => t.stop()); } catch { }
    try { cancelAnimationFrame(this.raf); this.ac && this.ac.close(); } catch { }
    try { this.wakeLock && this.wakeLock.release(); } catch { }
    document.removeEventListener('visibilitychange', this._onVis);
  }

  async discard() {
    if (this.state !== 'stopped') await this.stop();
    await deleteRecording(this.recId);
  }
}

export async function pendingRecordings() {
  try { return (await idb.all('recMeta')) || []; } catch { return []; }
}

export async function assembleRecording(recId) {
  const meta = await idb.get('recMeta', recId);
  const parts = (await idb.byIndex('recChunks', 'recId', recId)) || [];
  parts.sort((a, b) => a.seq - b.seq);
  return { meta, blob: new Blob(parts.map(p => p.blob), { type: (meta && meta.mime) || 'audio/webm' }) };
}

export async function deleteRecording(recId) {
  try { await idb.delByIndex('recChunks', 'recId', recId); } catch { }
  try { await idb.del('recMeta', recId); } catch { }
}
