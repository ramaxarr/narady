// Завантаження аудіо на Google Диск і назад — напряму з браузера,
// з короткочасним токеном, який видає бекенд.

import { call } from './api.js';

export async function driveToken(meetingId) {
  return call('driveToken', { meetingId });
}

export async function uploadBlob(blob, name, meetingId, onProgress, auth) {
  const { token, folderId } = auth || await driveToken(meetingId);
  const type = blob.type || 'audio/mpeg';
  const init = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name,size', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json; charset=UTF-8',
      'X-Upload-Content-Type': type,
      'X-Upload-Content-Length': String(blob.size)
    },
    body: JSON.stringify({ name, parents: [folderId], mimeType: type })
  });
  if (!init.ok) throw new Error('Google Диск не прийняв файл (' + init.status + '): ' + (await init.text()).slice(0, 200));
  const url = init.headers.get('Location');
  if (!url) throw new Error('Google Диск не дав адресу для завантаження');

  const file = await new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('PUT', url);
    xhr.setRequestHeader('Content-Type', type);
    // Не чекаємо вічно, якщо мережа «застигла»: щонайменше 3 хв, для великих файлів — з розрахунку 30 КБ/с.
    xhr.timeout = Math.max(180000, Math.round(blob.size / 30));
    xhr.ontimeout = () => reject(new Error('Завантаження на Google Диск зависло. Перевір інтернет і спробуй ще раз — запис збережено на пристрої.'));
    xhr.upload.onprogress = (e) => { if (e.lengthComputable && onProgress) onProgress(e.loaded / e.total); };
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) { try { resolve(JSON.parse(xhr.responseText)); } catch { reject(new Error('Незрозуміла відповідь Диска')); } }
      else reject(new Error('Помилка завантаження на Диск (' + xhr.status + ')'));
    };
    xhr.onerror = () => reject(new Error('Мережа обірвалась під час завантаження на Диск'));
    xhr.send(blob);
  });
  return file;
}

/** Кілька файлів (частини запису) з одним токеном; спроба повторюється один раз. */
export async function uploadPieces(pieces, baseName, meetingId, onProgress) {
  const auth = await driveToken(meetingId);
  const total = pieces.reduce((a, p) => a + p.blob.size, 0) || 1;
  let done = 0;
  const out = [];
  for (let i = 0; i < pieces.length; i++) {
    const p = pieces[i];
    const name = pieces.length > 1 ? `${baseName} — ${String(i + 1).padStart(2, '0')}.mp3` : `${baseName}.mp3`;
    const prog = (f) => onProgress && onProgress((done + f * p.blob.size) / total, i, pieces.length);
    let file;
    try { file = await uploadBlob(p.blob, name, meetingId, prog, auth); }
    catch (e) { await new Promise(r => setTimeout(r, 3000)); file = await uploadBlob(p.blob, name, meetingId, prog, auth); }
    done += p.blob.size;
    out.push({ id: file.id, o: Math.round(p.offset * 100) / 100, d: Math.round(p.duration * 100) / 100 });
  }
  return out;
}

export async function downloadFile(fileId, onProgress, auth) {
  const { token } = auth || await driveToken();
  const res = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
    headers: { Authorization: 'Bearer ' + token }
  });
  if (!res.ok) throw new Error('Не вдалося отримати аудіо з Диска (' + res.status + ')');
  const total = Number(res.headers.get('Content-Length')) || 0;
  if (!res.body || !total || !onProgress) return await res.blob();
  const reader = res.body.getReader();
  const parts = []; let got = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    parts.push(value); got += value.length; onProgress(got / total);
  }
  return new Blob(parts, { type: res.headers.get('Content-Type') || 'audio/mpeg' });
}
