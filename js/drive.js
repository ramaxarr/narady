// Завантаження аудіо на Google Диск і назад — напряму з браузера,
// з короткочасним токеном, який видає бекенд.

import { call } from './api.js';

export async function driveToken(meetingId) {
  return call('driveToken', { meetingId });
}

export async function uploadBlob(blob, name, meetingId, onProgress) {
  const { token, folderId } = await driveToken(meetingId);
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

export async function downloadFile(fileId, onProgress) {
  const { token } = await driveToken();
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
