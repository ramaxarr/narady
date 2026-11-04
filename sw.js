// Service worker: офлайн-кеш оболонки додатку і прийом файлів через «Поділитися» на Android.

const VERSION = 'narady-1.3.0';
const ASSETS = [
  './', './index.html', './app.css', './manifest.webmanifest',
  './js/app.js', './js/ui.js', './js/api.js', './js/store.js', './js/audio.js', './js/mp3worker.js',
  './js/recorder.js', './js/engines.js', './js/pipeline.js', './js/drive.js', './js/tasks.js', './js/search.js', './js/speakers.js',
  './vendor/lame.min.js', './icons/icon-192.png', './icons/icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('narady', 2);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('recChunks')) d.createObjectStore('recChunks', { autoIncrement: true }).createIndex('recId', 'recId');
      if (!d.objectStoreNames.contains('recMeta')) d.createObjectStore('recMeta', { keyPath: 'recId' });
      if (!d.objectStoreNames.contains('prepared')) d.createObjectStore('prepared');
      if (!d.objectStoreNames.contains('shared')) d.createObjectStore('shared', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('search')) d.createObjectStore('search', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('kv')) d.createObjectStore('kv');
    };
    req.onsuccess = () => { const d = req.result; d.onversionchange = () => d.close(); resolve(d); };
    req.onerror = () => reject(req.error);
    req.onblocked = () => setTimeout(() => reject(new Error('blocked')), 2500);
  });
}

async function handleShare(request) {
  try {
    const form = await request.formData();
    const files = form.getAll('audio').filter(f => f && f.size);
    if (files.length) {
      const db = await openDb();
      await new Promise((resolve, reject) => {
        const tx = db.transaction('shared', 'readwrite');
        const st = tx.objectStore('shared');
        files.forEach((f, i) => st.put({ id: 's_' + Date.now() + '_' + i, file: f, name: f.name, type: f.type, lastModified: f.lastModified, size: f.size }));
        tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
      });
      db.close();
    }
  } catch (e) { /* відкриємо додаток у будь-якому разі */ }
  return Response.redirect(new URL('./#/import', self.registration.scope).href, 303);
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method === 'POST' && url.origin === location.origin && url.pathname.endsWith('/share-target')) {
    e.respondWith(handleShare(e.request));
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  // Спершу мережа (щоб оновлення приходили одразу), без мережі — кеш.
  e.respondWith(fetch(e.request).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(e.request, copy)); }
    return res;
  }).catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('./index.html'))));
});
