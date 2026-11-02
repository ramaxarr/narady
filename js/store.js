// Локальне сховище в браузері (IndexedDB): шматки запису, підготовлене аудіо, файли «Поділитися».

const DB_NAME = 'narady';
const DB_VER = 2;
let dbp = null;

function db() {
  if (dbp) return dbp;
  dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VER);
    req.onupgradeneeded = () => {
      const d = req.result;
      if (!d.objectStoreNames.contains('recChunks')) {
        const s = d.createObjectStore('recChunks', { autoIncrement: true });
        s.createIndex('recId', 'recId');
      }
      if (!d.objectStoreNames.contains('recMeta')) d.createObjectStore('recMeta', { keyPath: 'recId' });
      if (!d.objectStoreNames.contains('prepared')) d.createObjectStore('prepared');
      if (!d.objectStoreNames.contains('shared')) d.createObjectStore('shared', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('search')) d.createObjectStore('search', { keyPath: 'id' });
      if (!d.objectStoreNames.contains('kv')) d.createObjectStore('kv');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbp;
}

function tx(store, mode, fn) {
  return db().then(d => new Promise((resolve, reject) => {
    const t = d.transaction(store, mode);
    const s = t.objectStore(store);
    let result;
    Promise.resolve(fn(s)).then(r => { result = r; });
    t.oncomplete = () => resolve(result && result.__req ? result.__req.result : result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  }));
}

const wrap = (req) => ({ __req: req });

export const idb = {
  put: (store, value, key) => tx(store, 'readwrite', s => wrap(key === undefined ? s.put(value) : s.put(value, key))),
  get: (store, key) => tx(store, 'readonly', s => wrap(s.get(key))),
  del: (store, key) => tx(store, 'readwrite', s => wrap(s.delete(key))),
  all: (store) => tx(store, 'readonly', s => wrap(s.getAll())),
  byIndex: (store, index, key) => tx(store, 'readonly', s => wrap(s.index(index).getAll(key))),
  delByIndex: (store, index, key) => tx(store, 'readwrite', s => new Promise((res) => {
    const r = s.index(index).openKeyCursor(IDBKeyRange.only(key));
    r.onsuccess = () => { const c = r.result; if (c) { s.delete(c.primaryKey); c.continue(); } else res(true); };
  }))
};

// Підготовлене аудіо зберігаємо для останніх нарад, щоб не качати з Диска повторно.
export async function savePrepared(meetingId, prepared) {
  try { await idb.put('prepared', { ...prepared, savedAt: Date.now() }, meetingId); } catch (e) { console.warn('prepared cache', e); }
}
export async function getPrepared(meetingId) {
  try { return await idb.get('prepared', meetingId); } catch { return null; }
}
export async function trimPrepared(keep = 6) {
  try {
    const d = await db();
    const keys = await new Promise((res) => { const r = d.transaction('prepared').objectStore('prepared').getAllKeys(); r.onsuccess = () => res(r.result); });
    if (keys.length <= keep) return;
    const items = await Promise.all(keys.map(async k => ({ k, t: (await idb.get('prepared', k))?.savedAt || 0 })));
    items.sort((a, b) => b.t - a.t);
    for (const it of items.slice(keep)) await idb.del('prepared', it.k);
  } catch { }
}
