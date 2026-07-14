// IndexedDB-backed queue for field-feed items (photos/videos/notes) that
// couldn't be uploaded immediately — e.g. no signal in the field. Items sit
// here until connectivity returns, then get flushed in order.

const DB_NAME = 'nir-project-outbox';
const STORE = 'pending';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: 'id' });
        store.createIndex('projectId', 'projectId', { unique: false });
      }
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

/**
 * Queues an item for later upload.
 * @param {object} entry - { projectId, folderPath, kind: 'photo'|'video'|'note', file?, text?, caption?, createdAt }
 */
export async function addToOutbox(entry) {
  const db = await openDB();
  const item = { id: uid(), ...entry };
  await new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).add(item);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
  return item;
}

export async function getOutbox(projectId) {
  const db = await openDB();
  const items = await new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const idx = tx.objectStore(STORE).index('projectId');
    const req = idx.getAll(IDBKeyRange.only(projectId));
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
  db.close();
  return items.sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
}

export async function removeFromOutbox(id) {
  const db = await openDB();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

export async function countOutbox(projectId) {
  return (await getOutbox(projectId)).length;
}
