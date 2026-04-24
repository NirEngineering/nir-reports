import { precacheAndRoute, cleanupOutdatedCaches } from 'workbox-precaching';
import { clientsClaim } from 'workbox-core';

// ── Share Target handler ──────────────────────────────────────────────────────
// MUST be registered BEFORE precacheAndRoute so our handler intercepts
// POST /nir-reports/share before workbox falls through to network (→ 405).

const DB_NAME  = 'nir-share';
const DB_STORE = 'pending-images';

function openShareDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(DB_STORE, { autoIncrement: true });
    req.onsuccess  = () => resolve(req.result);
    req.onerror    = () => reject(req.error);
  });
}

async function storeSharedFiles(files) {
  const db = await openShareDb();
  return new Promise((resolve, reject) => {
    const tx    = db.transaction(DB_STORE, 'readwrite');
    const store = tx.objectStore(DB_STORE);
    files.forEach(f => store.add(f));
    tx.oncomplete = resolve;
    tx.onerror    = () => reject(tx.error);
  });
}

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method === 'POST' && url.pathname === '/nir-reports/share') {
    event.respondWith(
      (async () => {
        try {
          const fd    = await event.request.formData();
          const files = fd.getAll('images');
          if (files.length) await storeSharedFiles(files);
        } catch (e) {
          console.error('[SW] share-target error', e);
        }
        return Response.redirect('/nir-reports/?shared=1', 303);
      })()
    );
  }
});

// ── Workbox setup (after share handler) ──────────────────────────────────────

self.skipWaiting();
clientsClaim();
cleanupOutdatedCaches();
precacheAndRoute(self.__WB_MANIFEST);
