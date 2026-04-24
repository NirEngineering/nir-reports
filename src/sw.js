import { precacheAndRoute, cleanupOutdatedCaches } from 'workbox-precaching';
import { clientsClaim } from 'workbox-core';

self.skipWaiting();
clientsClaim();

cleanupOutdatedCaches();
precacheAndRoute(self.__WB_MANIFEST);

// ── Share Target handler ──────────────────────────────────────────────────────
// Android calls POST /nir-reports/share when user shares images to this PWA.
// We extract the files, store them in IndexedDB, then redirect to the app.

const DB_NAME  = 'nir-share';
const DB_STORE = 'pending-images';
const DB_VER   = 1;

function openShareDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VER);
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
  if (url.pathname === '/nir-reports/share' && event.request.method === 'POST') {
    event.respondWith(
      (async () => {
        try {
          const fd    = await event.request.formData();
          const files = fd.getAll('images');
          await storeSharedFiles(files);
        } catch (e) {
          console.error('[SW] share-target error', e);
        }
        return Response.redirect('/nir-reports/?shared=1', 303);
      })()
    );
  }
});
