import { precacheAndRoute } from 'workbox-precaching';

// Injected by VitePWA at build time
precacheAndRoute(self.__WB_MANIFEST);

const DB_NAME = 'nir-share-db';
const STORE_NAME = 'shared-images';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = (e) => {
      e.target.result.createObjectStore(STORE_NAME, { autoIncrement: true });
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}

function storeImages(db, blobs) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    blobs.forEach((blob) => store.add(blob));
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Handle Web Share Target POST
  if (
    event.request.method === 'POST' &&
    url.pathname === '/nir-reports/share-target'
  ) {
    event.respondWith(
      (async () => {
        try {
          const formData = await event.request.formData();
          const files = formData.getAll('images');
          const imageBlobs = files.filter((f) => f instanceof File && f.type.startsWith('image/'));

          if (imageBlobs.length > 0) {
            const db = await openDB();
            await storeImages(db, imageBlobs);
            db.close();
          }
        } catch (err) {
          console.error('[SW] share-target error:', err);
        }
        // Redirect to app root
        return Response.redirect('/nir-reports/?shared=1', 303);
      })()
    );
  }
});
