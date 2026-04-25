import { initializeApp } from 'firebase/app';
import {
  initializeFirestore, collection, doc,
  setDoc, deleteDoc, getDoc, onSnapshot,
} from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyDOv9qOVJw5jDEUK6usJDBOO_kf_1fpqLE',
  authDomain: 'nir-reports-2.firebaseapp.com',
  projectId: 'nir-reports-2',
  storageBucket: 'nir-reports-2.firebasestorage.app',
  messagingSenderId: '143729984389',
  appId: '1:143729984389:web:4daeae4f71b39c550bfaec',
};

const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(app, {
  experimentalForceLongPolling: true,
});


// ── Helpers ───────────────────────────────────────────────────────────────────

function col(syncCode, name) {
  return collection(db, 'syncs', syncCode, name);
}
function ref(syncCode, name, id) {
  return doc(db, 'syncs', syncCode, name, id);
}

export function generateSyncCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from({ length: 20 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
}

// ── Field notes sync ──────────────────────────────────────────────────────────

export function subscribeFieldNotes(syncCode, onUpdate) {
  return onSnapshot(col(syncCode, 'fieldnotes'), snap => {
    const remote = [], removed = [];
    snap.docChanges().forEach(ch => {
      if (ch.type === 'removed') removed.push(ch.doc.id);
      else remote.push(ch.doc.data());
    });
    onUpdate({ remote, removed });
  }, err => console.warn('[sync] fieldnotes error', err));
}

export async function pushFieldNote(syncCode, note) {
  if (!syncCode || !note?.id) return;
  await setDoc(ref(syncCode, 'fieldnotes', String(note.id)), {
    ...note, updatedAt: new Date().toISOString(),
  });
}

export async function deleteFieldNote(syncCode, id) {
  if (!syncCode) return;
  await deleteDoc(ref(syncCode, 'fieldnotes', id));
}

// ── Drafts sync ───────────────────────────────────────────────────────────────

export function subscribeDrafts(syncCode, onUpdate) {
  return onSnapshot(col(syncCode, 'drafts'), snap => {
    const remote = [], removed = [];
    snap.docChanges().forEach(ch => {
      if (ch.type === 'removed') removed.push(ch.doc.id);
      else remote.push(ch.doc.data());
    });
    onUpdate({ remote, removed });
  }, err => console.warn('[sync] drafts error', err));
}

export async function pushDraft(syncCode, draft) {
  if (!syncCode || !draft.id) return;
  await setDoc(ref(syncCode, 'drafts', draft.id), {
    ...draft, updatedAt: new Date().toISOString(),
  });
}

export async function deleteDraft(syncCode, id) {
  if (!syncCode) return;
  await deleteDoc(ref(syncCode, 'drafts', id));
}

// ── Photo sync via Firestore ──────────────────────────────────────────────────
// One Firestore document per photo: syncs/{syncCode}/photos/{noteId}_{index}
// { noteId, index, data: base64, caption }
// Each photo is compressed to stay under Firestore's 1 MB document limit.

// Re-compress a data URL until it fits within maxBytes
async function compressForFirestore(dataUrl, maxBytes = 700_000) {
  if (dataUrl.length <= maxBytes) return dataUrl;
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      let { width, height } = img;
      // Reduce dimensions first
      const MAX_DIM = 900;
      if (width > MAX_DIM) { height = Math.round(height * MAX_DIM / width); width = MAX_DIM; }
      else if (height > MAX_DIM) { width = Math.round(width * MAX_DIM / height); height = MAX_DIM; }
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      let quality = 0.72;
      let result = canvas.toDataURL('image/jpeg', quality);
      while (result.length > maxBytes && quality > 0.3) {
        quality = Math.round((quality - 0.1) * 10) / 10;
        result = canvas.toDataURL('image/jpeg', quality);
      }
      resolve(result);
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

export async function pushNotePhotos(syncCode, noteId, photos) {
  if (!syncCode || !noteId || !photos?.length) return;
  await Promise.all(
    photos.map(async (ph, i) => {
      if (!ph?.data) return;
      const data = await compressForFirestore(ph.data);
      return setDoc(ref(syncCode, 'photos', `${String(noteId)}_${i}`), {
        noteId: String(noteId),
        index: i,
        data,
        caption: ph.caption || '',
        updatedAt: new Date().toISOString(),
      });
    })
  );
}

// Real-time listener — fires with all existing photo docs on connect,
// then fires again whenever any photo document is added or changed.
export function subscribePhotos(syncCode, onUpdate) {
  return onSnapshot(col(syncCode, 'photos'), snap => {
    const changes = [];
    snap.docChanges().forEach(ch => {
      if (ch.type !== 'removed') changes.push(ch.doc.data());
    });
    if (changes.length > 0) onUpdate(changes);
  }, err => console.warn('[sync] photos error', err));
}

export async function deleteNotePhotos(syncCode, noteId) {
  if (!syncCode || !noteId) return;
  const id = String(noteId);
  for (let i = 0; i < 30; i++) {
    try {
      const snap = await getDoc(ref(syncCode, 'photos', `${id}_${i}`));
      if (!snap.exists()) break;
      await deleteDoc(ref(syncCode, 'photos', `${id}_${i}`));
    } catch (_) { break; }
  }
}
