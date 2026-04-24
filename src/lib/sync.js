import { initializeApp } from 'firebase/app';
import {
  getFirestore, collection, doc,
  setDoc, deleteDoc, onSnapshot,
} from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyA56SQ26YzWmq2TxEObrdfc_vjtx1Br9hw',
  authDomain: 'nir-reports.firebaseapp.com',
  projectId: 'nir-reports',
  storageBucket: 'nir-reports.firebasestorage.app',
  messagingSenderId: '774977940834',
  appId: '1:774977940834:web:7e762aef36133ca467503b',
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

// ── Helpers ───────────────────────────────────────────────────────────────────

function col(syncCode, name) {
  return collection(db, 'syncs', syncCode, name);
}
function ref(syncCode, name, id) {
  return doc(db, 'syncs', syncCode, name, id);
}

// Generate a random 20-char sync code
export function generateSyncCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from({ length: 20 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
}

// ── Field notes sync (photos excluded — too large for Firestore) ──────────────

export function subscribeFieldNotes(syncCode, onUpdate) {
  return onSnapshot(col(syncCode, 'fieldnotes'), snap => {
    const remote = [];
    const removed = [];
    snap.docChanges().forEach(ch => {
      if (ch.type === 'removed') removed.push(ch.doc.id);
      else remote.push(ch.doc.data());
    });
    onUpdate({ remote, removed });
  }, err => console.warn('[sync] fieldnotes error', err));
}

export async function pushFieldNote(syncCode, note) {
  if (!syncCode) return;
  const { photos, ...rest } = note; // exclude photos
  await setDoc(ref(syncCode, 'fieldnotes', note.id), {
    ...rest, updatedAt: new Date().toISOString(),
  });
}

export async function deleteFieldNote(syncCode, id) {
  if (!syncCode) return;
  await deleteDoc(ref(syncCode, 'fieldnotes', id));
}

// ── Drafts sync ───────────────────────────────────────────────────────────────

export function subscribeDrafts(syncCode, onUpdate) {
  return onSnapshot(col(syncCode, 'drafts'), snap => {
    const remote = [];
    const removed = [];
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
