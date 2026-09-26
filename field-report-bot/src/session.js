// Accumulates one field visit's worth of WhatsApp messages (text + photos)
// between a reset/generate and the next one. Single in-memory session — this
// bot watches exactly one WhatsApp chat, so there's only ever one visit "in flight".
//
// `structured` holds wizard-collected data too rigid to trust to Claude's
// free-text extraction (e.g. group11's 15-item checklist, where every status
// must land on the exact right row) — classify.js merges it directly into
// the generated document's payload instead of re-deriving it from prose.
let session = { texts: [], photos: [], structured: {} };

export function getSession() {
  return session;
}

export function addText(text) {
  session.texts.push({ text, ts: Date.now() });
}

export function addPhoto(data, caption) {
  session.photos.push({ data, caption: caption || '' });
}

export function setStructured(key, value) {
  session.structured[key] = value;
}

export function resetSession() {
  session = { texts: [], photos: [], structured: {} };
}

export function isEmpty() {
  return session.texts.length === 0 && session.photos.length === 0;
}
