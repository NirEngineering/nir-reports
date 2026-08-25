// Accumulates one field visit's worth of WhatsApp messages (text + photos)
// between a reset/generate and the next one. Single in-memory session — this
// bot watches exactly one WhatsApp chat, so there's only ever one visit "in flight".
let session = { texts: [], photos: [] };

export function getSession() {
  return session;
}

export function addText(text) {
  session.texts.push({ text, ts: Date.now() });
}

export function addPhoto(data, caption) {
  session.photos.push({ data, caption: caption || '' });
}

export function resetSession() {
  session = { texts: [], photos: [] };
}

export function isEmpty() {
  return session.texts.length === 0 && session.photos.length === 0;
}
