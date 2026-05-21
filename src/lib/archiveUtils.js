const ARCHIVE_KEY = 'nir_v2_archive';
const MAX_ENTRIES = 300;

export function loadArchive() {
  try {
    return JSON.parse(localStorage.getItem(ARCHIVE_KEY) || '[]');
  } catch {
    return [];
  }
}

export function saveToArchive(entry) {
  try {
    const archive = loadArchive();
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    archive.unshift({ id, createdAt: Date.now(), ...entry });
    localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive.slice(0, MAX_ENTRIES)));
  } catch (_) {}
}

export function deleteFromArchive(id) {
  try {
    const archive = loadArchive().filter(e => e.id !== id);
    localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive));
  } catch (_) {}
}
