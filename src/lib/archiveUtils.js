const ARCHIVE_KEY = 'nir_v2_archive';
const DATA_PREFIX  = 'nir_v2_arc_d_';
const MAX_ENTRIES  = 300;

export function loadArchive() {
  try {
    return JSON.parse(localStorage.getItem(ARCHIVE_KEY) || '[]');
  } catch {
    return [];
  }
}

/**
 * Save an archive entry.
 * @param {object} entry   - metadata (type, filename, client, …)
 * @param {object|null} payload - full data to regenerate the doc.
 *   Photo image data (base64) is stripped to keep localStorage small.
 */
export function saveToArchive(entry, payload = null) {
  try {
    const archive = loadArchive();
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    archive.unshift({ id, createdAt: Date.now(), ...entry, hasData: payload !== null });
    localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive.slice(0, MAX_ENTRIES)));
    if (payload !== null) {
      const slim = {
        ...payload,
        photos: (payload.photos || []).map(p => ({ caption: p.caption || '' })),
      };
      localStorage.setItem(DATA_PREFIX + id, JSON.stringify(slim));
    }
  } catch (_) {}
}

export function loadDocData(id) {
  try {
    const raw = localStorage.getItem(DATA_PREFIX + id);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export function deleteFromArchive(id) {
  try {
    const archive = loadArchive().filter(e => e.id !== id);
    localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive));
    localStorage.removeItem(DATA_PREFIX + id);
  } catch (_) {}
}
