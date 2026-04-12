import { useRef, useState, useEffect } from 'react';

// Compress image: resize to max 1200px, 82% JPEG quality (~150-300KB per photo)
async function compressImage(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const MAX = 1200;
      let { width, height } = img;
      if (width > MAX) { height = Math.round(height * MAX / width); width = MAX; }
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', 0.82));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

// ── Paste Modal – contentEditable for mobile long-press paste ─────────────────
function PasteModal({ onClose, onPaste }) {
  const divRef = useRef(null);

  useEffect(() => {
    setTimeout(() => divRef.current?.focus(), 100);
  }, []);

  const handlePasteEvent = async (e) => {
    const items = (e.clipboardData || e.originalEvent?.clipboardData)?.items;
    if (!items) return;
    const imgItems = Array.from(items).filter(it => it.type.startsWith('image/'));
    if (imgItems.length === 0) return;
    e.preventDefault();
    const files = imgItems.map(it => it.getAsFile()).filter(Boolean);
    await onPaste(files);
    onClose();
  };

  const handleInput = async (e) => {
    // Catch images pasted as <img> nodes on mobile
    const imgs = divRef.current?.querySelectorAll('img');
    if (imgs && imgs.length > 0) {
      const src = imgs[0].src;
      if (src && src.startsWith('data:image')) {
        divRef.current.innerHTML = '';
        const compressed = await compressImage(src);
        await onPaste([{ _isDataUrl: true, dataUrl: compressed }]);
        onClose();
      }
    }
  };

  return (
    <div className="paste-modal-overlay" onClick={onClose}>
      <div className="paste-modal-box" onClick={e => e.stopPropagation()}>
        <div className="paste-modal-title">📋 הדבקת תמונה</div>
        <div className="paste-modal-hint">
          לחץ לחיצה ארוכה בתיבה למטה ובחר "הדבק"
        </div>
        <div
          ref={divRef}
          contentEditable
          suppressContentEditableWarning
          className="paste-modal-area"
          onPaste={handlePasteEvent}
          onInput={handleInput}
          dir="rtl"
        >
          הדבק תמונה כאן
        </div>
        <button className="btn btn-outline" style={{ marginTop: 12 }} onClick={onClose}>ביטול</button>
      </div>
    </div>
  );
}

export default function PhotoUpload({ photos, onChange }) {
  const fileInput = useRef(null);
  const cameraInput = useRef(null);
  const [preview, setPreview] = useState(null);
  const [showPasteModal, setShowPasteModal] = useState(false);

  // Paste support (Ctrl+V from clipboard – desktop)
  useEffect(() => {
    const handlePaste = async (e) => {
      const items = (e.clipboardData || e.originalEvent?.clipboardData)?.items;
      if (!items) return;
      const imgItems = Array.from(items).filter(it => it.type.startsWith('image/'));
      if (imgItems.length === 0) return;
      e.preventDefault();
      const files = imgItems.map(it => it.getAsFile()).filter(Boolean);
      await handleFiles(files);
    };
    document.addEventListener('paste', handlePaste);
    return () => document.removeEventListener('paste', handlePaste);
  }, [photos]);

  const handleFiles = async (files) => {
    // Support both File objects and {_isDataUrl, dataUrl} from PasteModal
    const arr = Array.from(files);
    if (!arr.length) return;
    const newPhotos = await Promise.all(arr.map(file => new Promise((resolve) => {
      if (file._isDataUrl) {
        resolve({ data: file.dataUrl, caption: '' });
        return;
      }
      if (!file.type?.startsWith('image/')) { resolve(null); return; }
      const reader = new FileReader();
      reader.onload = async (e) => {
        const compressed = await compressImage(e.target.result);
        resolve({ data: compressed, caption: '' });
      };
      reader.readAsDataURL(file);
    })));
    onChange([...photos, ...newPhotos.filter(Boolean)]);
  };

  const updateCaption = (idx, caption) => {
    onChange(photos.map((p, i) => i === idx ? { ...p, caption } : p));
  };

  const removePhoto = (idx) => {
    onChange(photos.filter((_, i) => i !== idx));
  };

  const movePhoto = (idx, dir) => {
    const arr = [...photos];
    const swap = idx + dir;
    if (swap < 0 || swap >= arr.length) return;
    [arr[idx], arr[swap]] = [arr[swap], arr[idx]];
    onChange(arr);
  };

  return (
    <div>
      {/* Paste Modal */}
      {showPasteModal && (
        <PasteModal
          onClose={() => setShowPasteModal(false)}
          onPaste={handleFiles}
        />
      )}

      {/* Full-screen preview modal */}
      {preview !== null && (
        <div className="photo-modal" onClick={() => setPreview(null)}>
          <div className="photo-modal-inner" onClick={e => e.stopPropagation()}>
            <button className="photo-modal-close" onClick={() => setPreview(null)}>✕</button>
            <img src={photos[preview]?.data} alt="תצוגה" className="photo-modal-img" />
            {photos[preview]?.caption && (
              <div className="photo-modal-caption">{photos[preview].caption}</div>
            )}
          </div>
        </div>
      )}

      {/* Photo grid */}
      {photos.length > 0 && (
        <div className="photo-grid">
          {photos.map((photo, idx) => (
            <div key={idx} className="photo-item">
              {/* Number badge */}
              <div className="photo-number">#{idx + 1}</div>

              {/* Image */}
              <img
                src={photo.data}
                alt={photo.caption || `תמונה ${idx + 1}`}
                onClick={() => setPreview(idx)}
              />

              {/* Top actions bar */}
              <div className="photo-actions-top">
                {idx > 0 && (
                  <button className="photo-action-btn" onClick={() => movePhoto(idx, -1)} title="הזז שמאלה">‹</button>
                )}
                <button className="photo-action-btn danger" onClick={() => removePhoto(idx)} title="מחק">✕</button>
                {idx < photos.length - 1 && (
                  <button className="photo-action-btn" onClick={() => movePhoto(idx, 1)} title="הזז ימינה">›</button>
                )}
              </div>

              {/* Caption input */}
              <input
                className="photo-caption-input"
                type="text"
                value={photo.caption}
                onChange={e => updateCaption(idx, e.target.value)}
                placeholder="כיתוב..."
              />
            </div>
          ))}
        </div>
      )}

      {photos.length === 0 && (
        <div className="photo-empty">
          <div style={{ fontSize: 40, marginBottom: 8 }}>📷</div>
          <div>לא נבחרו תמונות</div>
          <div style={{ fontSize: 12, opacity: 0.6, marginTop: 4 }}>ניתן להוסיף עד 30 תמונות</div>
        </div>
      )}

      {/* Upload buttons */}
      <div className="photo-upload-btns">
        <button
          type="button"
          className="btn btn-outline photo-btn"
          onClick={() => cameraInput.current.click()}
        >
          <span>📷</span>
          <span>צלם</span>
        </button>
        <button
          type="button"
          className="btn btn-outline photo-btn"
          onClick={() => fileInput.current.click()}
        >
          <span>🖼️</span>
          <span>גלריה</span>
        </button>
        <button
          type="button"
          className="btn btn-outline photo-btn"
          onClick={() => setShowPasteModal(true)}
        >
          <span>📋</span>
          <span>הדבק</span>
        </button>
      </div>

      {photos.length > 0 && (
        <div style={{ textAlign: 'center', marginTop: 8 }}>
          <button
            type="button"
            className="btn btn-sm"
            style={{ background: '#fee2e2', color: '#991b1b', fontSize: 12 }}
            onClick={() => { if (confirm('למחוק את כל התמונות?')) onChange([]); }}
          >
            מחק הכל
          </button>
        </div>
      )}

      {/* Hidden inputs */}
      <input
        ref={cameraInput}
        type="file"
        accept="image/*"
        capture="environment"
        style={{ display: 'none' }}
        onChange={e => { handleFiles(e.target.files); e.target.value = ''; }}
      />
      <input
        ref={fileInput}
        type="file"
        accept="image/*"
        multiple
        style={{ display: 'none' }}
        onChange={e => { handleFiles(e.target.files); e.target.value = ''; }}
      />
    </div>
  );
}
