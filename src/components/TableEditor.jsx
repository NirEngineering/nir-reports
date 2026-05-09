import { useRef, useState, useEffect } from 'react';
import { STATUS_OPTIONS, PRIORITY_OPTIONS, PRIORITY_OPTIONS_GAP } from '../constants';

// ── Image compression ────────────────────────────────────────────────────────
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

const SPECIAL_COLS = {
  'תקין/לא תקין': STATUS_OPTIONS,
  'קדימות ליקויים': PRIORITY_OPTIONS,
  'קדימות': PRIORITY_OPTIONS,
  'קדימות ליקוי': PRIORITY_OPTIONS,
  'קדימות הליקוי': PRIORITY_OPTIONS_GAP,  // group2: סקר פערי בטיחות (levels 0/1/2)
};

const isElementCol = (col) =>
  ['האלמנט/המתקן', 'האלמנט', 'המתקן', 'האלמנט/המבנה הנבדק', 'סוג הסככה', 'סוג התקרה'].includes(col);

const isFindingsCol = (col) =>
  ['נתונים וממצאים', 'נתונים ופירוט', "מידות (מ') ונתונים", 'ממצאים וליקויים', 'ממצאי הסיור', 'ממצאי ליקויים ודרישות'].includes(col);

const isLocationCol = (col) =>
  ['מיקום', 'מיקום/חדר', 'מיקום הסככה'].includes(col);

const isNotesCol = (col) =>
  ['הערות', 'פירוט והערות', 'טיפול נדרש'].includes(col);

// ── Paste image modal ────────────────────────────────────────────────────────
// Works on desktop (Ctrl+V) AND mobile (long-press → Paste on the contenteditable area)
function PasteModal({ onPaste, onClose }) {
  const editableRef = useRef(null);
  const [hint, setHint] = useState('');

  useEffect(() => {
    // Auto-focus so keyboard shortcuts work immediately on desktop
    setTimeout(() => editableRef.current?.focus(), 100);
  }, []);

  const handlePasteEvent = async (e) => {
    // Works for both desktop Ctrl+V and mobile long-press Paste
    const items = (e.clipboardData || e.originalEvent?.clipboardData)?.items;
    if (!items) return;
    const imgItems = Array.from(items).filter(it => it.type.startsWith('image/'));
    if (!imgItems.length) {
      setHint('לא נמצאה תמונה בלוח. נסה להעתיק תמונה ולהדביק שוב.');
      return;
    }
    e.preventDefault();
    const files = imgItems.map(it => it.getAsFile()).filter(Boolean);
    onPaste(files);
  };

  return (
    <div className="archive-overlay" onClick={onClose}>
      <div
        className="archive-panel"
        style={{ maxWidth: 380, padding: 0 }}
        onClick={e => e.stopPropagation()}
      >
        <div className="archive-header">
          <h2 style={{ fontSize: 15, direction: 'rtl' }}>📋 הדבקת תמונה</h2>
          <button className="archive-close" onClick={onClose}>✕</button>
        </div>

        <div style={{ padding: '16px 20px' }}>
          {/* Contenteditable paste target – works on mobile (long-press → Paste) */}
          <div
            ref={editableRef}
            contentEditable="true"
            suppressContentEditableWarning
            onPaste={handlePasteEvent}
            dir="rtl"
            style={{
              border: '2px dashed #93c5fd',
              borderRadius: 10,
              minHeight: 120,
              padding: '20px 16px',
              textAlign: 'center',
              background: '#eff6ff',
              outline: 'none',
              cursor: 'text',
              color: '#1d4ed8',
              fontSize: 15,
              lineHeight: 1.6,
              userSelect: 'text',
              WebkitUserSelect: 'text',
            }}
          >
            <div style={{ pointerEvents: 'none' }}>
              <div style={{ fontSize: 36, marginBottom: 8 }}>📋</div>
              <div style={{ fontWeight: 700, marginBottom: 4 }}>
                לחץ לכאן, אחר כך הדבק
              </div>
              <div style={{ fontSize: 13, color: '#3b82f6' }}>
                מחשב: Ctrl+V &nbsp;|&nbsp; טלפון: לחיצה ארוכה ← הדבק
              </div>
            </div>
          </div>

          {hint && (
            <div style={{ marginTop: 10, color: '#dc2626', fontSize: 13, textAlign: 'center', direction: 'rtl' }}>
              ⚠️ {hint}
            </div>
          )}

          <div style={{ marginTop: 12, fontSize: 12, color: '#9ca3af', textAlign: 'center', direction: 'rtl', lineHeight: 1.5 }}>
            💡 בטלפון: פתח WhatsApp → לחץ לחיצה ארוכה על תמונה → העתק →
            חזור לאפליקציה → לחץ לחיצה ארוכה בתיבה למעלה → הדבק
          </div>
        </div>

        <div style={{ padding: '0 20px 16px', textAlign: 'center' }}>
          <button className="btn btn-outline btn-sm" onClick={onClose}>ביטול</button>
        </div>
      </div>
    </div>
  );
}

// ── Photo section inside a row card ─────────────────────────────────────────
function CardPhotoSection({ photos, onChange }) {
  const cameraRef = useRef(null);
  const galleryRef = useRef(null);
  const [preview, setPreview] = useState(null);
  const [showPasteModal, setShowPasteModal] = useState(false);

  const handleFiles = async (files) => {
    const arr = Array.from(files).filter(f => f.type.startsWith('image/'));
    if (!arr.length) return;
    const newPhotos = await Promise.all(arr.map(file => new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const compressed = await compressImage(e.target.result);
        resolve({ data: compressed, caption: '' });
      };
      reader.readAsDataURL(file);
    })));
    onChange([...photos, ...newPhotos]);
  };

  const handlePasteFiles = async (files) => {
    await handleFiles(files);
    setShowPasteModal(false);
  };

  const removePhoto = (idx) => onChange(photos.filter((_, i) => i !== idx));
  const updateCaption = (idx, caption) =>
    onChange(photos.map((ph, i) => i === idx ? { ...ph, caption } : ph));

  return (
    <div className="card-photo-section">
      {/* Paste modal */}
      {showPasteModal && (
        <PasteModal
          onPaste={handlePasteFiles}
          onClose={() => setShowPasteModal(false)}
        />
      )}

      {/* Preview modal */}
      {preview !== null && (
        <div className="photo-modal" onClick={() => setPreview(null)}>
          <div className="photo-modal-inner" onClick={e => e.stopPropagation()}>
            <button className="photo-modal-close" onClick={() => setPreview(null)}>✕</button>
            <img src={photos[preview]?.data} alt="" className="photo-modal-img" />
          </div>
        </div>
      )}

      {/* Thumbnail strip */}
      {photos.length > 0 && (
        <div className="card-photo-strip">
          {photos.map((ph, idx) => (
            <div key={idx} className="card-photo-thumb">
              <img src={ph.data} alt="" onClick={() => setPreview(idx)} />
              <button className="card-photo-remove" onClick={() => removePhoto(idx)}>✕</button>
              <textarea
                className="card-photo-caption"
                value={ph.caption}
                onChange={e => updateCaption(idx, e.target.value)}
                placeholder="כיתוב התמונה..."
                rows={2}
                dir="rtl"
              />
            </div>
          ))}
        </div>
      )}

      {/* Add photo buttons */}
      <div className="card-photo-btns">
        <button type="button" className="card-photo-btn" onClick={() => cameraRef.current?.click()}>
          📷 צלם
        </button>
        <button type="button" className="card-photo-btn" onClick={() => galleryRef.current?.click()}>
          🖼️ גלריה
        </button>
        <button type="button" className="card-photo-btn" onClick={() => setShowPasteModal(true)}>
          📋 הדבק
        </button>
        {photos.length > 0 && (
          <span style={{ fontSize: 12, color: '#6b7280', marginRight: 'auto' }}>
            {photos.length} תמונות
          </span>
        )}
      </div>

      <input ref={cameraRef} type="file" accept="image/*" capture="environment" style={{ display: 'none' }}
        onChange={e => { handleFiles(e.target.files); e.target.value = ''; }} />
      <input ref={galleryRef} type="file" accept="image/*" multiple style={{ display: 'none' }}
        onChange={e => { handleFiles(e.target.files); e.target.value = ''; }} />
    </div>
  );
}

// ── Single row card ──────────────────────────────────────────────────────────
function RowCard({
  rowIdx, columns, row, onCellChange, photos, onPhotosChange,
  onDelete, onMoveUp, onMoveDown, isFirst, isLast,
  elementOptions, elementFindings, prevLocation,
}) {
  // Find element value to look up suggestions
  const elementColIdx = columns.findIndex(c => isElementCol(c));
  const elementVal = elementColIdx >= 0 ? (row[elementColIdx] || '') : '';

  // Look up suggestions for current element
  const getSuggestions = (val) => {
    if (!val.trim() || !elementFindings) return [];
    const key = Object.keys(elementFindings).find(
      k => k.trim().toLowerCase() === val.trim().toLowerCase()
    );
    return key ? (elementFindings[key] || []).slice(0, 8) : [];
  };

  const suggestions = getSuggestions(elementVal);

  const renderField = (col, colIdx) => {
    const val = row[colIdx] || '';

    // Status / Priority → select
    const specialOpts = SPECIAL_COLS[col];
    if (specialOpts) {
      return (
        <select
          className="card-field-select"
          value={val}
          onChange={e => onCellChange(colIdx, e.target.value)}
          dir="rtl"
        >
          {specialOpts.map(opt => <option key={opt} value={opt}>{opt}</option>)}
        </select>
      );
    }

    // Number column ("מס'") → small input
    if (col === "מס'") {
      return (
        <input
          type="text"
          className="card-field-input"
          value={val}
          onChange={e => onCellChange(colIdx, e.target.value)}
          placeholder={col}
          style={{ maxWidth: 60 }}
          dir="rtl"
        />
      );
    }

    // Element col → input + datalist (datalist needs <input>, not <textarea>, for mobile autocomplete)
    if (isElementCol(col)) {
      return (
        <div>
          <input
            type="text"
            className="card-field-input"
            value={val}
            onChange={e => onCellChange(colIdx, e.target.value)}
            placeholder={`${col} – הקלד לחיפוש...`}
            list="el-list-main"
            dir="rtl"
            style={{ width: '100%' }}
          />
        </div>
      );
    }

    // Findings col → dropdown suggestions + textarea
    if (isFindingsCol(col)) {
      return (
        <div>
          {suggestions.length > 0 && (
            <div style={{ marginBottom: 6 }}>
              <select
                className="card-field-select"
                value=""
                onChange={e => {
                  if (!e.target.value) return;
                  const current = val.trim();
                  const newVal = current ? `${current}\n• ${e.target.value}` : `• ${e.target.value}`;
                  onCellChange(colIdx, newVal);
                  e.target.value = '';
                }}
                dir="rtl"
              >
                <option value="">💡 בחר ממצא מהרשימה ({suggestions.length} אפשרויות)...</option>
                {suggestions.map((s, i) => (
                  <option key={i} value={s}>{s}</option>
                ))}
              </select>
            </div>
          )}
          <textarea
            className="card-field-textarea"
            value={val}
            onChange={e => onCellChange(colIdx, e.target.value)}
            placeholder={col}
            rows={3}
            dir="rtl"
          />
        </div>
      );
    }

    // Location col → textarea + hint from previous row
    if (isLocationCol(col)) {
      const showHint = !val.trim() && prevLocation;
      return (
        <div>
          <textarea
            className={`card-field-textarea${showHint ? ' card-field-inherited' : ''}`}
            value={val}
            onChange={e => onCellChange(colIdx, e.target.value)}
            placeholder={showHint ? `כמו שורה קודמת: ${prevLocation}` : col}
            rows={2}
            dir="rtl"
          />
          {showHint && (
            <div style={{ fontSize: 11, color: '#6b7280', marginTop: 3, direction: 'rtl' }}>
              ריק = מיקום זהה לשורה הקודמת ({prevLocation}) – שורות אלו ימוזגו בדוח
            </div>
          )}
        </div>
      );
    }

    // Notes / long text → textarea
    if (isNotesCol(col)) {
      return (
        <textarea
          className="card-field-textarea"
          value={val}
          onChange={e => onCellChange(colIdx, e.target.value)}
          placeholder={col}
          rows={3}
          dir="rtl"
        />
      );
    }

    // Default → textarea
    return (
      <textarea
        className="card-field-textarea"
        value={val}
        onChange={e => onCellChange(colIdx, e.target.value)}
        placeholder={col}
        rows={2}
        dir="rtl"
      />
    );
  };

  // Build a short label for the card header (first non-number, non-empty cell)
  const cardTitle = columns
    .map((c, i) => ({ col: c, val: row[i] || '' }))
    .filter(({ col, val }) => col !== "מס'" && val.trim())
    .map(({ val }) => val)[0] || `שורה ${rowIdx + 1}`;

  return (
    <div className="row-card" dir="rtl">
      {/* Card header */}
      <div className="row-card-header">
        <span className="row-card-num">#{rowIdx + 1}</span>
        <span className="row-card-title" title={cardTitle}>
          {cardTitle.length > 35 ? cardTitle.slice(0, 35) + '…' : cardTitle}
        </span>
        <div className="row-card-actions">
          {!isFirst && (
            <button className="btn-icon" onClick={onMoveUp} title="הזז למעלה">↑</button>
          )}
          {!isLast && (
            <button className="btn-icon" onClick={onMoveDown} title="הזז למטה">↓</button>
          )}
          <button className="btn-icon del" onClick={onDelete} title="מחק שורה">🗑</button>
        </div>
      </div>

      {/* Fields */}
      <div className="row-card-fields">
        {columns.map((col, colIdx) => {
          // photo columns are handled by CardPhotoSection below
          if (col === 'תמונה' || col === 'תמונות' || col === 'תמונות הליקוי') return null;
          return (
            <div key={colIdx} className="row-card-field">
              <label className="row-card-field-label">{col}</label>
              {renderField(col, colIdx)}
            </div>
          );
        })}
      </div>

      {/* Photos */}
      <CardPhotoSection
        photos={photos}
        onChange={onPhotosChange}
      />
    </div>
  );
}

// ── TableEditor (main export) ────────────────────────────────────────────────
// Accepts `elements` and `findings` as props — no fetch() calls.
export default function TableEditor({
  columns, rows, onRowsChange, onChange: onChangeLegacy, docType,
  rowPhotos = [], onRowPhotosChange,
  elements = [], findings = {},
}) {
  // Support both `onRowsChange` (new API) and `onChange` (legacy)
  const onChange = onRowsChange || onChangeLegacy;

  // ── Row operations ─────────────────────────────────────────────────────────
  const addRow = () => {
    onChange([...rows, columns.map(() => '')]);
    onRowPhotosChange?.([...rowPhotos, []]);
  };

  const updateCell = (rowIdx, colIdx, value) => {
    onChange(rows.map((row, ri) =>
      ri === rowIdx ? row.map((cell, ci) => ci === colIdx ? value : cell) : row
    ));
  };

  const removeRow = (idx) => {
    onChange(rows.filter((_, i) => i !== idx));
    onRowPhotosChange?.(rowPhotos.filter((_, i) => i !== idx));
  };

  const moveRow = (idx, dir) => {
    const newRows = [...rows];
    const newPhotos = [...rowPhotos];
    const swap = idx + dir;
    if (swap < 0 || swap >= newRows.length) return;
    [newRows[idx], newRows[swap]] = [newRows[swap], newRows[idx]];
    [newPhotos[idx], newPhotos[swap]] = [newPhotos[swap], newPhotos[idx]];
    onChange(newRows);
    onRowPhotosChange?.(newPhotos);
  };

  const updateRowPhotos = (rowIdx, photos) => {
    const updated = [...rowPhotos];
    while (updated.length <= rowIdx) updated.push([]);
    updated[rowIdx] = photos;
    onRowPhotosChange?.(updated);
  };

  return (
    <div>
      {/* Datalist for element autocomplete — populated from props, no fetch */}
      <datalist id="el-list-main">
        {elements.map((opt, i) => <option key={i} value={opt} />)}
      </datalist>

      {/* Row cards */}
      <div className="row-cards-list">
        {rows.length === 0 && (
          <div className="table-empty">
            <div style={{ fontSize: 32, marginBottom: 8 }}>📋</div>
            <div>אין שורות עדיין</div>
            <div style={{ fontSize: 12, opacity: 0.6, marginTop: 4 }}>לחץ "הוסף שורה" להתחיל</div>
          </div>
        )}
        {rows.map((row, rowIdx) => {
          // Compute the effective location of the previous row (for the hint)
          const locColIdx = columns.findIndex(c => isLocationCol(c));
          let prevLocation = '';
          if (locColIdx >= 0 && rowIdx > 0) {
            for (let r = rowIdx - 1; r >= 0; r--) {
              const v = (rows[r][locColIdx] || '').trim();
              if (v) { prevLocation = v; break; }
            }
          }
          return (
            <RowCard
              key={rowIdx}
              rowIdx={rowIdx}
              columns={columns}
              row={row}
              onCellChange={(colIdx, val) => updateCell(rowIdx, colIdx, val)}
              photos={rowPhotos[rowIdx] || []}
              onPhotosChange={(photos) => updateRowPhotos(rowIdx, photos)}
              onDelete={() => removeRow(rowIdx)}
              onMoveUp={() => moveRow(rowIdx, -1)}
              onMoveDown={() => moveRow(rowIdx, 1)}
              isFirst={rowIdx === 0}
              isLast={rowIdx === rows.length - 1}
              elementOptions={elements}
              elementFindings={findings}
              prevLocation={prevLocation}
            />
          );
        })}
      </div>

      {/* Bottom actions */}
      <div className="table-actions">
        <button type="button" className="btn btn-outline btn-sm" onClick={addRow}>
          + הוסף שורה
        </button>
        {rows.length > 0 && (
          <button
            type="button" className="btn btn-sm"
            style={{ background: '#fee2e2', color: '#991b1b' }}
            onClick={() => { if (confirm('למחוק את כל השורות?')) { onChange([]); onRowPhotosChange?.([]); } }}
          >
            נקה הכל
          </button>
        )}
      </div>
    </div>
  );
}
