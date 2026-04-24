import { useState, useEffect, useRef } from 'react';
import TableEditor from './components/TableEditor';
import PhotoUpload from './components/PhotoUpload';
import SearchDropdown from './components/SearchDropdown';
import FieldNotesEditor from './components/FieldNotesEditor';
import {
  DOC_TYPES_CONFIG, TABLE_COLUMNS, DEFECTS_COLUMNS, KNOWN_ORGANIZATIONS,
  DRAFT_KEY, DRAFTS_KEY, ARCHIVE_KEY, FIELDNOTES_KEY,
} from './constants';
import { generateDocument, generateFieldNotesDocument } from './lib/docGenerator';
import { importDocx, exportDocx } from './lib/docImporter';
import elementsData from './data/elements_by_type.json';
import findingsData from './data/findings_by_type.json';
import clientsData from './data/clients.json';
import './index.css';

// ── Helpers ──────────────────────────────────────────────────────────────────

const isoToDisplay = (iso) => {
  if (!iso || !iso.includes('-')) return iso || '';
  const [y, m, d] = iso.split('-');
  return `${parseInt(d)}.${parseInt(m)}.${y}`;
};

const nowLabel = () => {
  const d = new Date();
  return `${d.getDate()}.${d.getMonth()+1}.${d.getFullYear()} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
};

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

async function compressImage(dataUrl, maxW = 1200, quality = 0.82) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      let { width, height } = img;
      if (width > maxW) { height = Math.round(height * maxW / width); width = maxW; }
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

const defaultForm = () => ({
  date: new Date().toISOString().split('T')[0],
  inspection_date: new Date().toISOString().split('T')[0],
  client: '',
  organization: "החברה למוסדות חינוך ותרבות ת''א",
  location: '',
  address: '',
  subject: '',
  intro_extra: '',
  has_defects: false,
  conclusion_custom: '',
  notes_custom: '',
  // group6 fields
  building_stability_data: '',
  building_stability_status: 'לא תקין',
  building_stability_priority: '2',
});

// ── localStorage helpers ──────────────────────────────────────────────────────

function lsGet(key, fallback) {
  try { const r = localStorage.getItem(key); return r ? JSON.parse(r) : fallback; }
  catch (_) { return fallback; }
}

function lsSet(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
}

function lsRemove(key) {
  try { localStorage.removeItem(key); } catch (_) {}
}

// ── StepBar (clickable) ───────────────────────────────────────────────────────
const STEP_LABELS = ['פרטים', 'טבלה', 'תמונות', 'יצור'];

function StepBar({ step, onStepClick }) {
  // step prop = 0-based index of active label (0=פרטים … 3=יצור)
  return (
    <div className="stepbar">
      {STEP_LABELS.map((l, i) => (
        <div
          key={i}
          className={`stepbar-item${i === step ? ' active' : i < step ? ' done' : ''}${onStepClick ? ' clickable' : ''}`}
          onClick={() => onStepClick?.(i + 1)}
          style={{ cursor: onStepClick ? 'pointer' : 'default' }}
        >
          <div className="stepbar-circle">{i < step ? '✓' : i + 1}</div>
          <div className="stepbar-label">{l}</div>
        </div>
      ))}
    </div>
  );
}

// ── Field wrapper ─────────────────────────────────────────────────────────────
function Field({ label, required, children }) {
  return (
    <div className="form-group">
      <label className="form-label">
        {label}{required && <span className="required"> *</span>}
      </label>
      {children}
    </div>
  );
}

// ── RichTextarea ──────────────────────────────────────────────────────────────
function RichTextarea({ value, onChange, placeholder, rows = 4 }) {
  const taRef = useRef(null);

  const insert = (prefix) => {
    const ta = taRef.current;
    if (!ta) return;
    const { selectionStart: s, selectionEnd: e, value: v } = ta;
    const line = prefix + (s === e ? '' : v.slice(s, e));
    const next = v.slice(0, s) + line + '\n' + v.slice(e);
    onChange(next);
    requestAnimationFrame(() => {
      ta.selectionStart = ta.selectionEnd = s + line.length + 1;
      ta.focus();
    });
  };

  const autoResize = () => {
    const ta = taRef.current;
    if (ta) { ta.style.height = 'auto'; ta.style.height = ta.scrollHeight + 'px'; }
  };

  return (
    <div className="rich-textarea-wrapper">
      <div className="rich-toolbar">
        <button type="button" className="rich-btn" onClick={() => insert('• ')}>• נקודה</button>
        <button type="button" className="rich-btn" onClick={() => insert('1. ')}>1. ממוספר</button>
        <button type="button" className="rich-btn" onClick={() => {
          const ta = taRef.current;
          if (!ta) return;
          const v = ta.value;
          const pos = ta.selectionStart;
          const next = v.slice(0, pos) + '\n' + v.slice(pos);
          onChange(next);
          requestAnimationFrame(() => { ta.selectionStart = ta.selectionEnd = pos + 1; ta.focus(); });
        }}>↵ שורה</button>
      </div>
      <textarea
        ref={taRef}
        className="form-textarea rich"
        value={value}
        onChange={e => { onChange(e.target.value); autoResize(); }}
        onInput={autoResize}
        placeholder={placeholder}
        rows={rows}
        dir="rtl"
      />
    </div>
  );
}

// ── DraftNameModal ────────────────────────────────────────────────────────────
function DraftNameModal({ defaultName, onSave, onClose }) {
  const [name, setName] = useState(defaultName);
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={e => e.stopPropagation()}>
        <div className="modal-title">שמור טיוטה</div>
        <input
          className="form-input"
          value={name}
          onChange={e => setName(e.target.value)}
          dir="rtl"
          autoFocus
        />
        <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
          <button className="btn btn-outline" style={{ flex: 1 }} onClick={onClose}>ביטול</button>
          <button className="btn btn-primary" style={{ flex: 1 }} onClick={() => onSave(name.trim() || defaultName)}>שמור</button>
        </div>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  // mode: 'home' | 'new' | 'edit' | 'drafts' | 'archive'
  const [mode, setMode]     = useState('home');
  const [step, setStep]     = useState(0);   // 0=type,1=details,2=table,3=photos,4=generate
  const [docType, setDocType]       = useState('');
  const [form, setForm]             = useState(defaultForm());
  const [tableRows, setTableRows]   = useState([]);
  const [defectsRows, setDefectsRows] = useState([]);
  const [rowPhotos, setRowPhotos]   = useState([]);
  const [defectsRowPhotos, setDefectsRowPhotos] = useState([]);
  const [photos, setPhotos]         = useState([]);
  const [loading, setLoading]       = useState(false);
  const [error, setError]           = useState('');
  const [success, setSuccess]       = useState('');

  // Draft / archive lists
  const [hasDraft, setHasDraft]     = useState(false);
  const [drafts, setDrafts]         = useState([]);
  const [archive, setArchive]       = useState([]);
  const [showDraftModal, setShowDraftModal] = useState(false);

  // Edit mode
  const [editImport, setEditImport] = useState(null);
  const [editEdits, setEditEdits]   = useState({});
  const [editLoading, setEditLoading] = useState(false);
  const [editFile, setEditFile]     = useState(null);
  const fileRef    = useRef(null);
  const importRef  = useRef(null);

  // Archive search
  const [archiveSearch, setArchiveSearch] = useState('');

  // Field Notes state
  const [fieldNotes, setFieldNotes]   = useState([]);
  const [fnMode, setFnMode]           = useState('list'); // 'list' | 'edit'
  const [fnCurrent, setFnCurrent]     = useState(null);   // current note being edited
  const [fnLoading, setFnLoading]     = useState(false);

  // ── Init: load draft / drafts / archive ───────────────────────────────────
  useEffect(() => {
    const d = lsGet(DRAFT_KEY, null);
    if (d?.mode === 'new') setHasDraft(true);
    setDrafts(lsGet(DRAFTS_KEY, []));
    setArchive(lsGet(ARCHIVE_KEY, []));
    setFieldNotes(lsGet(FIELDNOTES_KEY, []));
  }, []);

  // ── Auto-save current draft ───────────────────────────────────────────────
  useEffect(() => {
    if (mode === 'new' && docType) {
      lsSet(DRAFT_KEY, { mode: 'new', step, docType, form, tableRows, defectsRows });
    }
  }, [mode, step, docType, form, tableRows, defectsRows]);

  // ── Draft helpers ──────────────────────────────────────────────────────────
  const applyDraftData = (d) => {
    if (d.docType)    setDocType(d.docType);
    if (d.form)       setForm(d.form);
    if (d.tableRows)  setTableRows(d.tableRows);
    if (d.defectsRows) setDefectsRows(d.defectsRows);
    setRowPhotos([]); setDefectsRowPhotos([]); setPhotos([]);
    setStep(d.step || 1);
    setMode('new');
  };

  const restoreAutoDraft = () => {
    const d = lsGet(DRAFT_KEY, null);
    if (d) { applyDraftData(d); setHasDraft(false); }
  };

  const discardAutoDraft = () => {
    lsRemove(DRAFT_KEY);
    setHasDraft(false);
  };

  const saveDraftManual = (name) => {
    const entry = {
      id: uid(),
      name,
      updated_at: nowLabel(),
      docType,
      form,
      tableRows,
      defectsRows,
      step,
    };
    const list = [entry, ...lsGet(DRAFTS_KEY, [])].slice(0, 30);
    lsSet(DRAFTS_KEY, list);
    setDrafts(list);
    setShowDraftModal(false);
    setSuccess('✅ הטיוטה נשמרה!');
    setTimeout(() => setSuccess(''), 2500);
  };

  const restoreDraftFromList = (d) => {
    applyDraftData(d);
    setError(''); setSuccess('');
  };

  const deleteDraft = (id) => {
    const list = drafts.filter(d => d.id !== id);
    lsSet(DRAFTS_KEY, list);
    setDrafts(list);
  };

  // ── Archive helpers ────────────────────────────────────────────────────────
  const addToArchive = (meta) => {
    const list = [{ ...meta, id: uid(), created_at: nowLabel() }, ...lsGet(ARCHIVE_KEY, [])].slice(0, 100);
    lsSet(ARCHIVE_KEY, list);
    setArchive(list);
  };

  const deleteArchiveEntry = (id) => {
    const list = archive.filter(a => a.id !== id);
    lsSet(ARCHIVE_KEY, list);
    setArchive(list);
  };

  // ── JSON backup / restore ──────────────────────────────────────────────────
  const handleExportJSON = () => {
    const payload = {
      version: 2,
      exported_at: new Date().toISOString(),
      drafts:  lsGet(DRAFTS_KEY, []),
      archive: lsGet(ARCHIVE_KEY, []),
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    const d    = new Date();
    a.download = `ניר-הנדסה-גיבוי-${d.getDate()}.${d.getMonth()+1}.${d.getFullYear()}.json`;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setSuccess('✅ הגיבוי הורד!');
    setTimeout(() => setSuccess(''), 2500);
  };

  const handleImportJSON = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        let imported = 0;
        if (Array.isArray(data.drafts)) {
          const merged = [...data.drafts, ...lsGet(DRAFTS_KEY, [])];
          const unique = [...new Map(merged.map(d => [d.id, d])).values()].slice(0, 30);
          lsSet(DRAFTS_KEY, unique);
          setDrafts(unique);
          imported += data.drafts.length;
        }
        if (Array.isArray(data.archive)) {
          const merged = [...data.archive, ...lsGet(ARCHIVE_KEY, [])];
          const unique = [...new Map(merged.map(a => [a.id, a])).values()].slice(0, 100);
          lsSet(ARCHIVE_KEY, unique);
          setArchive(unique);
          imported += data.archive.length;
        }
        setSuccess(`✅ יובאו ${imported} פריטים!`);
        setTimeout(() => setSuccess(''), 3000);
      } catch {
        setError('שגיאה: קובץ JSON לא תקין');
      }
    };
    reader.readAsText(file);
  };

  const restoreFromArchive = (entry) => {
    if (entry.form)       setForm(entry.form);
    if (entry.docType)    setDocType(entry.docType);
    if (entry.tableRows)  setTableRows(entry.tableRows);
    if (entry.defectsRows) setDefectsRows(entry.defectsRows);
    setRowPhotos([]); setDefectsRowPhotos([]); setPhotos([]);
    setStep(1);
    setMode('new');
    setError(''); setSuccess('');
  };

  // ── Generate ───────────────────────────────────────────────────────────────
  const handleGenerate = async () => {
    setLoading(true); setError(''); setSuccess('');
    try {
      const mergedPhotos = [];
      tableRows.forEach((row, idx) => {
        (rowPhotos[idx] || []).forEach(ph => {
          mergedPhotos.push({ data: ph.data, caption: ph.caption || [row[0], row[1]].filter(Boolean).join(' – ') });
        });
      });
      defectsRows.forEach((row, idx) => {
        (defectsRowPhotos[idx] || []).forEach(ph => {
          mergedPhotos.push({ data: ph.data, caption: ph.caption || `ליקוי ${idx + 1}: ${row[1] || row[0] || ''}` });
        });
      });
      photos.forEach(ph => mergedPhotos.push({ data: ph.data, caption: ph.caption || '' }));

      const payload = {
        doc_type: docType,
        ...form,
        inspection_date: isoToDisplay(form.inspection_date),
        inspection_date_raw: form.inspection_date,
        table_rows: tableRows,
        defects_rows: defectsRows,
        photos: mergedPhotos,
        notes_custom: typeof form.notes_custom === 'string'
          ? form.notes_custom.split('\n').filter(s => s.trim())
          : (form.notes_custom || []),
      };

      const blob = await generateDocument(payload);
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      const DOC_TYPE_PREFIX = {
        group1: 'דו"ח יציבות',
        group3: 'סקר תקרות',
        group4: 'סקר יציבות',
        group5: 'בדיקת סככות',
        group6: 'סקר מבנים',
      };
      a.download = `${DOC_TYPE_PREFIX[form.doc_type] || form.subject || 'דוח'} - ${form.client}.docx`;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      // Save to archive (no photos – too large)
      addToArchive({
        doc_type: docType,
        doc_type_name: DOC_TYPES_CONFIG[docType]?.name?.replace('\n', ' ') ?? docType,
        client:    form.client,
        subject:   form.subject,
        location:  form.location,
        inspection_date: isoToDisplay(form.inspection_date),
        form, tableRows, defectsRows,
      });

      setSuccess('✅ המסמך נוצר בהצלחה!');
      lsRemove(DRAFT_KEY);
    } catch (e) {
      setError(`שגיאה: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  // ── Edit export ────────────────────────────────────────────────────────────
  const handleEditExport = async () => {
    if (!editImport) return;
    setEditLoading(true); setError('');
    try {
      const blob = await exportDocx(editImport, editEdits);
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = `${editFile?.name?.replace('.docx', '') || 'מסמך'} - ערוך.docx`;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setSuccess('✅ המסמך הערוך הורד!');
    } catch (e) {
      setError(`שגיאה: ${e.message}`);
    } finally {
      setEditLoading(false);
    }
  };

  // ── Edit upload ────────────────────────────────────────────────────────────
  const handleEditUpload = async (file) => {
    if (!file?.name.endsWith('.docx')) return setError('יש לבחור קובץ .docx');
    setEditFile(file); setEditLoading(true); setError('');
    try {
      const result = await importDocx(file);
      setEditImport(result);
      setEditEdits({ paragraphs: {} });
    } catch (e) {
      setError(e.message);
    } finally {
      setEditLoading(false);
    }
  };

  const setField = (key, val) => setForm(f => {
    const next = { ...f, [key]: val };
    // Location defaults to client name unless user has typed something different
    if (key === 'client' && (f.location === '' || f.location === f.client)) {
      next.location = val;
    }
    return next;
  });

  // ── Field Notes helpers ────────────────────────────────────────────────────
  const defaultFn = () => ({
    id: uid(),
    created_at: new Date().toISOString().split('T')[0],
    updated_at: nowLabel(),
    client:   '',
    location: '',
    title:    '',
    content:  '',
    photos:   [],
  });

  const fnSave = (note) => {
    const updated = { ...note, updated_at: nowLabel() };
    const list = [updated, ...lsGet(FIELDNOTES_KEY, []).filter(n => n.id !== updated.id)].slice(0, 50);
    lsSet(FIELDNOTES_KEY, list);
    setFieldNotes(list);
    return updated;
  };

  const fnDelete = (id) => {
    const list = fieldNotes.filter(n => n.id !== id);
    lsSet(FIELDNOTES_KEY, list);
    setFieldNotes(list);
  };

  const fnGenerate = async (note) => {
    setFnLoading(true); setError('');
    try {
      const blob = await generateFieldNotesDocument({
        date:     note.created_at,
        title:    note.title,
        client:   note.client,
        location: note.location,
        content:  note.content,
        photos:   note.photos,
      });
      const url = URL.createObjectURL(blob);
      const a   = document.createElement('a');
      a.href    = url;
      a.download = `${note.title || 'יומן שטח'} - ${note.client || note.location || 'ניר הנדסה'}.docx`;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setSuccess('✅ המסמך נוצר!');
      setTimeout(() => setSuccess(''), 2500);
    } catch (e) {
      setError(`שגיאה: ${e.message}`);
    } finally {
      setFnLoading(false);
    }
  };

  // ── Reset / navigation ─────────────────────────────────────────────────────
  const goHome = () => {
    setMode('home'); setStep(0); setError(''); setSuccess('');
    setDocType(''); setEditImport(null); setEditFile(null); setEditEdits({});
  };

  // Step navigation bar click handler
  const handleStepBarClick = (targetStep) => {
    if (!docType) return;
    setError('');
    setStep(targetStep);
  };

  // Step nav buttons (next/back) with save draft button
  const StepNav = ({ back, next, backLabel = '◀ חזור', nextLabel = 'המשך ▶', onNext }) => (
    <div style={{ display: 'flex', gap: 8, marginTop: 12, flexWrap: 'wrap' }}>
      {back != null && (
        <button className="btn btn-outline" onClick={() => { setError(''); setStep(back); }}>
          {backLabel}
        </button>
      )}
      <button
        className="btn btn-sm btn-outline"
        onClick={() => setShowDraftModal(true)}
        title="שמור טיוטה"
        style={{ padding: '8px 12px' }}
      >
        💾 שמור
      </button>
      {next != null && (
        <button
          className="btn btn-primary"
          style={{ flex: 1 }}
          onClick={onNext || (() => { setError(''); setStep(next); })}
        >
          {nextLabel}
        </button>
      )}
    </div>
  );

  // ═══════════════════════════════════════════════════════════════════════════
  // Render
  // ═══════════════════════════════════════════════════════════════════════════
  return (
    <div>
      {/* ── Header ── */}
      <div className="app-header">
        <div className="header-inner">
          {mode !== 'home' && (
            <button className="back-btn" onClick={goHome}>◀</button>
          )}
          <div className="header-title">
            <img
              src={`${import.meta.env.BASE_URL}logo.png`}
              alt="ניר הנדסה"
              className="header-logo-img"
              onError={e => { e.target.onerror = null; e.target.src = `${import.meta.env.BASE_URL}logo.svg`; }}
            />
            <span className="header-app-name">ניר הנדסה</span>
          </div>
          {docType && mode === 'new' && (
            <span className="header-doctype">
              {DOC_TYPES_CONFIG[docType]?.icon} {DOC_TYPES_CONFIG[docType]?.name.split('\n')[0]}
            </span>
          )}
          <span className="offline-badge">✅ אופליין</span>
        </div>
      </div>

      {/* ── Draft name modal ── */}
      {showDraftModal && (
        <DraftNameModal
          defaultName={`${form.client || 'טיוטה'} – ${DOC_TYPES_CONFIG[docType]?.name?.replace('\n',' ') || ''}`}
          onSave={saveDraftManual}
          onClose={() => setShowDraftModal(false)}
        />
      )}

      {/* ── Auto-draft banner (home only) ── */}
      {hasDraft && mode === 'home' && (
        <div className="draft-banner">
          <div className="draft-banner-text"><span>📝</span><span>נמצאה טיוטה שמורה</span></div>
          <div className="draft-banner-actions">
            <button className="btn btn-sm btn-outline" style={{ color:'white', borderColor:'rgba(255,255,255,0.5)' }} onClick={restoreAutoDraft}>שחזר</button>
            <button className="btn btn-sm" style={{ background:'rgba(255,255,255,0.2)', color:'white' }} onClick={discardAutoDraft}>מחק</button>
          </div>
        </div>
      )}

      {/* ── Alerts ── */}
      {error && <div className="app-body" style={{ paddingBottom: 0 }}><div className="alert alert-error">{error}</div></div>}
      {success && <div className="app-body" style={{ paddingBottom: 0 }}><div className="alert alert-success">{success}</div></div>}

      {/* ── Loading overlay ── */}
      {(loading || editLoading) && (
        <div className="loading-overlay">
          <div className="spinner-lg" />
          <div>{loading ? 'יוצר מסמך...' : 'טוען קובץ...'}</div>
        </div>
      )}

      {/* ── Hidden JSON import input ── */}
      <input
        ref={importRef}
        type="file"
        accept=".json"
        style={{ display: 'none' }}
        onChange={e => { if (e.target.files[0]) handleImportJSON(e.target.files[0]); e.target.value = ''; }}
      />

      {/* ══════════════════════════════════════════════════════════════════════
          HOME SCREEN
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'home' && (
        <div className="home-screen">
          <div className="home-logo">
            <img
              src={`${import.meta.env.BASE_URL}logo.png`}
              alt="ניר הנדסה"
              onError={e => { e.target.onerror = null; e.target.src = `${import.meta.env.BASE_URL}logo.svg`; }}
            />
            <span className="home-logo-name">ניר הנדסה</span>
          </div>

          <div className="home-cards">
            <div className="home-card" onClick={() => { setMode('fieldnotes'); setFnMode('list'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">📓</div>
              <div className="home-card-title">יומן שטח</div>
              <div className="home-card-sub">
                {fieldNotes.length > 0 ? `${fieldNotes.length} רשומות שמורות` : 'רשימות ותמונות מהשטח'}
              </div>
            </div>

            <div className="home-card" onClick={() => { setMode('new'); setStep(0); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">➕</div>
              <div className="home-card-title">מסמך חדש</div>
              <div className="home-card-sub">צור דוח הנדסי מאפס</div>
            </div>

            <div className="home-card" onClick={() => { setMode('edit'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">📂</div>
              <div className="home-card-title">ערוך מסמך</div>
              <div className="home-card-sub">טען קובץ Word ועדכן</div>
            </div>

            <div className="home-card" onClick={() => { setMode('drafts'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">📝</div>
              <div className="home-card-title">טיוטות</div>
              <div className="home-card-sub">
                {drafts.length > 0 ? `${drafts.length} טיוטות שמורות` : 'אין טיוטות'}
              </div>
            </div>

            <div className="home-card" onClick={() => { setMode('archive'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">📁</div>
              <div className="home-card-title">ארכיון</div>
              <div className="home-card-sub">
                {archive.length > 0 ? `${archive.length} מסמכים שנוצרו` : 'אין מסמכים'}
              </div>
            </div>
          </div>

          <div className="home-backup-row">
            <button className="btn btn-outline btn-sm" onClick={handleExportJSON}>
              ⬇️ יצא גיבוי JSON
            </button>
            <button className="btn btn-outline btn-sm" onClick={() => importRef.current?.click()}>
              ⬆️ ייבא גיבוי JSON
            </button>
          </div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          DRAFTS SCREEN
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'drafts' && (
        <div className="app-body">
          <div className="section-header">
            <span className="section-icon">📝</span>
            <span className="section-title">טיוטות שמורות</span>
          </div>

          {drafts.length === 0 ? (
            <div className="empty-state">
              <div style={{ fontSize: 40 }}>📝</div>
              <div>אין טיוטות שמורות</div>
              <div style={{ fontSize: 13, color: '#94a3b8', marginTop: 4 }}>
                בזמן יצירת מסמך לחץ "שמור" כדי לשמור טיוטה
              </div>
            </div>
          ) : (
            <div className="archive-list">
              {drafts.map(d => (
                <div key={d.id} className="archive-item">
                  <div className="archive-item-info">
                    <div className="archive-item-title">{d.name || d.form?.client || 'ללא שם'}</div>
                    <div className="archive-item-sub">
                      {DOC_TYPES_CONFIG[d.docType]?.name?.replace('\n',' ')} • {d.updated_at}
                    </div>
                  </div>
                  <div className="archive-item-actions">
                    <button className="btn btn-sm btn-primary" onClick={() => restoreDraftFromList(d)}>שחזר</button>
                    <button className="btn btn-sm btn-outline" style={{ color:'#ef4444', borderColor:'#ef4444' }} onClick={() => deleteDraft(d.id)}>מחק</button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          ARCHIVE SCREEN
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'archive' && (
        <div className="app-body">
          <div className="section-header">
            <span className="section-icon">📁</span>
            <span className="section-title">ארכיון מסמכים</span>
          </div>

          {archive.length > 0 && (
            <input
              className="form-input"
              placeholder="🔍 חיפוש לפי לקוח, מיקום, נושא..."
              value={archiveSearch}
              onChange={e => setArchiveSearch(e.target.value)}
              dir="rtl"
              style={{ marginBottom: 12 }}
            />
          )}

          {(() => {
            const q = archiveSearch.trim().toLowerCase();
            const filtered = q
              ? archive.filter(a =>
                  [a.client, a.location, a.subject, a.doc_type_name, a.inspection_date]
                    .some(f => (f || '').toLowerCase().includes(q))
                )
              : archive;
            if (archive.length === 0) return (
              <div className="empty-state">
                <div style={{ fontSize: 40 }}>📁</div>
                <div>אין מסמכים בארכיון</div>
                <div style={{ fontSize: 13, color: '#94a3b8', marginTop: 4 }}>לאחר יצירת מסמך הוא יופיע כאן</div>
              </div>
            );
            if (filtered.length === 0) return (
              <div className="empty-state">
                <div style={{ fontSize: 32 }}>🔍</div>
                <div>אין תוצאות עבור "{archiveSearch}"</div>
              </div>
            );
            return (
              <div className="archive-list">
                {filtered.map(a => (
                  <div key={a.id} className="archive-item">
                    <div className="archive-item-info">
                      <div className="archive-item-title">{a.client || '—'}</div>
                      <div className="archive-item-sub">
                        {a.doc_type_name} • {a.inspection_date} • {a.created_at}
                      </div>
                      {a.subject && <div className="archive-item-subject">{a.subject}</div>}
                      {a.location && <div className="archive-item-subject" style={{ color: '#64748b' }}>{a.location}</div>}
                    </div>
                    <div className="archive-item-actions">
                      <button className="btn btn-sm btn-outline" onClick={() => restoreFromArchive(a)} title="צור מחדש">♻️ יצור</button>
                      <button className="btn btn-sm btn-outline" style={{ color:'#ef4444', borderColor:'#ef4444' }} onClick={() => deleteArchiveEntry(a.id)}>מחק</button>
                    </div>
                  </div>
                ))}
                {q && <div style={{ textAlign: 'center', fontSize: 12, color: '#94a3b8', marginTop: 8 }}>{filtered.length} מתוך {archive.length} מסמכים</div>}
              </div>
            );
          })()}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          NEW DOCUMENT MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'new' && (
        <div className="app-body">

          {/* Step 0 — בחר סוג מסמך */}
          {step === 0 && (
            <div>
              <div className="section-header">
                <span className="section-icon">📄</span>
                <span className="section-title">בחר סוג מסמך</span>
              </div>
              <div className="doc-type-grid">
                {Object.entries(DOC_TYPES_CONFIG).map(([key, cfg]) => (
                  <div
                    key={key}
                    className={`doc-type-card${docType === key ? ' selected' : ''}`}
                    onClick={() => {
                      setDocType(key);
                      setForm(f => ({ ...f, subject: cfg.subject_default }));
                      setTableRows([]);
                      setDefectsRows([]);
                      setRowPhotos([]);
                      setDefectsRowPhotos([]);
                      setStep(1);
                    }}
                  >
                    <div className="doc-type-check">✓</div>
                    <div className="doc-type-icon">{cfg.icon}</div>
                    <div className="doc-type-name">{cfg.name.replace('\n', ' ')}</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Step 1 — פרטי הדוח */}
          {step === 1 && (
            <div>
              <StepBar step={0} onStepClick={handleStepBarClick} />

              <div className="card">
                <div className="card-title">📋 פרטי הדוח</div>

                <Field label="לקוח" required>
                  <SearchDropdown value={form.client} onChange={v => setField('client', v)} options={clientsData} placeholder="שם הלקוח..." />
                </Field>

                <Field label="גוף / ארגון">
                  <SearchDropdown value={form.organization} onChange={v => setField('organization', v)} options={KNOWN_ORGANIZATIONS} placeholder="בחר ארגון..." />
                </Field>

                <Field label="מיקום" required>
                  <input className="form-input" value={form.location} onChange={e => setField('location', e.target.value)} placeholder="שם המוסד / מיקום..." dir="rtl" />
                </Field>

                <Field label="כתובת">
                  <input className="form-input" value={form.address} onChange={e => setField('address', e.target.value)} placeholder="כתובת מלאה..." dir="rtl" />
                </Field>

                <Field label="נושא הדוח">
                  <input className="form-input" value={form.subject} onChange={e => setField('subject', e.target.value)} dir="rtl" />
                </Field>

                <Field label="תאריך הדוח">
                  <input type="date" className="form-input" value={form.date} onChange={e => setField('date', e.target.value)} />
                </Field>

                <Field label="תאריך הביקור / הבדיקה">
                  <input type="date" className="form-input" value={form.inspection_date} onChange={e => setField('inspection_date', e.target.value)} />
                </Field>

                <Field label="הכנס ליקויים נפרדים">
                  <label className="toggle-row">
                    <span className="toggle">
                      <input type="checkbox" checked={form.has_defects} onChange={e => setField('has_defects', e.target.checked)} />
                      <span className="toggle-slider" />
                    </span>
                    <span className="toggle-label">{form.has_defects ? 'כן' : 'לא'}</span>
                  </label>
                </Field>

                {/* group6: building stability section */}
                {docType === 'group6' && (
                  <>
                    <div className="card-title" style={{ marginTop: 12, marginBottom: 4, fontSize: 14 }}>🏗️ יציבות מבנה</div>
                    <Field label="נתוני יציבות המבנה" required>
                      <RichTextarea
                        value={form.building_stability_data}
                        onChange={v => setField('building_stability_data', v)}
                        placeholder="תאר את ממצאי יציבות המבנה..."
                        rows={4}
                      />
                    </Field>
                    <div style={{ display: 'flex', gap: 8 }}>
                      <Field label="סטטוס">
                        <select className="form-input" value={form.building_stability_status} onChange={e => setField('building_stability_status', e.target.value)} dir="rtl">
                          <option>תקין</option>
                          <option>לא תקין</option>
                          <option>תקין - דורש מעקב</option>
                        </select>
                      </Field>
                      <Field label="קדימות">
                        <select className="form-input" value={form.building_stability_priority} onChange={e => setField('building_stability_priority', e.target.value)} dir="rtl">
                          <option>--</option>
                          <option>1</option>
                          <option>2</option>
                          <option>3</option>
                        </select>
                      </Field>
                    </div>
                  </>
                )}

                <Field label="הערות מסכמות">
                  <RichTextarea value={form.notes_custom} onChange={v => setField('notes_custom', v)} placeholder="הערות נוספות לסיכום..." />
                </Field>
              </div>

              <StepNav back={0} backLabel="◀ סוג מסמך" next={2} />
            </div>
          )}

          {/* Step 2 — טבלת ממצאים */}
          {step === 2 && (
            <div>
              <StepBar step={1} onStepClick={handleStepBarClick} />

              <div className="card">
                <div className="card-title">📊 טבלת ממצאים ראשית</div>
                <TableEditor
                  columns={TABLE_COLUMNS[docType] || []}
                  rows={tableRows}
                  onRowsChange={setTableRows}
                  rowPhotos={rowPhotos}
                  onRowPhotosChange={setRowPhotos}
                  docType={docType}
                  elements={elementsData[docType] || []}
                  findings={findingsData[docType] || {}}
                />
              </div>

              {form.has_defects && DEFECTS_COLUMNS[docType] && (
                <div className="card">
                  <div className="card-title">⚠️ טבלת ליקויים</div>
                  <TableEditor
                    columns={DEFECTS_COLUMNS[docType]}
                    rows={defectsRows}
                    onRowsChange={setDefectsRows}
                    rowPhotos={defectsRowPhotos}
                    onRowPhotosChange={setDefectsRowPhotos}
                    docType={docType}
                    elements={elementsData[docType] || []}
                    findings={findingsData[docType] || {}}
                  />
                </div>
              )}

              <StepNav back={1} next={3} />
            </div>
          )}

          {/* Step 3 — תמונות כלליות */}
          {step === 3 && (
            <div>
              <StepBar step={2} onStepClick={handleStepBarClick} />

              <div className="card">
                <div className="card-title">🖼️ תמונות כלליות</div>
                <PhotoUpload photos={photos} onChange={setPhotos} />
              </div>

              <StepNav back={2} next={4} />
            </div>
          )}

          {/* Step 4 — יצירה */}
          {step === 4 && (
            <div>
              <StepBar step={3} onStepClick={handleStepBarClick} />

              <div className="card">
                <div className="card-title">📄 סיכום ויצירת מסמך</div>
                <div className="summary-grid">
                  <div className="summary-item"><span className="summary-label">לקוח</span><span className="summary-value">{form.client || '—'}</span></div>
                  <div className="summary-item"><span className="summary-label">מיקום</span><span className="summary-value">{form.location || '—'}</span></div>
                  <div className="summary-item"><span className="summary-label">שורות בטבלה</span><span className="summary-value">{tableRows.length}</span></div>
                  <div className="summary-item">
                    <span className="summary-label">תמונות</span>
                    <span className="summary-value">{photos.length + rowPhotos.reduce((s, arr) => s + (arr?.length || 0), 0)}</span>
                  </div>
                  <div className="summary-item"><span className="summary-label">תאריך ביקור</span><span className="summary-value">{isoToDisplay(form.inspection_date)}</span></div>
                  <div className="summary-item"><span className="summary-label">נושא</span><span className="summary-value" style={{ fontSize: 12 }}>{form.subject || '—'}</span></div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: 10, flexDirection: 'column' }}>
                <button className="btn btn-success btn-lg generate-btn" onClick={handleGenerate} disabled={loading}>
                  {loading ? '⏳ יוצר...' : '⬇️ צור והורד מסמך'}
                </button>
                <StepNav back={3} />
              </div>
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          FIELD NOTES MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'fieldnotes' && (
        <div className="app-body">
          {fnLoading && (
            <div className="loading-overlay">
              <div className="spinner-lg" />
              <div>יוצר מסמך...</div>
            </div>
          )}

          {/* List of field notes */}
          {fnMode === 'list' && (
            <div>
              <div className="section-header">
                <span className="section-icon">📓</span>
                <span className="section-title">יומן שטח</span>
              </div>

              <button
                className="btn btn-primary"
                style={{ width: '100%', marginBottom: 16 }}
                onClick={() => { setFnCurrent(defaultFn()); setFnMode('edit'); }}
              >
                ➕ רשומה חדשה
              </button>

              {fieldNotes.length === 0 ? (
                <div className="empty-state">
                  <div style={{ fontSize: 40 }}>📓</div>
                  <div>אין רשומות שמורות</div>
                  <div style={{ fontSize: 13, color: '#94a3b8', marginTop: 4 }}>
                    לחץ "רשומה חדשה" כדי להתחיל
                  </div>
                </div>
              ) : (
                <div className="archive-list">
                  {fieldNotes.map(note => (
                    <div key={note.id} className="archive-item">
                      <div className="archive-item-info">
                        <div className="archive-item-title">{note.title || note.client || 'ללא כותרת'}</div>
                        <div className="archive-item-sub">
                          {note.client && <span>{note.client} • </span>}
                          {note.created_at} • {note.updated_at}
                        </div>
                        {note.photos?.length > 0 && (
                          <div className="archive-item-subject" style={{ color: '#64748b' }}>
                            📷 {note.photos.length} תמונות
                          </div>
                        )}
                      </div>
                      <div className="archive-item-actions">
                        <button className="btn btn-sm btn-primary" onClick={() => { setFnCurrent({...note}); setFnMode('edit'); }}>ערוך</button>
                        <button className="btn btn-sm btn-outline" onClick={() => fnGenerate(note)} title="צור מסמך Word">⬇️</button>
                        <button className="btn btn-sm btn-outline" style={{ color:'#ef4444', borderColor:'#ef4444' }} onClick={() => fnDelete(note.id)}>מחק</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Edit / create a field note */}
          {fnMode === 'edit' && fnCurrent && (
            <div>
              <div className="section-header">
                <span className="section-icon">✏️</span>
                <span className="section-title">{fnCurrent.id && fieldNotes.some(n => n.id === fnCurrent.id) ? 'עריכת רשומה' : 'רשומה חדשה'}</span>
              </div>

              <div className="card">
                <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                  <input
                    className="form-input"
                    placeholder="כותרת..."
                    value={fnCurrent.title}
                    onChange={e => setFnCurrent(n => ({ ...n, title: e.target.value }))}
                    dir="rtl"
                    style={{ flex: 2 }}
                  />
                </div>
                <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
                  <SearchDropdown
                    value={fnCurrent.client}
                    onChange={v => setFnCurrent(n => {
                      const next = { ...n, client: v };
                      if (!n.location || n.location === n.client) next.location = v;
                      return next;
                    })}
                    options={clientsData}
                    placeholder="לקוח..."
                  />
                  <input
                    className="form-input"
                    placeholder="מיקום..."
                    value={fnCurrent.location}
                    onChange={e => setFnCurrent(n => ({ ...n, location: e.target.value }))}
                    dir="rtl"
                    style={{ flex: 1 }}
                  />
                </div>

                <FieldNotesEditor
                  content={fnCurrent.content}
                  onChange={v => setFnCurrent(n => ({ ...n, content: v }))}
                />
              </div>

              <div className="card">
                <div className="card-title">📷 תמונות</div>
                <PhotoUpload
                  photos={fnCurrent.photos || []}
                  onChange={photos => setFnCurrent(n => ({ ...n, photos }))}
                />
              </div>

              <div style={{ display: 'flex', gap: 8, flexDirection: 'column', marginTop: 4 }}>
                <button
                  className="btn btn-success btn-lg"
                  onClick={() => {
                    const saved = fnSave(fnCurrent);
                    setFnCurrent(saved);
                    setSuccess('✅ נשמר!');
                    setTimeout(() => setSuccess(''), 2000);
                  }}
                >
                  💾 שמור רשומה
                </button>
                <button
                  className="btn btn-primary"
                  onClick={() => fnGenerate(fnSave(fnCurrent))}
                >
                  ⬇️ שמור וצור מסמך Word
                </button>
                <button
                  className="btn btn-outline"
                  onClick={() => setFnMode('list')}
                >
                  ◀ חזור לרשימה
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          EDIT MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'edit' && (
        <div className="app-body">
          {!editImport ? (
            <div className="card" style={{ textAlign: 'center', padding: '32px 20px' }}>
              <div className="card-title">📂 טעינת מסמך Word</div>
              <div
                className="word-upload-zone"
                style={{ minHeight: 180, fontSize: 16 }}
                onClick={() => fileRef.current?.click()}
                onDragOver={e => e.preventDefault()}
                onDrop={e => {
                  e.preventDefault();
                  const f = e.dataTransfer.files[0];
                  if (f) handleEditUpload(f);
                }}
              >
                <input
                  ref={fileRef}
                  type="file"
                  accept=".docx"
                  style={{ display: 'none' }}
                  onChange={e => { if (e.target.files[0]) handleEditUpload(e.target.files[0]); e.target.value = ''; }}
                />
                {editLoading
                  ? <div className="spinner" style={{ borderTopColor: '#1a56db', borderColor: '#e2e8f0' }} />
                  : (
                    <>
                      <div style={{ fontSize: 48, marginBottom: 12 }}>📄</div>
                      <div style={{ fontWeight: 700, marginBottom: 6 }}>גרור קובץ Word לכאן</div>
                      <div style={{ fontSize: 13, color: '#94a3b8' }}>או לחץ לבחירה</div>
                    </>
                  )
                }
              </div>
              <div style={{ marginTop: 16 }}>
                <button className="btn btn-primary" onClick={() => fileRef.current?.click()} disabled={editLoading}>
                  {editLoading ? '⏳ טוען...' : '📁 בחר קובץ .docx'}
                </button>
              </div>
            </div>
          ) : (
            <div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
                <span className="edit-info-badge">📄 {editFile?.name}</span>
                {editImport.paragraphs && <span className="edit-info-badge">✏️ {editImport.paragraphs.length} פסקאות</span>}
                {editImport.images?.length > 0 && <span className="edit-info-badge">🖼️ {editImport.images.length} תמונות</span>}
                {editImport.tables?.length > 0 && <span className="edit-info-badge">📊 {editImport.tables.length} טבלאות</span>}
              </div>

              {editImport.images?.length > 0 && (
                <>
                  <div className="edit-section-title">תמונות במסמך</div>
                  <div className="edit-images-strip">
                    {editImport.images.slice(0, 8).map((img, i) => (
                      <img
                        key={i}
                        src={typeof img === 'string' ? img : `data:image/png;base64,${img.base64 || img}`}
                        alt={`תמונה ${i + 1}`}
                        className="edit-image-thumb"
                      />
                    ))}
                  </div>
                </>
              )}

              {editImport.tables?.length > 0 && (
                <>
                  <div className="edit-section-title">טבלאות (לקריאה בלבד)</div>
                  {editImport.tables.map((tbl, ti) => (
                    <div key={ti} className="edit-info-badge" style={{ marginBottom: 6 }}>
                      טבלה {ti + 1}: {tbl.length} שורות
                    </div>
                  ))}
                </>
              )}

              <div className="edit-section-title">עריכת פסקאות</div>
              <div className="edit-para-list">
                {(editImport.paragraphs || []).map((para, i) => {
                  const currentText = (editEdits.paragraphs || {})[para.id] ?? para.text;
                  const changed = currentText !== para.text;
                  return (
                    <div key={para.id || i} className="edit-para-item" style={changed ? { borderColor: '#1a56db' } : {}}>
                      <div className="edit-para-meta">
                        {para.style && <span>{para.style}</span>}
                        {para.bold && <span>מודגש</span>}
                        {para.heading && <span>כותרת</span>}
                        {para.centered && <span>מרכוז</span>}
                        {changed && <span style={{ background: '#dbeafe', color: '#1e40af' }}>שונה</span>}
                      </div>
                      <textarea
                        className="edit-para-textarea"
                        value={currentText}
                        onChange={e => setEditEdits(prev => ({
                          ...prev,
                          paragraphs: { ...(prev.paragraphs || {}), [para.id]: e.target.value },
                        }))}
                        dir="rtl"
                        style={para.bold ? { fontWeight: 700 } : {}}
                      />
                    </div>
                  );
                })}
              </div>

              <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
                <button className="btn btn-outline" onClick={() => { setEditImport(null); setEditFile(null); setEditEdits({}); }}>
                  🔄 טען קובץ אחר
                </button>
                <button className="btn btn-success" style={{ flex: 1 }} onClick={handleEditExport} disabled={editLoading}>
                  {editLoading ? '⏳ מייצא...' : '⬇️ הורד מסמך ערוך'}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
