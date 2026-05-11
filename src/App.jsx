import { useState, useEffect, useRef } from 'react';
import TableEditor from './components/TableEditor';
import PhotoUpload from './components/PhotoUpload';
import SearchDropdown from './components/SearchDropdown';
import FieldJournal from './components/FieldJournal';
import SmartPaste from './components/SmartPaste';
import { DOC_TYPES_CONFIG, TABLE_COLUMNS, DEFECTS_COLUMNS, KNOWN_ORGANIZATIONS, DRAFT_KEY } from './constants';
import { generateDocument } from './lib/docGenerator';
import { generateEventApproval } from './lib/eventApproval';
import { importDocx, exportDocx } from './lib/docImporter';
import elementsData from './data/elements_by_type.json';
import findingsData from './data/findings_by_type.json';
import clientsData from './data/clients.json';
import './index.css';

// ── Pre-processed global lists (computed once at module load) ─────────────────

/** All client names, excluding pure-number entries, sorted Hebrew-first */
const ALL_CLIENTS = clientsData
  .filter(c => !/^\d+$/.test(c.trim()))
  .sort((a, b) => a.localeCompare(b, 'he'));

/** All elements merged across ALL doc types, deduplicated, excluding pure-number entries */
const ALL_ELEMENTS = [...new Set(
  Object.values(elementsData).flat().filter(e => !/^\d+$/.test(e.trim()))
)].sort((a, b) => a.localeCompare(b, 'he'));

/** All findings merged across ALL doc types, keyed by element name */
const ALL_FINDINGS = (() => {
  const merged = {};
  Object.values(findingsData).forEach(typeFindings => {
    Object.entries(typeFindings).forEach(([key, arr]) => {
      if (!merged[key]) merged[key] = new Set();
      arr.forEach(f => merged[key].add(f));
    });
  });
  return Object.fromEntries(Object.entries(merged).map(([k, v]) => [k, [...v]]));
})();

// ── Asset paths (must include Vite base URL for GitHub Pages) ─────────────────
const LOGO_PNG = `${import.meta.env.BASE_URL}logo.png`;
const LOGO_SVG = `${import.meta.env.BASE_URL}logo.svg`;

// ── Helpers ──────────────────────────────────────────────────────────────────

const isoToDisplay = (iso) => {
  if (!iso || !iso.includes('-')) return iso || '';
  const [y, m, d] = iso.split('-');
  return `${parseInt(d)}.${parseInt(m)}.${y}`;
};

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
});

const defaultEventForm = () => ({
  to: '',
  date: new Date().toISOString().split('T')[0],
  address: '',
  owner: '',
  id_num: '',
  phone: '',
  structures: '',
  notes: '',
  validity: '',
});

// ── StepBar ───────────────────────────────────────────────────────────────────
function StepBar({ step, labels }) {
  return (
    <div className="stepbar">
      {labels.map((l, i) => (
        <div key={i} className={`stepbar-item${i === step ? ' active' : i < step ? ' done' : ''}`}>
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

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [mode, setMode] = useState('home');   // 'home' | 'new' | 'edit' | 'field' | 'event' | 'smartpaste'
  const [step, setStep] = useState(0);         // for 'new': 0=type,1=details,2=table,3=photos,4=generate
  const [docType, setDocType] = useState('');
  const [form, setForm] = useState(defaultForm());
  const [tableRows, setTableRows] = useState([]);
  const [defectsRows, setDefectsRows] = useState([]);
  const [rowPhotos, setRowPhotos] = useState([]);
  const [defectsRowPhotos, setDefectsRowPhotos] = useState([]);
  const [photos, setPhotos] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [lastBlob, setLastBlob] = useState(null);
  const [hasDraft, setHasDraft] = useState(false);

  // Edit mode state
  const [editImport, setEditImport] = useState(null);
  const [editEdits, setEditEdits] = useState({});
  const [editLoading, setEditLoading] = useState(false);

  // Event approval state
  const [eventForm, setEventForm] = useState(defaultEventForm());
  const setEventField = (key, val) => setEventForm(f => ({ ...f, [key]: val }));
  const [editFile, setEditFile] = useState(null);
  const fileRef = useRef(null);

  // ── Draft ──────────────────────────────────────────────────────────────────
  useEffect(() => {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (raw) {
        const d = JSON.parse(raw);
        if (d && d.mode === 'new') setHasDraft(true);
      }
    } catch (_) {}
  }, []);

  const saveDraft = () => {
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify({
        mode: 'new', step, docType, form,
        tableRows, defectsRows,
        rowPhotos: [], defectsRowPhotos: [], photos: [],
      }));
    } catch (_) {}
  };

  const restoreDraft = () => {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (!raw) return;
      const d = JSON.parse(raw);
      if (d.docType) setDocType(d.docType);
      if (d.form) setForm(d.form);
      if (d.tableRows) setTableRows(d.tableRows);
      if (d.defectsRows) setDefectsRows(d.defectsRows);
      setRowPhotos([]); setDefectsRowPhotos([]); setPhotos([]);
      setStep(d.step || 1);
      setMode('new');
      setHasDraft(false);
    } catch (_) {}
  };

  const discardDraft = () => {
    localStorage.removeItem(DRAFT_KEY);
    setHasDraft(false);
  };

  // Auto-save draft whenever form state changes in new mode
  useEffect(() => {
    if (mode === 'new' && docType) saveDraft();
  }, [mode, step, docType, form, tableRows, defectsRows]);

  // ── Event approval generate ───────────────────────────────────────────────
  const handleEventGenerate = async () => {
    if (!eventForm.to.trim()) { setError('יש למלא את שדה "לכבוד"'); return; }
    setLoading(true); setError(''); setSuccess('');
    try {
      const blob = await generateEventApproval({
        ...eventForm,
        date: isoToDisplay(eventForm.date),
      });
      setLastBlob(blob);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `אישור_אירוע_${eventForm.to || 'מסמך'}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      setSuccess('המסמך נוצר בהצלחה!');
    } catch (e) {
      setError(`שגיאה: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  // ── Auto-detect defects from main table rows ──────────────────────────────
  // Recalculate whenever tableRows changes
  const autoHasDefects = tableRows.some(row =>
    row.some(cell => String(cell || '').startsWith('לא תקין'))
  );

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

      // has_defects: auto-detected OR manually forced on
      const hasDefects = autoHasDefects || form.has_defects;

      const payload = {
        doc_type: docType,
        ...form,
        has_defects: hasDefects,
        inspection_date: isoToDisplay(form.inspection_date),
        table_rows: tableRows,
        defects_rows: defectsRows,
        photos: mergedPhotos,
        notes_custom: typeof form.notes_custom === 'string'
          ? form.notes_custom.split('\n').filter(s => s.trim())
          : (form.notes_custom || []),
      };

      const blob = await generateDocument(payload);
      setLastBlob(blob);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${form.subject || 'דוח'} - ${form.client}.docx`;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setSuccess('✅ המסמך נוצר בהצלחה!');
      localStorage.removeItem(DRAFT_KEY);
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
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
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

  // ── Share ──────────────────────────────────────────────────────────────────
  const handleShare = async () => {
    if (!lastBlob || !navigator.share) return;
    try {
      const fileName = `${form.subject || 'דוח'} - ${form.client}.docx`;
      await navigator.share({
        title: fileName,
        files: [new File([lastBlob], fileName, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })],
      });
    } catch (_) {}
  };

  // ── Validation ─────────────────────────────────────────────────────────────
  const validateDetails = () => {
    if (!form.client.trim()) return setError('יש למלא שם לקוח'), false;
    setError('');
    return true;
  };

  const setField = (key, val) => setForm(f => ({ ...f, [key]: val }));

  // ── Reset ──────────────────────────────────────────────────────────────────
  const goHome = () => {
    setMode('home'); setStep(0); setError(''); setSuccess('');
    setDocType(''); setEditImport(null); setEditFile(null); setEditEdits({});
  };

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
              src={LOGO_PNG}
              alt="ניר הנדסה"
              className="header-logo-img"
              onError={e => { e.target.onerror = null; e.target.src = LOGO_SVG; }}
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

      {/* ── Draft banner ── */}
      {hasDraft && mode === 'home' && (
        <div className="draft-banner">
          <div className="draft-banner-text">
            <span>📝</span>
            <span>נמצאה טיוטה שמורה</span>
          </div>
          <div className="draft-banner-actions">
            <button className="btn btn-sm btn-outline" style={{ color: 'white', borderColor: 'rgba(255,255,255,0.5)' }} onClick={restoreDraft}>
              שחזר טיוטה
            </button>
            <button className="btn btn-sm" style={{ background: 'rgba(255,255,255,0.2)', color: 'white' }} onClick={discardDraft}>
              מחק
            </button>
          </div>
        </div>
      )}

      {/* ── Alerts ── */}
      {error && (
        <div className="app-body" style={{ paddingBottom: 0 }}>
          <div className="alert alert-error">{error}</div>
        </div>
      )}
      {success && (
        <div className="app-body" style={{ paddingBottom: 0 }}>
          <div className="alert alert-success" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
            <span>{success}</span>
            {lastBlob && typeof navigator !== 'undefined' && navigator.share && navigator.canShare && (
              <button className="btn btn-sm btn-outline" style={{ flexShrink: 0 }} onClick={handleShare}>
                📤 שתף
              </button>
            )}
          </div>
        </div>
      )}

      {/* ── Loading overlay ── */}
      {(loading || editLoading) && (
        <div className="loading-overlay">
          <div className="spinner-lg" />
          <div>{loading ? 'יוצר מסמך...' : 'טוען קובץ...'}</div>
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          HOME SCREEN
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'home' && (
        <div className="home-screen">
          <div className="home-logo">
            <img src={LOGO_PNG} alt="ניר הנדסה" onError={e => { e.target.onerror = null; e.target.src = LOGO_SVG; }} />
            <span className="home-logo-name">ניר הנדסה</span>
          </div>

          <div className="home-cards">
            <div className="home-card" onClick={() => { setMode('new'); setStep(0); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">➕</div>
              <div className="home-card-title">מסמך חדש</div>
              <div className="home-card-sub">צור דוח הנדסי מאפס עם טפסים מובנים</div>
            </div>

            <div className="home-card" onClick={() => { setMode('edit'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">📂</div>
              <div className="home-card-title">ערוך מסמך קיים</div>
              <div className="home-card-sub">טען קובץ Word ועדכן את התוכן שלו</div>
            </div>

            <div className="home-card" onClick={() => { setEventForm(defaultEventForm()); setMode('event'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">🎪</div>
              <div className="home-card-title">אישור לאירוע</div>
              <div className="home-card-sub">אישור מבנים ומתקנים ארעיים/קבועים</div>
            </div>

            <div className="home-card" onClick={() => setMode('field')}>
              <div className="home-card-icon">📋</div>
              <div className="home-card-title">יומן שטח</div>
              <div className="home-card-sub">רשום הערות בשטח – נשמר אוטומטית</div>
            </div>

            <div className="home-card" onClick={() => { setMode('smartpaste'); setError(''); setSuccess(''); }}>
              <div className="home-card-icon">⚡</div>
              <div className="home-card-title">הדבק וייצא</div>
              <div className="home-card-sub">הדבק טקסט ותמונות – יוצר מסמך מסודר אוטומטית</div>
            </div>
          </div>
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
              <StepBar step={0} labels={['פרטים', 'טבלה', 'תמונות', 'יצור']} />

              <div className="card">
                <div className="card-title">📋 פרטי הדוח</div>

                <Field label="לקוח" required>
                  <SearchDropdown
                    value={form.client}
                    onChange={v => setField('client', v)}
                    options={ALL_CLIENTS}
                    placeholder="שם הלקוח..."
                  />
                </Field>

                <Field label="גוף / ארגון">
                  <SearchDropdown
                    value={form.organization}
                    onChange={v => setField('organization', v)}
                    options={KNOWN_ORGANIZATIONS}
                    placeholder="בחר ארגון..."
                  />
                </Field>

                <Field label="מיקום">
                  <input
                    className="form-input"
                    value={form.location}
                    onChange={e => setField('location', e.target.value)}
                    placeholder="שם המוסד / מיקום (ריק = שם הלקוח)..."
                    dir="rtl"
                  />
                </Field>

                <Field label="כתובת">
                  <input
                    className="form-input"
                    value={form.address}
                    onChange={e => setField('address', e.target.value)}
                    placeholder="כתובת מלאה..."
                    dir="rtl"
                  />
                </Field>

                <Field label="נושא הדוח">
                  <input
                    className="form-input"
                    value={form.subject}
                    onChange={e => setField('subject', e.target.value)}
                    dir="rtl"
                  />
                </Field>

                <Field label="תאריך הדוח">
                  <input
                    type="date"
                    className="form-input"
                    value={form.date}
                    onChange={e => setField('date', e.target.value)}
                  />
                </Field>

                <Field label="תאריך הביקור / הבדיקה">
                  <input
                    type="date"
                    className="form-input"
                    value={form.inspection_date}
                    onChange={e => setField('inspection_date', e.target.value)}
                  />
                </Field>

                <Field label='טבלת ליקויים נפרדת'>
                  <label className="toggle-row">
                    <span className="toggle">
                      <input
                        type="checkbox"
                        checked={form.has_defects}
                        onChange={e => setField('has_defects', e.target.checked)}
                      />
                      <span className="toggle-slider" />
                    </span>
                    <span className="toggle-label" style={{ fontSize: 13, color: '#64748b' }}>
                      {form.has_defects ? 'כן – תוצג בשלב הטבלה' : 'לא (מזוהה אוטומטית מהטבלה)'}
                    </span>
                  </label>
                </Field>
              </div>

              <div style={{ display: 'flex', gap: 10 }}>
                <button className="btn btn-outline" onClick={() => setStep(0)}>◀ חזור</button>
                <button
                  className="btn btn-primary"
                  style={{ flex: 1 }}
                  onClick={() => { if (validateDetails()) setStep(2); }}
                >
                  המשך ▶
                </button>
              </div>
            </div>
          )}

          {/* Step 2 — טבלת ממצאים */}
          {step === 2 && (
            <div>
              <StepBar step={1} labels={['פרטים', 'טבלה', 'תמונות', 'יצור']} />

              <div className="card">
                <div className="card-title">📊 טבלת ממצאים ראשית</div>
                <TableEditor
                  columns={TABLE_COLUMNS[docType] || []}
                  rows={tableRows}
                  onRowsChange={setTableRows}
                  rowPhotos={rowPhotos}
                  onRowPhotosChange={setRowPhotos}
                  docType={docType}
                  elements={ALL_ELEMENTS}
                  findings={ALL_FINDINGS}
                />
              </div>

              {/* Defects table – shown when auto-detected OR manually enabled */}
              {(autoHasDefects || form.has_defects) && DEFECTS_COLUMNS[docType] && (
                <div className="card">
                  <div className="card-title">
                    ⚠️ טבלת ליקויים
                    {autoHasDefects && (
                      <span style={{ fontSize: 11, fontWeight: 400, color: '#c55a11', marginRight: 8 }}>
                        (זוהתה אוטומטית)
                      </span>
                    )}
                  </div>
                  <TableEditor
                    columns={DEFECTS_COLUMNS[docType]}
                    rows={defectsRows}
                    onRowsChange={setDefectsRows}
                    rowPhotos={defectsRowPhotos}
                    onRowPhotosChange={setDefectsRowPhotos}
                    docType={docType}
                    elements={ALL_ELEMENTS}
                    findings={ALL_FINDINGS}
                  />
                </div>
              )}

              {/* Notes & conclusions quick-edit */}
              <div className="card">
                <div className="card-title">📝 הערות ומסקנות</div>
                <Field label="הערות והנחיות (אופציונלי)">
                  <RichTextarea
                    value={form.notes_custom}
                    onChange={v => setField('notes_custom', v)}
                    placeholder="השאר ריק לשימוש בהערות ברירת המחדל..."
                    rows={3}
                  />
                </Field>
                <Field label="מסקנות (אופציונלי)">
                  <textarea
                    className="form-textarea"
                    value={form.conclusion_custom}
                    onChange={e => setField('conclusion_custom', e.target.value)}
                    placeholder={autoHasDefects || form.has_defects ? 'ברירת מחדל: נמצאו ליקויים הדורשים טיפול...' : 'ברירת מחדל: הכל תקין...'}
                    rows={2}
                    dir="rtl"
                  />
                </Field>
              </div>

              <div style={{ display: 'flex', gap: 10, flexDirection: 'column' }}>
                <button
                  className="btn btn-success btn-lg generate-btn"
                  onClick={handleGenerate}
                  disabled={loading}
                >
                  {loading ? '⏳ יוצר...' : '⬇️ צור מסמך ללא תמונות'}
                </button>
                <div style={{ display: 'flex', gap: 10 }}>
                  <button className="btn btn-outline" onClick={() => setStep(1)}>◀ חזור</button>
                  <button className="btn btn-primary" style={{ flex: 1 }} onClick={() => setStep(3)}>
                    המשך עם תמונות ▶
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Step 3 — תמונות כלליות */}
          {step === 3 && (
            <div>
              <StepBar step={2} labels={['פרטים', 'טבלה', 'תמונות', 'יצור']} />

              <div className="card">
                <div className="card-title">🖼️ תמונות כלליות</div>
                <PhotoUpload photos={photos} onChange={setPhotos} />
              </div>

              <div style={{ display: 'flex', gap: 10, flexDirection: 'column' }}>
                <button
                  className="btn btn-success btn-lg generate-btn"
                  onClick={handleGenerate}
                  disabled={loading}
                >
                  {loading ? '⏳ יוצר...' : '⬇️ צור והורד מסמך'}
                </button>
                <button className="btn btn-outline" onClick={() => setStep(2)}>◀ חזור לטבלה</button>
              </div>
            </div>
          )}

          {/* Step 4 — יצירה */}
          {step === 4 && (
            <div>
              <StepBar step={3} labels={['פרטים', 'טבלה', 'תמונות', 'יצור']} />

              <div className="card">
                <div className="card-title">📄 סיכום</div>
                <div className="summary-grid">
                  <div className="summary-item">
                    <span className="summary-label">לקוח</span>
                    <span className="summary-value">{form.client || '—'}</span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">מיקום</span>
                    <span className="summary-value">{form.location || '—'}</span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">שורות בטבלה</span>
                    <span className="summary-value">{tableRows.length}</span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">תמונות</span>
                    <span className="summary-value">
                      {photos.length + rowPhotos.reduce((s, arr) => s + (arr?.length || 0), 0)}
                    </span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">ליקויים</span>
                    <span className="summary-value" style={{ color: (autoHasDefects || form.has_defects) ? '#c00000' : '#059669' }}>
                      {(autoHasDefects || form.has_defects) ? `כן (${defectsRows.length} שורות)` : 'לא נמצאו'}
                    </span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">תאריך ביקור</span>
                    <span className="summary-value">{isoToDisplay(form.inspection_date)}</span>
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: 10, flexDirection: 'column' }}>
                <button
                  className="btn btn-success btn-lg generate-btn"
                  onClick={handleGenerate}
                  disabled={loading}
                >
                  {loading ? '⏳ יוצר...' : '⬇️ צור והורד מסמך'}
                </button>
                <button className="btn btn-outline" onClick={() => setStep(3)}>◀ חזור</button>
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
            /* ── Upload zone ── */
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
                <button
                  className="btn btn-primary"
                  onClick={() => fileRef.current?.click()}
                  disabled={editLoading}
                >
                  {editLoading ? '⏳ טוען...' : '📁 בחר קובץ .docx'}
                </button>
              </div>
            </div>
          ) : (
            /* ── Edit paragraphs ── */
            <div>
              {/* File info */}
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 16 }}>
                <span className="edit-info-badge">📄 {editFile?.name}</span>
                {editImport.paragraphs && (
                  <span className="edit-info-badge">
                    ✏️ {editImport.paragraphs.length} פסקאות
                  </span>
                )}
                {editImport.images && editImport.images.length > 0 && (
                  <span className="edit-info-badge">🖼️ {editImport.images.length} תמונות</span>
                )}
                {editImport.tables && editImport.tables.length > 0 && (
                  <span className="edit-info-badge">📊 {editImport.tables.length} טבלאות</span>
                )}
              </div>

              {/* Images strip */}
              {editImport.images && editImport.images.length > 0 && (
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

              {/* Tables preview */}
              {editImport.tables && editImport.tables.length > 0 && (
                <>
                  <div className="edit-section-title">טבלאות (לקריאה בלבד)</div>
                  {editImport.tables.map((tbl, ti) => (
                    <div key={ti} className="edit-info-badge" style={{ marginBottom: 6 }}>
                      טבלה {ti + 1}: {tbl.length} שורות
                    </div>
                  ))}
                </>
              )}

              {/* Paragraphs */}
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

              {/* Actions */}
              <div style={{ display: 'flex', gap: 10, marginTop: 20 }}>
                <button
                  className="btn btn-outline"
                  onClick={() => { setEditImport(null); setEditFile(null); setEditEdits({}); }}
                >
                  🔄 טען קובץ אחר
                </button>
                <button
                  className="btn btn-success"
                  style={{ flex: 1 }}
                  onClick={handleEditExport}
                  disabled={editLoading}
                >
                  {editLoading ? '⏳ מייצא...' : '⬇️ הורד מסמך ערוך'}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          EVENT APPROVAL MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'event' && (
        <div className="app-body">
          <div className="section-header">
            <button className="btn btn-outline btn-sm" style={{ marginLeft: 12 }} onClick={() => setMode('home')}>◀ חזור</button>
            <span className="section-icon">🎪</span>
            <span className="section-title">אישור מבנים ומתקנים ארעיים/קבועים</span>
          </div>

          <div className="card">
            <div className="card-title">👤 פרטי הלקוח</div>

            <Field label="לכבוד" required>
              <input className="form-input" value={eventForm.to} onChange={e => setEventField('to', e.target.value)} placeholder="שם הגוף / הלקוח..." dir="rtl" />
            </Field>
            <Field label="תאריך המסמך">
              <input type="date" className="form-input" value={eventForm.date} onChange={e => setEventField('date', e.target.value)} />
            </Field>
            <Field label="כתובת העסק">
              <input className="form-input" value={eventForm.address} onChange={e => setEventField('address', e.target.value)} placeholder="כתובת מלאה..." dir="rtl" />
            </Field>
            <Field label="שם בעל העסק / המזמין">
              <input className="form-input" value={eventForm.owner} onChange={e => setEventField('owner', e.target.value)} placeholder="שם מלא..." dir="rtl" />
            </Field>
            <Field label="מספר ת.זהות">
              <input className="form-input" value={eventForm.id_num} onChange={e => setEventField('id_num', e.target.value)} placeholder="מספר ת.ז..." dir="rtl" />
            </Field>
            <Field label="טלפון סלולארי">
              <input className="form-input" value={eventForm.phone} onChange={e => setEventField('phone', e.target.value)} placeholder="050-0000000" dir="rtl" />
            </Field>
          </div>

          <div className="card">
            <div className="card-title">🏗️ מתקנים שנבדקו</div>
            <Field label="רשום מתקן אחד בכל שורה">
              <textarea
                className="form-textarea event-structures-textarea"
                value={eventForm.structures}
                onChange={e => setEventField('structures', e.target.value)}
                placeholder={'לדוגמה:\nבמה מוגבהת 10x6 מ\'\nגגון צללית 5x5 מ\'\nמאהל 3x3 מ\''}
                rows={6}
                dir="rtl"
              />
            </Field>
          </div>

          <div className="card">
            <div className="card-title">📝 הערות ותוקף</div>
            <Field label="הערות נוספות">
              <input className="form-input" value={eventForm.notes} onChange={e => setEventField('notes', e.target.value)} placeholder="הערות נוספות (אופציונלי)..." dir="rtl" />
            </Field>
            <Field label="תוקף האישור">
              <input className="form-input" value={eventForm.validity} onChange={e => setEventField('validity', e.target.value)} placeholder="לדוגמה: 31.12.2025" dir="rtl" />
            </Field>
          </div>

          {error && <div className="error-msg">{error}</div>}
          {success && <div className="success-msg">{success}</div>}

          <button
            className="btn btn-success btn-lg generate-btn"
            onClick={handleEventGenerate}
            disabled={loading}
          >
            {loading ? '⏳ יוצר מסמך...' : '⬇️ צור אישור לאירוע'}
          </button>
        </div>
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          FIELD JOURNAL MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'field' && (
        <FieldJournal onBack={goHome} />
      )}

      {/* ══════════════════════════════════════════════════════════════════════
          SMART PASTE MODE
      ══════════════════════════════════════════════════════════════════════ */}
      {mode === 'smartpaste' && (
        <SmartPaste onBack={goHome} />
      )}
    </div>
  );
}
