import { useState, useRef, useEffect } from 'react';
import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Header, Footer, AlignmentType, convertMillimetersToTwip as mm,
} from 'docx';
import { DOC_TYPES_CONFIG } from '../constants';

// Exact pixel sizes derived from sample Word document (914400 EMU = 1 inch = 96px)
const HEADER_W = 359, HEADER_H = 142;   // 95.12mm × 37.57mm
const FOOTER_W = 568, FOOTER_H = 39;    // 150.50mm × 10.28mm
const FONT = 'Arial';
const SP = { line: 360, lineRule: 'auto', after: 0 };

async function fetchBuf(paths) {
  for (const url of paths) {
    try { const r = await fetch(url); if (r.ok) return new Uint8Array(await r.arrayBuffer()); } catch {}
  }
  return new Uint8Array(0);
}

function b64ToUint8(b64) {
  const raw = b64.replace(/^data:[^,]+,/, '');
  const bin = atob(raw);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

// ── Smart text parser ─────────────────────────────────────────────────────────
// Detects common Hebrew field patterns and extracts form values
function parseHebrewText(text) {
  const result = { client: '', organization: '', address: '', date: '', subject: '', location: '', notes: '', tableRows: [] };

  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  for (const line of lines) {
    const lower = line.replace(/[\s:：]/g, '').toLowerCase();

    if (/^לכבוד/.test(line))           { result.client = line.replace(/^לכבוד[\s:：]*/,'').trim(); continue; }
    if (/^שם הלקוח/.test(line))        { result.client = line.replace(/^שם הלקוח[\s:：]*/,'').trim(); continue; }
    if (/^לקוח/.test(line))            { result.client = line.replace(/^לקוח[\s:：]*/,'').trim(); continue; }
    if (/^ארגון|^גוף/.test(line))      { result.organization = line.replace(/^[^:：]+[\s:：]*/,'').trim(); continue; }
    if (/^כתובת/.test(line))           { result.address = line.replace(/^כתובת[\s:：]*/,'').trim(); continue; }
    if (/^מיקום/.test(line))           { result.location = line.replace(/^מיקום[\s:：]*/,'').trim(); continue; }
    if (/^נושא|^הנדון/.test(line))     { result.subject = line.replace(/^[^:：]+[\s:：]*/,'').trim(); continue; }
    if (/^הערות/.test(line))           { result.notes = line.replace(/^הערות[\s:：]*/,'').trim(); continue; }

    // Date patterns: dd.mm.yyyy or dd/mm/yyyy
    const dateMatch = line.match(/(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})/);
    if (dateMatch && !result.date) {
      const [, d, m, y] = dateMatch;
      const year = y.length === 2 ? '20' + y : y;
      result.date = `${year}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
      continue;
    }

    // Lines that look like defect items (numbered or bulleted)
    if (/^[\d]+[.)]\s/.test(line) || /^[-•*]\s/.test(line)) {
      const clean = line.replace(/^[\d]+[.)]\s|^[-•*]\s/, '').trim();
      if (clean.length > 3) result.tableRows.push(['', clean, '', '', '']);
    }
  }

  return result;
}

// ── Main Component ────────────────────────────────────────────────────────────
export default function SmartPaste({ onBack }) {
  const [rawText,  setRawText]  = useState('');
  const [photos,   setPhotos]   = useState([]);
  const [parsed,   setParsed]   = useState(null);
  const [docType,  setDocType]  = useState('group1');
  const [loading,  setLoading]  = useState(false);
  const [success,  setSuccess]  = useState('');
  const [step,     setStep]     = useState(0); // 0=paste, 1=review, 2=done

  const fileRef   = useRef(null);
  const cameraRef = useRef(null);

  // Global document-level paste handler: captures images pasted anywhere on the page
  // (skips text inputs and textareas so normal text paste still works there)
  useEffect(() => {
    const handler = (e) => {
      const tag = e.target?.tagName?.toLowerCase();
      if (tag === 'input' || tag === 'textarea') return;
      const imgs = Array.from(e.clipboardData?.items || []).filter(it => it.type.startsWith('image/'));
      if (!imgs.length) return;
      e.preventDefault();
      imgs.forEach(it => { const f = it.getAsFile(); if (f) handlePhotoFiles([f]); });
    };
    document.addEventListener('paste', handler);
    return () => document.removeEventListener('paste', handler);
  }, []);

  // ── Handle photo paste into the textarea ──────────────────────────────────
  const handleTextareaPaste = (e) => {
    const items = Array.from(e.clipboardData?.items || []);
    const imgItems = items.filter(it => it.type.startsWith('image/'));
    if (imgItems.length === 0) return; // let text paste work normally
    e.preventDefault();
    imgItems.forEach(it => {
      const file = it.getAsFile();
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => setPhotos(prev => [...prev, { data: ev.target.result, caption: '' }]);
      reader.readAsDataURL(file);
    });
  };

  const handlePhotoFiles = (files) => {
    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = e => setPhotos(prev => [...prev, { data: e.target.result, caption: '' }]);
      reader.readAsDataURL(file);
    });
  };

  const removePhoto = (idx) => setPhotos(prev => prev.filter((_, i) => i !== idx));
  const updateCaption = (idx, v) => setPhotos(prev => prev.map((p, i) => i === idx ? { ...p, caption: v } : p));

  // ── Parse step ────────────────────────────────────────────────────────────
  const handleParse = () => {
    const p = parseHebrewText(rawText);
    // Merge photos into parsed
    setParsed({ ...p, photos });
    setStep(1);
  };

  // ── Generate Word doc ──────────────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!parsed) return;
    setLoading(true); setSuccess('');
    try {
      const base = import.meta.env.BASE_URL || '/';
      const [headerBuf, footerBuf] = await Promise.all([
        fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
        fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
      ]);

      const mk  = (text, { size = 11, bold = false } = {}) =>
        new TextRun({ text: String(text ?? ''), font: FONT, size: size * 2, bold, rtl: true });
      const mkP = (children, align = AlignmentType.RIGHT) =>
        new Paragraph({ children, alignment: align, spacing: SP, bidirectional: true });

      const headerSection = new Header({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: headerBuf.length > 100
          ? [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })]
          : [new TextRun('')],
      })]});
      const footerSection = new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: footerBuf.length > 100
          ? [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })]
          : [new TextRun('')],
      })]});

      const cfg = DOC_TYPES_CONFIG[docType];
      const body = [
        // Header block
        mkP([mk(`לכבוד: ${parsed.client}`, { size: 8 })]),
        new Paragraph({ alignment: AlignmentType.LEFT, spacing: SP, bidirectional: true,
          children: [mk(parsed.date || new Date().toLocaleDateString('he-IL'), { size: 8 })] }),
        mkP([mk('')]),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480, after: 0 }, bidirectional: true,
          children: [mk(`הנדון: ${parsed.subject || (cfg?.subject_default ?? '')}`, { size: 15, bold: true })] }),
        mkP([mk('')]),
        ...(parsed.organization ? [mkP([mk(parsed.organization, { size: 9 })])] : []),
        ...(parsed.address ? [mkP([mk(`כתובת: ${parsed.address}`, { size: 9 })])] : []),
        ...(parsed.location  ? [mkP([mk(`מיקום: ${parsed.location}`, { size: 9 })])] : []),
        mkP([mk('')]),
        // Body text
        ...rawText.split('\n').filter(l => l.trim() && !isMetaLine(l)).map((l, i) =>
          new Paragraph({
            children: [mk(l, { size: 9 })],
            alignment: AlignmentType.RIGHT,
            spacing: i === 0 ? { ...SP, before: 360 } : SP,
            bidirectional: true,
          })
        ),
        mkP([mk('')]),
        // Notes
        ...(parsed.notes ? [mkP([mk(`הערות: ${parsed.notes}`, { size: 10 })])] : []),
        mkP([mk('')]),
        // Photos
        ...((parsed.photos || []).flatMap(ph => [
          new Paragraph({ alignment: AlignmentType.CENTER,
            children: [new ImageRun({ data: b64ToUint8(ph.data), transformation: { width: 400, height: 300 } })],
          }),
          ...(ph.caption ? [mkP([mk(ph.caption, { size: 9 })], AlignmentType.CENTER)] : []),
          mkP([mk('')]),
        ])),
      ];

      const doc = new Document({
        styles: { paragraphStyles: [{
          id: 'Normal', name: 'Normal', quickFormat: true,
          paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT },
          run: { font: { name: FONT } },
        }]},
        sections: [{
          properties: {
            page: {
              size: { width: mm(210), height: mm(297) },
              margin: { top: mm(31.70), bottom: mm(12.51), left: mm(7.00), right: mm(7.00), header: mm(3.00), footer: mm(1.99) },
            },
            bidi: true,
          },
          headers: { default: headerSection },
          footers: { default: footerSection },
          children: body,
        }],
      });

      const blob = await Packer.toBlob(doc);
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url;
      a.download = `${parsed.subject || parsed.client || 'מסמך'}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      setSuccess('✅ המסמך נוצר בהצלחה!');
      setStep(2);
    } catch (e) {
      setSuccess(`שגיאה: ${e.message}`);
    } finally {
      setLoading(false);
    }
  };

  const reset = () => { setRawText(''); setPhotos([]); setParsed(null); setStep(0); setSuccess(''); };

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="app-body">
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="field-journal-title">⚡ הדבק וייצא</span>
        {step > 0 && <button className="btn btn-outline btn-sm" onClick={reset}>✕ נקה</button>}
      </div>

      {/* ── STEP 0: Paste ── */}
      {step === 0 && (
        <div>
          <div className="card">
            <div className="card-title">📋 הדבק טקסט</div>
            <p style={{ fontSize: 13, color: '#64748b', marginBottom: 10, direction: 'rtl' }}>
              הדבק טקסט מכל מקור — אימייל, הערות שטח, מסמך קודם. המערכת תזהה אוטומטית שמות, כתובות, תאריכים וליקויים.
            </p>
            <textarea
              className="form-textarea"
              style={{ minHeight: 220, fontSize: 15, lineHeight: 1.8 }}
              value={rawText}
              onChange={e => setRawText(e.target.value)}
              onPaste={handleTextareaPaste}
              placeholder="הדבק כאן את הטקסט שלך (Ctrl+V)..."
              dir="rtl"
            />
          </div>

          <div className="card">
            <div className="card-title">📷 תמונות</div>
            <p style={{ fontSize: 13, color: '#64748b', marginBottom: 10, direction: 'rtl' }}>
              הדבק תמונות בתוך תיבת הטקסט (Ctrl+V), גרור לכאן, או בחר מהגלריה.
            </p>
            <div className="sp-photo-drop"
              onDrop={e => { e.preventDefault(); handlePhotoFiles(e.dataTransfer.files); }}
              onDragOver={e => e.preventDefault()}
            >
              {photos.length === 0
                ? <span style={{ color: '#94a3b8' }}>גרור תמונות לכאן</span>
                : (
                  <div className="journal-photos">
                    {photos.map((ph, idx) => (
                      <div key={idx} className="journal-photo">
                        <img src={ph.data} alt={`תמונה ${idx+1}`} />
                        <button className="journal-photo-remove" onClick={() => removePhoto(idx)}>✕</button>
                        <input className="journal-photo-caption" value={ph.caption}
                          onChange={e => updateCaption(idx, e.target.value)}
                          placeholder="כיתוב..." dir="rtl" />
                      </div>
                    ))}
                  </div>
                )
              }
            </div>
            <div style={{ display: 'flex', gap: 8, marginTop: 10, flexWrap: 'wrap' }}>
              <button className="btn btn-outline" onClick={() => fileRef.current?.click()}>📎 בחר קובץ</button>
              <button className="btn btn-outline" onClick={() => cameraRef.current?.click()}>📸 צלם</button>
            </div>
            <input ref={fileRef} type="file" accept="image/*" multiple style={{ display: 'none' }}
              onChange={e => { handlePhotoFiles(e.target.files); e.target.value = ''; }} />
            <input ref={cameraRef} type="file" accept="image/*" capture="environment" style={{ display: 'none' }}
              onChange={e => { handlePhotoFiles(e.target.files); e.target.value = ''; }} />
          </div>

          <div className="card">
            <div className="card-title">📄 סוג מסמך</div>
            <div className="doc-type-grid">
              {Object.entries(DOC_TYPES_CONFIG).map(([key, cfg]) => (
                <div key={key}
                  className={`doc-type-card${docType === key ? ' selected' : ''}`}
                  onClick={() => setDocType(key)}
                >
                  <div className="doc-type-check">✓</div>
                  <div className="doc-type-icon">{cfg.icon}</div>
                  <div className="doc-type-name">{cfg.name.replace('\n',' ')}</div>
                </div>
              ))}
            </div>
          </div>

          <button className="btn btn-primary btn-lg" style={{ marginBottom: 20 }}
            onClick={handleParse} disabled={!rawText.trim() && photos.length === 0}>
            ⚡ ארגן אוטומטית ▶
          </button>
        </div>
      )}

      {/* ── STEP 1: Review parsed data ── */}
      {step === 1 && parsed && (
        <div>
          <div className="card">
            <div className="card-title">✅ נתונים שזוהו — ניתן לערוך לפני הייצוא</div>

            {[
              { label: 'לקוח / מוסד', key: 'client' },
              { label: 'ארגון', key: 'organization' },
              { label: 'כתובת', key: 'address' },
              { label: 'מיקום', key: 'location' },
              { label: 'נושא', key: 'subject' },
              { label: 'הערות', key: 'notes' },
            ].map(({ label, key }) => (
              <div className="form-group" key={key}>
                <label className="form-label">{label}</label>
                <input className="form-input"
                  value={parsed[key] || ''}
                  onChange={e => setParsed(p => ({ ...p, [key]: e.target.value }))}
                  dir="rtl"
                  placeholder={`${label}...`}
                />
              </div>
            ))}

            <div className="form-group">
              <label className="form-label">תאריך</label>
              <input type="date" className="form-input"
                value={parsed.date || ''}
                onChange={e => setParsed(p => ({ ...p, date: e.target.value }))}
              />
            </div>
          </div>

          {(parsed.tableRows || []).length > 0 && (
            <div className="card">
              <div className="card-title">📋 ממצאים שזוהו ({parsed.tableRows.length})</div>
              {parsed.tableRows.map((row, i) => (
                <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'center', direction: 'rtl' }}>
                  <span style={{ width: 20, color: '#64748b', flexShrink: 0 }}>{i+1}.</span>
                  <input className="form-input"
                    value={row[1] || ''}
                    onChange={e => setParsed(p => {
                      const rows = [...p.tableRows];
                      rows[i] = [...rows[i]]; rows[i][1] = e.target.value;
                      return { ...p, tableRows: rows };
                    })}
                    dir="rtl"
                  />
                  <button className="btn-icon del"
                    onClick={() => setParsed(p => ({ ...p, tableRows: p.tableRows.filter((_, j) => j !== i) }))}>
                    🗑
                  </button>
                </div>
              ))}
            </div>
          )}

          {photos.length > 0 && (
            <div className="card">
              <div className="card-title">📷 תמונות ({photos.length})</div>
              <div className="journal-photos">
                {photos.map((ph, idx) => (
                  <div key={idx} className="journal-photo">
                    <img src={ph.data} alt={`תמונה ${idx+1}`} />
                    <button className="journal-photo-remove" onClick={() => removePhoto(idx)}>✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {success && <div className={`alert ${success.startsWith('שגיאה') ? 'alert-error' : 'alert-success'}`}>{success}</div>}

          <div style={{ display: 'flex', gap: 10, marginBottom: 20 }}>
            <button className="btn btn-outline" onClick={() => setStep(0)}>◀ חזור</button>
            <button className="btn btn-success btn-lg generate-btn" style={{ flex: 1 }}
              onClick={handleGenerate} disabled={loading}>
              {loading ? '⏳ יוצר מסמך...' : '⬇️ צור מסמך Word'}
            </button>
          </div>
        </div>
      )}

      {/* ── STEP 2: Done ── */}
      {step === 2 && (
        <div style={{ textAlign: 'center', padding: '40px 20px', direction: 'rtl' }}>
          <div style={{ fontSize: 56, marginBottom: 16 }}>✅</div>
          <p style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>המסמך נוצר בהצלחה!</p>
          <p style={{ color: '#64748b', marginBottom: 24 }}>הקובץ הורד למכשירך</p>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
            <button className="btn btn-primary" onClick={reset}>+ מסמך חדש</button>
            <button className="btn btn-outline" onClick={onBack}>חזור לבית</button>
          </div>
        </div>
      )}
    </div>
  );
}

// Lines that are "meta" (known labels) — skip them from raw body lines
function isMetaLine(line) {
  return /^(לכבוד|שם הלקוח|לקוח|ארגון|גוף|כתובת|מיקום|נושא|הנדון|הערות)[\s:：]/.test(line);
}
