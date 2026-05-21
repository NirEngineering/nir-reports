import { useState, useRef, useEffect } from 'react';
import { DOC_TYPES_CONFIG, TABLE_COLUMNS, DEFECTS_COLUMNS } from '../constants';
import { generateDocument } from '../lib/docGenerator';
import { saveToArchive } from '../lib/archiveUtils';

const API_KEY_STORAGE = 'nir_anthropic_key';
const MODEL = 'claude-sonnet-4-6';

// ── System prompt ─────────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `אתה מנתח דוחות הנדסיים מומחה לחברת ניר הנדסה, חברה לייעוץ הנדסי בישראל.

תפקידך: לקרוא בעיון את כל הטקסט והתמונות שמסופקות, ולסווג את המידע לפי הקטגוריות הנכונות.

━━━ כללי סיווג ━━━
• נתונים כלליים (data_section) — עובדות, נתוני רקע על האתר/המבנה: גיל, חומרים, מידות, שימוש, תיאור כללי, מי ביקש את הבדיקה
• ממצאים (findings) — ליקוי/תקלה/חריגה שנצפתה בבדיקה — ממצא אחד לכל ליקוי, לא לאחד
• מסקנות (conclusions) — הכרעות מקצועיות סופיות: האם המבנה תקין? מה נדרש לביצוע?
• הערות (notes) — הנחיות, מגבלות, תנאים, תוקף האישור, אחריות

━━━ JSON נדרש ━━━
החזר JSON תקני בלבד (ללא markdown, ללא \`\`\`json, רק JSON נקי):
{
  "client": "שם הלקוח/המוסד",
  "organization": "שם הארגון (אם קיים)",
  "location": "שם המיקום/הנכס",
  "address": "כתובת מלאה",
  "date": "YYYY-MM-DD",
  "inspection_date": "YYYY-MM-DD",
  "subject": "נושא הדוח בקצרה (עד 8 מילים)",
  "intro": "פסקת פתיחה — תאריך, מיקום, מטרת הביקור, מי ביצע",
  "data_section": [
    "נתון כללי 1 על האתר/המבנה",
    "נתון כללי 2 — מידות, חומרים, גיל, שימוש"
  ],
  "findings": [
    {
      "element": "שם האלמנט הספציפי שנבדק",
      "location_detail": "מיקום מדויק בשטח (קומה, חדר, כיוון אוריינטציה)",
      "description": "תיאור מפורט של הממצא — מה נצפה, מה הבעיה",
      "recommendation": "המלצה לטיפול — מה יש לעשות ומתי",
      "priority": "1",
      "status": "לא תקין"
    }
  ],
  "conclusions": [
    "מסקנה מקצועית 1",
    "מסקנה מקצועית 2"
  ],
  "notes": [
    "הערה/הנחיה כללית 1",
    "הערה 2"
  ]
}

━━━ כללים קפדניים ━━━
- priority: "1"=דחוף (טיפול מיידי), "2"=בינוני (עד 3 חודשים), "3"=נמוך (עד 6 חודשים)
- status: "תקין" / "לא תקין" / "תקין - דורש מעקב"
- ממצא אחד לכל ליקוי — אל תאחד ממצאים שונים לשורה אחת
- נתח תמונות שצורפו ותוסיף ממצאים ויזואליים שנראים בהן
- כל הטקסט בעברית הנדסית מקצועית
- שדות ריקים = מחרוזת ריקה, רשימות ריקות = []
- data_section: לפחות 2 נתונים כלליים
- conclusions: 1-3 מסקנות תמציתיות וברורות`;

// ── Map Claude findings → tableRows per doc type ──────────────────────────────
function toTableRows(findings, docType) {
  const cols = TABLE_COLUMNS[docType];
  if (!cols) return findings.map(f => [f.location_detail, f.element, f.description, f.recommendation].filter((_, i) => i < 4));

  return findings.map(f => {
    const row = Array(cols.length).fill('');
    switch (docType) {
      case 'group1':
        row[0] = f.location_detail || '';
        row[1] = f.element || '';
        row[2] = f.description || '';
        row[3] = f.recommendation || '';
        break;
      case 'group2':
        row[0] = '';
        row[1] = f.element || '';
        row[2] = '';
        row[3] = f.description || '';
        row[4] = `${f.description || ''}${f.location_detail ? ` — ${f.location_detail}` : ''}`;
        row[5] = f.priority || '1';
        break;
      case 'group3':
        row[0] = f.location_detail || '';
        row[1] = f.element || '';
        row[2] = f.status || 'לא תקין';
        row[3] = f.priority || '1';
        row[4] = f.recommendation || '';
        break;
      case 'group4':
        row[0] = f.element || '';
        row[1] = f.description || '';
        row[2] = f.recommendation || '';
        row[3] = f.status || 'לא תקין';
        row[4] = f.priority || '1';
        break;
      case 'group5':
        row[0] = '';
        row[1] = f.location_detail || '';
        row[2] = f.element || '';
        row[3] = f.description || '';
        row[4] = f.status || 'לא תקין';
        row[5] = f.priority || '1';
        row[6] = '';
        break;
      default:
        row[0] = f.location_detail || f.element || '';
        row[1] = f.description || '';
        if (row.length > 2) row[2] = f.recommendation || '';
        break;
    }
    return row;
  });
}

function toDefectsRows(findings, docType) {
  const cols = DEFECTS_COLUMNS[docType];
  if (!cols) return [];
  return findings.filter(f => f.status !== 'תקין').map(f => {
    const row = Array(cols.length).fill('');
    switch (docType) {
      case 'group3':
        row[0] = f.location_detail || '';
        row[1] = f.description || '';
        row[2] = f.recommendation || '';
        row[3] = '';
        row[4] = f.priority || '1';
        break;
      case 'group4':
        row[0] = '';
        row[1] = f.location_detail || '';
        row[2] = f.description || '';
        row[3] = f.recommendation || '';
        row[4] = '';
        row[5] = f.priority || '1';
        break;
      case 'group5':
        row[0] = '';
        row[1] = f.location_detail || '';
        row[2] = f.description || '';
        row[3] = '';
        row[4] = f.priority || '1';
        break;
    }
    return row;
  });
}

// ── Main component ────────────────────────────────────────────────────────────
export default function AIWriter({ onBack }) {
  const [apiKey, setApiKey]         = useState(() => localStorage.getItem(API_KEY_STORAGE) || '');
  const [keyInput, setKeyInput]     = useState('');
  const [showKeySetup, setShowKeySetup] = useState(!localStorage.getItem(API_KEY_STORAGE));

  const [rawText, setRawText]       = useState('');
  const [photos, setPhotos]         = useState([]);
  const [docType, setDocType]       = useState('group1');

  const [isGenerating, setIsGenerating] = useState(false);
  const [parsed, setParsed]         = useState(null);
  const [error, setError]           = useState('');
  const [step, setStep]             = useState(0); // 0=input 1=review 2=done
  const [isExporting, setIsExporting] = useState(false);
  const [sampleContext, setSampleContext] = useState(''); // loaded sample text for Claude

  const fileRef   = useRef(null);
  const cameraRef = useRef(null);

  // ── API key setup ─────────────────────────────────────────────────────────
  const saveKey = () => {
    const k = keyInput.trim();
    if (!k.startsWith('sk-ant-')) {
      setError('מפתח לא תקני — חייב להתחיל ב-sk-ant-');
      return;
    }
    localStorage.setItem(API_KEY_STORAGE, k);
    setApiKey(k);
    setShowKeySetup(false);
    setError('');
  };

  // ── Load sample reports when docType changes ──────────────────────────────
  useEffect(() => {
    let cancelled = false;
    async function loadSamples() {
      try {
        const base = import.meta.env.BASE_URL || '/';
        const res = await fetch(`${base}samples/index.json`);
        if (!res.ok) return;
        const idx = await res.json();
        const files = (idx[docType] || idx['*'] || []);
        if (!files.length) { setSampleContext(''); return; }
        const texts = await Promise.all(
          files.map(async (path) => {
            try {
              const r = await fetch(`${base}${path}`);
              return r.ok ? await r.text() : '';
            } catch { return ''; }
          })
        );
        if (!cancelled) setSampleContext(texts.filter(Boolean).join('\n\n---\n\n'));
      } catch { /* samples folder not found — no context */ }
    }
    setSampleContext('');
    loadSamples();
    return () => { cancelled = true; };
  }, [docType]);

  // ── Photo handling ────────────────────────────────────────────────────────
  const addPhotos = (files) => {
    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = e => setPhotos(prev => [...prev, { data: e.target.result, caption: '' }]);
      reader.readAsDataURL(file);
    });
  };

  const handlePaste = (e) => {
    const imgs = Array.from(e.clipboardData?.items || []).filter(it => it.type.startsWith('image/'));
    if (!imgs.length) return;
    e.preventDefault();
    imgs.forEach(it => { const f = it.getAsFile(); if (f) addPhotos([f]); });
  };

  // ── Call Claude API ───────────────────────────────────────────────────────
  const generate = async () => {
    if (!rawText.trim() && photos.length === 0) {
      setError('הדבק טקסט או צרף תמונות לפני הניתוח');
      return;
    }
    setIsGenerating(true);
    setError('');

    try {
      const content = [
        // Up to 10 photos as vision inputs
        ...photos.slice(0, 10).map(ph => {
          const match = ph.data.match(/^data:([^;]+);base64,(.+)$/);
          if (!match) return null;
          return { type: 'image', source: { type: 'base64', media_type: match[1], data: match[2] } };
        }).filter(Boolean),
        {
          type: 'text',
          text: [
            `סוג הדוח המבוקש: ${DOC_TYPES_CONFIG[docType].name.replace('\n', ' ')}`,
            sampleContext ? `\n\nלהלן דוגמאות לדוחות מאותו סוג — למד מסגנון הכתיבה, המבנה וניסוח הממצאים:\n\n${sampleContext}\n\n--- סוף הדוגמאות ---` : '',
            `\n\n${rawText || 'נתח את התמונות שצורפו וכתוב דוח מלא.'}`,
          ].join(''),
        },
      ];

      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: MODEL,
          max_tokens: 4096,
          system: SYSTEM_PROMPT,
          messages: [{ role: 'user', content }],
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err?.error?.message || `שגיאת API: ${res.status}`);
      }

      const data = await res.json();
      const raw = data.content?.[0]?.text || '';
      // Strip any markdown code fences
      const clean = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
      const result = JSON.parse(clean);
      setParsed({ ...result, photos });
      setStep(1);
    } catch (e) {
      setError(`שגיאה: ${e.message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  // ── Export to Word ────────────────────────────────────────────────────────
  const exportWord = async () => {
    if (!parsed) return;
    setIsExporting(true);
    try {
      const tableRows   = toTableRows(parsed.findings || [], docType);
      const defectsRows = toDefectsRows(parsed.findings || [], docType);
      const hasDefects  = (parsed.findings || []).some(f => f.status !== 'תקין');

      // Build intro_extra: AI intro + data_section items
      const introParts = [parsed.intro || '', ...(parsed.data_section || [])].filter(Boolean);
      const introExtra = introParts.join('\n');

      // conclusions array → newline string; notes array → array for doc generator
      const conclusionCustom = (parsed.conclusions || []).join('\n');
      const notesCustom      = parsed.notes || [];

      const data = {
        doc_type:          docType,
        client:            parsed.client || '',
        organization:      parsed.organization || '',
        location:          parsed.location || '',
        address:           parsed.address || '',
        date:              parsed.date || new Date().toISOString().split('T')[0],
        inspection_date:   parsed.inspection_date || parsed.date || new Date().toISOString().split('T')[0],
        subject:           parsed.subject || DOC_TYPES_CONFIG[docType].subject_default,
        intro_extra:       introExtra,
        has_defects:       hasDefects,
        conclusion_custom: conclusionCustom,
        notes_custom:      notesCustom,
        table_rows:        tableRows,
        defects_rows:      defectsRows,
        photos:            (parsed.photos || []).map(p => ({ data: p.data, caption: p.caption || '' })),
      };

      const blob = await generateDocument(data);
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url;
      const aiFilename = `${parsed.client || 'דוח'} — ${parsed.subject || DOC_TYPES_CONFIG[docType].name.replace('\n',' ')}.docx`;
      a.download = aiFilename;
      a.click();
      URL.revokeObjectURL(url);
      saveToArchive({ type: 'ai', docType, filename: aiFilename, client: parsed.client, subject: parsed.subject, date: parsed.date });
      setStep(2);
    } catch (e) {
      setError(`שגיאת ייצוא: ${e.message}`);
    } finally {
      setIsExporting(false);
    }
  };

  // ── Helpers for review UI ─────────────────────────────────────────────────
  const updateFinding = (i, patch) =>
    setParsed(p => { const fs = [...p.findings]; fs[i] = { ...fs[i], ...patch }; return { ...p, findings: fs }; });

  const updateList = (key, i, val) =>
    setParsed(p => { const arr = [...(p[key] || [])]; arr[i] = val; return { ...p, [key]: arr }; });

  const removeListItem = (key, i) =>
    setParsed(p => ({ ...p, [key]: (p[key] || []).filter((_, j) => j !== i) }));

  const addListItem = (key, val = '') =>
    setParsed(p => ({ ...p, [key]: [...(p[key] || []), val] }));

  const reset = () => { setRawText(''); setPhotos([]); setParsed(null); setStep(0); setError(''); };

  // ── Render: API key setup ─────────────────────────────────────────────────
  if (showKeySetup) {
    return (
      <div className="app-body">
        <div className="field-journal-header">
          <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
          <span className="field-journal-title">🤖 כתיבת דוח AI</span>
        </div>
        <div className="card" style={{ maxWidth: 480, margin: '32px auto' }}>
          <div className="card-title" style={{ fontSize: 18, marginBottom: 8 }}>🔑 הגדרת מפתח API</div>
          <p style={{ fontSize: 13, color: '#64748b', direction: 'rtl', marginBottom: 16, lineHeight: 1.6 }}>
            הפיצ'ר משתמש ב-Claude AI של Anthropic.
            קבל מפתח חינמי ב-<strong>console.anthropic.com</strong> ← API Keys ← Create Key.
          </p>
          <input
            className="form-input" type="password" placeholder="sk-ant-..."
            value={keyInput} onChange={e => setKeyInput(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && saveKey()}
            dir="ltr" style={{ fontFamily: 'monospace', letterSpacing: 1 }} autoFocus
          />
          {error && <div className="alert alert-error" style={{ marginTop: 8 }}>{error}</div>}
          <button className="btn btn-primary btn-lg" style={{ marginTop: 12, width: '100%' }} onClick={saveKey}>שמור ▶</button>
        </div>
      </div>
    );
  }

  // ── Render: Done ──────────────────────────────────────────────────────────
  if (step === 2) {
    return (
      <div className="app-body" style={{ textAlign: 'center', padding: '40px 20px', direction: 'rtl' }}>
        <div style={{ fontSize: 56, marginBottom: 16 }}>✅</div>
        <p style={{ fontSize: 18, fontWeight: 700, marginBottom: 8 }}>הדוח נוצר בהצלחה!</p>
        <p style={{ color: '#64748b', marginBottom: 24 }}>הקובץ הורד למכשירך</p>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'center' }}>
          <button className="btn btn-primary" onClick={reset}>+ דוח חדש</button>
          <button className="btn btn-outline" onClick={onBack}>חזור לבית</button>
        </div>
      </div>
    );
  }

  // ── Render: Review ────────────────────────────────────────────────────────
  if (step === 1 && parsed) {
    const PRIORITY_LABELS = { '1': '1 — דחוף', '2': '2 — בינוני', '3': '3 — נמוך' };
    const STATUS_OPTIONS  = ['תקין', 'לא תקין', 'תקין - דורש מעקב'];

    const EditableList = ({ title, icon, itemKey, placeholder, rows = 1 }) => (
      <div className="card">
        <div className="card-title">{icon} {title} ({(parsed[itemKey] || []).length})</div>
        {(parsed[itemKey] || []).map((item, i) => (
          <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 8, direction: 'rtl' }}>
            {rows > 1
              ? <textarea className="form-textarea" rows={rows} style={{ flex: 1, marginBottom: 0 }}
                  value={item} onChange={e => updateList(itemKey, i, e.target.value)} dir="rtl" />
              : <input className="form-input" style={{ flex: 1 }}
                  value={item} onChange={e => updateList(itemKey, i, e.target.value)} dir="rtl" />
            }
            <button className="btn-icon del" style={{ flexShrink: 0 }}
              onClick={() => removeListItem(itemKey, i)}>🗑</button>
          </div>
        ))}
        <button className="btn btn-outline btn-sm" onClick={() => addListItem(itemKey, '')}>
          + הוסף {placeholder}
        </button>
      </div>
    );

    return (
      <div className="app-body">
        <div className="field-journal-header">
          <button className="btn btn-outline btn-sm" onClick={() => setStep(0)}>◀ חזור</button>
          <span className="field-journal-title">✏️ עריכה לפני ייצוא</span>
          <button className="btn btn-outline btn-sm" style={{ fontSize: 11 }}
            onClick={() => { localStorage.removeItem(API_KEY_STORAGE); setApiKey(''); setShowKeySetup(true); }}>
            🔑
          </button>
        </div>

        {/* ── נתונים מזהים ── */}
        <div className="card">
          <div className="card-title">📋 נתונים מזהים</div>
          {[['לקוח/מוסד','client'],['ארגון','organization'],['מיקום','location'],['כתובת','address'],['נושא הדוח','subject']].map(([lbl, key]) => (
            <div className="form-group" key={key}>
              <label className="form-label">{lbl}</label>
              <input className="form-input" value={parsed[key] || ''}
                onChange={e => setParsed(p => ({ ...p, [key]: e.target.value }))} dir="rtl" />
            </div>
          ))}
          <div style={{ display: 'flex', gap: 10 }}>
            <div className="form-group" style={{ flex: 1 }}>
              <label className="form-label">תאריך דוח</label>
              <input type="date" className="form-input" value={parsed.date || ''}
                onChange={e => setParsed(p => ({ ...p, date: e.target.value }))} />
            </div>
            <div className="form-group" style={{ flex: 1 }}>
              <label className="form-label">תאריך ביקור</label>
              <input type="date" className="form-input" value={parsed.inspection_date || parsed.date || ''}
                onChange={e => setParsed(p => ({ ...p, inspection_date: e.target.value }))} />
            </div>
          </div>
        </div>

        {/* ── פסקת פתיחה ── */}
        {parsed.intro !== undefined && (
          <div className="card">
            <div className="card-title">📝 פסקת פתיחה</div>
            <textarea className="form-textarea" rows={3} dir="rtl"
              value={parsed.intro || ''}
              onChange={e => setParsed(p => ({ ...p, intro: e.target.value }))} />
          </div>
        )}

        {/* ── נתונים כלליים ── */}
        <EditableList title="נתונים כלליים" icon="📊" itemKey="data_section" placeholder="נתון" />

        {/* ── ממצאים ── */}
        <div className="card">
          <div className="card-title">⚠️ ממצאים ({(parsed.findings || []).length})</div>
          {(parsed.findings || []).map((f, i) => (
            <div key={i} style={{ border: '1px solid #e5e7eb', borderRadius: 8, padding: 12, marginBottom: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, direction: 'rtl' }}>
                <span style={{ fontWeight: 700, color: '#64748b' }}>ממצא {i + 1}</span>
                <button className="btn-icon del" onClick={() => setParsed(p => ({ ...p, findings: p.findings.filter((_, j) => j !== i) }))}>🗑</button>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 8 }}>
                <div>
                  <label className="form-label">אלמנט</label>
                  <input className="form-input" value={f.element || ''} dir="rtl"
                    onChange={e => updateFinding(i, { element: e.target.value })} />
                </div>
                <div>
                  <label className="form-label">מיקום</label>
                  <input className="form-input" value={f.location_detail || ''} dir="rtl"
                    onChange={e => updateFinding(i, { location_detail: e.target.value })} />
                </div>
              </div>
              <div className="form-group">
                <label className="form-label">תיאור הממצא</label>
                <textarea className="form-textarea" rows={2} dir="rtl" value={f.description || ''}
                  onChange={e => updateFinding(i, { description: e.target.value })} />
              </div>
              <div className="form-group">
                <label className="form-label">המלצה לטיפול</label>
                <input className="form-input" value={f.recommendation || ''} dir="rtl"
                  onChange={e => updateFinding(i, { recommendation: e.target.value })} />
              </div>
              <div style={{ display: 'flex', gap: 10 }}>
                <div style={{ flex: 1 }}>
                  <label className="form-label">סטטוס</label>
                  <select className="form-select" value={f.status || 'לא תקין'}
                    onChange={e => updateFinding(i, { status: e.target.value })}>
                    {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div style={{ flex: 1 }}>
                  <label className="form-label">קדימות</label>
                  <select className="form-select" value={f.priority || '1'}
                    onChange={e => updateFinding(i, { priority: e.target.value })}>
                    {Object.entries(PRIORITY_LABELS).map(([v, lbl]) => <option key={v} value={v}>{lbl}</option>)}
                  </select>
                </div>
              </div>
            </div>
          ))}
          <button className="btn btn-outline btn-sm"
            onClick={() => addListItem('findings', { element: '', location_detail: '', description: '', recommendation: '', status: 'לא תקין', priority: '2' })}>
            + הוסף ממצא
          </button>
        </div>

        {/* ── מסקנות ── */}
        <EditableList title="מסקנות הבדיקה" icon="✅" itemKey="conclusions" placeholder="מסקנה" rows={2} />

        {/* ── הערות ── */}
        <EditableList title="הערות כלליות" icon="📌" itemKey="notes" placeholder="הערה" />

        {error && <div className="alert alert-error">{error}</div>}

        <button className="btn btn-success btn-lg" style={{ marginBottom: 20, width: '100%' }}
          onClick={exportWord} disabled={isExporting}>
          {isExporting ? '⏳ יוצר קובץ Word...' : '⬇️ ייצא ל-Word'}
        </button>
      </div>
    );
  }

  // ── Render: Input ─────────────────────────────────────────────────────────
  return (
    <div className="app-body">
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="field-journal-title">🤖 כתיבת דוח AI</span>
        <button className="btn btn-outline btn-sm" style={{ fontSize: 11 }}
          onClick={() => { localStorage.removeItem(API_KEY_STORAGE); setApiKey(''); setShowKeySetup(true); }}>
          🔑 מפתח
        </button>
      </div>

      {/* Doc type */}
      <div className="card">
        <div className="card-title">
          📄 סוג הדוח
          {sampleContext && (
            <span style={{ fontSize: 11, fontWeight: 400, color: '#059669', marginRight: 8 }}>
              ✓ נטענו דוחות לדוגמה
            </span>
          )}
        </div>
        <div className="doc-type-grid">
          {Object.entries(DOC_TYPES_CONFIG).map(([key, cfg]) => (
            <div key={key} className={`doc-type-card${docType === key ? ' selected' : ''}`}
              onClick={() => setDocType(key)}>
              <div className="doc-type-check">✓</div>
              <div className="doc-type-icon">{cfg.icon}</div>
              <div className="doc-type-name">{cfg.name.replace('\n', ' ')}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Text input */}
      <div className="card">
        <div className="card-title">📝 הערות שטח</div>
        <p style={{ fontSize: 13, color: '#64748b', direction: 'rtl', marginBottom: 10 }}>
          הדבק הערות מהביקור — תאריך, לקוח, כתובת, ממצאים. Claude יזהה הכל אוטומטית.
        </p>
        <textarea
          className="form-textarea"
          style={{ minHeight: 200, fontSize: 15, lineHeight: 1.8 }}
          value={rawText}
          onChange={e => setRawText(e.target.value)}
          onPaste={handlePaste}
          placeholder={`לדוגמה:\nלקוח: בית ספר ביאליק\nכתובת: רחוב הרצל 5, תל אביב\nתאריך ביקור: 15.5.2025\n\nממצאים:\n- גג בנין ב': קורת פלדה חלודה\n- חדר כיתה 3א: רצפה שקועה בפינה המזרחית...`}
          dir="rtl"
        />
      </div>

      {/* Photos */}
      <div className="card">
        <div className="card-title">📷 תמונות ({photos.length})</div>
        <p style={{ fontSize: 13, color: '#64748b', direction: 'rtl', marginBottom: 10 }}>
          Claude ינתח את התמונות ויוסיף ממצאים שנראים בהן.
        </p>
        {photos.length > 0 && (
          <div className="journal-photos" style={{ marginBottom: 12 }}>
            {photos.map((ph, idx) => (
              <div key={idx} className="journal-photo">
                <img src={ph.data} alt={`תמונה ${idx + 1}`} />
                <button className="journal-photo-remove" onClick={() => setPhotos(p => p.filter((_, i) => i !== idx))}>✕</button>
              </div>
            ))}
          </div>
        )}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn btn-outline" onClick={() => fileRef.current?.click()}>📎 בחר קבצים</button>
          <button className="btn btn-outline" onClick={() => cameraRef.current?.click()}>📸 צלם</button>
        </div>
        <input ref={fileRef} type="file" accept="image/*" multiple hidden
          onChange={e => { addPhotos(e.target.files); e.target.value = ''; }} />
        <input ref={cameraRef} type="file" accept="image/*" capture="environment" hidden
          onChange={e => { addPhotos(e.target.files); e.target.value = ''; }} />
      </div>

      {error && <div className="alert alert-error">{error}</div>}

      <button
        className="btn btn-primary btn-lg generate-btn"
        style={{ marginBottom: 20, width: '100%', fontSize: 17 }}
        onClick={generate}
        disabled={isGenerating || (!rawText.trim() && photos.length === 0)}
      >
        {isGenerating
          ? <span>⏳ Claude מנתח ומכתב... <small style={{ fontSize: 13, opacity: 0.8 }}>(5-15 שניות)</small></span>
          : '🤖 צור דוח עם AI ▶'}
      </button>

      {isGenerating && (
        <div style={{ textAlign: 'center', padding: '20px', direction: 'rtl', color: '#64748b' }}>
          <div className="spinner-lg" style={{ margin: '0 auto 12px' }} />
          <p>Claude קורא את ההערות{photos.length > 0 ? ` ומנתח ${photos.length} תמונות` : ''}...</p>
          <p style={{ fontSize: 12, marginTop: 4 }}>מזהה לקוח, ממצאים, המלצות</p>
        </div>
      )}
    </div>
  );
}
