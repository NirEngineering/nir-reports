import { useState, useEffect, useRef, useCallback } from 'react';
import { isConnected } from '../lib/oneDriveSync';
import {
  ensureProjectFolder, uploadProjectFile, listProjectFiles,
  loadProjectNotes, saveProjectNotes, fetchFileBytes,
} from '../lib/projectStorage';
import { addToOutbox, getOutbox, removeFromOutbox } from '../lib/offlineOutbox';
import { saveToArchive } from '../lib/archiveUtils';
import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Header, Footer, AlignmentType, convertMillimetersToTwip as mm,
} from 'docx';

const PROJECTS_KEY = 'nir_v2_projects';
const FONT = 'Arial';
const HEADER_W = 359, HEADER_H = 142;
const FOOTER_W = 568, FOOTER_H = 39;
const SP = { line: 360, lineRule: 'auto', after: 0 };

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

async function fetchBuf(paths) {
  for (const url of paths) {
    try {
      const r = await fetch(url);
      if (r.ok) return new Uint8Array(await r.arrayBuffer());
    } catch { /* try next */ }
  }
  return new Uint8Array(0);
}

function loadProjects() {
  try { return JSON.parse(localStorage.getItem(PROJECTS_KEY) || '[]'); } catch { return []; }
}
function persistProjects(list) {
  try { localStorage.setItem(PROJECTS_KEY, JSON.stringify(list)); } catch { /* ignore */ }
}

function fmtTime(ts) {
  if (!ts) return '';
  const d = new Date(ts);
  return d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });
}
function fmtDay(ts) {
  return new Date(ts).toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// ═════════════════════════════════════════════════════════════════════════
// Main
// ═════════════════════════════════════════════════════════════════════════
export default function ProjectVault({ onBack }) {
  const [projects, setProjects] = useState(() => loadProjects());
  const [activeId, setActiveId] = useState(null);
  const [creating, setCreating] = useState(false);
  const [newName, setNewName] = useState('');
  const [newClient, setNewClient] = useState('');
  const [busy, setBusy] = useState(false);
  const connected = isConnected();

  useEffect(() => { persistProjects(projects); }, [projects]);

  const activeProject = projects.find(p => p.id === activeId) || null;

  const createProject = async () => {
    const name = newName.trim();
    if (!name) return;
    setBusy(true);
    try {
      const folderPath = await ensureProjectFolder(name);
      const project = { id: uid(), name, client: newClient.trim(), folderPath, createdAt: Date.now() };
      setProjects(prev => [project, ...prev]);
      setNewName(''); setNewClient(''); setCreating(false);
      setActiveId(project.id);
    } catch (e) {
      alert('שגיאה ביצירת הפרויקט: ' + e.message);
    } finally {
      setBusy(false);
    }
  };

  const renameProject = (id, name) => {
    setProjects(prev => prev.map(p => p.id === id ? { ...p, name } : p));
  };

  const removeProject = (id) => {
    if (!window.confirm('להסיר פרויקט זה מהרשימה? (הקבצים ב-OneDrive לא יימחקו)')) return;
    setProjects(prev => prev.filter(p => p.id !== id));
    if (activeId === id) setActiveId(null);
  };

  if (activeProject) {
    return (
      <ProjectFeed
        project={activeProject}
        onBack={() => setActiveId(null)}
        onRename={(name) => renameProject(activeProject.id, name)}
      />
    );
  }

  return (
    <div className="app-body">
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="field-journal-title">🗂️ פרויקטים בשטח</span>
        <button className="btn btn-primary btn-sm" onClick={() => setCreating(true)}>+ חדש</button>
      </div>

      {!connected && (
        <div className="alert alert-error" style={{ marginBottom: 12 }}>
          ⚠️ יש להתחבר קודם ל-OneDrive (כפתור ☁️ בראש המסך) כדי שהקבצים יישמרו בענן בלי הגבלת נפח.
        </div>
      )}

      {creating && (
        <div className="journal-new-card">
          <div className="journal-new-card-label">שם הפרויקט</div>
          <input
            className="journal-new-card-input"
            placeholder="לדוגמה: בית ספר יסודי - רמת גן"
            value={newName}
            onChange={e => setNewName(e.target.value)}
            dir="rtl"
            autoFocus
          />
          <div className="journal-new-card-label" style={{ marginTop: 8 }}>לקוח (אופציונלי)</div>
          <input
            className="journal-new-card-input"
            placeholder="שם הלקוח..."
            value={newClient}
            onChange={e => setNewClient(e.target.value)}
            dir="rtl"
          />
          <div className="journal-new-card-actions">
            <button className="btn btn-primary" onClick={createProject} disabled={busy || !connected}>
              {busy ? '⏳ יוצר תיקייה...' : 'צור פרויקט ←'}
            </button>
            <button className="btn btn-outline" onClick={() => setCreating(false)}>ביטול</button>
          </div>
        </div>
      )}

      {projects.length === 0 && !creating ? (
        <div className="journal-empty">
          <div className="journal-empty-icon">🗂️</div>
          <p>אין עדיין פרויקטים</p>
          <p style={{ fontSize: 13, color: '#94a3b8' }}>כל פרויקט מקבל תיקייה משלו ב-OneDrive — תמונות, סרטונים והערות נשמרים בה מיד</p>
          <button className="btn btn-primary btn-lg" style={{ marginTop: 16 }} onClick={() => setCreating(true)}>
            + צור פרויקט ראשון
          </button>
        </div>
      ) : (
        <div className="journal-items">
          {projects.map(p => (
            <div key={p.id} className="journal-item" onClick={() => setActiveId(p.id)}>
              <div className="journal-item-body">
                <div className="journal-item-title">{p.name}</div>
                <div className="journal-item-date">{p.client || 'ללא לקוח'} · נוצר {fmtDay(p.createdAt)}</div>
              </div>
              <div className="journal-item-actions">
                <button
                  className="btn-icon del"
                  onClick={e => { e.stopPropagation(); removeProject(p.id); }}
                  title="הסר מהרשימה"
                >🗑</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// Project feed — chat-style: scrollable history + fixed compose bar
// ═════════════════════════════════════════════════════════════════════════
function ProjectFeed({ project, onBack, onRename }) {
  const [items, setItems] = useState([]);       // merged, chronological
  const [loading, setLoading] = useState(true);
  const [text, setText] = useState('');
  const [listening, setListening] = useState(false);
  const [showReport, setShowReport] = useState(false);
  const [renaming, setRenaming] = useState(false);
  const [renameVal, setRenameVal] = useState(project.name);

  const notesRef = useRef([]);   // raw notes.json array kept in sync for saves
  const feedEndRef = useRef(null);
  const photoRef = useRef(null);
  const videoRef = useRef(null);

  const scrollToEnd = () => feedEndRef.current?.scrollIntoView({ block: 'end' });

  // ── Load feed: remote notes + remote files + local outbox ────────────────
  const loadFeed = useCallback(async () => {
    setLoading(true);
    try {
      const [notes, files, outbox] = await Promise.all([
        loadProjectNotes(project.folderPath).catch(() => []),
        listProjectFiles(project.folderPath).catch(() => []),
        getOutbox(project.id),
      ]);
      notesRef.current = notes;

      const noteItems = notes.map(n => ({ kind: 'note', id: n.id, text: n.text, ts: n.ts, status: 'sent' }));
      const fileItems = files.map(f => ({
        kind: f.isVideo ? 'video' : 'photo',
        id: f.id, name: f.name, url: f.downloadUrl, thumb: f.thumbnailUrl,
        ts: new Date(f.createdDateTime).getTime(), status: 'sent',
      }));
      const outboxItems = outbox.map(o => ({
        kind: o.kind, id: o.id, text: o.text,
        url: o.file ? URL.createObjectURL(o.file) : null,
        ts: o.createdAt, status: 'pending', outboxId: o.id,
      }));

      const merged = [...noteItems, ...fileItems, ...outboxItems].sort((a, b) => (a.ts || 0) - (b.ts || 0));
      setItems(merged);
    } finally {
      setLoading(false);
      setTimeout(scrollToEnd, 50);
    }
  }, [project.folderPath, project.id]);

  useEffect(() => { loadFeed(); }, [loadFeed]);

  // ── Flush outbox when connectivity returns ────────────────────────────────
  const flushOutbox = useCallback(async () => {
    const pending = await getOutbox(project.id);
    for (const entry of pending) {
      try {
        if (entry.kind === 'note') {
          const list = [...notesRef.current, { id: entry.id, text: entry.text, ts: entry.createdAt }];
          await saveProjectNotes(project.folderPath, list);
          notesRef.current = list;
        } else {
          await uploadProjectFile(project.folderPath, entry.file);
        }
        await removeFromOutbox(entry.id);
      } catch {
        break; // still offline (or failing) — stop, retry next time
      }
    }
    loadFeed();
  }, [project.folderPath, project.id, loadFeed]);

  useEffect(() => {
    window.addEventListener('online', flushOutbox);
    return () => window.removeEventListener('online', flushOutbox);
  }, [flushOutbox]);

  // ── Sending ────────────────────────────────────────────────────────────
  const sendNote = async () => {
    const t = text.trim();
    if (!t) return;
    setText('');
    const entry = { id: uid(), text: t, ts: Date.now() };
    setItems(prev => [...prev, { kind: 'note', ...entry, status: 'sending' }]);
    setTimeout(scrollToEnd, 30);
    try {
      const list = [...notesRef.current, entry];
      await saveProjectNotes(project.folderPath, list);
      notesRef.current = list;
      setItems(prev => prev.map(it => it.id === entry.id ? { ...it, status: 'sent' } : it));
    } catch {
      await addToOutbox({ projectId: project.id, folderPath: project.folderPath, kind: 'note', text: t, createdAt: entry.ts });
      setItems(prev => prev.map(it => it.id === entry.id ? { ...it, status: 'pending' } : it));
    }
  };

  const sendFiles = async (files, kind) => {
    for (const file of files) {
      const localId = uid();
      const ts = Date.now();
      const previewUrl = URL.createObjectURL(file);
      setItems(prev => [...prev, { kind, id: localId, url: previewUrl, name: file.name, ts, status: 'sending' }]);
      setTimeout(scrollToEnd, 30);
      try {
        await uploadProjectFile(project.folderPath, file);
        setItems(prev => prev.map(it => it.id === localId ? { ...it, status: 'sent' } : it));
      } catch {
        await addToOutbox({ projectId: project.id, folderPath: project.folderPath, kind, file, createdAt: ts });
        setItems(prev => prev.map(it => it.id === localId ? { ...it, status: 'pending' } : it));
      }
    }
  };

  // ── Voice-to-text (Hebrew) — fills the input, user reviews then sends ────
  const startVoice = () => {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) return alert('הדפדפן שלך אינו תומך בהקלטת קול');
    const recog = new SR();
    recog.lang = 'he-IL';
    recog.interimResults = false;
    recog.onresult = (e) => setText(prev => (prev ? prev + ' ' : '') + e.results[0][0].transcript);
    recog.onerror = () => setListening(false);
    recog.onend = () => setListening(false);
    recog.start();
    setListening(true);
  };

  const commitRename = () => {
    const v = renameVal.trim();
    if (v) onRename(v);
    setRenaming(false);
  };

  const pendingCount = items.filter(i => i.status === 'pending').length;

  return (
    <div className="pv-feed-wrap">
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ פרויקטים</button>
        {renaming ? (
          <input
            className="journal-rename-input"
            value={renameVal}
            autoFocus
            dir="rtl"
            onChange={e => setRenameVal(e.target.value)}
            onBlur={commitRename}
            onKeyDown={e => { if (e.key === 'Enter') commitRename(); if (e.key === 'Escape') setRenaming(false); }}
          />
        ) : (
          <span className="field-journal-title" onDoubleClick={() => setRenaming(true)} title="לחץ פעמיים לשינוי שם">
            {project.name}
          </span>
        )}
        <button className="btn btn-primary btn-sm" onClick={() => setShowReport(true)}>📄 הפק דו"ח</button>
      </div>

      {pendingCount > 0 && (
        <div className="pv-outbox-banner">🕐 {pendingCount} ממתינים לשליחה — יעלו אוטומטית כשתהיה רשת</div>
      )}

      <div className="pv-feed">
        {loading && <div className="pv-feed-loading">⏳ טוען...</div>}
        {!loading && items.length === 0 && (
          <div className="journal-empty">
            <div className="journal-empty-icon">💬</div>
            <p>עדיין אין כלום כאן</p>
            <p style={{ fontSize: 13, color: '#94a3b8' }}>כתוב הערה, צלם תמונה או סרטון — בדיוק כמו הודעה בווצאפ</p>
          </div>
        )}
        {items.map(it => <FeedBubble key={it.id} item={it} />)}
        <div ref={feedEndRef} />
      </div>

      <div className="pv-input-bar">
        <button className="pv-input-btn" onClick={() => photoRef.current?.click()} title="צלם / בחר תמונה">📷</button>
        <button className="pv-input-btn" onClick={() => videoRef.current?.click()} title="צלם / בחר סרטון">🎥</button>
        <button
          className={`pv-input-btn${listening ? ' active' : ''}`}
          onClick={startVoice}
          disabled={listening}
          title="הכתבה קולית"
        >{listening ? '🔴' : '🎙️'}</button>
        <input
          className="pv-input-text"
          value={text}
          onChange={e => setText(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendNote(); } }}
          placeholder="כתוב הערה..."
          dir="rtl"
        />
        <button className="pv-input-send" onClick={sendNote} title="שלח">➤</button>

        <input ref={photoRef} type="file" accept="image/*" capture="environment" multiple hidden
          onChange={e => { if (e.target.files.length) sendFiles(Array.from(e.target.files), 'photo'); e.target.value = ''; }} />
        <input ref={videoRef} type="file" accept="video/*" capture="environment" multiple hidden
          onChange={e => { if (e.target.files.length) sendFiles(Array.from(e.target.files), 'video'); e.target.value = ''; }} />
      </div>

      {showReport && (
        <ReportModal
          project={project}
          items={items.filter(i => i.status !== 'pending')}
          onClose={() => setShowReport(false)}
        />
      )}
    </div>
  );
}

function FeedBubble({ item }) {
  const statusIcon = item.status === 'sending' ? '⏳' : item.status === 'pending' ? '🕐' : '✓';
  if (item.kind === 'note') {
    return (
      <div className="pv-bubble pv-bubble-note">
        <div className="pv-bubble-text">{item.text}</div>
        <div className="pv-bubble-meta">{fmtTime(item.ts)} <span className="pv-bubble-status">{statusIcon}</span></div>
      </div>
    );
  }
  return (
    <div className="pv-bubble pv-bubble-media">
      {item.kind === 'video'
        ? <video src={item.url} poster={item.thumb || undefined} controls className="pv-bubble-media-el" />
        : <img src={item.thumb || item.url} alt="" className="pv-bubble-media-el" />}
      <div className="pv-bubble-meta">{fmtTime(item.ts)} <span className="pv-bubble-status">{statusIcon}</span></div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════
// Report generation — pick what goes into the Word document
// ═════════════════════════════════════════════════════════════════════════
function ReportModal({ project, items, onClose }) {
  const media = items.filter(i => i.kind === 'photo' || i.kind === 'video');
  const notesText = items.filter(i => i.kind === 'note').map(i => `[${fmtTime(i.ts)}] ${i.text}`).join('\n');

  const [selected, setSelected] = useState(() => new Set(media.filter(m => m.kind === 'photo').map(m => m.id)));
  const [bodyText, setBodyText] = useState(notesText);
  const [generating, setGenerating] = useState(false);

  const toggle = (id) => setSelected(prev => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });

  const generate = async () => {
    setGenerating(true);
    try {
      const base = import.meta.env.BASE_URL || '/';
      const [headerBuf, footerBuf] = await Promise.all([
        fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
        fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
      ]);

      const mk = (t, { size = 9, bold = false } = {}) => new TextRun({ text: String(t ?? ''), font: FONT, size: size * 2, bold });
      const mkP = (children, align = AlignmentType.RIGHT) => new Paragraph({ children, alignment: align, spacing: SP, bidirectional: true });

      const headerSection = new Header({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: headerBuf.length > 100 ? [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })] : [new TextRun('')],
      })]});
      const footerSection = new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: footerBuf.length > 100 ? [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })] : [new TextRun('')],
      })]});

      const photoParas = [];
      const chosenPhotos = media.filter(m => m.kind === 'photo' && selected.has(m.id));
      const chosenVideos = media.filter(m => m.kind === 'video' && selected.has(m.id));
      for (const ph of chosenPhotos) {
        try {
          const bytes = await fetchFileBytes(ph.url);
          photoParas.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({ data: bytes, transformation: { width: 400, height: 300 } })],
          }));
        } catch { /* skip photo that failed to fetch (expired link) */ }
      }
      for (const v of chosenVideos) {
        photoParas.push(mkP([mk(`🎥 וידאו מצורף: ${v.name} (ראה בתיקיית הפרויקט ב-OneDrive)`, { size: 8 })], AlignmentType.CENTER));
      }

      const doc = new Document({
        styles: {
          paragraphStyles: [{
            id: 'Normal', name: 'Normal', quickFormat: true,
            paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT },
            run: { font: { name: FONT } },
          }],
        },
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
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER, spacing: { before: 480, after: 0 }, bidirectional: true,
              children: [mk(project.name, { size: 15, bold: true })],
            }),
            new Paragraph({
              alignment: AlignmentType.LEFT, spacing: SP, bidirectional: true,
              children: [mk(new Date().toLocaleDateString('he-IL'), { size: 8 })],
            }),
            mkP([mk('')]),
            ...bodyText.split('\n').filter(Boolean).map((line, i) => new Paragraph({
              children: [mk(line, { size: 9 })],
              alignment: AlignmentType.RIGHT,
              spacing: i === 0 ? { ...SP, before: 360 } : SP,
              bidirectional: true,
            })),
            ...(photoParas.length ? [mkP([mk('')]), ...photoParas] : []),
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const filename = `${project.name} - דוח שטח.docx`;
      a.href = url; a.download = filename; a.click();
      URL.revokeObjectURL(url);
      saveToArchive({ type: 'journal', filename, title: project.name, date: new Date().toISOString().split('T')[0] });
      onClose();
    } catch (e) {
      alert('שגיאה ביצירת המסמך: ' + e.message);
    } finally {
      setGenerating(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-sheet" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span>📄 הפקת דו"ח — {project.name}</span>
          <button className="btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="form-group">
            <label className="form-label">טקסט הדו"ח (מבוסס על ההערות שנרשמו)</label>
            <textarea className="form-textarea" rows={6} value={bodyText} onChange={e => setBodyText(e.target.value)} dir="rtl" />
          </div>

          {media.length > 0 && (
            <div className="form-group">
              <label className="form-label">בחר מדיה לכלול בדו"ח ({selected.size}/{media.length})</label>
              <div className="pv-report-grid">
                {media.map(m => (
                  <label key={m.id} className={`pv-report-thumb${selected.has(m.id) ? ' selected' : ''}`}>
                    <input type="checkbox" checked={selected.has(m.id)} onChange={() => toggle(m.id)} />
                    {m.kind === 'video'
                      ? <video src={m.url} className="pv-report-thumb-el" muted preload="metadata" playsInline />
                      : <img src={m.thumb || m.url} alt="" className="pv-report-thumb-el" />}
                    {m.kind === 'video' && <span className="pv-report-thumb-badge">🎥</span>}
                  </label>
                ))}
              </div>
            </div>
          )}

          <button className="btn btn-success btn-lg" style={{ width: '100%' }} onClick={generate} disabled={generating}>
            {generating ? '⏳ יוצר מסמך...' : '⬇️ צור והורד מסמך Word'}
          </button>
        </div>
      </div>
    </div>
  );
}
