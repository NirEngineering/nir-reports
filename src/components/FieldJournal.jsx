import { useState, useRef, useEffect, useCallback } from 'react';
import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Header, Footer, AlignmentType, convertMillimetersToTwip as mm,
  UnderlineType,
} from 'docx';

// Exact pixel sizes derived from sample Word document (914400 EMU = 1 inch = 96px)
const HEADER_W = 359, HEADER_H = 142;   // 95.12mm × 37.57mm
const FOOTER_W = 568, FOOTER_H = 39;    // 150.50mm × 10.28mm
const FONT = 'Arial';
const SP = { line: 360, lineRule: 'auto', after: 0 };

async function fetchBuf(paths) {
  for (const url of paths) {
    try {
      const r = await fetch(url);
      if (r.ok) return new Uint8Array(await r.arrayBuffer());
    } catch { /* try next */ }
  }
  return new Uint8Array(0);
}

const JOURNALS_KEY = 'nir_v2_journals';
const FONTS = ['Arial', 'Times New Roman', 'Courier New', 'David', 'FrankRuehl'];
const SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32];

// execCommand font size: HTML 1–7 mapped from px
const pxToHtmlSize = (px) => {
  if (px <= 10) return 1;
  if (px <= 12) return 2;
  if (px <= 14) return 3;
  if (px <= 16) return 4;
  if (px <= 20) return 5;
  if (px <= 24) return 6;
  return 7;
};

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function newJournal() {
  const now = new Date();
  return {
    id: uid(),
    title: `יומן שטח ${now.toLocaleDateString('he-IL')}`,
    date: now.toISOString().split('T')[0],
    content: '',
    photos: [],
    createdAt: now.getTime(),
  };
}

function stripHtml(html) {
  if (!html) return '';
  return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

// Parse HTML into blocks with alignment, preserving Ctrl+R / Ctrl+L formatting.
// For <w:bidi/> paragraphs Word reverses physical alignment:
//   AlignmentType.LEFT = physical right, AlignmentType.RIGHT = physical left.
// So we swap: editor "right" → LEFT (physical right), editor "left" → RIGHT (physical left).
function parseHtmlBlocks(html) {
  if (!html) return [];
  const container = document.createElement('div');
  container.innerHTML = html;
  const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE']);
  const results = [];

  function getAlign(el) {
    const ta = (el.style?.textAlign || '').toLowerCase();
    if (ta === 'left') return AlignmentType.RIGHT;    // physical left
    if (ta === 'center') return AlignmentType.CENTER;
    return AlignmentType.LEFT;                         // physical right (default)
  }

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = node.textContent.trim();
      if (t) results.push({ text: t, align: AlignmentType.LEFT }); // default physical right
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      if (BLOCK.has(node.tagName)) {
        const t = node.textContent.trim();
        if (t) results.push({ text: t, align: getAlign(node) });
      } else {
        for (const c of node.childNodes) walk(c);
      }
    }
  }

  for (const c of container.childNodes) walk(c);
  return results;
}

function base64ToUint8Array(b64) {
  const raw = b64.replace(/^data:[^,]+,/, '');
  const bin = atob(raw);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

// ─────────────────────────────────────────────────────────────────────────────
// Main component
// ─────────────────────────────────────────────────────────────────────────────
export default function FieldJournal({ onBack }) {
  const [journals, setJournals] = useState(() => {
    try { return JSON.parse(localStorage.getItem(JOURNALS_KEY) || '[]'); } catch { return []; }
  });
  const [activeId, setActiveId] = useState(null);

  // Toolbar state (for editor)
  const [font, setFont]         = useState('Arial');
  const [fontSize, setFontSize] = useState(16);
  const [textDir, setTextDir]   = useState('rtl');

  // Drag state (for list)
  const [dragIdx, setDragIdx]       = useState(null);
  const [dragOverIdx, setDragOverIdx] = useState(null);

  // Inline rename state (for list)
  const [renamingId, setRenamingId]   = useState(null);
  const [renameVal,  setRenameVal]    = useState('');

  // New-journal creation card state
  const [creatingNew, setCreatingNew] = useState(false);
  const [newTitle,    setNewTitle]    = useState('');
  const newTitleRef = useRef(null);

  const editorRef  = useRef(null);
  const savedRange = useRef(null);
  const fileRef    = useRef(null);

  // Persist journals to localStorage on every change
  useEffect(() => {
    try { localStorage.setItem(JOURNALS_KEY, JSON.stringify(journals)); } catch (_) {}
  }, [journals]);

  // Load content into editor when opening a journal
  useEffect(() => {
    if (!editorRef.current || !activeId) return;
    const j = journals.find(x => x.id === activeId);
    if (j) editorRef.current.innerHTML = j.content || '';
    editorRef.current.focus();
  }, [activeId]);

  // ── Helpers ──────────────────────────────────────────────────────────────
  const activeJournal = journals.find(j => j.id === activeId) || null;

  const updateJournal = useCallback((id, updates) => {
    setJournals(js => js.map(j => j.id === id ? { ...j, ...updates } : j));
  }, []);

  const saveContent = useCallback(() => {
    if (!editorRef.current || !activeId) return;
    updateJournal(activeId, { content: editorRef.current.innerHTML });
  }, [activeId, updateJournal]);

  // Selection management — preserves cursor when toolbar button is clicked
  const saveSelection = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) savedRange.current = sel.getRangeAt(0).cloneRange();
  };

  const restoreSelection = () => {
    if (!savedRange.current || !editorRef.current) return;
    editorRef.current.focus();
    const sel = window.getSelection();
    if (sel) { sel.removeAllRanges(); sel.addRange(savedRange.current); }
  };

  const execCmd = (cmd, value) => {
    restoreSelection();
    document.execCommand(cmd, false, value ?? null);
    saveContent();
    editorRef.current?.focus();
  };

  // ── Journal list operations ───────────────────────────────────────────────
  const addJournal = () => {
    setNewTitle('');
    setCreatingNew(true);
    setTimeout(() => {
      newTitleRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      newTitleRef.current?.focus();
    }, 80);
  };

  const confirmNewJournal = () => {
    const title = newTitle.trim() || `יומן שטח ${new Date().toLocaleDateString('he-IL')}`;
    const j = { ...newJournal(), title };
    setJournals(prev => [j, ...prev]);
    setFont('Arial'); setFontSize(16); setTextDir('rtl');
    setCreatingNew(false);
    setNewTitle('');
    setActiveId(j.id);
  };

  const cancelNewJournal = () => { setCreatingNew(false); setNewTitle(''); };

  const openJournal = (id) => {
    saveContent();
    setFont('Arial'); setFontSize(16); setTextDir('rtl');
    setActiveId(id);
  };

  const deleteJournal = (id) => {
    if (!window.confirm('למחוק יומן זה לצמיתות?')) return;
    setJournals(js => js.filter(j => j.id !== id));
    if (activeId === id) setActiveId(null);
  };

  const moveUp = (idx) => {
    if (idx === 0) return;
    setJournals(js => { const a = [...js]; [a[idx-1], a[idx]] = [a[idx], a[idx-1]]; return a; });
  };

  const moveDown = (idx) => {
    setJournals(js => {
      if (idx >= js.length - 1) return js;
      const a = [...js]; [a[idx], a[idx+1]] = [a[idx+1], a[idx]]; return a;
    });
  };

  // HTML5 drag-and-drop reorder
  const handleDragStart = (e, idx) => { e.dataTransfer.effectAllowed = 'move'; setDragIdx(idx); };
  const handleDragOver  = (e, idx) => { e.preventDefault(); setDragOverIdx(idx); };
  const handleDrop      = (idx) => {
    if (dragIdx !== null && dragIdx !== idx) {
      setJournals(js => {
        const a = [...js];
        const [item] = a.splice(dragIdx, 1);
        a.splice(idx, 0, item);
        return a;
      });
    }
    setDragIdx(null); setDragOverIdx(null);
  };
  const handleDragEnd   = () => { setDragIdx(null); setDragOverIdx(null); };

  // ── Photos ───────────────────────────────────────────────────────────────
  const handlePhotoAdd = (files) => {
    const id = activeId;
    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = e => {
        setJournals(js => js.map(j =>
          j.id === id ? { ...j, photos: [...j.photos, { data: e.target.result, caption: '' }] } : j
        ));
      };
      reader.readAsDataURL(file);
    });
  };

  const removePhoto = (idx) => {
    setJournals(js => js.map(j =>
      j.id === activeId ? { ...j, photos: j.photos.filter((_, i) => i !== idx) } : j
    ));
  };

  const updateCaption = (idx, caption) => {
    setJournals(js => js.map(j => {
      if (j.id !== activeId) return j;
      const photos = [...j.photos];
      photos[idx] = { ...photos[idx], caption };
      return { ...j, photos };
    }));
  };

  // Global document-level paste handler: captures images pasted anywhere on the page
  // (only active when a journal is open; skips text inputs and textareas)
  useEffect(() => {
    if (!activeId) return;
    const handler = (e) => {
      const tag = e.target?.tagName?.toLowerCase();
      if (tag === 'input' || tag === 'textarea') return;
      const imgs = Array.from(e.clipboardData?.items || []).filter(it => it.type.startsWith('image/'));
      if (!imgs.length) return;
      e.preventDefault();
      imgs.forEach(it => { const f = it.getAsFile(); if (f) handlePhotoAdd([f]); });
    };
    document.addEventListener('paste', handler);
    return () => document.removeEventListener('paste', handler);
  }, [activeId]);

  // Intercept Ctrl+V with image clipboard data
  const handleEditorPaste = (e) => {
    const items = Array.from(e.clipboardData?.items || []);
    const imageItems = items.filter(it => it.type.startsWith('image/'));
    if (imageItems.length === 0) return; // let normal text paste proceed
    e.preventDefault();
    handlePhotoAdd(imageItems.map(it => it.getAsFile()).filter(Boolean));
  };

  // Drop images onto the editor area
  const handleEditorDrop = (e) => {
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (files.length === 0) return;
    e.preventDefault();
    handlePhotoAdd(files);
  };

  // ── Toolbar actions ───────────────────────────────────────────────────────
  const insertTimestamp = () => {
    restoreSelection();
    const now = new Date();
    const ts = `[${now.toLocaleDateString('he-IL')} ${now.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' })}] `;
    document.execCommand('insertText', false, ts);
    saveContent();
  };

  const changeFont = (f) => {
    setFont(f);
    restoreSelection();
    document.execCommand('fontName', false, f);
    saveContent();
    editorRef.current?.focus();
  };

  const changeFontSize = (sz) => {
    setFontSize(sz);
    restoreSelection();
    document.execCommand('fontSize', false, pxToHtmlSize(sz));
    saveContent();
    editorRef.current?.focus();
  };

  const toggleDir = (d) => {
    setTextDir(d);
    if (editorRef.current) editorRef.current.dir = d;
    execCmd(d === 'rtl' ? 'justifyRight' : 'justifyLeft');
  };

  // ── Word export (company template: logos, RTL, proper formatting) ──────────
  const exportToWord = async () => {
    if (!activeJournal) return;
    saveContent();

    const base = import.meta.env.BASE_URL || '/';
    const [headerBuf, footerBuf] = await Promise.all([
      fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
      fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
    ]);

    const mk = (text, { size = 11, bold = false } = {}) =>
      new TextRun({ text: String(text ?? ''), font: FONT, size: size * 2, bold });

    // For <w:bidi/> paragraphs: LEFT = physical right, RIGHT = physical left
    const mkP = (children, align = AlignmentType.LEFT) =>
      new Paragraph({ children, alignment: align, spacing: SP, bidirectional: true });

    // Header / footer images
    const headerSection = new Header({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: headerBuf.length > 100
          ? [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })]
          : [new TextRun('')],
      })],
    });
    const footerSection = new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: footerBuf.length > 100
          ? [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })]
          : [new TextRun('')],
      })],
    });

    // Parse HTML content preserving per-paragraph alignment (Ctrl+R / Ctrl+L)
    const rawHtml = editorRef.current?.innerHTML || activeJournal.content || '';
    const blocks = parseHtmlBlocks(rawHtml);

    // Photo paragraphs
    const photoParas = [];
    for (const ph of activeJournal.photos || []) {
      photoParas.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: base64ToUint8Array(ph.data), transformation: { width: 400, height: 300 } })],
      }));
      if (ph.caption) {
        photoParas.push(mkP([mk(ph.caption, { size: 8 })], AlignmentType.CENTER));
      }
    }

    const doc = new Document({
      styles: {
        paragraphStyles: [{
          id: 'Normal', name: 'Normal', quickFormat: true,
          paragraph: { bidirectional: true, alignment: AlignmentType.LEFT },
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
          // Title
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 480, after: 0 },
            bidirectional: true,
            children: [mk(activeJournal.title, { size: 15, bold: true })],
          }),
          // Date — RIGHT = physical left for <w:bidi/> paragraphs
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: SP,
            bidirectional: true,
            children: [mk(activeJournal.date, { size: 8 })],
          }),
          mkP([mk('')]),

          // Body lines — alignment preserved from editor (Ctrl+R / Ctrl+L)
          ...blocks.map((block, i) => new Paragraph({
            children: [mk(block.text, { size: 9 })],
            alignment: block.align,
            spacing: i === 0 ? { ...SP, before: 360 } : SP,
            bidirectional: true,
          })),

          // Photos
          ...(photoParas.length ? [mkP([mk('')]), ...photoParas] : []),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${activeJournal.title || 'יומן שטח'}.docx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ═══════════════════════════════════════════════════════════════════════════
  // RENDER — LIST
  // ═══════════════════════════════════════════════════════════════════════════
  if (!activeId) {
    return (
      <div className="app-body">
        <div className="field-journal-header">
          <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
          <span className="field-journal-title">📋 יומני שטח</span>
          <button className="btn btn-primary btn-sm" onClick={addJournal}>+ חדש</button>
        </div>

        {/* ── New journal creation card ── */}
        {creatingNew && (
          <div className="journal-new-card">
            <div className="journal-new-card-label">שם היומן</div>
            <input
              ref={newTitleRef}
              className="journal-new-card-input"
              placeholder="לדוגמה: ביקור שטח מוסד חינוך..."
              value={newTitle}
              onChange={e => setNewTitle(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') confirmNewJournal();
                if (e.key === 'Escape') cancelNewJournal();
              }}
              dir="rtl"
              autoComplete="off"
              autoFocus
            />
            <div className="journal-new-card-actions">
              <button className="btn btn-primary" onClick={confirmNewJournal}>
                פתח עורך ←
              </button>
              <button className="btn btn-outline" onClick={cancelNewJournal}>
                ביטול
              </button>
            </div>
          </div>
        )}

        {journals.length === 0 && !creatingNew ? (
          <div className="journal-empty">
            <div className="journal-empty-icon">📋</div>
            <p>אין יומנים עדיין</p>
            <button className="btn btn-primary btn-lg" style={{ marginTop: 16 }} onClick={addJournal}>
              + צור יומן שטח ראשון
            </button>
          </div>
        ) : (
          <div className="journal-items">
            {journals.map((j, idx) => (
              <div
                key={j.id}
                className={`journal-item${dragOverIdx === idx ? ' drag-over' : ''}`}
                draggable
                onDragStart={e => handleDragStart(e, idx)}
                onDragOver={e => handleDragOver(e, idx)}
                onDrop={() => handleDrop(idx)}
                onDragEnd={handleDragEnd}
              >
                <span className="journal-drag-handle">⠿</span>

                <div className="journal-item-body" onClick={() => renamingId !== j.id && openJournal(j.id)}>
                  {renamingId === j.id ? (
                    <input
                      autoFocus
                      className="journal-rename-input"
                      value={renameVal}
                      dir="rtl"
                      onChange={e => setRenameVal(e.target.value)}
                      onBlur={() => { updateJournal(j.id, { title: renameVal }); setRenamingId(null); }}
                      onKeyDown={e => {
                        if (e.key === 'Enter') { updateJournal(j.id, { title: renameVal }); setRenamingId(null); }
                        if (e.key === 'Escape') setRenamingId(null);
                        e.stopPropagation();
                      }}
                      onClick={e => e.stopPropagation()}
                    />
                  ) : (
                    <div className="journal-item-title">{j.title}</div>
                  )}
                  <div className="journal-item-date">{j.date}</div>
                  {j.content && (
                    <div className="journal-item-preview">{stripHtml(j.content).slice(0, 80)}</div>
                  )}
                  {(j.photos || []).length > 0 && (
                    <div className="journal-item-photos-hint">📷 {j.photos.length} תמונות</div>
                  )}
                </div>

                <div className="journal-item-actions">
                  <button
                    className="btn-icon"
                    onClick={e => { e.stopPropagation(); setRenamingId(j.id); setRenameVal(j.title); }}
                    title="שנה שם"
                  >✏️</button>
                  <button
                    className="btn-icon"
                    onClick={e => { e.stopPropagation(); moveUp(idx); }}
                    disabled={idx === 0}
                    title="הזז למעלה"
                  >↑</button>
                  <button
                    className="btn-icon"
                    onClick={e => { e.stopPropagation(); moveDown(idx); }}
                    disabled={idx === journals.length - 1}
                    title="הזז למטה"
                  >↓</button>
                  <button
                    className="btn-icon del"
                    onClick={e => { e.stopPropagation(); deleteJournal(j.id); }}
                    title="מחק יומן"
                  >🗑</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // RENDER — EDITOR
  // ═══════════════════════════════════════════════════════════════════════════
  return (
    <div className="app-body">

      {/* ── Top bar ── */}
      <div className="field-journal-header">
        <button
          className="btn btn-outline btn-sm"
          onClick={() => { saveContent(); setActiveId(null); }}
        >◀ רשימה</button>

        <input
          className="journal-title-input"
          value={activeJournal?.title || ''}
          onChange={e => updateJournal(activeId, { title: e.target.value })}
          dir="rtl"
          placeholder="שם היומן..."
        />

        <input
          type="date"
          className="journal-date-input"
          value={activeJournal?.date || ''}
          onChange={e => updateJournal(activeId, { date: e.target.value })}
        />
      </div>

      {/* ── Toolbar ── */}
      <div className="journal-toolbar">

        {/* Font family */}
        <select
          className="journal-toolbar-select"
          value={font}
          onMouseDown={saveSelection}
          onChange={e => changeFont(e.target.value)}
        >
          {FONTS.map(f => <option key={f} value={f}>{f}</option>)}
        </select>

        {/* Font size */}
        <select
          className="journal-toolbar-select journal-toolbar-size"
          value={fontSize}
          onMouseDown={saveSelection}
          onChange={e => changeFontSize(parseInt(e.target.value))}
        >
          {SIZES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>

        <div className="journal-toolbar-sep" />

        {/* Bold */}
        <button className="journal-toolbar-btn" onMouseDown={saveSelection} onClick={() => execCmd('bold')} title="מודגש">
          <b>B</b>
        </button>
        {/* Italic */}
        <button className="journal-toolbar-btn" onMouseDown={saveSelection} onClick={() => execCmd('italic')} title="נטוי">
          <i>I</i>
        </button>
        {/* Underline */}
        <button className="journal-toolbar-btn" onMouseDown={saveSelection} onClick={() => execCmd('underline')} title="קו תחתון">
          <u>U</u>
        </button>

        <div className="journal-toolbar-sep" />

        {/* RTL (right-align) */}
        <button
          className={`journal-toolbar-btn${textDir === 'rtl' ? ' active' : ''}`}
          onMouseDown={saveSelection}
          onClick={() => toggleDir('rtl')}
          title="יישור ימינה (RTL)"
        >⇤</button>

        {/* Center */}
        <button
          className="journal-toolbar-btn"
          onMouseDown={saveSelection}
          onClick={() => execCmd('justifyCenter')}
          title="מרכוז"
        >≡</button>

        {/* LTR (left-align) */}
        <button
          className={`journal-toolbar-btn${textDir === 'ltr' ? ' active' : ''}`}
          onMouseDown={saveSelection}
          onClick={() => toggleDir('ltr')}
          title="יישור שמאלה (LTR)"
        >⇥</button>

        <div className="journal-toolbar-sep" />

        {/* Unordered list / bullets */}
        <button
          className="journal-toolbar-btn"
          onMouseDown={saveSelection}
          onClick={() => execCmd('insertUnorderedList')}
          title="רשימת נקודות"
        >• —</button>

        {/* Ordered list / numbers */}
        <button
          className="journal-toolbar-btn"
          onMouseDown={saveSelection}
          onClick={() => execCmd('insertOrderedList')}
          title="רשימה ממוספרת"
        >1.</button>

        <div className="journal-toolbar-sep" />

        {/* Timestamp */}
        <button
          className="journal-toolbar-btn"
          onMouseDown={saveSelection}
          onClick={insertTimestamp}
          title="הוסף תאריך ושעה"
        >🕐</button>

        {/* Photo */}
        <button
          className="journal-toolbar-btn"
          onClick={() => fileRef.current?.click()}
          title="צרף תמונה / צלם"
        >📷</button>

        <input
          ref={fileRef}
          type="file"
          accept="image/*"
          multiple
          capture="environment"
          style={{ display: 'none' }}
          onChange={e => { handlePhotoAdd(e.target.files); e.target.value = ''; }}
        />

        {/* Export */}
        <button
          className="journal-toolbar-btn journal-export-btn"
          onClick={exportToWord}
          title="ייצא לקובץ Word"
        >⬇ Word</button>
      </div>

      {/* ── Content editable area ── */}
      <div
        ref={editorRef}
        className="journal-editor"
        contentEditable
        suppressContentEditableWarning
        dir={textDir}
        style={{ fontFamily: font, fontSize: `${fontSize}px` }}
        onInput={saveContent}
        onKeyDown={e => {
          if (e.ctrlKey && e.key === "r") { e.preventDefault(); execCmd("justifyRight"); }
          if (e.ctrlKey && e.key === "l") { e.preventDefault(); execCmd("justifyLeft"); }
        }}
        onKeyUp={saveSelection}
        onMouseUp={saveSelection}
        onSelect={saveSelection}
        onPaste={handleEditorPaste}
        onDrop={handleEditorDrop}
        onDragOver={e => e.preventDefault()}
      />

      {/* ── Photo zone ── always visible, prominent ── */}
      <div className="journal-photo-zone">
        <div className="journal-photo-zone-buttons">
          <button
            className="journal-photo-zone-btn"
            onClick={() => { fileRef.current.removeAttribute('capture'); fileRef.current.click(); }}
          >
            <span>🖼️</span> בחר מהגלריה
          </button>
          <button
            className="journal-photo-zone-btn"
            onClick={() => { fileRef.current.setAttribute('capture', 'environment'); fileRef.current.click(); }}
          >
            <span>📷</span> צלם תמונה
          </button>
        </div>
        <div className="journal-photo-zone-hint">או גרור תמונה לכאן · הדבק מהלוח (Ctrl+V)</div>

        {(activeJournal?.photos || []).length > 0 && (
          <div className="journal-photos">
            {activeJournal.photos.map((ph, idx) => (
              <div key={idx} className="journal-photo">
                <img src={ph.data} alt={`תמונה ${idx + 1}`} />
                <button className="journal-photo-remove" onClick={() => removePhoto(idx)}>✕</button>
                <input
                  className="journal-photo-caption"
                  value={ph.caption}
                  onChange={e => updateCaption(idx, e.target.value)}
                  placeholder="כיתוב לתמונה..."
                  dir="rtl"
                />
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ── Footer ── */}
      <div className="field-journal-footer">
        <span style={{ fontSize: 12, color: '#64748b' }}>✓ נשמר אוטומטית</span>
        <button className="btn btn-success" onClick={exportToWord}>
          ⬇️ ייצא לוורד
        </button>
      </div>
    </div>
  );
}
