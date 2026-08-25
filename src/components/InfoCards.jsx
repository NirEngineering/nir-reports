import { useState, useRef, useEffect, useCallback } from 'react';
import { saveToArchive } from '../lib/archiveUtils';
import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Header, Footer, AlignmentType, convertMillimetersToTwip as mm,
  Table as DocxTable, TableRow, TableCell, WidthType,
} from 'docx';

// ── Constants ─────────────────────────────────────────────────────────────────
const CARDS_KEY = 'nir_v2_info_cards';
const FONT = 'Arial';
const HEADER_W = 337, HEADER_H = 133; // verified against real Drive documents
const FOOTER_W = 569, FOOTER_H = 39; // verified against real Drive documents
const SP = { line: 360, lineRule: 'auto', after: 0 };
const FONTS = ['Arial', 'Times New Roman', 'David', 'FrankRuehl', 'Courier New'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32];

const pxToHtmlSize = (px) => {
  if (px <= 10) return 1; if (px <= 12) return 2; if (px <= 14) return 3;
  if (px <= 16) return 4; if (px <= 20) return 5; if (px <= 24) return 6;
  return 7;
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

function newCard() {
  const now = new Date();
  return {
    id: uid(),
    title: `כרטיסייה ${now.toLocaleDateString('he-IL')}`,
    date: now.toISOString().split('T')[0],
    content: '',
    photos: [],
    createdAt: now.getTime(),
  };
}

function stripHtml(html) {
  if (!html) return '';
  const div = document.createElement('div');
  div.innerHTML = html;
  return (div.innerText || div.textContent || '').trim();
}

// Parse HTML into blocks with alignment, preserving Ctrl+R / Ctrl+L formatting.
function parseHtmlBlocks(html) {
  if (!html) return [];
  const container = document.createElement('div');
  container.innerHTML = html;
  const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE']);
  const results = [];

  function getAlign(el) {
    const ta = (el.style?.textAlign || '').toLowerCase();
    if (ta === 'left') return AlignmentType.LEFT;
    if (ta === 'center') return AlignmentType.CENTER;
    return AlignmentType.RIGHT;                        // default: right
  }

  function walk(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = node.textContent.trim();
      if (t) results.push({ text: t, align: AlignmentType.RIGHT }); // default right
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

function b64ToUint8(b64) {
  const raw = b64.replace(/^data:[^,]+,/, '');
  const bin = atob(raw);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

async function fetchBuf(paths) {
  for (const url of paths) {
    try { const r = await fetch(url); if (r.ok) return new Uint8Array(await r.arrayBuffer()); } catch {}
  }
  return new Uint8Array(0);
}

// ── Main Component ────────────────────────────────────────────────────────────
export default function InfoCards({ onBack }) {
  const [cards, setCards] = useState(() => {
    try { return JSON.parse(localStorage.getItem(CARDS_KEY) || '[]'); } catch { return []; }
  });
  const [activeId, setActiveId]         = useState(null);
  const [search, setSearch]             = useState('');
  const [dragIdx, setDragIdx]           = useState(null);
  const [dragOverIdx, setDragOverIdx]   = useState(null);
  const [renamingId, setRenamingId]     = useState(null);
  const [renameVal, setRenameVal]       = useState('');
  const [showTableDlg, setShowTableDlg] = useState(false);
  const [tRows, setTRows]               = useState(3);
  const [tCols, setTCols]               = useState(3);
  const [exporting, setExporting]       = useState(false);
  const [font, setFont]                 = useState('Arial');
  const [fontSize, setFontSize]         = useState(14);

  const editorRef  = useRef(null);
  const fileRef    = useRef(null);
  const cameraRef  = useRef(null);
  const savedRange = useRef(null);

  const activeCard = cards.find(c => c.id === activeId) || null;

  // Persist to localStorage
  useEffect(() => {
    try { localStorage.setItem(CARDS_KEY, JSON.stringify(cards)); } catch {}
  }, [cards]);

  // Load content into editor when switching cards
  useEffect(() => {
    if (!activeId || !editorRef.current) return;
    const card = cards.find(c => c.id === activeId);
    if (card) { editorRef.current.innerHTML = card.content || ''; editorRef.current.focus(); }
  }, [activeId]); // eslint-disable-line

  const saveContent = useCallback(() => {
    if (!activeId || !editorRef.current) return;
    const html = editorRef.current.innerHTML;
    setCards(prev => prev.map(c => c.id === activeId ? { ...c, content: html } : c));
  }, [activeId]);

  // ── Card CRUD ─────────────────────────────────────────────────────────────
  const createCard = () => {
    const card = newCard();
    setCards(prev => [card, ...prev]);
    setActiveId(card.id);
  };

  const deleteCard = (id, e) => {
    e.stopPropagation();
    if (!window.confirm('למחוק כרטיסייה זו?')) return;
    setCards(prev => prev.filter(c => c.id !== id));
    if (activeId === id) setActiveId(null);
  };

  const updateCard = useCallback((id, updates) => {
    setCards(prev => prev.map(c => c.id === id ? { ...c, ...updates } : c));
  }, []);

  // ── Drag to reorder ───────────────────────────────────────────────────────
  const onDragStart = (e, i) => { e.dataTransfer.effectAllowed = 'move'; setDragIdx(i); };
  const onDragOver  = (e, i) => { e.preventDefault(); setDragOverIdx(i); };
  const onDrop      = (e, i) => {
    e.preventDefault();
    if (dragIdx === null || dragIdx === i) return;
    const next = [...cards];
    const [item] = next.splice(dragIdx, 1);
    next.splice(i, 0, item);
    setCards(next); setDragIdx(null); setDragOverIdx(null);
  };

  // ── Rename ────────────────────────────────────────────────────────────────
  const startRename = (e, card) => { e.stopPropagation(); setRenamingId(card.id); setRenameVal(card.title); };
  const commitRename = () => {
    if (renamingId && renameVal.trim()) updateCard(renamingId, { title: renameVal.trim() });
    setRenamingId(null);
  };

  // ── Rich text toolbar ─────────────────────────────────────────────────────
  const saveRange = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) savedRange.current = sel.getRangeAt(0).cloneRange();
  };

  const exec = useCallback((cmd, val = null) => {
    const ed = editorRef.current;
    if (!ed) return;
    if (!ed.contains(document.activeElement)) {
      ed.focus();
      if (savedRange.current) {
        const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(savedRange.current);
      }
    }
    document.execCommand(cmd, false, val);
    setTimeout(saveContent, 0);
  }, [saveContent]);

  const applyFont = (f) => { setFont(f); exec('fontName', f); };
  const applySize = (s) => { setFontSize(+s); exec('fontSize', pxToHtmlSize(+s)); };

  const insertTable = () => {
    let html = `<table style="border-collapse:collapse;width:100%;margin:8px 0"><tbody>`;
    for (let r = 0; r < tRows; r++) {
      html += '<tr>';
      for (let c = 0; c < tCols; c++) {
        html += `<td style="border:1px solid #aab;padding:6px 10px;min-width:55px;text-align:right" contenteditable="true">&nbsp;</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody></table><br>';
    exec('insertHTML', html);
    setShowTableDlg(false);
  };

  // ── Photos ────────────────────────────────────────────────────────────────
  const addPhotos = useCallback((files) => {
    Array.from(files).forEach(file => {
      const reader = new FileReader();
      reader.onload = ev => {
        setCards(prev => prev.map(c =>
          c.id === activeId
            ? { ...c, photos: [...c.photos, { id: uid(), data: ev.target.result, caption: '' }] }
            : c
        ));
      };
      reader.readAsDataURL(file);
    });
  }, [activeId]);

  // Global document-level paste handler: captures images pasted anywhere on the page
  // (only active when a card is open; skips text inputs and textareas)
  useEffect(() => {
    if (!activeId) return;
    const handler = (e) => {
      const tag = e.target?.tagName?.toLowerCase();
      if (tag === 'input' || tag === 'textarea') return;
      const imgs = Array.from(e.clipboardData?.items || []).filter(it => it.type.startsWith('image/'));
      if (!imgs.length) return;
      e.preventDefault();
      imgs.forEach(it => { const f = it.getAsFile(); if (f) addPhotos([f]); });
    };
    document.addEventListener('paste', handler);
    return () => document.removeEventListener('paste', handler);
  }, [activeId, addPhotos]);

  const handleEditorPaste = (e) => {
    const imgs = Array.from(e.clipboardData?.items || []).filter(it => it.type.startsWith('image/'));
    if (imgs.length === 0) return;
    e.preventDefault();
    imgs.forEach(it => { const f = it.getAsFile(); if (f) addPhotos([f]); });
  };

  const handlePhotoDrop = (e) => {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (files.length) addPhotos(files);
  };

  const removePhoto = (photoId) => {
    if (!activeCard) return;
    updateCard(activeId, { photos: activeCard.photos.filter(p => p.id !== photoId) });
  };

  const updateCaption = useCallback((photoId, caption) => {
    setCards(prev => prev.map(c =>
      c.id === activeId
        ? { ...c, photos: c.photos.map(p => p.id === photoId ? { ...p, caption } : p) }
        : c
    ));
  }, [activeId]);

  // ── Word export ───────────────────────────────────────────────────────────
  const exportToWord = async () => {
    if (!activeCard || exporting) return;
    saveContent();
    setExporting(true);
    try {
      const base = import.meta.env.BASE_URL || '/';
      const [headerBuf, footerBuf] = await Promise.all([
        fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
        fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
      ]);

      const mkR = (text, opts = {}) => new TextRun({
        text: String(text ?? ''), font: FONT, size: (opts.size || 9) * 2,
        bold: !!opts.bold, italics: !!opts.italic,
      });
      const mkP = (children, opts = {}) => new Paragraph({
        children,
        alignment: opts.align ?? AlignmentType.RIGHT,
        spacing: opts.spacing ?? SP,
        bidirectional: true,
      });

      const mkHeader = () => new Header({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: headerBuf.length > 100
          ? [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })]
          : [new TextRun('')],
      })]});
      const mkFooter = () => new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: footerBuf.length > 100
          ? [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })]
          : [new TextRun('')],
      })]});

      const rawHtml = editorRef.current
        ? editorRef.current.innerHTML
        : (activeCard.content || '');
      const blocks = parseHtmlBlocks(rawHtml);

      // Photos 2-per-row in a table
      const photos = activeCard.photos || [];
      const photoItems = [];
      for (let i = 0; i < photos.length; i += 2) {
        const pair = photos.slice(i, i + 2);
        const makeEmptyCell = () => new TableCell({
          children: [new Paragraph({ children: [] })],
          width: { size: 50, type: WidthType.PERCENTAGE },
        });
        const imgCells = pair.map(ph => new TableCell({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({ data: b64ToUint8(ph.data), transformation: { width: 255, height: 191 } })],
          })],
          width: { size: 50, type: WidthType.PERCENTAGE },
          margins: { top: 40, bottom: 40, left: 40, right: 40 },
        }));
        const capCells = pair.map(ph => new TableCell({
          children: [mkP([mkR(ph.caption || '', { size: 8 })], { align: AlignmentType.CENTER })],
          width: { size: 50, type: WidthType.PERCENTAGE },
        }));

        if (pair.length < 2) { imgCells.push(makeEmptyCell()); capCells.push(makeEmptyCell()); }

        photoItems.push(
          new DocxTable({
            alignment: AlignmentType.CENTER,
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({ children: imgCells }),
              new TableRow({ children: capCells }),
            ],
          }),
          mkP([mkR('')])
        );
      }

      const body = [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 480, after: 0 },
          bidirectional: true,
          children: [mkR(activeCard.title, { size: 15, bold: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: SP,
          bidirectional: true,
          children: [mkR(activeCard.date, { size: 8 })],
        }),
        mkP([mkR('')]),
        ...blocks.map((block, idx) => new Paragraph({
          children: [mkR(block.text, { size: 9 })],
          alignment: block.align,
          spacing: idx === 0 ? { ...SP, before: 360 } : SP,
          bidirectional: true,
        })),
        ...(photoItems.length ? [mkP([mkR('')]), ...photoItems] : []),
      ];

      const doc = new Document({
        styles: {
          default: { document: { run: { font: { name: FONT } } } },
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
              margin: { top: 2268, bottom: 568, left: 1260, right: 1800, header: 142, footer: 0 } // verified twips from real docs
            },
            bidi: true,
          },
          headers: { default: mkHeader() },
          footers: { default: mkFooter() },
          children: body,
        }],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const cardFilename = `${activeCard.title || 'כרטיסייה'}.docx`;
      a.href = url; a.download = cardFilename; a.click();
      URL.revokeObjectURL(url);
      saveToArchive({
        type: 'card',
        filename: cardFilename,
        title: activeCard.title,
        date: activeCard.date,
      });
    } finally {
      setExporting(false);
    }
  };

  // ── Share to WhatsApp ─────────────────────────────────────────────────────
  const shareCard = async () => {
    if (!activeCard || exporting) return;
    saveContent();
    setExporting(true);
    try {
      const textContent = editorRef.current
        ? stripHtml(editorRef.current.innerHTML)
        : stripHtml(activeCard.content);
      const shareText = `*${activeCard.title}*\n${activeCard.date}\n\n${textContent}`;

      if (navigator.canShare) {
        // Try file share (docx)
        const base = import.meta.env.BASE_URL || '/';
        const [headerBuf, footerBuf] = await Promise.all([
          fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
          fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
        ]);
        const shareBlocks = parseHtmlBlocks(editorRef.current ? editorRef.current.innerHTML : (activeCard.content || ''));
        const doc = new Document({
          styles: {
            default: { document: { run: { font: { name: FONT } } } },
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
                margin: { top: 2268, bottom: 568, left: 1260, right: 1800, header: 142, footer: 0 } // verified twips from real docs
              },
              bidi: true,
            },
            headers: {
              default: new Header({ children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: headerBuf.length > 100
                  ? [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })]
                  : [new TextRun('')],
              })]}),
            },
            footers: {
              default: new Footer({ children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: footerBuf.length > 100
                  ? [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })]
                  : [new TextRun('')],
              })]}),
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { before: 480, after: 0 }, bidirectional: true,
                children: [new TextRun({ text: activeCard.title, bold: true, size: 30, font: FONT })],
              }),
              new Paragraph({
                alignment: AlignmentType.LEFT, spacing: SP, bidirectional: true,
                children: [new TextRun({ text: activeCard.date, size: 16, font: FONT })],
              }),
              new Paragraph({ children: [] }),
              ...shareBlocks.map((block, i) => new Paragraph({
                alignment: block.align, bidirectional: true,
                spacing: i === 0 ? { ...SP, before: 360 } : SP,
                children: [new TextRun({ text: block.text, font: FONT, size: 18 })],
              })),
            ],
          }],
        });
        const blob = await Packer.toBlob(doc);
        const file = new File([blob], `${activeCard.title}.docx`, { type: blob.type });
        if (navigator.canShare({ files: [file] })) {
          try { await navigator.share({ title: activeCard.title, files: [file] }); return; } catch {}
        }
        try { await navigator.share({ title: activeCard.title, text: shareText }); return; } catch {}
      }
      // Fallback: copy to clipboard
      await navigator.clipboard.writeText(shareText).catch(() => {});
      alert('הטקסט הועתק ללוח');
    } finally {
      setExporting(false);
    }
  };

  // ── Filtered list ─────────────────────────────────────────────────────────
  const filtered = search.trim()
    ? cards.filter(c => c.title.includes(search) || stripHtml(c.content).includes(search))
    : cards;

  // ════════════════════════════════════════════════════════════════════════════
  // RENDER — Editor view
  // ════════════════════════════════════════════════════════════════════════════
  if (activeCard) {
    return (
      <div className="ic-wrap">

        {/* Header row */}
        <div className="ic-editor-header">
          <button className="btn-icon" onClick={() => { saveContent(); setActiveId(null); }} title="חזור לרשימה">←</button>
          {renamingId === activeCard.id ? (
            <input
              className="ic-rename-input"
              value={renameVal}
              onChange={e => setRenameVal(e.target.value)}
              onBlur={commitRename}
              onKeyDown={e => { if (e.key === 'Enter') commitRename(); if (e.key === 'Escape') setRenamingId(null); }}
              autoFocus
            />
          ) : (
            <h2 className="ic-editor-title" onDoubleClick={e => startRename(e, activeCard)} title="לחץ פעמיים לשינוי שם">
              {activeCard.title}
            </h2>
          )}
          <input
            type="date"
            className="ic-date-input"
            value={activeCard.date}
            onChange={e => updateCard(activeId, { date: e.target.value })}
          />
        </div>

        {/* Rich-text toolbar */}
        <div className="ic-toolbar" onMouseDown={e => { e.preventDefault(); saveRange(); }}>
          <select value={font} onChange={e => applyFont(e.target.value)} className="ic-tb-select">
            {FONTS.map(f => <option key={f} value={f}>{f}</option>)}
          </select>
          <select value={fontSize} onChange={e => applySize(e.target.value)} className="ic-tb-select-sm">
            {SIZES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>

          <span className="ic-tb-sep" />
          <button className="ic-tb-btn" onMouseDown={() => exec('bold')} title="הדגשה (Ctrl+B)"><strong>B</strong></button>
          <button className="ic-tb-btn" onMouseDown={() => exec('italic')} title="נטוי (Ctrl+I)"><em>I</em></button>
          <button className="ic-tb-btn" onMouseDown={() => exec('underline')} title="קו תחתון (Ctrl+U)"><u>U</u></button>
          <button className="ic-tb-btn" onMouseDown={() => exec('strikeThrough')} title="קו חוצה"><s>S</s></button>

          <span className="ic-tb-sep" />
          <button className="ic-tb-btn" onMouseDown={() => exec('justifyRight')} title="יישור ימין">⇥</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('justifyCenter')} title="מרכז">≡</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('justifyLeft')} title="יישור שמאל">⇤</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('justifyFull')} title="יישור מלא">☰</button>

          <span className="ic-tb-sep" />
          <button className="ic-tb-btn" onMouseDown={() => exec('insertOrderedList')} title="רשימה ממוספרת">1.</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('insertUnorderedList')} title="רשימת נקודות">•</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('outdent')} title="הפחת הזחה">⟵</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('indent')} title="הגדל הזחה">⟶</button>

          <span className="ic-tb-sep" />
          <button className="ic-tb-btn" onMouseDown={e => { e.preventDefault(); setShowTableDlg(true); }} title="הוסף טבלה">⊞</button>

          <span className="ic-tb-sep" />
          <button className="ic-tb-btn" onMouseDown={() => exec('undo')} title="בטל (Ctrl+Z)">↩</button>
          <button className="ic-tb-btn" onMouseDown={() => exec('redo')} title="חזור (Ctrl+Y)">↪</button>
        </div>

        {/* Content editor */}
        <div
          ref={editorRef}
          className="ic-editor-body"
          contentEditable
          dir="rtl"
          suppressContentEditableWarning
          onInput={saveContent}
          onKeyDown={e => {
            if (e.ctrlKey && e.key === "r") { e.preventDefault(); exec("justifyRight"); }
            if (e.ctrlKey && e.key === "l") { e.preventDefault(); exec("justifyLeft"); }
          }}
          onPaste={handleEditorPaste}
          onKeyUp={saveRange}
          onMouseUp={saveRange}
          onBlur={saveContent}
          data-placeholder="הקלד כאן את הנתונים, המידות, הטקסט..."
        />

        {/* Photos section */}
        <div
          className="ic-photos-section"
          onDragOver={e => e.preventDefault()}
          onDrop={handlePhotoDrop}
        >
          <div className="ic-photos-header">
            <span className="ic-photos-label">📷 תמונות ({(activeCard.photos || []).length})</span>
            <div className="ic-photos-btns">
              <button className="btn btn-sm btn-outline" onClick={() => fileRef.current?.click()}>📎 הוסף קובץ</button>
              <button className="btn btn-sm btn-outline" onClick={() => cameraRef.current?.click()}>📸 צלם</button>
            </div>
          </div>

          <input
            ref={fileRef}
            type="file"
            accept="image/*"
            multiple
            hidden
            onChange={e => { addPhotos(e.target.files); e.target.value = ''; }}
          />
          <input
            ref={cameraRef}
            type="file"
            accept="image/*"
            capture="environment"
            hidden
            onChange={e => { addPhotos(e.target.files); e.target.value = ''; }}
          />

          {(activeCard.photos || []).length === 0 && (
            <p className="ic-photos-empty">
              גרור תמונות לכאן · הדבק (Ctrl+V) · או השתמש בכפתורים<br />
              <small>מהווצאפ: פתח WhatsApp Web, לחץ ימני על תמונה → העתק, ואז הדבק כאן</small>
            </p>
          )}

          <div className="ic-photos-grid">
            {(activeCard.photos || []).map(ph => (
              <div key={ph.id} className="ic-photo-item">
                <div className="ic-photo-img-wrap">
                  <img src={ph.data} alt="" className="ic-photo-img" />
                  <button className="ic-photo-remove" onClick={() => removePhoto(ph.id)} title="הסר תמונה">✕</button>
                </div>
                <input
                  className="ic-photo-caption"
                  placeholder="כיתוב לתמונה..."
                  value={ph.caption}
                  onChange={e => updateCaption(ph.id, e.target.value)}
                />
              </div>
            ))}
          </div>
        </div>

        {/* Action bar */}
        <div className="ic-editor-actions">
          <button className="btn btn-primary" onClick={exportToWord} disabled={exporting}>
            {exporting ? '...' : '📝 ייצוא וורד'}
          </button>
          <button className="btn btn-success" onClick={shareCard} disabled={exporting}>
            📤 שתף
          </button>
          <button className="btn btn-outline" onClick={() => { saveContent(); setActiveId(null); }}>
            💾 שמור וסגור
          </button>
        </div>

        {/* Table dialog */}
        {showTableDlg && (
          <div className="ic-modal-overlay" onClick={() => setShowTableDlg(false)}>
            <div className="ic-modal" onClick={e => e.stopPropagation()}>
              <h3 className="ic-modal-title">הוסף טבלה</h3>
              <div className="ic-modal-row">
                <label>שורות</label>
                <input type="number" min={1} max={20} value={tRows} onChange={e => setTRows(+e.target.value)} className="ic-modal-input" />
              </div>
              <div className="ic-modal-row">
                <label>עמודות</label>
                <input type="number" min={1} max={10} value={tCols} onChange={e => setTCols(+e.target.value)} className="ic-modal-input" />
              </div>
              <div className="ic-modal-actions">
                <button className="btn btn-primary" onClick={insertTable}>הוסף</button>
                <button className="btn btn-outline" onClick={() => setShowTableDlg(false)}>ביטול</button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ════════════════════════════════════════════════════════════════════════════
  // RENDER — List view
  // ════════════════════════════════════════════════════════════════════════════
  return (
    <div className="ic-wrap">
      <div className="ic-list-header">
        <button className="btn-icon" onClick={onBack} title="חזור">←</button>
        <h2 className="ic-list-title">כרטיסיות מידע</h2>
        <button className="btn btn-primary btn-sm" onClick={createCard}>+ חדש</button>
      </div>

      <div className="ic-search-wrap">
        <input
          className="ic-search"
          placeholder="🔍 חיפוש בכרטיסיות..."
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        {search && (
          <button className="ic-search-clear" onClick={() => setSearch('')} title="נקה חיפוש">✕</button>
        )}
      </div>

      {filtered.length === 0 && (
        <div className="ic-empty">
          {search ? `אין תוצאות עבור "${search}"` : 'אין כרטיסיות עדיין — לחץ + חדש כדי להתחיל'}
        </div>
      )}

      <div className="ic-list">
        {filtered.map((card, i) => (
          <div
            key={card.id}
            className={`ic-card${dragOverIdx === i ? ' ic-card--dragover' : ''}`}
            draggable
            onDragStart={e => onDragStart(e, i)}
            onDragOver={e => onDragOver(e, i)}
            onDragEnd={() => { setDragIdx(null); setDragOverIdx(null); }}
            onDrop={e => onDrop(e, i)}
            onClick={() => setActiveId(card.id)}
          >
            <div className="ic-card-top">
              {renamingId === card.id ? (
                <input
                  className="ic-rename-input"
                  value={renameVal}
                  onClick={e => e.stopPropagation()}
                  onChange={e => setRenameVal(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={e => {
                    if (e.key === 'Enter') commitRename();
                    if (e.key === 'Escape') setRenamingId(null);
                  }}
                  autoFocus
                />
              ) : (
                <strong className="ic-card-title">{card.title}</strong>
              )}
              <span className="ic-card-date">{card.date}</span>
            </div>

            <p className="ic-card-preview">{stripHtml(card.content).slice(0, 100) || '(ריק)'}</p>

            {(card.photos || []).length > 0 && (
              <span className="ic-card-badge">📷 {card.photos.length}</span>
            )}

            <div className="ic-card-actions" onClick={e => e.stopPropagation()}>
              <button className="ic-act-btn" onClick={e => startRename(e, card)} title="שנה שם">✏</button>
              <button className="ic-act-btn ic-act-del" onClick={e => deleteCard(card.id, e)} title="מחק">🗑</button>
              <span className="ic-drag-handle" title="גרור לסידור מחדש">⠿</span>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
