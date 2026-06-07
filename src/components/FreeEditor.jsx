import { useEditor, EditorContent } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle, FontSize, FontFamily as FontFamilyExt } from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import FontFamily from '@tiptap/extension-font-family';
import { Table } from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import Image from '@tiptap/extension-image';
import { tiptapToDocx } from '../lib/tiptapToDocx';
import { saveToArchive } from '../lib/archiveUtils';
import { isConnected, upload } from '../lib/oneDriveSync';
import { useRef, useState } from 'react';

const FONT_FAMILIES = ['Arial', 'David', 'Times New Roman', 'Calibri', 'Tahoma', 'Courier New'];
const FONT_SIZES    = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];
const BASE = import.meta.env.BASE_URL || '/';

function readAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = e => resolve(e.target.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

// ── Toolbar button ─────────────────────────────────────────────────────────────
function TBtn({ active, onClick, title, children, disabled }) {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onMouseDown={e => { e.preventDefault(); onClick(); }}
      className={`editor-tbtn${active ? ' active' : ''}`}
    >
      {children}
    </button>
  );
}

function TDivider() { return <span className="editor-tdivider" />; }

// ── Main component ─────────────────────────────────────────────────────────────
// ── Name-before-export dialog ─────────────────────────────────────────────────
function NameDialog({ defaultValue, onConfirm, onCancel }) {
  const [val, setVal] = useState(defaultValue || '');
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-sheet" style={{ maxWidth: 360 }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span>שם המסמך</span>
          <button className="btn-icon" onClick={onCancel}>✕</button>
        </div>
        <div className="modal-body">
          <input
            className="form-input"
            value={val}
            onChange={e => setVal(e.target.value)}
            placeholder="הזן שם מסמך..."
            dir="rtl"
            autoFocus
            onKeyDown={e => { if (e.key === 'Enter') onConfirm(val.trim() || 'מסמך חופשי'); }}
          />
          <button
            className="btn btn-primary"
            style={{ width: '100%', marginTop: 12 }}
            onClick={() => onConfirm(val.trim() || 'מסמך חופשי')}
          >
            ⬇️ שמור והורד
          </button>
        </div>
      </div>
    </div>
  );
}

export default function FreeEditor({ onBack }) {
  const [title, setTitle]         = useState('');
  const [saving, setSaving]       = useState(false);
  const [msg, setMsg]             = useState('');
  const [showNameDlg, setShowNameDlg] = useState(false);
  const fileInputRef   = useRef(null);
  const cameraInputRef = useRef(null);

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ heading: { levels: [1, 2, 3] } }),
      Underline,
      TextAlign.configure({ types: ['heading', 'paragraph'], defaultAlignment: 'right' }),
      TextStyle,
      Color,
      FontFamily,
      FontSize,
      Table.configure({ resizable: true }),
      TableRow,
      TableHeader,
      TableCell,
      Image.configure({ inline: false, allowBase64: true }),
    ],
    content: '<p dir="rtl"></p>',
    editorProps: {
      attributes: { dir: 'rtl', class: 'editor-content-area' },
      handlePaste(view, event) {
        const items = event.clipboardData?.items;
        if (!items) return false;
        for (const item of items) {
          if (item.type.startsWith('image/')) {
            const file = item.getAsFile();
            if (file) {
              readAsDataURL(file).then(src => {
                view.dispatch(view.state.tr.replaceSelectionWith(
                  view.state.schema.nodes.image.create({ src })
                ));
              });
              return true;
            }
          }
        }
        return false;
      },
    },
  });

  if (!editor) return null;

  // ── Image insertion ────────────────────────────────────────────────────────
  const insertImageFromFile = async (file) => {
    if (!file || !file.type.startsWith('image/')) return;
    const src = await readAsDataURL(file);
    editor.chain().focus().setImage({ src }).run();
  };

  // ── Export (opens name dialog first) ─────────────────────────────────────
  const handleExport = () => setShowNameDlg(true);

  const doExport = async (name) => {
    setShowNameDlg(false);
    setTitle(name);
    setSaving(true); setMsg('');
    try {
      const json = editor.getJSON();
      const blob = await tiptapToDocx(json, { title: name });
      const fname = `${name}.docx`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fname; a.click();
      URL.revokeObjectURL(url);
      saveToArchive({ type: 'report', filename: fname, client: name, docType: 'free' }, null);
      if (isConnected()) upload().catch(() => {});
      setMsg('✅ המסמך הורד!');
    } catch (e) {
      setMsg('❌ שגיאה: ' + e.message);
    } finally {
      setSaving(false);
    }
  };

  // ── Table helpers ──────────────────────────────────────────────────────────
  const insertTable = () =>
    editor.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();

  // ── Current font size/family display ──────────────────────────────────────
  // fontSize stored as "14pt" string; parse to number for select comparison
  const curFontSize = parseInt(editor.getAttributes('textStyle').fontSize) || 10;
  const curFont     = editor.getAttributes('textStyle').fontFamily || 'Arial';

  return (
    <div className="free-editor-wrap">
      {/* Hidden file inputs */}
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        style={{ display: 'none' }}
        onChange={e => { insertImageFromFile(e.target.files?.[0]); e.target.value = ''; }}
      />
      <input
        ref={cameraInputRef}
        type="file"
        accept="image/*"
        capture="environment"
        style={{ display: 'none' }}
        onChange={e => { insertImageFromFile(e.target.files?.[0]); e.target.value = ''; }}
      />

      {/* ── Header ── */}
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="editor-title-label" dir="rtl">
          {title || 'עורך חופשי'}
        </span>
        <button
          className="btn btn-success btn-sm"
          onClick={handleExport}
          disabled={saving}
        >
          {saving ? '⏳' : '⬇️ ייצא'}
        </button>
      </div>
      {msg && <div className="sync-msg" style={{ margin: '0 12px 8px' }}>{msg}</div>}

      {/* ── Toolbar ── */}
      <div className="editor-toolbar">
        {/* Undo/Redo */}
        <TBtn onClick={() => editor.chain().focus().undo().run()} title="בטל" disabled={!editor.can().undo()}>↩</TBtn>
        <TBtn onClick={() => editor.chain().focus().redo().run()} title="חזור" disabled={!editor.can().redo()}>↪</TBtn>
        <TDivider />

        {/* Paragraph style */}
        <select
          className="editor-select"
          value={editor.isActive('heading', { level: 1 }) ? 'h1'
               : editor.isActive('heading', { level: 2 }) ? 'h2'
               : editor.isActive('heading', { level: 3 }) ? 'h3' : 'p'}
          onChange={e => {
            const v = e.target.value;
            if (v === 'p') editor.chain().focus().setParagraph().run();
            else editor.chain().focus().toggleHeading({ level: parseInt(v[1]) }).run();
          }}
        >
          <option value="p">רגיל</option>
          <option value="h1">כותרת 1</option>
          <option value="h2">כותרת 2</option>
          <option value="h3">כותרת 3</option>
        </select>
        <TDivider />

        {/* Font family */}
        <select
          className="editor-select"
          value={curFont}
          onChange={e => editor.chain().focus().setFontFamily(e.target.value).run()}
          style={{ minWidth: 90 }}
        >
          {FONT_FAMILIES.map(f => <option key={f} value={f}>{f}</option>)}
        </select>

        {/* Font size */}
        <select
          className="editor-select editor-select--sm"
          value={curFontSize}
          onChange={e => editor.chain().focus().setFontSize(`${e.target.value}pt`).run()}
        >
          {FONT_SIZES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <TDivider />

        {/* Formatting */}
        <TBtn active={editor.isActive('bold')} onClick={() => editor.chain().focus().toggleBold().run()} title="מודגש"><strong>B</strong></TBtn>
        <TBtn active={editor.isActive('italic')} onClick={() => editor.chain().focus().toggleItalic().run()} title="נטוי"><em>I</em></TBtn>
        <TBtn active={editor.isActive('underline')} onClick={() => editor.chain().focus().toggleUnderline().run()} title="קו תחתון"><u>U</u></TBtn>
        <TBtn active={editor.isActive('strike')} onClick={() => editor.chain().focus().toggleStrike().run()} title="קו חוצה"><s>S</s></TBtn>
        <TDivider />

        {/* Color */}
        <label className="editor-color-label" title="צבע טקסט">
          <span style={{ borderBottom: `3px solid ${editor.getAttributes('textStyle').color || '#000'}` }}>A</span>
          <input type="color" className="editor-color-input"
            value={editor.getAttributes('textStyle').color || '#000000'}
            onChange={e => editor.chain().focus().setColor(e.target.value).run()} />
        </label>
        <TDivider />

        {/* Alignment */}
        <TBtn active={editor.isActive({ textAlign: 'right' })} onClick={() => editor.chain().focus().setTextAlign('right').run()} title="יישור ימין">⮕</TBtn>
        <TBtn active={editor.isActive({ textAlign: 'center' })} onClick={() => editor.chain().focus().setTextAlign('center').run()} title="מרכוז">≡</TBtn>
        <TBtn active={editor.isActive({ textAlign: 'left' })} onClick={() => editor.chain().focus().setTextAlign('left').run()} title="יישור שמאל">⬅</TBtn>
        <TBtn active={editor.isActive({ textAlign: 'justify' })} onClick={() => editor.chain().focus().setTextAlign('justify').run()} title="מוצדק">▤</TBtn>
        <TDivider />

        {/* Lists */}
        <TBtn active={editor.isActive('bulletList')} onClick={() => editor.chain().focus().toggleBulletList().run()} title="רשימת נקודות">• —</TBtn>
        <TBtn active={editor.isActive('orderedList')} onClick={() => editor.chain().focus().toggleOrderedList().run()} title="רשימה ממוספרת">1. —</TBtn>
        <TDivider />

        {/* Table */}
        <TBtn onClick={insertTable} title="הוסף טבלה">⊞</TBtn>
        {editor.isActive('table') && (
          <>
            <TBtn onClick={() => editor.chain().focus().addRowAfter().run()} title="הוסף שורה למטה">+שורה</TBtn>
            <TBtn onClick={() => editor.chain().focus().addColumnAfter().run()} title="הוסף עמודה">+עמודה</TBtn>
            <TBtn onClick={() => editor.chain().focus().deleteRow().run()} title="מחק שורה">−שורה</TBtn>
            <TBtn onClick={() => editor.chain().focus().deleteColumn().run()} title="מחק עמודה">−עמודה</TBtn>
            <label className="editor-color-label" title="צבע תא">
              <span style={{ fontSize: 11 }}>תא</span>
              <input type="color" className="editor-color-input"
                defaultValue="#ffffff"
                onChange={e => editor.chain().focus().setCellAttribute('backgroundColor', e.target.value).run()} />
            </label>
            <TBtn onClick={() => editor.chain().focus().deleteTable().run()} title="מחק טבלה">🗑</TBtn>
          </>
        )}
        <TDivider />

        {/* Images */}
        <TBtn onClick={() => fileInputRef.current?.click()} title="הוסף תמונה מהגלריה">🖼</TBtn>
        <TBtn onClick={() => cameraInputRef.current?.click()} title="צלם תמונה">📷</TBtn>
      </div>

      {/* ── Editor canvas ── */}
      <div className="editor-page-wrapper">
        <div className="editor-page">

          {/* Header logo preview */}
          <div className="editor-logo-area editor-logo-area--header">
            <div className="editor-page-num-preview">1/…</div>
            <img
              src={`${BASE}header-logo.jpg`}
              alt="כותרת"
              className="editor-logo-img"
              onError={e => { e.currentTarget.style.display = 'none'; }}
            />
          </div>

          <div className="editor-logo-divider" />

          {/* Content */}
          <EditorContent editor={editor} />

          <div className="editor-logo-divider" />

          {/* Footer logo preview */}
          <div className="editor-logo-area editor-logo-area--footer">
            <img
              src={`${BASE}footer-logo.png`}
              alt="כותרת תחתונה"
              className="editor-logo-img editor-logo-img--footer"
              onError={e => { e.currentTarget.style.display = 'none'; }}
            />
          </div>

        </div>
      </div>

      {/* ── Name dialog ── */}
      {showNameDlg && (
        <NameDialog
          defaultValue={title}
          onConfirm={doExport}
          onCancel={() => setShowNameDlg(false)}
        />
      )}
    </div>
  );
}
