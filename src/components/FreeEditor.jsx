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
import { tiptapToDocx } from '../lib/tiptapToDocx';
import { saveToArchive } from '../lib/archiveUtils';
import { isConnected, upload } from '../lib/oneDriveSync';
import { useState } from 'react';

const FONT_FAMILIES = ['Arial', 'David', 'Times New Roman', 'Calibri', 'Tahoma', 'Courier New'];
const FONT_SIZES    = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];

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
export default function FreeEditor({ onBack }) {
  const [title, setTitle]   = useState('');
  const [saving, setSaving] = useState(false);
  const [msg, setMsg]       = useState('');

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
    ],
    content: '<p dir="rtl"></p>',
    editorProps: {
      attributes: { dir: 'rtl', class: 'editor-content-area' },
    },
  });

  if (!editor) return null;

  // ── Export ─────────────────────────────────────────────────────────────────
  const handleExport = async () => {
    setSaving(true); setMsg('');
    try {
      const json = editor.getJSON();
      const blob = await tiptapToDocx(json, { title });
      const fname = `${title || 'מסמך חופשי'}.docx`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fname; a.click();
      URL.revokeObjectURL(url);
      saveToArchive({ type: 'report', filename: fname, client: title, docType: 'free' }, null);
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

  // ── Current font size display ──────────────────────────────────────────────
  const curFontSize = editor.getAttributes('textStyle').fontSize || 10;
  const curFont     = editor.getAttributes('textStyle').fontFamily || 'Arial';

  return (
    <div className="free-editor-wrap">
      {/* ── Header ── */}
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <input
          className="editor-title-input"
          value={title}
          onChange={e => setTitle(e.target.value)}
          placeholder="שם המסמך..."
          dir="rtl"
        />
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
          onChange={e => editor.chain().focus().setFontSize(parseInt(e.target.value)).run()}
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
      </div>

      {/* ── Editor canvas ── */}
      <div className="editor-page-wrapper">
        <div className="editor-page">
          <EditorContent editor={editor} />
        </div>
      </div>
    </div>
  );
}
