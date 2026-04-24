import { useRef, useEffect, useCallback } from 'react';

const FONT_SIZES = [
  { label: 'קטן',    value: '1' },
  { label: 'רגיל',   value: '3' },
  { label: 'גדול',   value: '5' },
  { label: 'כותרת',  value: '7' },
];

export default function FieldNotesEditor({ content, onChange }) {
  const editorRef = useRef(null);
  const savedRange = useRef(null);

  // Sync external content into editor on mount only
  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== content) {
      editorRef.current.innerHTML = content || '';
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const saveRange = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) savedRange.current = sel.getRangeAt(0).cloneRange();
  };

  const restoreRange = () => {
    const sel = window.getSelection();
    if (sel && savedRange.current) {
      sel.removeAllRanges();
      sel.addRange(savedRange.current);
    }
  };

  const exec = useCallback((cmd, value) => {
    restoreRange();
    document.execCommand(cmd, false, value ?? null);
    editorRef.current?.focus();
    onChange(editorRef.current?.innerHTML || '');
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [onChange]);

  const handleInput = () => {
    saveRange();
    onChange(editorRef.current?.innerHTML || '');
  };

  const ToolBtn = ({ title, cmd, value, children }) => (
    <button
      type="button"
      className="fn-tool-btn"
      title={title}
      onMouseDown={e => { e.preventDefault(); saveRange(); exec(cmd, value); }}
    >
      {children}
    </button>
  );

  return (
    <div className="fn-editor-wrap">
      <div className="fn-toolbar">
        <ToolBtn cmd="bold" title="הדגש"><b>B</b></ToolBtn>
        <ToolBtn cmd="underline" title="קו תחתון"><u>U</u></ToolBtn>
        <div className="fn-tool-sep" />
        <ToolBtn cmd="justifyRight"  title="יישר ימינה">⟹</ToolBtn>
        <ToolBtn cmd="justifyCenter" title="מרכז">≡</ToolBtn>
        <ToolBtn cmd="justifyLeft"   title="יישר שמאלה">⟸</ToolBtn>
        <div className="fn-tool-sep" />
        <ToolBtn cmd="insertUnorderedList" title="רשימת נקודות">•</ToolBtn>
        <ToolBtn cmd="insertOrderedList"   title="רשימה ממוספרת">1.</ToolBtn>
        <div className="fn-tool-sep" />
        <select
          className="fn-tool-select"
          title="גודל גופן"
          onMouseDown={saveRange}
          onChange={e => exec('fontSize', e.target.value)}
          defaultValue="3"
        >
          {FONT_SIZES.map(f => (
            <option key={f.value} value={f.value}>{f.label}</option>
          ))}
        </select>
        <div className="fn-tool-sep" />
        <ToolBtn cmd="removeFormat" title="נקה עיצוב">✕</ToolBtn>
      </div>
      <div
        ref={editorRef}
        contentEditable
        suppressContentEditableWarning
        dir="rtl"
        className="fn-editor-content"
        onInput={handleInput}
        onKeyUp={saveRange}
        onMouseUp={saveRange}
        data-placeholder="כתוב כאן ממצאים, מידות, הערות שטח..."
      />
    </div>
  );
}
