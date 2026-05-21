import { useState } from 'react';
import { loadArchive, loadDocData, deleteFromArchive } from '../lib/archiveUtils';
import { DOC_TYPES_CONFIG } from '../constants';
import { generateDocument } from '../lib/docGenerator';
import { generateEventApproval } from '../lib/eventApproval';

const TYPE_ICON = {
  report: '📄',
  event: '🏗️',
  ai: '🤖',
  journal: '📋',
  card: '🗂',
};

const TYPE_LABEL = {
  report: 'דוח',
  event: 'אישור מבנים',
  ai: 'דוח AI',
  journal: 'יומן שטח',
  card: 'כרטיסייה',
};

function formatDate(ts) {
  if (!ts) return '';
  const d = new Date(ts);
  return (
    d.toLocaleDateString('he-IL') +
    ' ' +
    d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' })
  );
}

export default function Archive({ onBack }) {
  const [entries, setEntries] = useState(() => loadArchive());
  const [regenerating, setRegenerating] = useState(null);

  const handleDelete = (id) => {
    deleteFromArchive(id);
    setEntries(prev => prev.filter(e => e.id !== id));
  };

  const handleClearAll = () => {
    if (!window.confirm('למחוק את כל הרשומות בארכיון?')) return;
    entries.forEach(e => deleteFromArchive(e.id));
    setEntries([]);
  };

  const handleRegenerate = async (entry) => {
    const payload = loadDocData(entry.id);
    if (!payload) return;
    setRegenerating(entry.id);
    try {
      let blob;
      if (entry.type === 'event') {
        blob = await generateEventApproval(payload);
      } else {
        blob = await generateDocument(payload);
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = entry.filename || 'מסמך.docx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert(`שגיאה ביצירת המסמך: ${e.message}`);
    } finally {
      setRegenerating(null);
    }
  };

  return (
    <div className="app-body">
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="field-journal-title">🗄️ ארכיון מסמכים</span>
        {entries.length > 0 && (
          <button
            className="btn btn-outline btn-sm"
            style={{ color: '#ef4444' }}
            onClick={handleClearAll}
          >
            🗑 נקה הכל
          </button>
        )}
      </div>

      {entries.length === 0 ? (
        <div className="journal-empty">
          <div className="journal-empty-icon">🗄️</div>
          <p>הארכיון ריק</p>
          <p style={{ fontSize: 13, color: '#94a3b8', marginTop: 8 }}>
            כל מסמך שתיצור יישמר כאן אוטומטית
          </p>
        </div>
      ) : (
        <div className="journal-items">
          {entries.map(entry => {
            const icon      = TYPE_ICON[entry.type] || '📄';
            const typeLabel = TYPE_LABEL[entry.type] || entry.type;
            const docLabel  = entry.docType
              ? DOC_TYPES_CONFIG[entry.docType]?.name?.replace('\n', ' ')
              : '';
            const nameLabel   = entry.client || entry.title || '';
            const canRegen    = entry.hasData && ['report','event','ai'].includes(entry.type);
            const isRegening  = regenerating === entry.id;

            return (
              <div key={entry.id} className="journal-item" style={{ alignItems: 'flex-start' }}>
                <div style={{ fontSize: 24, lineHeight: 1, marginTop: 2 }}>{icon}</div>

                <div className="journal-item-body" style={{ cursor: 'default' }}>
                  <div className="journal-item-title" style={{ fontSize: 14 }}>
                    {nameLabel || typeLabel}
                  </div>
                  <div className="journal-item-date">{formatDate(entry.createdAt)}</div>
                  <div className="journal-item-preview">
                    {[typeLabel, docLabel].filter(Boolean).join(' · ')}
                    {entry.filename ? ` · ${entry.filename}` : ''}
                  </div>
                  {entry.type === 'journal' && (
                    <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}>
                      פתח "יומן שטח" כדי לייצא מחדש
                    </div>
                  )}
                  {entry.type === 'card' && (
                    <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}>
                      פתח "כרטיסיות מידע" כדי לייצא מחדש
                    </div>
                  )}
                </div>

                <div className="journal-item-actions">
                  {canRegen && (
                    <button
                      className="btn btn-sm btn-outline"
                      style={{ fontSize: 12, padding: '4px 8px' }}
                      onClick={() => handleRegenerate(entry)}
                      disabled={isRegening}
                      title="צור והורד מחדש (ללא תמונות)"
                    >
                      {isRegening ? '⏳' : '⬇️ צור מחדש'}
                    </button>
                  )}
                  <button
                    className="btn-icon del"
                    onClick={() => handleDelete(entry.id)}
                    title="מחק רשומה"
                  >🗑</button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
