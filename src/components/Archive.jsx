import { useState } from 'react';
import { loadArchive, deleteFromArchive } from '../lib/archiveUtils';
import { DOC_TYPES_CONFIG } from '../constants';

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

  const handleDelete = (id) => {
    deleteFromArchive(id);
    setEntries(prev => prev.filter(e => e.id !== id));
  };

  const handleClearAll = () => {
    if (!window.confirm('למחוק את כל הרשומות בארכיון?')) return;
    entries.forEach(e => deleteFromArchive(e.id));
    setEntries([]);
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
            const icon = TYPE_ICON[entry.type] || '📄';
            const typeLabel = TYPE_LABEL[entry.type] || entry.type;
            const docLabel = entry.docType
              ? DOC_TYPES_CONFIG[entry.docType]?.name?.replace('\n', ' ')
              : '';
            const nameLabel = entry.client || entry.title || '';

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
                </div>

                <div className="journal-item-actions">
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
