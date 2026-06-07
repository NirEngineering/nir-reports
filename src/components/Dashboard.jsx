import { loadArchive } from '../lib/archiveUtils';
import { DOC_TYPES_CONFIG } from '../constants';

export default function Dashboard({ onBack }) {
  const entries = loadArchive();

  // ── Basic counts ──────────────────────────────────────────────────────────
  const now = new Date();

  const thisMonth = entries.filter(e => {
    const d = new Date(e.createdAt);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  }).length;

  const thisYear = entries.filter(e =>
    new Date(e.createdAt).getFullYear() === now.getFullYear()
  ).length;

  const total = entries.length;

  // ── Count by docType ──────────────────────────────────────────────────────
  const byDocType = {};
  entries.forEach(e => {
    const key = e.docType || 'unknown';
    byDocType[key] = (byDocType[key] || 0) + 1;
  });

  // Top 5 doc types sorted by count desc
  const topDocTypes = Object.entries(byDocType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  const maxDocTypeCount = topDocTypes.length > 0 ? topDocTypes[0][1] : 1;

  // ── Top 5 clients ─────────────────────────────────────────────────────────
  const byClient = {};
  entries.forEach(e => {
    const name = (e.client || '').trim();
    if (name) byClient[name] = (byClient[name] || 0) + 1;
  });

  const topClients = Object.entries(byClient)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  // ── Recent 5 documents ────────────────────────────────────────────────────
  const recent5 = entries.slice(0, 5);

  const formatDate = (ts) => {
    if (!ts) return '';
    const d = new Date(ts);
    return `${d.getDate()}.${d.getMonth() + 1}.${d.getFullYear()}`;
  };

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="app-body">
      {/* Header */}
      <div className="field-journal-header">
        <button className="btn btn-outline btn-sm" onClick={onBack}>◀ חזור</button>
        <span className="field-journal-title">📊 לוח בקרה</span>
      </div>

      {/* Stat grid */}
      <div className="card">
        <div className="card-title">📈 סטטיסטיקות</div>
        <div className="dash-stat-grid">
          <div className="dash-stat-card">
            <div className="dash-stat-num">{thisMonth}</div>
            <div className="dash-stat-label">החודש</div>
          </div>
          <div className="dash-stat-card">
            <div className="dash-stat-num">{thisYear}</div>
            <div className="dash-stat-label">השנה</div>
          </div>
          <div className="dash-stat-card">
            <div className="dash-stat-num">{total}</div>
            <div className="dash-stat-label">סה"כ</div>
          </div>
        </div>
      </div>

      {/* Top 5 doc types */}
      {topDocTypes.length > 0 && (
        <div className="card">
          <div className="card-title">📋 לפי סוג מסמך (top 5)</div>
          {topDocTypes.map(([key, count]) => {
            const cfg = DOC_TYPES_CONFIG[key];
            const label = cfg ? cfg.name.replace('\n', ' ') : key;
            const icon = cfg ? cfg.icon : '📄';
            const pct = Math.round((count / maxDocTypeCount) * 100);
            return (
              <div key={key} className="dash-row">
                <span style={{ fontSize: 18 }}>{icon}</span>
                <span className="dash-row-label">{label}</span>
                <div className="dash-bar">
                  <div className="dash-bar-fill" style={{ width: `${pct}%` }} />
                </div>
                <span className="dash-row-count">{count}</span>
              </div>
            );
          })}
        </div>
      )}

      {/* Top 5 clients */}
      {topClients.length > 0 && (
        <div className="card">
          <div className="card-title">🏢 לקוחות מובילים</div>
          {topClients.map(([name, count]) => (
            <div key={name} className="dash-row">
              <span className="dash-row-label" style={{ minWidth: 0, flex: 1 }}>{name}</span>
              <span className="dash-row-count">{count}</span>
            </div>
          ))}
        </div>
      )}

      {/* Recent 5 documents */}
      {recent5.length > 0 && (
        <div className="card">
          <div className="card-title">🕐 מסמכים אחרונים</div>
          {recent5.map(e => {
            const cfg = DOC_TYPES_CONFIG[e.docType];
            const icon = cfg ? cfg.icon : '📄';
            const name = e.filename || e.client || '—';
            return (
              <div key={e.id} className="dash-row" style={{ alignItems: 'flex-start' }}>
                <span style={{ fontSize: 16, marginTop: 1 }}>{icon}</span>
                <div style={{ flex: 1, minWidth: 0, direction: 'rtl' }}>
                  <div style={{ fontSize: 13, fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</div>
                  <div style={{ fontSize: 11, color: '#94a3b8' }}>{formatDate(e.createdAt)}</div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {total === 0 && (
        <div className="card" style={{ textAlign: 'center', color: '#94a3b8', padding: '32px 20px' }}>
          <div style={{ fontSize: 40, marginBottom: 8 }}>📭</div>
          <div>אין מסמכים בארכיון עדיין</div>
        </div>
      )}
    </div>
  );
}
