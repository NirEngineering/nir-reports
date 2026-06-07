import { useState, useEffect } from 'react';
import {
  isConnected, getClientId, getLastSync,
  startAuth, disconnect, upload, downloadAndMerge,
} from '../lib/oneDriveSync';

function fmtDate(ts) {
  if (!ts) return '—';
  return new Date(ts).toLocaleString('he-IL', {
    day: '2-digit', month: '2-digit', year: '2-digit',
    hour: '2-digit', minute: '2-digit',
  });
}

export default function CloudSync({ onClose, onSynced }) {
  const [connected, setConnected] = useState(isConnected());
  const [busy, setBusy]           = useState('');
  const [msg, setMsg]             = useState('');
  const [lastSync, setLastSync]   = useState(getLastSync());

  useEffect(() => { setConnected(isConnected()); setLastSync(getLastSync()); }, []);

  const handleConnect = async () => {
    await startAuth(getClientId());
  };

  const handleUpload = async () => {
    setBusy('upload'); setMsg('');
    try {
      await upload();
      setLastSync(getLastSync());
      setMsg('✅ הנתונים הועלו ל-OneDrive');
    } catch (e) {
      setMsg('❌ ' + e.message);
    } finally { setBusy(''); }
  };

  const handleDownload = async () => {
    setBusy('download'); setMsg('');
    try {
      const ts = await downloadAndMerge();
      setLastSync(getLastSync());
      setMsg(ts ? `✅ סנכרון הושלם (גיבוי מ-${fmtDate(ts)})` : '⚠️ לא נמצא גיבוי ב-OneDrive');
      onSynced?.();
    } catch (e) {
      setMsg('❌ ' + e.message);
    } finally { setBusy(''); }
  };

  const handleDisconnect = () => {
    disconnect();
    setConnected(false);
    setLastSync(null);
    setMsg('');
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-sheet" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span>☁️ סנכרון OneDrive</span>
          <button className="btn-icon" onClick={onClose}>✕</button>
        </div>

        {!connected ? (
          <div className="modal-body">
            <p className="sync-hint">
              סנכרן ארכיון, יומן שטח וכרטיסיות בין כל המכשירים שלך דרך OneDrive האישי שלך.
            </p>
            <p className="sync-hint" style={{ marginTop: 0 }}>
              לחץ "התחבר" — תועבר לדף הכניסה של Microsoft, ואחרי אישור תחזור אוטומטית.
            </p>

            {msg && <div className="sync-msg">{msg}</div>}

            <button
              className="btn btn-primary"
              style={{ width: '100%', marginTop: 8 }}
              onClick={handleConnect}
            >
              🔑 התחבר ל-OneDrive
            </button>
          </div>
        ) : (
          <div className="modal-body">
            <div className="sync-status-row">
              <span className="sync-dot connected" />
              <span>מחובר ל-OneDrive</span>
            </div>
            <div className="sync-last">סנכרון אחרון: {fmtDate(lastSync)}</div>

            {msg && <div className="sync-msg">{msg}</div>}

            <div style={{ display: 'flex', gap: 8, flexDirection: 'column', marginTop: 16 }}>
              <button
                className="btn btn-primary"
                onClick={handleUpload}
                disabled={!!busy}
              >
                {busy === 'upload' ? '⏳ מעלה...' : '⬆️ העלה לעכשיו'}
              </button>
              <button
                className="btn btn-outline"
                onClick={handleDownload}
                disabled={!!busy}
              >
                {busy === 'download' ? '⏳ מוריד...' : '⬇️ סנכרן מ-OneDrive'}
              </button>
              <button
                className="btn btn-outline"
                style={{ color: '#ef4444', borderColor: '#ef4444', marginTop: 8 }}
                onClick={handleDisconnect}
              >
                התנתק
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
