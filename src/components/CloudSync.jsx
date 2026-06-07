import { useState, useEffect } from 'react';
import {
  isConnected, getClientId, saveClientId, getLastSync,
  startAuth, disconnect, upload, downloadAndMerge,
} from '../lib/oneDriveSync';

function fmtDate(ts) {
  if (!ts) return '—';
  return new Date(ts).toLocaleString('he-IL', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });
}

export default function CloudSync({ onClose, onSynced }) {
  const [clientId, setClientId]   = useState(getClientId());
  const [connected, setConnected] = useState(isConnected());
  const [busy, setBusy]           = useState('');
  const [msg, setMsg]             = useState('');
  const [lastSync, setLastSync]   = useState(getLastSync());

  useEffect(() => { setConnected(isConnected()); setLastSync(getLastSync()); }, []);

  const handleConnect = async () => {
    if (!clientId.trim()) return setMsg('יש להזין Client ID');
    saveClientId(clientId.trim());
    await startAuth(clientId.trim());
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
              מאפשר לסנכרן ארכיון, יומן שטח וכרטיסיות בין כל המכשירים שלך דרך OneDrive.
            </p>

            <div className="sync-steps">
              <p className="sync-steps-title">הגדרה חד-פעמית (5 דקות):</p>
              <ol>
                <li>היכנס ל-<strong>portal.azure.com</strong></li>
                <li>Azure Active Directory ← App registrations ← New registration</li>
                <li>שם: כל שם, סוג: <em>Personal Microsoft accounts only</em></li>
                <li>Redirect URI → <em>Single-page application (SPA)</em> →<br/>
                  <code>https://nirengineering.github.io/nir-reports/</code>
                </li>
                <li>API permissions → Add → Microsoft Graph → Delegated → <strong>Files.ReadWrite</strong> + <strong>offline_access</strong></li>
                <li>העתק את <strong>Application (client) ID</strong> והדבק כאן:</li>
              </ol>
            </div>

            <div className="form-group" style={{ marginTop: 12 }}>
              <label className="form-label">Client ID</label>
              <input
                className="form-input"
                value={clientId}
                onChange={e => setClientId(e.target.value)}
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                style={{ fontFamily: 'monospace', fontSize: 13 }}
              />
            </div>

            {msg && <div className="sync-msg">{msg}</div>}

            <button
              className="btn btn-primary"
              style={{ width: '100%', marginTop: 12 }}
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
