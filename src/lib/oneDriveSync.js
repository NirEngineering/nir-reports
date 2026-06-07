const DEFAULT_CLIENT_ID = '408784d0-cbff-4b33-a9ac-871879a6e86e';
const OD_CLIENT_ID_KEY = 'nir_od_client_id';
const OD_TOKENS_KEY    = 'nir_od_tokens';
const OD_LAST_SYNC_KEY = 'nir_od_last_sync';
const BACKUP_FILE = 'nir-reports-backup.json';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const AUTH  = 'https://login.microsoftonline.com/common/oauth2/v2.0';
const SCOPE = 'Files.ReadWrite offline_access';

// ── PKCE ──────────────────────────────────────────────────────────────────────
function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}
async function makePKCE() {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const digest   = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  return { verifier, challenge: b64url(digest) };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
const redirectUri = () =>
  window.location.origin + window.location.pathname.replace(/\/?$/, '/');

export const getClientId  = ()  => localStorage.getItem(OD_CLIENT_ID_KEY) || DEFAULT_CLIENT_ID;
export const saveClientId = (v) => localStorage.setItem(OD_CLIENT_ID_KEY, v.trim());
export const getLastSync  = ()  => { const t = localStorage.getItem(OD_LAST_SYNC_KEY); return t ? Number(t) : null; };
export const isConnected  = ()  => !!getTokens()?.refresh_token;

export function getTokens() {
  try { return JSON.parse(localStorage.getItem(OD_TOKENS_KEY)); } catch { return null; }
}

export function disconnect() {
  localStorage.removeItem(OD_TOKENS_KEY);
  localStorage.removeItem(OD_LAST_SYNC_KEY);
}

// ── OAuth PKCE flow ───────────────────────────────────────────────────────────
export async function startAuth(clientId) {
  const { verifier, challenge } = await makePKCE();
  sessionStorage.setItem('od_v', verifier);
  sessionStorage.setItem('od_cid', clientId);
  const p = new URLSearchParams({
    client_id: clientId, response_type: 'code',
    redirect_uri: redirectUri(), scope: SCOPE,
    code_challenge: challenge, code_challenge_method: 'S256',
    response_mode: 'query',
  });
  window.location.href = `${AUTH}/authorize?${p}`;
}

export async function handleCallback() {
  const params   = new URLSearchParams(location.search);
  const code     = params.get('code');
  if (!code) return false;

  const verifier = sessionStorage.getItem('od_v');
  const clientId = sessionStorage.getItem('od_cid') || getClientId();
  if (!verifier) return false;

  sessionStorage.removeItem('od_v');
  sessionStorage.removeItem('od_cid');
  history.replaceState(null, '', location.pathname);

  const resp = await fetch(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId, grant_type: 'authorization_code',
      code, redirect_uri: redirectUri(), code_verifier: verifier,
    }),
  });
  if (!resp.ok) throw new Error('Auth failed: ' + resp.status);
  const d = await resp.json();
  saveClientId(clientId);
  localStorage.setItem(OD_TOKENS_KEY, JSON.stringify({
    access_token:  d.access_token,
    refresh_token: d.refresh_token,
    expires_at:    Date.now() + d.expires_in * 1000,
  }));
  return true;
}

async function refreshToken() {
  const t  = getTokens();
  const id = getClientId();
  if (!t?.refresh_token || !id) throw new Error('Not connected');
  const resp = await fetch(`${AUTH}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: id, grant_type: 'refresh_token',
      refresh_token: t.refresh_token, scope: SCOPE,
    }),
  });
  if (!resp.ok) { disconnect(); throw new Error('Session expired — please reconnect'); }
  const d = await resp.json();
  const updated = {
    access_token:  d.access_token,
    refresh_token: d.refresh_token || t.refresh_token,
    expires_at:    Date.now() + d.expires_in * 1000,
  };
  localStorage.setItem(OD_TOKENS_KEY, JSON.stringify(updated));
  return updated.access_token;
}

async function getToken() {
  const t = getTokens();
  if (!t) throw new Error('Not connected');
  return Date.now() < t.expires_at - 60_000 ? t.access_token : refreshToken();
}

// ── Backup helpers ────────────────────────────────────────────────────────────
export function collectBackup() {
  const data = {};
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (k && k.startsWith('nir_') && !k.startsWith('nir_od_') && !k.startsWith('nir_anthropic_')) {
      data[k] = localStorage.getItem(k);
    }
  }
  return { _v: 2, _ts: Date.now(), data };
}

function mergeKey(key, remoteVal) {
  // For JSON arrays (archive index, field log): merge by entry id, keep local
  try {
    const local  = JSON.parse(localStorage.getItem(key) || '[]');
    const remote = JSON.parse(remoteVal || '[]');
    if (!Array.isArray(local) || !Array.isArray(remote)) return;
    const ids    = new Set(local.map(e => e.id));
    const merged = [...local, ...remote.filter(e => !ids.has(e.id))];
    merged.sort((a, b) => (b.createdAt ?? b.ts ?? 0) - (a.createdAt ?? a.ts ?? 0));
    localStorage.setItem(key, JSON.stringify(merged.slice(0, 500)));
  } catch {
    if (!localStorage.getItem(key)) localStorage.setItem(key, remoteVal);
  }
}

export function applyBackup(backup) {
  if (!backup?.data) return;
  Object.entries(backup.data).forEach(([k, v]) => {
    // Archive payload entries: add only if missing
    if (k.startsWith('nir_v2_arc_d_')) {
      if (!localStorage.getItem(k)) localStorage.setItem(k, v);
      return;
    }
    // Mergeable list keys
    if (k === 'nir_v2_archive' || k === 'nir_v2_field_log' || k === 'nir_v2_info_cards') {
      mergeKey(k, v);
      return;
    }
    // Draft: keep local
    if (k === 'nir_v2_draft') return;
    // Everything else: add if missing locally
    if (!localStorage.getItem(k)) localStorage.setItem(k, v);
  });
}

// ── Graph API calls ───────────────────────────────────────────────────────────
export async function upload() {
  const token  = await getToken();
  const backup = collectBackup();
  const resp   = await fetch(
    `${GRAPH}/me/drive/root:/${BACKUP_FILE}:/content`,
    {
      method:  'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify(backup),
    }
  );
  if (!resp.ok) throw new Error('Upload failed: ' + resp.status);
  localStorage.setItem(OD_LAST_SYNC_KEY, Date.now().toString());
}

export async function downloadAndMerge() {
  const token = await getToken();
  const resp  = await fetch(
    `${GRAPH}/me/drive/root:/${BACKUP_FILE}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (resp.status === 404) return null;
  if (!resp.ok) throw new Error('Download failed: ' + resp.status);
  const backup = await resp.json();
  applyBackup(backup);
  localStorage.setItem(OD_LAST_SYNC_KEY, Date.now().toString());
  return backup._ts ?? null;
}
