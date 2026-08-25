import 'dotenv/config';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import qrcodeTerminal from 'qrcode-terminal';
import QRCode from 'qrcode';
import pkg from 'whatsapp-web.js';

import { DOC_TYPES, matchTypeHint, listTypesMessage } from './docTypes.js';
import { getSession, addText, addPhoto, resetSession, isEmpty } from './session.js';
import { classifyAndBuild } from './classify.js';
import { generateDocument } from './docGenerator.js';

const { Client, LocalAuth, MessageMedia } = pkg;
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, '..', 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Config ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3211;
const TARGET_CHAT = (process.env.WHATSAPP_TARGET_CHAT || '').trim();
const GENERATE_KEYWORD = (process.env.GENERATE_KEYWORD || 'צור דוח').trim();
const RESET_KEYWORD = (process.env.RESET_KEYWORD || 'דוח חדש').trim();
const HELP_KEYWORDS = new Set(['עזרה', 'סוגי מסמכים', 'help']);

if (!TARGET_CHAT) {
  console.error(
    '\n❌ WHATSAPP_TARGET_CHAT is not set.\n' +
    '   Copy .env.example to .env, set it to the exact WhatsApp chat name to watch, and re-run.\n'
  );
  process.exit(1);
}

// ── WhatsApp client ─────────────────────────────────────────────────────────
let latestQr = null;
let isReady = false;

// Chromium leaves a SingletonLock (and friends) in the profile dir and only
// removes it on a clean exit. If the container was killed (crash, `docker
// compose down` timeout) instead of shutting down gracefully, the lock
// survives in the persistent volume and blocks the next start with
// "profile appears to be in use by another process". Since only one instance
// of this container ever runs against this profile, any lock found here at
// startup is necessarily stale — safe to remove before launching.
function clearStaleChromiumLock() {
  const sessionDir = path.join('.wwebjs_auth', 'session');
  for (const name of ['SingletonLock', 'SingletonCookie', 'SingletonSocket']) {
    try { fs.rmSync(path.join(sessionDir, name), { force: true }); } catch (_) { /* ignore */ }
  }
}
clearStaleChromiumLock();

const client = new Client({
  authStrategy: new LocalAuth({ dataPath: '.wwebjs_auth' }),
  puppeteer: {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    // Set by the Docker image to use the system Chromium instead of downloading one.
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
  },
});

client.on('qr', (qr) => {
  latestQr = qr;
  console.log('\nScan this QR code with WhatsApp → Linked Devices → Link a Device:\n');
  qrcodeTerminal.generate(qr, { small: true });
});

client.on('ready', () => {
  isReady = true;
  latestQr = null;
  console.log(`✅ Field report bot connected. Watching chat: "${TARGET_CHAT}"`);
});

client.on('disconnected', (reason) => {
  isReady = false;
  console.warn('⚠️ WhatsApp disconnected:', reason);
});

function parseCommand(body) {
  const trimmed = body.trim();
  if (!trimmed) return null;
  if (trimmed === RESET_KEYWORD) return { type: 'reset' };
  if (HELP_KEYWORDS.has(trimmed)) return { type: 'help' };
  if (trimmed === GENERATE_KEYWORD) return { type: 'generate', typeHint: null };
  if (trimmed.startsWith(GENERATE_KEYWORD)) {
    const rest = trimmed.slice(GENERATE_KEYWORD.length).trim().replace(/^:\s*/, '');
    return { type: 'generate', typeHint: rest || null };
  }
  return null;
}

async function handleGenerate(msg, typeHint) {
  let forcedTypeId = null;
  if (typeHint) {
    forcedTypeId = matchTypeHint(typeHint);
    if (!forcedTypeId) {
      await msg.reply(`❓ לא זיהיתי את הסוג "${typeHint}".\n\n${listTypesMessage()}`);
      return;
    }
  }

  if (isEmpty()) {
    await msg.reply('אין עדיין הערות או תמונות לניתוח — שלח טקסט ותמונות מהביקור, ואז שלח שוב "' + GENERATE_KEYWORD + '".');
    return;
  }

  await msg.reply('⏳ מנתח את ההערות והתמונות ומייצר דוח... (עד כדקה)');

  try {
    const session = getSession();
    const payload = await classifyAndBuild(session, forcedTypeId);
    const buffer = await generateDocument(payload);

    const typeMeta = DOC_TYPES[payload.doc_type] || DOC_TYPES.group7;
    const safeClient = (payload.client || 'דוח').replace(/[\\/:*?"<>|]/g, ' ').trim();
    const filename = `${typeMeta.name} - ${safeClient} - ${payload.date}.docx`;
    fs.writeFileSync(path.join(DATA_DIR, filename), buffer);

    const media = new MessageMedia(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer.toString('base64'),
      filename
    );
    await msg.reply(media, undefined, { caption: `✅ ${typeMeta.name}${payload.client ? ' — ' + payload.client : ''}` });

    resetSession();
  } catch (e) {
    console.error('Report generation failed:', e);
    await msg.reply(`❌ שגיאה ביצירת הדוח: ${e.message}\n\nההערות והתמונות נשמרו — נסה שוב "${GENERATE_KEYWORD}".`);
  }
}

client.on('message', async (msg) => {
  if (!isReady) return;
  try {
    const chat = await msg.getChat();
    if ((chat.name || '').trim() !== TARGET_CHAT) return;

    const body = (msg.body || '').trim();
    const command = parseCommand(body);

    if (command?.type === 'reset') {
      resetSession();
      await msg.reply('🆕 האוסף אופס. שלח טקסט ותמונות מהביקור, ואז "' + GENERATE_KEYWORD + '" (אפשר גם "' + GENERATE_KEYWORD + ': סככות" וכו\').');
      return;
    }

    if (command?.type === 'help') {
      await msg.reply(listTypesMessage());
      return;
    }

    if (command?.type === 'generate') {
      await handleGenerate(msg, command.typeHint);
      return;
    }

    // Not a command — buffer as field data
    if (msg.hasMedia) {
      const media = await msg.downloadMedia().catch(() => null);
      if (media?.mimetype?.startsWith('image/')) {
        addPhoto(`data:${media.mimetype};base64,${media.data}`, body || '');
      }
    } else if (body) {
      addText(body);
    }
  } catch (e) {
    console.error('Failed to handle incoming WhatsApp message:', e);
  }
});

client.initialize();

// Close Chromium cleanly on stop/restart so it removes its own lock file —
// the primary defense against the stale-lock problem clearStaleChromiumLock()
// works around above (belt and suspenders).
let shuttingDown = false;
async function shutdown(signal) {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`Received ${signal}, shutting down...`);
  try { await client.destroy(); } catch (_) { /* already gone */ }
  process.exit(0);
}
process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));

// ── Small status HTTP server (QR code viewing / health check only) ─────────
const app = express();

app.get('/status', (req, res) => {
  const session = getSession();
  res.json({ ready: isReady, targetChat: TARGET_CHAT, pendingTexts: session.texts.length, pendingPhotos: session.photos.length });
});

app.get('/qr', async (req, res) => {
  if (isReady) return res.json({ ready: true });
  if (!latestQr) return res.status(404).json({ error: 'QR not generated yet — check the server console.' });
  const dataUrl = await QRCode.toDataURL(latestQr);
  res.json({ ready: false, qr: dataUrl });
});

app.listen(PORT, () => {
  console.log(`Field report bot status API listening on http://localhost:${PORT}`);
});
