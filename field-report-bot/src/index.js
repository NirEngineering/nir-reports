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
import { isWizardActive, startWizard, cancelWizard, answerWizard, getLastKnownDocType, clearLastKnownDocType } from './wizard.js';

const { Client, LocalAuth, MessageMedia } = pkg;
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, '..', 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Config ──────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3211;
const TARGET_CHAT = (process.env.WHATSAPP_TARGET_CHAT || '').trim();
const GENERATE_KEYWORD = (process.env.GENERATE_KEYWORD || 'צור דוח').trim();
const RESET_KEYWORD = (process.env.RESET_KEYWORD || 'דוח חדש').trim();
const WIZARD_KEYWORD = (process.env.WIZARD_KEYWORD || 'שאלון').trim();
const CANCEL_WIZARD_KEYWORDS = new Set(['בטל שאלון', 'בטל']);
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

// Resolved once at startup by reading the raw internal chat store directly
// in the page context — NOT via client.getChats()/getChatById(), both of
// which route through whatsapp-web.js's chat.serialize(), which is broken
// against the current WhatsApp Web build (WhatsApp renamed an internal WID
// property in a July 2026 update; see index.js history for details). This
// touches only .id and a name field on the raw store models, never
// serialize(), so it avoids that broken code path entirely.
let targetChatId = null;

async function resolveTargetChat() {
  const rawChats = await client.pupPage.evaluate(() => {
    const chats = window.require('WAWebCollections').Chat.getModelsArray();
    return chats.map((c) => ({
      id: c.id?._serialized || c.id?.$1 || (typeof c.id === 'string' ? c.id : null),
      name: c.name || c.formattedTitle || c.contact?.name || c.contact?.pushname || '',
    }));
  });

  const match = rawChats.find((c) => (c.name || '').trim() === TARGET_CHAT);
  if (match?.id) {
    targetChatId = match.id;
    console.log(`✅ Field report bot connected. Watching chat: "${TARGET_CHAT}"`);
  } else {
    console.error(
      `❌ No chat named exactly "${TARGET_CHAT}" was found. Available chat names:\n` +
      rawChats.slice(0, 30).map((c) => `   - "${c.name}" (id: ${c.id})`).join('\n') +
      '\nFix WHATSAPP_TARGET_CHAT in .env to match exactly, then restart.'
    );
  }
}

client.on('ready', async () => {
  isReady = true;
  latestQr = null;
  try {
    await resolveTargetChat();
  } catch (e) {
    console.error('Failed to resolve target chat:', e);
  }
});

client.on('disconnected', (reason) => {
  isReady = false;
  console.warn('⚠️ WhatsApp disconnected:', reason);
});

// Tracks message IDs the bot itself sent, so message_create (which fires for
// our own outgoing messages too) doesn't loop back and treat our replies as
// new field data.
const ownMessageIds = new Set();

async function reply(msg, ...args) {
  const sent = await msg.reply(...args);
  const id = sent?.id?.id;
  if (id) {
    ownMessageIds.add(id);
    setTimeout(() => ownMessageIds.delete(id), 5 * 60 * 1000).unref?.();
  }
  return sent;
}

function parseCommand(body) {
  const trimmed = body.trim();
  if (!trimmed) return null;
  if (trimmed === RESET_KEYWORD) return { type: 'reset' };
  if (HELP_KEYWORDS.has(trimmed)) return { type: 'help' };
  if (CANCEL_WIZARD_KEYWORDS.has(trimmed)) return { type: 'cancel_wizard' };
  if (trimmed === GENERATE_KEYWORD) return { type: 'generate', typeHint: null };
  if (trimmed.startsWith(GENERATE_KEYWORD)) {
    const rest = trimmed.slice(GENERATE_KEYWORD.length).trim().replace(/^:\s*/, '');
    return { type: 'generate', typeHint: rest || null };
  }
  if (trimmed === WIZARD_KEYWORD) return { type: 'wizard', typeHint: null };
  if (trimmed.startsWith(WIZARD_KEYWORD)) {
    const rest = trimmed.slice(WIZARD_KEYWORD.length).trim().replace(/^:\s*/, '');
    return { type: 'wizard', typeHint: rest || null };
  }
  return null;
}

async function handleGenerate(msg, typeHint) {
  let forcedTypeId = null;
  if (typeHint) {
    forcedTypeId = matchTypeHint(typeHint);
    if (!forcedTypeId) {
      await reply(msg, `❓ לא זיהיתי את הסוג "${typeHint}".\n\n${listTypesMessage()}`);
      return;
    }
  } else {
    // No type given in the command itself — if a wizard run already settled
    // on a type (even one that finished a while ago, until reset/generate),
    // use it directly instead of asking Claude to re-guess it from the notes.
    forcedTypeId = getLastKnownDocType();
  }

  if (isEmpty()) {
    await reply(msg, 'אין עדיין הערות או תמונות לניתוח — שלח טקסט ותמונות מהביקור, ואז שלח שוב "' + GENERATE_KEYWORD + '".');
    return;
  }

  await reply(msg, '⏳ מנתח את ההערות והתמונות ומייצר דוח... (עד כדקה)');

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
    await reply(msg, media, undefined, { caption: `✅ ${typeMeta.name}${payload.client ? ' — ' + payload.client : ''}` });

    resetSession();
    clearLastKnownDocType();
  } catch (e) {
    console.error('Report generation failed:', e);
    await reply(msg, `❌ שגיאה ביצירת הדוח: ${e.message}\n\nההערות והתמונות נשמרו — נסה שוב "${GENERATE_KEYWORD}".`);
  }
}

// message_create (not message) — the watched chat is expected to be a group
// with only the field engineer in it, so every real message they send has
// fromMe: true. The plain 'message' event only fires for messages from OTHER
// people and would never see them at all.
client.on('message_create', async (msg) => {
  if (!isReady) return;
  try {
    const id = msg.id?.id;
    if (id && ownMessageIds.has(id)) { ownMessageIds.delete(id); return; }

    // msg.id.remote is the chat JID the message belongs to, already present
    // on the message payload — comparing it avoids calling msg.getChat() /
    // getChatById(), which currently throws against this WhatsApp Web build.
    if (!targetChatId || msg.id?.remote !== targetChatId) return;

    const body = (msg.body || '').trim();
    const command = parseCommand(body);

    if (command?.type === 'cancel_wizard') {
      if (isWizardActive()) {
        cancelWizard();
        await reply(msg, '❌ השאלון בוטל. מה שכבר נענה נשאר שמור, אפשר להמשיך בכתיבה חופשית.');
      }
      return;
    }

    // While a wizard is in progress, every message is an answer to the
    // current question — free text always works too, not just numbers.
    if (isWizardActive() && !msg.hasMedia) {
      const { prompt, done } = answerWizard(body);
      if (prompt) await reply(msg, prompt);
      return;
    }

    if (command?.type === 'reset') {
      resetSession();
      cancelWizard();
      clearLastKnownDocType();
      await reply(msg, '🆕 האוסף אופס. שלח טקסט ותמונות מהביקור, ואז "' + GENERATE_KEYWORD + '", או "' + WIZARD_KEYWORD + '" לשאלון מודרך.');
      return;
    }

    if (command?.type === 'help') {
      await reply(msg, listTypesMessage() + `\n\nטיפ: אפשר גם לענות על שאלות מודרכות במקום לכתוב חופשי — שלח "${WIZARD_KEYWORD}" (הבוט ישאל גם איזה סוג מסמך).`);
      return;
    }

    if (command?.type === 'wizard') {
      // No type in the command — the wizard itself asks which document type
      // first, then tailors every question that follows to that type (no
      // defect-priority questions for a חוות דעת הנדסית, etc.).
      const result = startWizard(command.typeHint);
      if (!result.ok) {
        await reply(msg, `❓ לא זיהיתי את הסוג "${command.typeHint}".\n\n${listTypesMessage()}`);
        return;
      }
      await reply(msg, result.prompt);
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
