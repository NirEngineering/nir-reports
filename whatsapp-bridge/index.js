import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import qrcodeTerminal from 'qrcode-terminal';
import QRCode from 'qrcode';
import pkg from 'whatsapp-web.js';

const { Client, LocalAuth } = pkg;

// ── Config ──────────────────────────────────────────────────────────────────
// TARGET_CHAT must match the exact name of the WhatsApp chat to watch — e.g. a
// dedicated group called "דוחות שטח" that the field engineer sends notes/photos
// to, or a saved contact name. Only messages from this one chat are read.
const PORT = process.env.PORT || 3210;
const TARGET_CHAT = (process.env.WHATSAPP_TARGET_CHAT || '').trim();
const RESET_KEYWORD = (process.env.RESET_KEYWORD || 'דוח חדש').trim();

if (!TARGET_CHAT) {
  console.error(
    '\n❌ WHATSAPP_TARGET_CHAT is not set.\n' +
    '   Copy .env.example to .env and set it to the exact WhatsApp chat name to watch\n' +
    '   (a group or contact name), then run again.\n'
  );
  process.exit(1);
}

// ── State ───────────────────────────────────────────────────────────────────
let latestQr = null;
let isReady = false;
let inbox = { texts: [], photos: [] }; // texts: [{text, ts}], photos: [{data: dataURL, caption, ts}]

// ── WhatsApp client ─────────────────────────────────────────────────────────
const client = new Client({
  authStrategy: new LocalAuth({ dataPath: '.wwebjs_auth' }),
  puppeteer: { headless: true, args: ['--no-sandbox', '--disable-setuid-sandbox'] },
});

client.on('qr', (qr) => {
  latestQr = qr;
  console.log('\nScan this QR code with WhatsApp → Linked Devices → Link a Device:\n');
  qrcodeTerminal.generate(qr, { small: true });
});

client.on('ready', () => {
  isReady = true;
  latestQr = null;
  console.log(`✅ WhatsApp bridge connected. Watching chat: "${TARGET_CHAT}"`);
});

client.on('disconnected', (reason) => {
  isReady = false;
  console.warn('⚠️ WhatsApp disconnected:', reason);
});

client.on('message', async (msg) => {
  if (!isReady) return;
  try {
    const chat = await msg.getChat();
    if ((chat.name || '').trim() !== TARGET_CHAT) return;

    const body = (msg.body || '').trim();

    if (body === RESET_KEYWORD) {
      inbox = { texts: [], photos: [] };
      await msg.reply(
        `🆕 קליטה חדשה נפתחה. שלח כאן טקסט ותמונות, ואז ב-"ניר הנדסה" ← 🤖 כתיבת דוח AI ← ייבוא מוואטסאפ.`
      );
      return;
    }

    if (msg.hasMedia) {
      const media = await msg.downloadMedia().catch(() => null);
      if (media?.mimetype?.startsWith('image/')) {
        inbox.photos.push({
          data: `data:${media.mimetype};base64,${media.data}`,
          caption: body || '',
          ts: msg.timestamp,
        });
      }
    } else if (body) {
      inbox.texts.push({ text: body, ts: msg.timestamp });
    }
  } catch (e) {
    console.error('Failed to handle incoming WhatsApp message:', e);
  }
});

client.initialize();

// ── HTTP API for the nir-reports PWA ───────────────────────────────────────
const app = express();
app.use(cors());
app.use(express.json({ limit: '1mb' }));

app.get('/status', (req, res) => {
  res.json({
    ready: isReady,
    targetChat: TARGET_CHAT,
    pendingTexts: inbox.texts.length,
    pendingPhotos: inbox.photos.length,
  });
});

app.get('/qr', async (req, res) => {
  if (isReady) return res.json({ ready: true });
  if (!latestQr) return res.status(404).json({ error: 'QR not generated yet — check the server console.' });
  const dataUrl = await QRCode.toDataURL(latestQr);
  res.json({ ready: false, qr: dataUrl });
});

app.get('/inbox', (req, res) => {
  res.json({
    text: inbox.texts.map((t) => t.text).join('\n'),
    photos: inbox.photos.map((p) => ({ data: p.data, caption: p.caption })),
  });
});

app.post('/inbox/clear', (req, res) => {
  inbox = { texts: [], photos: [] };
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`WhatsApp bridge API listening on http://localhost:${PORT}`);
});
