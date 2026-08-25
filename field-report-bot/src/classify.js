// Turns the raw WhatsApp session (text lines + photos) into the structured
// payload docGenerator.generateDocument() expects. Uses the same Claude
// extraction approach as nir-reports' AIWriter.jsx, extended to (a) pick the
// document type itself when the field engineer didn't name one, and (b)
// cover all 8 document types, not just the 5 table-based ones.
import { DOC_TYPES } from './docTypes.js';

const MODEL = 'claude-sonnet-4-6';
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

function buildSystemPrompt(forcedTypeId) {
  const typeList = Object.values(DOC_TYPES)
    .map((t) => `- ${t.id} — "${t.name}"${t.kind === 'table' ? ` (טבלת ממצאים: ${t.tableColumns.filter(Boolean).join(' | ')})` : ' (טקסט חופשי, ללא טבלה)'}`)
    .join('\n');

  const typeInstruction = forcedTypeId
    ? `סוג המסמך נקבע מראש ע"י המשתמש: "${forcedTypeId}". השתמש בו בשדה doc_type ואל תשנה אותו.`
    : `בחר את סוג המסמך המתאים ביותר לתוכן מתוך הרשימה, והחזר את המזהה (למשל "group5") בשדה doc_type.`;

  return `אתה עוזר מקצועי לחברת "ניר הנדסה", חברה לייעוץ הנדסי בישראל.
תפקידך לנתח הערות שטח (טקסט חופשי) ותמונות שנשלחו בוואטסאפ במהלך ביקור/סקר הנדסי,
ולהחזיר JSON מובנה בלבד שישמש ליצירת מסמך Word רשמי.

סוגי המסמכים האפשריים:
${typeList}

${typeInstruction}

החזר JSON תקני בלבד (ללא markdown, ללא \`\`\`json, רק JSON נקי) עם המבנה הבא:
{
  "doc_type": "group1..group8",
  "client": "שם הלקוח או המוסד",
  "organization": "שם הארגון (אם קיים, אחרת השאר ריק)",
  "location": "שם המיקום/הנכס",
  "address": "כתובת מלאה",
  "date": "YYYY-MM-DD",
  "subject": "נושא הדוח בקצרה",
  "intro_extra": "משפט פתיחה נוסף (אופציונלי, אחרת מחרוזת ריקה)",
  "findings": [
    {
      "element": "שם האלמנט/הרכיב/הסככה שנבדק",
      "location_detail": "מיקום ספציפי בשטח",
      "description": "תיאור מפורט של הממצא",
      "recommendation": "המלצה לתיקון",
      "priority": "1",
      "status": "לא תקין",
      "dimensions": "מידות (רלוונטי רק לסככות)"
    }
  ],
  "conclusion": "פסקת סיכום מקצועית (לדוחות טבלה) — או מסקנות ממוספרות מופרדות בשורות (לחוות דעת/אישור/מסמך כללי)",
  "notes": "הערות נוספות חופשיות (שורה = פריט נפרד; אופציונלי)"
}

הנחיות:
- כל הטקסט בעברית מקצועית, כמו שכותב מהנדס בפועל.
- priority: "1"=דחוף, "2"=בינוני, "3"=נמוך.
- status: "תקין" / "לא תקין" / "תקין - דורש מעקב".
- נתח גם תמונות שצורפו והוסף ממצאים שנראים בהן (מיקום, אלמנט, תיאור, המלצה).
- אם doc_type הוא group6/group7/group8 (ללא טבלה) — אל תמלא findings; מלא את "notes" בפריטי ממצא/פסקאות (שורה אחת = פריט אחד) ואת "conclusion" במסקנות.
- אם פרט לא קיים בהקלט — השאר מחרוזת ריקה, אל תמציא נתונים.
- findings יכול להיות מערך ריק אם אין ממצאים בטבלה.`;
}

// ── Map generic findings[] → per-type table rows (ported from nir-reports AIWriter.jsx) ──
function toTableRows(findings, docType) {
  switch (docType) {
    case 'group1':
      return findings.map((f) => [f.location_detail || '', f.element || '', f.description || '', f.recommendation || '']);
    case 'group2':
      return findings.map((f) => [
        '', f.element || '', '', f.description || '',
        `${f.description || ''}${f.location_detail ? ` — ${f.location_detail}` : ''}`,
        f.priority || '1',
      ]);
    case 'group3':
      return findings.map((f) => [f.location_detail || '', f.element || '', f.status || 'לא תקין', f.priority || '1', f.recommendation || '']);
    case 'group4':
      return findings.map((f) => [f.element || '', f.description || '', f.recommendation || '', f.status || 'לא תקין', f.priority || '1']);
    case 'group5':
      return findings.map((f, i) => [
        String(i + 1), f.location_detail || '', f.element || '', f.dimensions || '',
        f.status || 'לא תקין', f.priority || '1', '',
      ]);
    default:
      return findings.map((f) => [f.location_detail || f.element || '', f.description || '', f.recommendation || '']);
  }
}

function toDefectsRows(findings, docType) {
  const defects = findings.filter((f) => f.status && f.status !== 'תקין');
  switch (docType) {
    case 'group3':
      return defects.map((f) => [f.location_detail || '', f.description || '', f.recommendation || '', '', f.priority || '1']);
    case 'group5':
      return defects.map((f, i) => [String(i + 1), f.location_detail || '', f.description || '', '', f.priority || '1']);
    default:
      return [];
  }
}

async function callClaude(systemPrompt, textContent, photos) {
  const content = [
    ...photos.slice(0, 15).map((ph) => {
      const match = ph.data.match(/^data:([^;]+);base64,(.+)$/);
      if (!match) return null;
      return { type: 'image', source: { type: 'base64', media_type: match[1], data: match[2] } };
    }).filter(Boolean),
    { type: 'text', text: textContent || 'נתח את התמונות שצורפו וכתוב דוח מלא על סמך מה שנראה בהן.' },
  ];

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 4096,
      system: systemPrompt,
      messages: [{ role: 'user', content }],
    }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Claude API error: ${res.status}`);
  }

  const data = await res.json();
  const raw = data.content?.[0]?.text || '';
  const clean = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  return JSON.parse(clean);
}

/**
 * @param {{texts: {text:string, ts:number}[], photos: {data:string, caption:string}[]}} session
 * @param {string|null} forcedTypeId — doc type id if the user named one explicitly
 * @returns {Promise<object>} payload ready for docGenerator.generateDocument()
 */
export async function classifyAndBuild(session, forcedTypeId) {
  if (!ANTHROPIC_API_KEY) {
    throw new Error('ANTHROPIC_API_KEY לא מוגדר — הוסף אותו לקובץ .env');
  }

  const textContent = (session.texts || []).map((t) => t.text).join('\n');
  const photos = session.photos || [];

  const systemPrompt = buildSystemPrompt(forcedTypeId);
  const result = await callClaude(systemPrompt, textContent, photos);

  const docType = forcedTypeId || result.doc_type || 'group7';
  const typeMeta = DOC_TYPES[docType] || DOC_TYPES.group7;

  const findings = Array.isArray(result.findings) ? result.findings : [];
  const notesLines = typeof result.notes === 'string'
    ? result.notes.split('\n').map((s) => s.trim()).filter(Boolean)
    : (Array.isArray(result.notes) ? result.notes : []);

  const payload = {
    doc_type: docType,
    client: result.client || '',
    organization: result.organization || '',
    location: result.location || '',
    address: result.address || '',
    date: result.date || new Date().toISOString().split('T')[0],
    inspection_date: result.date || new Date().toISOString().split('T')[0],
    subject: result.subject || typeMeta.subjectDefault || '',
    intro_extra: result.intro_extra || '',
    conclusion_custom: result.conclusion || '',
    notes_custom: notesLines,
    photos: photos.map((p) => ({ data: p.data, caption: p.caption || '' })),
  };

  if (typeMeta.kind === 'table') {
    payload.table_rows = toTableRows(findings, docType);
    payload.defects_rows = toDefectsRows(findings, docType);
    payload.has_defects = findings.some((f) => f.status && f.status !== 'תקין');
  } else {
    payload.table_rows = [];
    payload.defects_rows = [];
    payload.has_defects = false;
  }

  return payload;
}
