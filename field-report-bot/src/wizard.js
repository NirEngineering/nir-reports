// Guided question-and-answer flow: an alternative to free typing, for a field
// engineer who'd rather pick from a numbered list than type a full sentence.
// Any prompt can still be answered with free text at any time — picking a
// number is just a shortcut, not a requirement. Every answer is written into
// the session as a plain labeled text line (e.g. "סטטוס: לא תקין"), so it
// flows into the exact same Claude extraction step as ordinary free-typed
// notes — no changes needed anywhere else in the pipeline.
import { DOC_TYPES, matchTypeHint } from './docTypes.js';
import { addText } from './session.js';

// Same canonical option lists used in the nir-reports manual app's dropdowns
// (src/constants.js STATUS_OPTIONS / PRIORITY_OPTIONS / PRIORITY_OPTIONS_GAP).
const STATUS_OPTIONS = ['תקין', 'לא תקין', 'תקין - דורש מעקב', 'דורש בדיקה חוזרת', 'הוצא מכלל שימוש'];
const PRIORITY_OPTIONS = ['1', '2', '3'];
const PRIORITY_OPTIONS_GAP = ['0', '1', '2']; // group2 (סקר פערי בטיחות) uses a 0-2 scale

// Institutions/companies most new reports are addressed to, so the field
// engineer can pick a number instead of retyping the same long name every
// visit. Sourced from the client's own real Drive reports — these two
// account for the large majority of documents. Free text always still works
// for anything else, and "אחר – פרט" makes that option visible up front too.
const OTHER_LABEL = 'אחר – פרט';
const CLIENT_OPTIONS = [
  "החברה למוסדות חינוך ותרבות ת''א",
  "החברה לתרבות פנאי וספורט בת ים",
  "מגלקום פתרונות טכנולוגיים בע''מ",
  OTHER_LABEL,
];

const HEADER_FIELDS = [
  { key: 'לקוח', label: 'מי מזמין העבודה / הלקוח?', options: CLIENT_OPTIONS },
  { key: 'מיקום', label: 'מה שם המיקום/המתחם?' },
  { key: 'כתובת', label: 'מה הכתובת?' },
  { key: 'תאריך', label: 'תאריך הביקור? (למשל 20.8.2026 — או "-" להיום)' },
];

// Fields relevant only to table-based types (group1-5) — a finding/row has a
// status and (when not fine) a priority and a fix recommendation. Opinion and
// freeform types don't have a row table at all, so none of this applies there
// — asking about defect priority for a חוות דעת הנדסית makes no sense.
function rowFields(docTypeId) {
  const priorityOptions = docTypeId === 'group2' ? PRIORITY_OPTIONS_GAP : PRIORITY_OPTIONS;
  return [
    { key: 'מיקום', label: 'איפה נמצא הממצא? (מיקום ספציפי בשטח)' },
    { key: 'אלמנט', label: 'איזה אלמנט/רכיב/סככה זה?' },
    { key: 'תיאור', label: 'מה הממצא או התיאור?' },
    { key: 'סטטוס', label: 'מה הסטטוס?', options: STATUS_OPTIONS },
    { key: 'קדימות', label: 'מה הקדימות לטיפול?', options: priorityOptions, skipIf: (a) => a['סטטוס'] === 'תקין' },
    { key: 'המלצה', label: 'מה ההמלצה לתיקון?', skipIf: (a) => a['סטטוס'] === 'תקין' },
  ];
}

// group7 (מסמך כללי, shown as "אחר") moved to the end of the list, so it
// reads as the catch-all "other" option rather than sitting in the middle.
const TYPE_LIST = [
  ...Object.values(DOC_TYPES).filter((t) => t.id !== 'group7'),
  ...Object.values(DOC_TYPES).filter((t) => t.id === 'group7'),
];

function promptFor(field) {
  let text = `❓ ${field.label}`;
  if (field.options) {
    text += '\n' + field.options.map((o, i) => `${i + 1}) ${o}`).join('\n');
    text += '\n(אפשר גם להקליד תשובה חופשית במקום לבחור מספר)';
  }
  return text;
}

// group7 (מסמך כללי) is the catch-all for anything that isn't one of the
// other 7 defined types, so it's labeled "אחר" here — the last option in
// the list, same spot a field engineer would expect an "other" choice.
function typeLabel(t) {
  return t.id === 'group7' ? 'אחר' : t.name;
}

function typePrompt() {
  return '❓ איזה סוג מסמך?\n' +
    TYPE_LIST.map((t, i) => `${i + 1}) ${typeLabel(t)}`).join('\n') +
    '\n(אפשר גם להקליד את הסוג בעצמו)';
}

function resolveAnswer(field, raw) {
  const trimmed = raw.trim();
  if (field.options) {
    const n = parseInt(trimmed, 10);
    if (!isNaN(n) && n >= 1 && n <= field.options.length && String(n) === trimmed) {
      return field.options[n - 1];
    }
  }
  return trimmed;
}

function resolveTypeAnswer(raw) {
  const trimmed = raw.trim();
  const n = parseInt(trimmed, 10);
  if (!isNaN(n) && n >= 1 && n <= TYPE_LIST.length && String(n) === trimmed) {
    return TYPE_LIST[n - 1].id;
  }
  return matchTypeHint(trimmed);
}

function firstPromptForType(docTypeId) {
  const type = DOC_TYPES[docTypeId];
  return `📝 שאלון עבור "${type.name}":\n\n${promptFor(HEADER_FIELDS[0])}`;
}

// Remembers the last document type the wizard settled on, so that a bare
// "צור דוח" (no ": <type>") after finishing the wizard uses it directly
// instead of relying on the AI to re-guess it from the notes.
let lastKnownDocType = null;
export function getLastKnownDocType() { return lastKnownDocType; }
export function clearLastKnownDocType() { lastKnownDocType = null; }

// Single active wizard, matching the single-session design of session.js —
// this bot watches exactly one chat, so there's only ever one wizard "in flight".
let wizard = null;

export function isWizardActive() {
  return wizard !== null;
}

/** @returns {{ok: boolean, prompt?: string}} */
export function startWizard(typeHint) {
  const docTypeId = typeHint ? matchTypeHint(typeHint) : null;

  if (docTypeId) {
    lastKnownDocType = docTypeId;
    wizard = { docTypeId, stage: 'header', headerIndex: 0, rowIndex: 1, fieldIndex: 0, rowAnswers: {}, findingsCount: 0, conclusionsCount: 0 };
    return { ok: true, prompt: firstPromptForType(docTypeId) };
  }

  if (typeHint) return { ok: false }; // hint given but not recognized

  wizard = { docTypeId: null, stage: 'doctype' };
  return { ok: true, prompt: `📝 שאלון מודרך — ${typePrompt()}` };
}

export function cancelWizard() {
  wizard = null;
}

// Shared by the 'header' and 'header-other' stages: records one header
// field's final answer, then either asks the next header question or hands
// off to whatever comes after the header (row/findings/freeform), depending
// on the chosen document type's kind.
function recordHeaderAnswer(field, answer) {
  if (answer && answer !== '-') addText(`${field.key}: ${answer}`);

  wizard.headerIndex++;
  if (wizard.headerIndex < HEADER_FIELDS.length) {
    return { prompt: promptFor(HEADER_FIELDS[wizard.headerIndex]) };
  }

  const type = DOC_TYPES[wizard.docTypeId];
  if (type.kind === 'table') {
    wizard.stage = 'row';
    wizard.fieldIndex = 0;
    wizard.rowAnswers = {};
    return { prompt: `📋 ממצא מס' ${wizard.rowIndex}:\n\n${promptFor(rowFields(wizard.docTypeId)[0])}` };
  }
  if (type.kind === 'opinion') {
    wizard.stage = 'findings';
    return { prompt: '📋 ממצא/נתון ראשון (תיאור חופשי) — או שלח "סיום" כדי לעבור למסקנות:' };
  }
  // freeform (group7) — no structured fields at all, hand off to free text
  wizard = null;
  return {
    done: true,
    prompt: '✅ הפרטים הכלליים נקלטו. סוג המסמך הזה הוא טקסט חופשי — המשך לכתוב את תוכן המסמך כטקסט רגיל, ואז שלח "צור דוח".',
  };
}

/** Feed the user's reply to the current question. @returns {{prompt?: string, done?: boolean}} */
export function answerWizard(raw) {
  if (!wizard) return {};

  // ── Stage: which document type ─────────────────────────────────────────
  if (wizard.stage === 'doctype') {
    const docTypeId = resolveTypeAnswer(raw);
    if (!docTypeId) {
      return { prompt: `❓ לא זיהיתי את הסוג. ${typePrompt()}` };
    }
    lastKnownDocType = docTypeId;
    wizard = { docTypeId, stage: 'header', headerIndex: 0, rowIndex: 1, fieldIndex: 0, rowAnswers: {}, findingsCount: 0, conclusionsCount: 0 };
    return { prompt: firstPromptForType(docTypeId) };
  }

  // ── Stage: header fields (client / location / address / date) ──────────
  if (wizard.stage === 'header') {
    const field = HEADER_FIELDS[wizard.headerIndex];
    const answer = resolveAnswer(field, raw);
    if (field.options && answer === OTHER_LABEL) {
      wizard.stage = 'header-other';
      return { prompt: `✍️ כתוב את השם:` };
    }
    return recordHeaderAnswer(field, answer);
  }

  // ── Stage: free-text follow-up after picking "אחר – פרט" on a header field ─
  if (wizard.stage === 'header-other') {
    const field = HEADER_FIELDS[wizard.headerIndex];
    wizard.stage = 'header';
    return recordHeaderAnswer(field, raw.trim());
  }

  // ── Stage: table-based finding rows (group1-5) ──────────────────────────
  if (wizard.stage === 'row') {
    const fields = rowFields(wizard.docTypeId);
    const field = fields[wizard.fieldIndex];
    const answer = resolveAnswer(field, raw);
    wizard.rowAnswers[field.key] = answer;
    addText(`ממצא ${wizard.rowIndex} - ${field.key}: ${answer}`);

    wizard.fieldIndex++;
    while (wizard.fieldIndex < fields.length && fields[wizard.fieldIndex].skipIf?.(wizard.rowAnswers)) {
      wizard.fieldIndex++;
    }
    if (wizard.fieldIndex < fields.length) {
      return { prompt: promptFor(fields[wizard.fieldIndex]) };
    }
    wizard.stage = 'more';
    return { prompt: '➕ להוסיף ממצא נוסף?\n1) כן\n2) לא, זהו' };
  }

  if (wizard.stage === 'more') {
    const trimmed = raw.trim();
    const wantsMore = trimmed === '1' || /^כ/.test(trimmed);
    if (wantsMore) {
      wizard.rowIndex++;
      wizard.stage = 'row';
      wizard.fieldIndex = 0;
      wizard.rowAnswers = {};
      return { prompt: `📋 ממצא מס' ${wizard.rowIndex}:\n\n${promptFor(rowFields(wizard.docTypeId)[0])}` };
    }
    wizard = null;
    return {
      done: true,
      prompt: '✅ סיימנו את השאלון! אפשר עכשיו גם להוסיף תמונות או הערות בכתיבה חופשית, ואז לשלוח "צור דוח".',
    };
  }

  // ── Stage: free-text findings / conclusions (group6, group8 — "opinion") ─
  if (wizard.stage === 'findings') {
    const trimmed = raw.trim();
    if (trimmed === 'סיום' || trimmed === 'סיים') {
      wizard.stage = 'conclusions';
      return { prompt: '📝 מסקנה/הערה ראשונה — או שלח "סיום" לסיים את השאלון:' };
    }
    wizard.findingsCount++;
    addText(`נתון/ממצא: ${trimmed}`);
    return { prompt: `📋 ממצא/נתון נוסף — או שלח "סיום" כדי לעבור למסקנות:` };
  }

  if (wizard.stage === 'conclusions') {
    const trimmed = raw.trim();
    if (trimmed === 'סיום' || trimmed === 'סיים') {
      wizard = null;
      return {
        done: true,
        prompt: '✅ סיימנו את השאלון! אפשר עכשיו גם להוסיף תמונות או הערות בכתיבה חופשית, ואז לשלוח "צור דוח".',
      };
    }
    wizard.conclusionsCount++;
    addText(`מסקנה: ${trimmed}`);
    return { prompt: '📝 מסקנה/הערה נוספת — או שלח "סיום" לסיים את השאלון:' };
  }

  return {};
}
