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

const HEADER_FIELDS = [
  { key: 'לקוח', label: 'מי הלקוח/המוסד?' },
  { key: 'מיקום', label: 'מה שם המיקום/המתחם?' },
  { key: 'כתובת', label: 'מה הכתובת?' },
  { key: 'תאריך', label: 'תאריך הביקור? (למשל 20.8.2026 — או "-" להיום)' },
];

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

function promptFor(field) {
  let text = `❓ ${field.label}`;
  if (field.options) {
    text += '\n' + field.options.map((o, i) => `${i + 1}) ${o}`).join('\n');
    text += '\n(אפשר גם להקליד תשובה חופשית במקום לבחור מספר)';
  }
  return text;
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

// Single active wizard, matching the single-session design of session.js —
// this bot watches exactly one chat, so there's only ever one wizard "in flight".
let wizard = null;

export function isWizardActive() {
  return wizard !== null;
}

/** @returns {{ok: boolean, prompt?: string, notSupported?: boolean}} */
export function startWizard(typeHint) {
  const docTypeId = matchTypeHint(typeHint);
  if (!docTypeId) return { ok: false };

  const type = DOC_TYPES[docTypeId];
  if (type.kind !== 'table') {
    return { ok: false, notSupported: true };
  }

  wizard = { docTypeId, stage: 'header', headerIndex: 0, rowIndex: 1, fieldIndex: 0, rowAnswers: {} };
  return { ok: true, prompt: `📝 שאלון עבור "${type.name}" — שאלה 1:\n\n${promptFor(HEADER_FIELDS[0])}` };
}

export function cancelWizard() {
  wizard = null;
}

/** Feed the user's reply to the current question. @returns {{prompt?: string, done?: boolean}} */
export function answerWizard(raw) {
  if (!wizard) return {};

  if (wizard.stage === 'header') {
    const field = HEADER_FIELDS[wizard.headerIndex];
    const answer = resolveAnswer(field, raw);
    if (answer && answer !== '-') addText(`${field.key}: ${answer}`);

    wizard.headerIndex++;
    if (wizard.headerIndex < HEADER_FIELDS.length) {
      return { prompt: promptFor(HEADER_FIELDS[wizard.headerIndex]) };
    }
    wizard.stage = 'row';
    wizard.fieldIndex = 0;
    wizard.rowAnswers = {};
    return { prompt: `📋 ממצא מס' ${wizard.rowIndex}:\n\n${promptFor(rowFields(wizard.docTypeId)[0])}` };
  }

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
      prompt: '✅ סיימנו את השאלון! אפשר עכשיו גם להוסיף תמונות או הערות בכתיבה חופשית, ואז לשלוח "צור דוח" (אפשר גם "צור דוח: סככות" וכו\').',
    };
  }

  return {};
}
