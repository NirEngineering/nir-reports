// Guided question-and-answer flow: an alternative to free typing, for a field
// engineer who'd rather pick from a numbered list than type a full sentence.
// Any prompt can still be answered with free text at any time — picking a
// number is just a shortcut, not a requirement. Every answer is written into
// the session as a plain labeled text line (e.g. "סטטוס: לא תקין"), so it
// flows into the exact same Claude extraction step as ordinary free-typed
// notes — no changes needed anywhere else in the pipeline.
import { DOC_TYPES, matchTypeHint, CAMP_CHECKLIST_ITEMS, CAMP_CHECKLIST_STATUS_OPTIONS } from './docTypes.js';
import { addText, setStructured } from './session.js';
import { findPastReportsByClient } from './sheetLookup.js';

// Checklist statuses (group11) that flag a problem needing a follow-up note —
// "קיים" and "לא רלוונטי, לא קיים" are fine-as-is, everything else isn't.
const CAMP_CHECKLIST_OK_STATUSES = new Set(['קיים', 'לא רלוונטי, לא קיים']);

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
// Shown after every multiple-choice prompt so the field engineer always sees
// the way out, not just the numbered options — picking a number is always a
// shortcut, never a requirement, and "בטל שאלון" hands the whole session back
// to plain free-text note-taking (already handled globally in index.js).
const CANCEL_HINT = '\nבכל שלב אפשר לשלוח "בטל שאלון" כדי לעבור לכתיבה חופשית.';

// Every numbered-option prompt also carries { poll: {question, options} } so
// index.js can additionally send a real WhatsApp poll (tap to answer) next
// to the plain-text version — the text prompt (with its own numbered list
// and free-text/cancel hint) is still sent either way, so nothing regresses
// if a poll fails to send or the engineer just prefers typing.
// WhatsApp native polls support at most 12 tappable options and cap each
// option's text length — pollFor() returns null above that so the caller
// falls back to text-only instead of sending a malformed/rejected poll.
const MAX_POLL_OPTIONS = 12;
const MAX_POLL_OPTION_CHARS = 95;
const MAX_POLL_QUESTION_CHARS = 250;
function pollFor(question, options) {
  if (!Array.isArray(options) || options.length < 2 || options.length > MAX_POLL_OPTIONS) return null;
  const truncate = (s, max) => (s.length > max ? s.slice(0, max - 1) + '…' : s);
  return {
    question: truncate(question, MAX_POLL_QUESTION_CHARS),
    options: options.map((o) => truncate(String(o), MAX_POLL_OPTION_CHARS)),
  };
}
function fieldPoll(field) {
  return field.options ? pollFor(field.label, field.options) : null;
}

const OTHER_LABEL = 'אחר – פרט';
const CLIENT_OPTIONS = [
  "החברה למוסדות חינוך ותרבות ת''א",
  "החברה לתרבות פנאי וספורט בת ים",
  "מגלקום פתרונות טכנולוגיים בע''מ",
  "עזריאל ע.י. סחור ושיווק בע''מ",
  "רשת קהילה ופנאי חולון",
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
  return text + CANCEL_HINT;
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
    '\n(אפשר גם להקליד את הסוג בעצמו)' + CANCEL_HINT;
}
function typePoll() {
  return pollFor('איזה סוג מסמך?', TYPE_LIST.map(typeLabel));
}

// group11 (אישור בטיחות לקייטנה) — one survey question per fixed checklist item.
function checklistPrompt(index) {
  return `☑️ פריט ${index + 1}/${CAMP_CHECKLIST_ITEMS.length}:\n${CAMP_CHECKLIST_ITEMS[index]}\n` +
    CAMP_CHECKLIST_STATUS_OPTIONS.map((o, i) => `${i + 1}) ${o}`).join('\n') +
    '\n(אפשר גם להקליד תשובה חופשית)' + CANCEL_HINT;
}
function checklistPoll(index) {
  return pollFor(CAMP_CHECKLIST_ITEMS[index], CAMP_CHECKLIST_STATUS_OPTIONS);
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
  return {
    prompt: `📝 שאלון עבור "${type.name}":\n\n${promptFor(HEADER_FIELDS[0])}`,
    poll: fieldPoll(HEADER_FIELDS[0]),
  };
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

/** @returns {{ok: boolean, prompt?: string, poll?: {question: string, options: string[]}}} */
export function startWizard(typeHint) {
  const docTypeId = typeHint ? matchTypeHint(typeHint) : null;

  if (docTypeId) {
    lastKnownDocType = docTypeId;
    wizard = { docTypeId, stage: 'header', headerIndex: 0, rowIndex: 1, fieldIndex: 0, rowAnswers: {}, findingsCount: 0, conclusionsCount: 0 };
    return { ok: true, ...firstPromptForType(docTypeId) };
  }

  if (typeHint) return { ok: false }; // hint given but not recognized

  wizard = { docTypeId: null, stage: 'doctype' };
  return { ok: true, prompt: `📝 שאלון מודרך — ${typePrompt()}`, poll: typePoll() };
}

export function cancelWizard() {
  wizard = null;
}

// Shared tail of the header stage: either asks the next header question or
// hands off to whatever comes after the header (row/findings/checklist/
// freeform), depending on the chosen document type's kind. Called both from
// the normal header flow and after a "prior report" pick/decline, which
// jumps straight here instead of re-asking a field already filled in.
function continueHeaderFlow() {
  if (wizard.headerIndex < HEADER_FIELDS.length) {
    return { prompt: promptFor(HEADER_FIELDS[wizard.headerIndex]), poll: fieldPoll(HEADER_FIELDS[wizard.headerIndex]) };
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
    return { prompt: '📋 ממצא/נתון ראשון (תיאור חופשי) — או שלח "סיום" כדי לעבור למסקנות:' + CANCEL_HINT };
  }
  if (type.kind === 'checklist') {
    wizard.stage = 'checklist';
    wizard.checklistIndex = 0;
    wizard.checklistAnswers = [];
    return { prompt: checklistPrompt(0), poll: checklistPoll(0) };
  }
  // freeform (group7) — no structured fields at all, hand off to free text
  wizard = null;
  return {
    done: true,
    prompt: '✅ הפרטים הכלליים נקלטו. סוג המסמך הזה הוא טקסט חופשי — המשך לכתוב את תוכן המסמך כטקסט רגיל, ואז שלח "צור דוח".',
  };
}

// Up to 5 non-archived past reports for the client just named, most recent
// first — queried from the same Google Sheet the document-validity dashboard
// (nir-dashboard-new) tracks. A hit lets the engineer confirm this visit is
// a continuation of a known site instead of typing its location from scratch.
const PRIOR_DECLINE_LABEL = 'לא, דוח חדש';

// The decline choice is the LAST positional option (not a special "0"), so a
// tapped poll vote (which only ever reports a 0-based position) and a typed
// number resolve through the exact same 1-based logic below.
function priorOptionLabels(matches) {
  return [
    ...matches.map((m) => `${m.location || '—'} — ${m.docType || '—'}${m.docDate ? ' (' + m.docDate + ')' : ''}`),
    PRIOR_DECLINE_LABEL,
  ];
}

function formatPriorPrompt(matches) {
  const labels = priorOptionLabels(matches);
  const lines = labels.map((l, i) => `${i + 1}) ${l}`);
  return '📂 נמצאו דוחות קודמים עבור הלקוח הזה — זה המשך לאחד מהם?\n' +
    lines.join('\n') + CANCEL_HINT;
}

// Shared by the 'header' and 'header-other' stages: records one header
// field's final answer, then either asks the next header question, offers a
// matching prior report (right after the client field), or hands off to
// whatever comes after the header, depending on the chosen document type's kind.
async function recordHeaderAnswer(field, answer) {
  if (answer && answer !== '-') addText(`${field.key}: ${answer}`);

  const justAnsweredClient = field.key === 'לקוח';
  wizard.headerIndex++;

  if (justAnsweredClient && answer && answer !== '-') {
    const matches = await findPastReportsByClient(answer);
    if (matches.length > 0) {
      wizard.stage = 'prior-pick';
      wizard.priorMatches = matches;
      return { prompt: formatPriorPrompt(matches), poll: pollFor('דוח קודם קיים — להמשיך אחד מהם?', priorOptionLabels(matches)) };
    }
  }

  return continueHeaderFlow();
}

/** Feed the user's reply to the current question. @returns {Promise<{prompt?: string, done?: boolean}>} */
export async function answerWizard(raw) {
  if (!wizard) return {};

  // ── Stage: which document type ─────────────────────────────────────────
  if (wizard.stage === 'doctype') {
    const docTypeId = resolveTypeAnswer(raw);
    if (!docTypeId) {
      return { prompt: `❓ לא זיהיתי את הסוג. ${typePrompt()}`, poll: typePoll() };
    }
    lastKnownDocType = docTypeId;
    wizard = { docTypeId, stage: 'header', headerIndex: 0, rowIndex: 1, fieldIndex: 0, rowAnswers: {}, findingsCount: 0, conclusionsCount: 0 };
    return firstPromptForType(docTypeId);
  }

  // ── Stage: header fields (client / location / address / date) ──────────
  if (wizard.stage === 'header') {
    const field = HEADER_FIELDS[wizard.headerIndex];
    const answer = resolveAnswer(field, raw);
    if (field.options && answer === OTHER_LABEL) {
      wizard.stage = 'header-other';
      return { prompt: `✍️ כתוב את השם:` + CANCEL_HINT };
    }
    return await recordHeaderAnswer(field, answer);
  }

  // ── Stage: free-text follow-up after picking "אחר – פרט" on a header field ─
  if (wizard.stage === 'header-other') {
    const field = HEADER_FIELDS[wizard.headerIndex];
    wizard.stage = 'header';
    return await recordHeaderAnswer(field, raw.trim());
  }

  // ── Stage: pick a matching prior report, or decline and continue normally ─
  if (wizard.stage === 'prior-pick') {
    const trimmed = raw.trim();
    const n = parseInt(trimmed, 10);
    const matches = wizard.priorMatches;
    // Positions 1..matches.length pick a match; position matches.length+1 is
    // the trailing "לא, דוח חדש" choice — same 1-based positions a tapped
    // poll option (converted from its 0-based localId) resolves through too.
    const validN = !isNaN(n) && n >= 1 && n <= matches.length + 1 && String(n) === trimmed;
    const picked = validN && n <= matches.length ? matches[n - 1] : null;

    if (picked) {
      addText(`מיקום: ${picked.location}`);
      addText(`המשך לדוח קודם מתאריך ${picked.docDate || '—'} (סוג: ${picked.docType || '—'}${picked.expiry ? ', בתוקף עד ' + picked.expiry : ''})`);
      wizard.headerIndex = 2; // skip "מיקום" (index 1) — already filled from the match
    }
    // The decline position, free text, or anything unrecognized: decline the
    // suggestion and fall through to asking "מיקום" normally — headerIndex
    // stays at 1, same as if no prior match had ever been found.
    wizard.stage = 'header';
    wizard.priorMatches = null;
    return continueHeaderFlow();
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
      return { prompt: promptFor(fields[wizard.fieldIndex]), poll: fieldPoll(fields[wizard.fieldIndex]) };
    }
    wizard.stage = 'more';
    return { prompt: '➕ להוסיף ממצא נוסף?\n1) כן\n2) לא, זהו' + CANCEL_HINT, poll: pollFor('להוסיף ממצא נוסף?', ['כן', 'לא, זהו']) };
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
      return { prompt: '📝 מסקנה/הערה ראשונה — או שלח "סיום" לסיים את השאלון:' + CANCEL_HINT };
    }
    wizard.findingsCount++;
    addText(`נתון/ממצא: ${trimmed}`);
    return { prompt: `📋 ממצא/נתון נוסף — או שלח "סיום" כדי לעבור למסקנות:` + CANCEL_HINT };
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
    return { prompt: '📝 מסקנה/הערה נוספת — או שלח "סיום" לסיים את השאלון:' + CANCEL_HINT };
  }

  // ── Stage: fixed 15-item safety checklist (group11 — "checklist") ────────
  if (wizard.stage === 'checklist') {
    const trimmed = raw.trim();
    const n = parseInt(trimmed, 10);
    const status = (!isNaN(n) && n >= 1 && n <= CAMP_CHECKLIST_STATUS_OPTIONS.length && String(n) === trimmed)
      ? CAMP_CHECKLIST_STATUS_OPTIONS[n - 1]
      : trimmed;

    wizard.checklistAnswers[wizard.checklistIndex] = { status };
    addText(`בדיקה ${wizard.checklistIndex + 1} (${CAMP_CHECKLIST_ITEMS[wizard.checklistIndex]}) - סטטוס: ${status}`);

    if (!CAMP_CHECKLIST_OK_STATUSES.has(status)) {
      wizard.stage = 'checklist-note';
      return { prompt: '✍️ מה הליקוי/הערה לגבי סעיף זה?' + CANCEL_HINT };
    }
    return advanceChecklist();
  }

  if (wizard.stage === 'checklist-note') {
    const note = raw.trim();
    wizard.checklistAnswers[wizard.checklistIndex].note = note;
    addText(`בדיקה ${wizard.checklistIndex + 1} - הערה: ${note}`);
    wizard.stage = 'checklist';
    return advanceChecklist();
  }

  return {};
}

// Shared by the 'checklist' and 'checklist-note' stages: moves to the next
// checklist item, or finalizes the wizard once all 15 are answered — the
// full answer set is handed to session.js as structured data (see its
// comment) rather than left for Claude to reconstruct from prose.
function advanceChecklist() {
  wizard.checklistIndex++;
  if (wizard.checklistIndex < CAMP_CHECKLIST_ITEMS.length) {
    return { prompt: checklistPrompt(wizard.checklistIndex), poll: checklistPoll(wizard.checklistIndex) };
  }
  setStructured('checklist_items', wizard.checklistAnswers);
  wizard = null;
  return {
    done: true,
    prompt: '✅ סיימנו את רשימת הבדיקה! אפשר עכשיו גם להוסיף תמונות או הערות בכתיבה חופשית, ואז לשלוח "צור דוח".',
  };
}
