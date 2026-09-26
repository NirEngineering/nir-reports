// Registry of the 8 document types this bot can produce. Mirrors
// nir-reports/src/constants.js (DOC_TYPES_CONFIG, TABLE_COLUMNS, DEFECTS_COLUMNS)
// and nir-reports/src/lib/docGenerator.js (DOC_TYPES) — the exact structures
// already reverse-engineered from the company's real Word templates.
//
// "table" types (group1–group5) collect findings into a row table.
// "opinion" types (group6, group8) are free-text: bullet findings + numbered conclusions.
// "freeform" type (group7) is fully free text.

export const DOC_TYPES = {
  group1: {
    id: 'group1',
    name: 'אלמנטים תלויים',
    kind: 'table',
    subjectDefault: 'בדיקת אלמנטים ומתקנים תלויים',
    aliases: ['אלמנטים', 'אלמנטים תלויים', 'מתקנים תלויים'],
    tableColumns: ['מיקום', 'האלמנט/המתקן', 'נתונים וממצאים', 'הערות'],
    hasDefectsTable: false,
  },
  group2: {
    id: 'group2',
    name: 'סקר פערי בטיחות',
    kind: 'table',
    subjectDefault: 'סקר פערי בטיחות',
    aliases: ['פערי בטיחות', 'סקר פערי בטיחות', 'סקר בטיחות'],
    tableColumns: ['', 'תחום הבדיקה', 'סעיף ברשימת המבדק', 'הדרישה', 'הממצא, מהותו ומיקומו', 'קדימות הליקוי'],
    hasDefectsTable: false,
  },
  group3: {
    id: 'group3',
    name: 'תקרות תותב',
    kind: 'table',
    subjectDefault: 'בדיקת תקרות תותב',
    aliases: ['תקרות תותב', 'תקרה תותבת', 'תקרות'],
    tableColumns: ['מיקום/חדר', 'סוג התקרה', 'תקין/לא תקין', 'קדימות ליקויים', 'הערות'],
    hasDefectsTable: true,
    defectsColumns: ['מיקום', 'ממצאים וליקויים', 'הדרישה', 'תמונות', 'קדימות ליקוי'],
  },
  group4: {
    id: 'group4',
    name: 'סקר תקופתי / שנתי',
    kind: 'table',
    subjectDefault: 'סקר בטיחות שנתי',
    aliases: ['סקר שנתי', 'סקר תקופתי', 'סקר'],
    tableColumns: ['האלמנט/המבנה הנבדק', 'נתונים ופירוט', 'הערות', 'תקין/לא תקין', 'קדימות ליקוי'],
    hasDefectsTable: false,
  },
  group5: {
    id: 'group5',
    name: 'סככות',
    kind: 'table',
    subjectDefault: 'בדיקת סככות',
    aliases: ['סככות', 'סככה'],
    tableColumns: ["מס'", 'מיקום', 'סוג הסככה', "מידות (מ') ונתונים", 'תקין/לא תקין', 'קדימות ליקויים', 'תמונה'],
    hasDefectsTable: true,
    defectsColumns: ['מספר סככה', 'מיקום', 'ממצאי ליקויים ודרישות', 'תמונות הליקוי', 'קדימות'],
  },
  group6: {
    id: 'group6',
    name: 'חוות דעת הנדסית',
    kind: 'opinion',
    subjectDefault: 'חוות דעת הנדסית',
    aliases: ['חוות דעת', 'חוות דעת הנדסית'],
  },
  group7: {
    id: 'group7',
    name: 'מסמך כללי',
    kind: 'freeform',
    subjectDefault: '',
    aliases: ['מסמך כללי', 'כללי'],
  },
  group8: {
    id: 'group8',
    name: 'אישור מבנים ארעיים',
    kind: 'opinion',
    subjectDefault: 'אישור יציבות מבנה/מתקן ארעי',
    aliases: ['מבנה ארעי', 'מבנים ארעיים', 'אישור מבנה ארעי', 'מתקן ארעי'],
  },
  group9: {
    id: 'group9',
    name: 'אישור ביצוע חיזוק',
    kind: 'opinion',
    subjectDefault: 'אישור ביצוע עבודות חיזוק',
    aliases: ['אישור ביצוע', 'אישור ביצוע חיזוק', 'ביצוע חיזוק', 'חיזוק'],
  },
  group10: {
    id: 'group10',
    name: 'אישור רשת הגנה',
    kind: 'opinion',
    subjectDefault: 'אישור יציבות רשת הגנה לנפילת חפצים',
    aliases: ['רשת הגנה', 'אישור רשת הגנה', 'רשת לנפילת חפצים'],
  },
  group11: {
    id: 'group11',
    name: 'אישור בטיחות שנתי לקייטנה',
    kind: 'checklist',
    subjectDefault: 'אישור בטיחות שנתי',
    aliases: ['אישור בטיחות קייטנה', 'בטיחות קייטנה', 'קייטנה', 'אישור בטיחות שנתי'],
  },
};

// Fixed 15-item safety checklist (group11, קייטנה) — identical wording in
// every real Drive sample found; only the status/הערות per item vary per
// site. Shared by wizard.js (survey prompts) and docGenerator.js (the table).
export const CAMP_CHECKLIST_ITEMS = [
  'קיים רכז בטיחות ותברואה לקייטנה. (בקייטנה קטנה אפשר שמנהל הקייטנה ישמש תפקיד זה).',
  "המדריכים קיבלו הדרכה בנושא הבטיחות ולמדו את ההוראות הייחודיות לכל סוג פעילות (טיולים, רחצה בבריכה, רחצה בפארק מים, מתקנים, מתנפחים וכו') הממפורטות בחוזר מנכ''ל.",
  'נקבעו בעלי תפקידים למקרה של שרפה והם תודרכו.',
  'בקייטנה מצוי תיק בטיחות ובו מרוכזים כל האישורים והוראות הבטיחות המחייבות.',
  "בקייטנה מצוי חוזר המנכ''ל המעודכן של משרד החינוך לבטיחות בקייטנות.",
  'קיים אישור טכנאי גז למערכות הגז במטבח (אם קיים גז).',
  'מתקני החשמל במבנה ובחצר נבדקו בידי חשמלאי מוסמך.',
  'כלי העבודה למלאכה ולאומנויות והחומרים שנעשה בהם שימוש בחוגים נמצאו תקינים.',
  'ציוד העזרה הראשונה תקין.',
  'ציוד כיבוי האש תקין ומתאים לסוג הפעילויות.',
  'קיימים מעקות ומאחזי יד בכל המקומות הנדרשים.',
  'בדלתות המבנה קיימים אמצעים המגנים מפני פגיעה באצבעות (אמצעי הגנה בין הדלת למשקוף, גלגל להאטת הדלת ותפס בסוף מהלך הפתיחה).',
  'הרחבות, השבילים ומקומות המשחק בשטח החצר נקיים מבורות ושקעים ואין מכשולים בקרבתם.',
  'מתקני המשחק תקינים (בדיקת המשחקים תיהיה ויזואלית).',
  'מתקני הספורט והמגרשים תקינים ושלמים, ללא שקעים ובורות, עם ריפוד והגנה.',
];
export const CAMP_CHECKLIST_STATUS_OPTIONS = ['קיים', 'לא רלוונטי, לא קיים', 'חסר', 'לא בוצע'];

const NORMALIZE = (s) => String(s || '').trim().toLowerCase().replace(/["'׳״]/g, '');

/** Match free-text typed by the field engineer (e.g. "סככות") to a doc type id. */
export function matchTypeHint(hint) {
  const norm = NORMALIZE(hint);
  if (!norm) return null;
  for (const type of Object.values(DOC_TYPES)) {
    if (NORMALIZE(type.name) === norm) return type.id;
    if (type.aliases.some((a) => NORMALIZE(a) === norm)) return type.id;
  }
  // Loose contains-match as a fallback (e.g. "תבנית סככות" or "דוח סככות")
  for (const type of Object.values(DOC_TYPES)) {
    if (type.aliases.some((a) => norm.includes(NORMALIZE(a)))) return type.id;
  }
  return null;
}

export function listTypesMessage() {
  const lines = Object.values(DOC_TYPES).map((t) => `• ${t.name}`);
  return `סוגי המסמכים הזמינים:\n${lines.join('\n')}\n\nכדי לבחור סוג, שלח: "${'צור דוח'}: <סוג>" (למשל "צור דוח: סככות").\nללא ציון סוג — הבוט ינחש לפי התוכן.`;
}
