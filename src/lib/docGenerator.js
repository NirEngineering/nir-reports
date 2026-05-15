import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  WidthType,
  BorderStyle,
  ShadingType,
  VerticalAlign,
  ImageRun,
  PageNumber,
  UnderlineType,
  convertMillimetersToTwip as mm,
  TableLayoutType,
} from 'docx';
import JSZip from 'jszip';

// ── Font constant ─────────────────────────────────────────────────────────────
const FONT = 'Arial';

// Exact pixel sizes derived from sample Word document (914400 EMU = 1 inch = 96px)
const HEADER_W = 359, HEADER_H = 142;   // 95.12mm × 37.57mm
const FOOTER_W = 568, FOOTER_H = 39;    // 150.50mm × 10.28mm

// ── Doc-type configuration ────────────────────────────────────────────────────
// Per-type font sizes and content extracted from 48 original Hebrew Word documents
const DOC_TYPES = {
  // group1: אלמנטים תלויים — 'simple' layout, first-person "ערכתי"
  group1: {
    name: 'אלמנטים תלויים',
    layout: 'simple',
    titleSize: 16,
    bodySize: 9,
    headingSize: 10,
    tableHdrSize: 10,
    tableDataSize: 9,
    titleSuffix: false,  // subject does NOT include "– {location}"
    introBold: true,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ערכתי סיור בדיקה ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את יציבות האלמנטים והמתקנים התלויים.`,
    tableColumns: ['מיקום', 'האלמנט/המתקן', 'נתונים וממצאים', 'הערות'],
    colWidths: [3.5, 4.5, 6.0, 3.0],
    defaultNotes: [
      "על כל שינוי קונסטרוקטיבי ועיוותים באופן חיבור/תליות האלמנטים (סדקים, עיוותים, שקיעות, ניתוקים, חלודה, אלמנטים רופפים, חוסרים/תוספות וכדומה) יש לדווח על כך לבדיקה חוזרת וטיפול מתאים עפ''י הממצאים.",
      'אין להעמיס עומסים על האלמנטים שנבדקו שאינם מיועדים לכך.',
      'יש להקפיד על הגבלת עומסים במתקנים שהוגבלו בעומסים מותרים.',
      'הבדיקה הינה ויזואלית ונכונה ליום הבדיקה.',
    ],
    hasValidityLine: true,  // validity line appended after notes
    conclusionOk: 'האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו מספר ליקויים והערות שיש לתקנם ולטפלם. שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: false,
  },

  // group2: סקר פערי בטיחות — 'gap-survey' layout
  group2: {
    name: 'סקר פערי בטיחות',
    layout: 'gap-survey',
    titleSize: 14,
    bodySize: 9,
    headingSize: 9,
    tableHdrSize: 9,
    tableDataSize: 9,
    titleSuffix: true,  // subject includes "– {location}"
    introBold: false,
    introTemplate: (d) => [
      { text: `להלן סקר בטיחות אשר נערך בתאריך ${d.inspection_date} ב${d.location}${d.address ? ', ' + d.address : ''}.`, bold: false },
      { text: 'בסקר הסיכונים נבדקו השטחים בהם ישהו עובדי המקום, ילדים ומבקרים.', bold: false },
      { text: "הסקר מבוסס על חוות דעת מקצועית של יועץ בטיחות בהתייחס לאופי המקום בהיבט של בטיחות קהל וציבור.", bold: false },
      { text: "מהות הבדיקה – התאמת תשתיות ואלמנטי המבנה לדרישות חוזר מנכ''ל משרד החינוך, חוק התכנון והבנייה, חוק החשמל, הוראות מכ''ר, פקודת הבטיחות בעבודה והתקנים השונים הרלוונטיים.", bold: false },
      { text: 'הממצאים ייאותרו מתוך השוואת המצב הקיים עם סטנדרטים נדרשים המפורטים ברשימות מנחות לעריכת מבדק בטיחות.', bold: false },
    ],
    tableColumns: ['', 'תחום הבדיקה', 'סעיף ברשימת המבדק', 'הדרישה', 'הממצא, מהותו ומיקומו', 'קדימות הליקוי'],
    colWidths: [1.0, 3.0, 2.5, 3.5, 5.5, 2.5],
    defaultNotes: [
      'מקרא קדימות לטיפול בליקויים:',
      "ליקויים בקדימות 0 – מתייחסת למפגע חמור במיוחד, המחייב להערכת עורך המבדק סגירה מידית של המקום/האתר במוסד החינוך ולאסור שימוש בו עד קבלת הודעה ממנהל הבטיחות ברשות או מנהל המוסד ויועץ בטיחות מטעם הבעלות.",
      "ליקויים בקדימות 1 – מתייחס למפגע בטיחותי אשר קיומו מחייב הסרתו המיידית.",
      "ליקויים בקדימות 2 – מתייחס לליקוי בטיחותי המחייב טיפול של הרשות המקומית/בעלות בתכנית עבודה סדורה.",
    ],
    authorizationLine: "אישור הבטיחות מותנה בהתאמת הממצאים לסעיפי הבדיקה וכן בהצגת כלל האישורים הנדרשים (מפורט בסוף הדו''ח).",
    validityLine: 'תוקף האישור הינו לשנה מיום הבדיקה, בכפוף למסקנות הדו\'\'ח.',
    conclusionHeading: 'מסקנות הדו"ח :',
    conclusionLines: [
      'נמצאו אי התאמות/התאמות לסעיפי הבדיקה.',
      "פירוט הממצאים והליקויים מפורטים בהמשך הדו''ח.",
      "יש להציג אישורים ומסמכים נדרשים רלוונטיים עפ''י המפורט בהמשך.",
    ],
    // Approvals checklist table (fixed template, appended after main table)
    approvalsTableHeaders: ["מס'", 'תחום הבדיקה', 'תדירות', 'הגוף המקצועי הבודק והמאשר', 'הוצג/לא הוצג'],
    approvalsColWidths: [1.0, 3.5, 2.5, 5.0, 3.0],
    hasDefectsTable: false,
  },

  // group3: תקרות תותב — 'survey' layout
  group3: {
    name: 'תקרות תותב',
    layout: 'survey',
    titleSize: 15,
    bodySize: 8.5,
    headingSize: 9,
    tableHdrSize: 8.5,
    tableDataSize: 8.5,
    titleSuffix: true,
    introBold: false,
    introTemplate: (d) => [
      {
        text: "סקר זה מתייחס למצב קיים של התשתיות במבנה (תשתית תקרות התותב שנבדקו) והצגת הפערים הקיימים בין המצב בשטח לבין דרישות התקנים הרלוונטיים.",
        bold: false,
      },
      {
        text: `הסקר בוצע לבקשת ${d.client} ב${d.location}${d.address ? ', ' + d.address : ''} בתאריך ${d.inspection_date}.`,
        bold: true,
      },
    ],
    tableColumns: ['מיקום/חדר', 'סוג התקרה', 'תקין/לא תקין', 'קדימות ליקויים', 'הערות'],
    colWidths: [3.5, 3.5, 3.0, 3.0, 4.0],
    defaultNotes: [
      "ממצאי סקר זה הם כפי שהועברו לח''מ ע''י בעלי התפקידים באתר ומציגים מצב קיים ביום הסיור בלבד, מזמין העבודה אחראי לביצוע תיקון הליקויים שנמצאו בפרק זמן שהוגדר, לא תשמע טענה כנגד הח''מ בגין ליקויים שהצביע עליהם במסגרת סקר זה ושאינם תוקנו בתוך מסגרת הזמן שנקבעה.",
      "מצורפת טבלת ליקויים בהמשך המסמך הכוללת את מיקום הליקוי, פירוט הממצא, דרישות/הנחיות לטיפול ותמונות הליקוי.",
      'מקרא קדימות לטיפול בליקויים:',
      "קדימות 1 – ''ליקוי חמור'' בהגדרתו - ליקוי/מפגע המחייב הסרתו/תיקונו וטיפולו המיידי, לאישור ובדיקה חוזרת.",
      "קדימות 2 – ''ליקוי בינוני'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 3 חודשים מתאריך דו''ח זה.",
      "קדימות 3 – ''ליקוי קל'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 6 חודשים מתאריך דו''ח זה.",
    ],
    hasValidityLine: false,
    defaultInstructions: [
      "על כל שינוי קונסטרוקטיבי בתקרות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה, כפף, שבר, סדק, תזוזות וכו' – יש לזמן לבדיקה חוזרת וטיפול מתאים עפ''י הממצאים.",
      'אין לטפס, להיתלות ו/או להעמיס עומסים על התקרות שנבדקו.',
      "תוקף הדו''ח הינו לחמש שנים מיום הבדיקה בכפוף למסקנות הבדיקה.",
    ],
    conclusionOk: 'התקרות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו ליקויים בתקרות התותב. יש לטפל בליקויים בהתאם לטבלת הליקויים המצורפת. שאר התקרות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: true,
    defectsColumns: ['מיקום', 'ממצאים וליקויים', 'הדרישה', 'תמונות', 'קדימות ליקוי'],
    defectsColWidths: [3.5, 5.5, 4.0, 2.0, 2.0],
  },

  // group4: סקר תקופתי — 'simple' layout, first-person "ביקרתי", has location in title
  group4: {
    name: 'סקר תקופתי',
    layout: 'simple',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    tableHdrSize: 8,
    tableDataSize: 8,
    titleSuffix: true,  // subject INCLUDES "– {location}"
    introBold: true,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ביקרתי ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את יציבות האלמנטים והמבנים.`,
    tableColumns: ['האלמנט/המבנה הנבדק', 'נתונים ופירוט', 'הערות', 'תקין/לא תקין', 'קדימות ליקוי'],
    colWidths: [4.5, 5.5, 3.0, 2.0, 2.0],
    defaultNotes: [
      "על כל שינוי קונסטרוקטיבי ועיוותים באופן חיבור/תליות האלמנטים (סדקים, עיוותים, שקיעות, ניתוקים, קורוזיה וכד') – לדווח לח''מ מיד.",
      'אין להעמיס עומסים על האלמנטים שנבדקו שאינם מיועדים לכך.',
      'הבדיקה הינה ויזואלית ונכונה ליום הבדיקה.',
    ],
    hasValidityLine: false,
    conclusionOk: 'האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו מספר ליקויים והערות שיש לתקנם ולטפלם. שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: false,
  },

  // group5: סקקות — 'survey' layout, canopy-specific
  group5: {
    name: 'סקקות',
    layout: 'survey',
    titleSize: 14,
    bodySize: 8.5,
    headingSize: 9,
    tableHdrSize: 9,
    tableDataSize: 8.5,
    titleSuffix: true,  // subject includes "– {location}"
    introBold: false,   // para1 plain, para2 bold
    introTemplate: (d) => [
      {
        text: "סקר זה מתייחס למצב קיים של תשתיות הסככות והצגת הפערים הקיימים בין המצב הקיים לבין הנדרש עפ''י תקנות הבנייה.",
        bold: false,
      },
      {
        text: `הסקר בוצע לבקשת ${d.client} ב${d.location}${d.address ? ', ' + d.address : ''} בתאריך ${d.inspection_date}.`,
        bold: true,
      },
    ],
    tableColumns: ["מס'", 'מיקום', 'סוג הסככה', "מידות (מ') ונתונים", 'תקין/לא תקין', 'קדימות ליקויים', 'תמונה'],
    colWidths: [1.5, 2.5, 3.0, 4.0, 2.5, 2.5, 1.5],
    defaultNotes: [
      'מקרא קדימות לטיפול בליקויים:',
      "קדימות 1 – ''ליקוי חמור'' בהגדרתו - ליקוי/מפגע המחייב הסרתו/תיקונו וטיפולו המיידי, לאישור ובדיקה חוזרת.",
      "קדימות 2 – ''ליקוי בינוני'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 3 חודשים מתאריך דו''ח זה.",
      "קדימות 3 – ''ליקוי קל'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 6 חודשים מתאריך דו''ח זה.",
    ],
    hasValidityLine: false,
    defaultInstructions: [
      'על כל שינוי קונסטרוקטיבי בסככות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה וכד\', יש לדווח לח\'\'מ מיד.',
      'אין לטפס, להיתלות, לפרוש שלטים/מפרשי רוח ו/או להעמיס עומסים על הסככות שנבדקו.',
      'הסככות נבדקו מבדיקה ויזואלית לתקינות ושלמות כללית.',
      "תוקף הדו''ח הינו לשנה מיום הבדיקה בכפוף למסקנות הבדיקה.",
    ],
    conclusionOk: 'הסככות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו ליקויים בסככות. יש לטפל בליקויים בהתאם לטבלת הליקויים המצורפת. שאר הסככות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: true,
    defectsColumns: ['מספר סככה', 'מיקום', 'ממצאי ליקויים ודרישות', 'תמונות הליקוי', 'קדימות'],
    defectsColWidths: [2.0, 3.0, 6.0, 4.0, 2.0],
  },
  // group6: חוות דעת הנדסיות — 'opinion' layout, free-text findings + numbered conclusions
  group6: {
    name: 'חוות דעת הנדסיות',
    layout: 'opinion',
    titleSize: 14,
    bodySize: 9,
    headingSize: 10,
    titleSuffix: true,
    introBold: true,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ערכתי ביקור ב${d.location}${d.address ? ', ' + d.address : ''} ובחנתי את יציבות ותקינות הנושא המפורט לעיל.`,
    findingsSectionHeading: 'נתונים כלליים וממצאים:',
    conclusionSectionHeading: 'הערות ומסקנות:',
    defaultFindings: [],
    defaultConclusions: [],
    hasDefectsTable: false,
  },

  // group7: מסמך כללי — 'freeform' layout, completely free text
  group7: {
    name: 'מסמך כללי',
    layout: 'freeform',
    titleSize: 14,
    bodySize: 9,
    headingSize: 10,
    titleSuffix: false,
    hasDefectsTable: false,
  },

  // group8: אישור מבנים ארעיים — 'opinion' layout, temporary structure approval
  group8: {
    name: 'אישור מבנים ארעיים',
    layout: 'opinion',
    titleSize: 14,
    bodySize: 9,
    headingSize: 10,
    titleSuffix: true,
    introBold: true,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ערכתי סיור בדיקה ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את יציבות ובטיחות המבנה/המתקן הארעי המפורט לעיל.`,
    findingsSectionHeading: 'נתונים טכניים וממצאים:',
    conclusionSectionHeading: 'תנאי האישור ומסקנות:',
    defaultFindings: [
      'סוג המבנה/המתקן:',
      'חומרים וחתכים:',
      'אופן עיגון לקרקע/תשתית:',
      'עומסי תכן:',
    ],
    defaultConclusions: [
      'המבנה/המתקן נמצא יציב ובטוח לשימוש נכון ליום הבדיקה.',
      'תוקף האישור מותנה בהתאמה למסקנות הבדיקה.',
    ],
    hasDefectsTable: false,
  },
};


// ── Helpers ───────────────────────────────────────────────────────────────────

/** Convert cm to twips (docx unit). 1 cm = 567 twips — used for table widths and page margins */
const cm = (v) => Math.round(v * 567);

/** Convert cm to pixels at 96 DPI — used ONLY for ImageRun.transformation (docx expects px there) */
const cmPx = (v) => Math.round(v / 2.54 * 96);

/** Format a date string YYYY-MM-DD → D.M.YYYY */
function formatDate(isoOrStr) {
  if (!isoOrStr) return '';
  const m = String(isoOrStr).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${parseInt(m[3])}.${parseInt(m[2])}.${m[1]}`;
  return isoOrStr;
}

/** Add N days to an ISO date string, return D.M.YYYY */
function addDays(isoOrStr, days) {
  if (!isoOrStr) return '';
  const d = new Date(isoOrStr);
  if (isNaN(d.getTime())) return '';
  d.setDate(d.getDate() + days);
  return `${d.getDate()}.${d.getMonth() + 1}.${d.getFullYear()}`;
}

/**
 * Create a TextRun with RTL, Arial font and specified options.
 * opts: { size, bold, color, underline, rtl, italic }
 */
function mkRun(text, opts = {}) {
  const {
    size   = 8.5,
    bold   = false,
    color  = undefined,
    underline = false,
    rtl    = true,
    italic = false,
  } = opts;

  return new TextRun({
    text,
    font: FONT,
    size: size * 2,           // docx uses half-points
    bold,
    color,
    rightToLeft: rtl,         // docx library property name is rightToLeft, not rtl
    italics: italic,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

/**
 * Create a RTL paragraph. bidirectional:true sets paragraph direction to RTL,
 * which is required for AlignmentType.RIGHT to anchor Hebrew text to the right margin.
 * Image paragraphs use plain new Paragraph() to avoid mirroring.
 */
function mkPara(children = [], opts = {}) {
  const {
    alignment = AlignmentType.RIGHT,
    spacing   = undefined,
    pageBreak = false,
  } = opts;

  return new Paragraph({
    children,
    alignment,
    spacing,
    pageBreakBefore: pageBreak,
    bidirectional: true,
  });
}

/** No-border spec used for photo tables */
const NO_BORDER = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

/** Thin border spec for data tables */
const THIN_BORDER = {
  top:    { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  left:   { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  right:  { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
};

/**
 * Create an RTL data table.
 * - Header row: always CENTER
 * - Data cells: always CENTER (matches original document style)
 * opts: { headerBg, hdrSize, dataSize }
 */
function mkTable(headers, rows, colWidthsCm, opts = {}) {
  const { headerBg = 'EAF1DD', hdrSize, dataSize } = opts;
  const FINDINGS_COLS = new Set(['נתונים וממצאים', 'נתונים ופירוט', "מידות (מ') ונתונים", 'ממצאים וליקויים', 'ממצאי הסיור', 'ממצאי ליקויים ודרישות', 'הממצא, מהותו ומיקומו', 'הערות/פירוט ליקוי כולל סעיף', 'ממצאי ליקויים ודרישות']);
  const colWidths = colWidthsCm.map((w) => cm(w));

  // Header row – always centered
  const headerCells = headers.map((h, i) =>
    new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, fill: headerBg },
      borders: THIN_BORDER,
      children: [mkPara(
        [mkRun(h, { size: hdrSize ?? 9, bold: true })],
        { alignment: AlignmentType.CENTER }
      )],
    })
  );

  const headerRow = new TableRow({ children: headerCells, tableHeader: true });

  // Data rows – all white (no alternating tint)
  const dataRows = rows.map((row) => {
    const cells = row.map((cellText, colIdx) => {
      const text = String(cellText ?? '');
      const isNotOk   = text === 'לא תקין'          || text.startsWith('לא תקין');
      const isWatchOk = text === 'תקין - דורש מעקב' || text.startsWith('תקין - דורש מעקב');

      let color;
      let bold = false;
      if (isNotOk)        { color = 'C00000'; bold = true; }
      else if (isWatchOk) { color = 'C55A11'; bold = true; }

      return new TableCell({
        width:         { size: colWidths[colIdx] ?? colWidths[colWidths.length - 1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading:       { type: ShadingType.CLEAR, fill: 'FFFFFF' },
        borders:       THIN_BORDER,
        children: [mkPara(
          [mkRun(text, { size: dataSize ?? 8.5, bold, color })],
          { alignment: FINDINGS_COLS.has(headers[colIdx]) ? AlignmentType.RIGHT : AlignmentType.CENTER }
        )],
      });
    });

    return new TableRow({ children: cells });
  });

  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.CENTER,
    rows: [headerRow, ...dataRows],
  });
}

/** Build the standard 3-line signature block paragraphs — left-aligned per Hebrew doc convention */
function mkSignatureBlock(bodySize) {
  return [
    mkPara([mkRun('ניר בן דוד', { size: bodySize, bold: true })], { alignment: AlignmentType.LEFT, spacing: { before: 240, after: 0 } }),
    mkPara([mkRun('מהנדס מבנים B.sc', { size: bodySize })], { alignment: AlignmentType.LEFT, spacing: { after: 0 } }),
    mkPara([mkRun('מ.ר 28566561', { size: bodySize })], { alignment: AlignmentType.LEFT, spacing: { after: 0 } }),
  ];
}

/** Convert base64 string to Uint8Array */
function base64ToUint8Array(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

/** Detect image dimensions from a base64 data-URL. Returns { width, height } in pixels. */
function getImageDimensions(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 1, height: 1 });
    img.src = dataUrl;
  });
}

// ── Main export ───────────────────────────────────────────────────────────────

export async function generateDocument(data) {
  const cfg = DOC_TYPES[data.doc_type] ?? DOC_TYPES.group4;

  // Use client name as fallback when location is empty
  const effectiveData = (data.location || '').trim()
    ? data
    : { ...data, location: (data.client || '').trim() };

  // ── 1. Fetch logos ────────────────────────────────────────────────────────
  // header-logo.jpg: 8.91 × 3.52 cm (landscape letterhead)
  // footer-logo.png: 15.05 × 1.03 cm (wide thin footer strip)
  const fetchBuf = async (paths) => {
    for (const path of paths) {
      try {
        const res = await fetch(path);
        if (res.ok) return new Uint8Array(await res.arrayBuffer());
      } catch (_) { /* try next */ }
    }
    return null;
  };

  const base = import.meta.env.BASE_URL;
  const [headerLogoBuffer, footerLogoBuffer] = await Promise.all([
    fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
    fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
  ]);

  // ── 2. Build header (page num + logo) and footer (logo only) ─────────────

  const pageNumPara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    spacing: { after: 0 },
    children: [
      new TextRun({ text: 'עמוד ', font: FONT, size: 7 * 2, rtl: true }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 7 * 2 }),
      new TextRun({ text: ' מתוך ', font: FONT, size: 7 * 2, rtl: true }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 7 * 2 }),
    ],
  });

  // No bidirectional on image paragraphs
  const headerLogoPara = headerLogoBuffer
    ? new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: headerLogoBuffer, transformation: { width: HEADER_W, height: HEADER_H } })],
      })
    : null;

  const docHeader = new Header({
    children: [pageNumPara, ...(headerLogoPara ? [headerLogoPara] : [])],
  });

  // Footer: logo only, no bidirectional on image paragraph
  const footerLogoPara = footerLogoBuffer
    ? new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: footerLogoBuffer, transformation: { width: FOOTER_W, height: FOOTER_H } })],
      })
    : new Paragraph({ children: [] });

  const docFooter = new Footer({ children: [footerLogoPara] });

  // ── 3. Spacing constants ──────────────────────────────────────────────────
  const SP_BODY    = { line: 360, lineRule: 'auto', after: 0 };
  const SP_SECTION = { line: 360, lineRule: 'auto', before: 360, after: 0 };

  // ── 4. Date paragraph (LEFT aligned — matches original documents)
  const dateStr  = formatDate(data.date);
  const datePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,
    spacing: { after: 0 },
    children: [mkRun(dateStr, { size: 10 })],
  });

  // ── 5. Client block ───────────────────────────────────────────────────────
  const toLabel    = mkPara([mkRun('לכבוד',                 { size: 8 })],                        { spacing: { after: 0 } });
  const clientPara = mkPara([mkRun(data.client ?? '',        { size: 8 })],                        { spacing: { after: 0 } });
  const orgPara    = mkPara([mkRun(data.organization ?? '',  { size: 8, underline: true })],       { spacing: { after: 0 } });

  // ── 6. Subject paragraph (CENTER, bold, cfg.titleSize) ───────────────────
  // Include "– {location}" suffix only when cfg.titleSuffix is true
  const subjectText = cfg.titleSuffix
    ? `הנדון: ${data.subject ?? ''} – ${effectiveData.location}`
    : `הנדון: ${data.subject ?? ''}`;

  const subjectPara = mkPara(
    [mkRun(subjectText, { size: cfg.titleSize, bold: true, underline: true })],
    { alignment: AlignmentType.CENTER, spacing: { before: 480, after: 0 } }
  );

  // ── 7. Table (shared between layouts) ────────────────────────────────────
  const tableRows = Array.isArray(data.table_rows) ? data.table_rows : [];
  const mainTable = tableRows.length > 0
    ? mkTable(cfg.tableColumns, tableRows, cfg.colWidths, {
        hdrSize: cfg.tableHdrSize,
        dataSize: cfg.tableDataSize,
      })
    : null;

  // ── 8. Build body based on cfg.layout ────────────────────────────────────
  const bodyChildren = [
    datePara,
    toLabel,
    clientPara,
    orgPara,
    subjectPara,
  ];

  if (cfg.layout === 'simple') {
    // ── 'simple' layout: group1, group4 ────────────────────────────────────
    // Intro paragraph(s) — bold if cfg.introBold
    const introRaw = cfg.introTemplate(effectiveData);
    const introText = data.intro_extra
      ? introRaw + '\n' + data.intro_extra
      : introRaw;

    const introParas = introText.split('\n').map((line, idx) =>
      mkPara(
        [mkRun(line, { size: cfg.bodySize, bold: cfg.introBold })],
        { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY }
      )
    );
    bodyChildren.push(...introParas);

    // "נתונים וממצאים:" heading (bold)
    bodyChildren.push(
      mkPara(
        [mkRun('נתונים וממצאים:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Main data table
    if (mainTable) bodyChildren.push(mainTable);

    bodyChildren.push(
      mkPara(
        [mkRun('מסקנות והערות:', { size: cfg.bodySize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Notes (plain, cfg.bodySize) — custom overrides defaults
    const notesSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
      ? data.notes_custom
      : [...cfg.defaultNotes];

    notesSource.forEach((note) => {
      bodyChildren.push(
        mkPara([mkRun(note, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    // Validity line (group1 only)
    if (cfg.hasValidityLine) {
      const validUntil = addDays(data.inspection_date, 365);
      bodyChildren.push(
        mkPara(
          [mkRun(`תוקף האישור לשנה מיום הבדיקה ועד לתאריך ${validUntil} בכפוף למסקנות.`, { size: cfg.bodySize })],
          { spacing: SP_BODY }
        )
      );
    }

    bodyChildren.push(
      mkPara(
        [mkRun('מסקנות:', { size: cfg.bodySize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Conclusion text (custom overrides standard)
    let conclusionText;
    if (data.conclusion_custom && String(data.conclusion_custom).trim()) {
      conclusionText = String(data.conclusion_custom).trim();
    } else if (data.has_defects) {
      conclusionText = cfg.conclusionDefects;
    } else {
      conclusionText = cfg.conclusionOk;
    }

    conclusionText.split('\n').forEach((line) => {
      bodyChildren.push(
        mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));

  } else if (cfg.layout === 'survey') {
    // ── 'survey' layout: group3, group5 ────────────────────────────────────

    // "הקדמה:" heading (bold, cfg.headingSize)
    bodyChildren.push(
      mkPara(
        [mkRun('הקדמה:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Intro paragraphs: introTemplate returns an array of { text, bold }
    const introParts = cfg.introTemplate(effectiveData);
    const effectiveParts = data.intro_extra
      ? [...introParts, { text: data.intro_extra, bold: false }]
      : introParts;

    effectiveParts.forEach((part, idx) => {
      bodyChildren.push(
        mkPara(
          [mkRun(part.text, { size: cfg.bodySize, bold: part.bold })],
          { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY }
        )
      );
    });

    // "הערות:" heading (bold, cfg.headingSize)
    bodyChildren.push(
      mkPara(
        [mkRun('הערות:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Notes paragraphs including priority legend (cfg.bodySize)
    // custom overrides defaults
    const notesSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
      ? data.notes_custom
      : [...cfg.defaultNotes];

    notesSource.forEach((note) => {
      bodyChildren.push(
        mkPara([mkRun(note, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    // "נתונים וממצאים:" heading (bold, cfg.headingSize)
    bodyChildren.push(
      mkPara(
        [mkRun('נתונים וממצאים:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Main data table
    if (mainTable) bodyChildren.push(mainTable);

    // "הנחיות:" heading (bold, cfg.headingSize)
    bodyChildren.push(
      mkPara(
        [mkRun('הנחיות:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Instructions paragraphs (cfg.bodySize)
    const instructionsSource = Array.isArray(data.instructions_custom) && data.instructions_custom.length > 0
      ? data.instructions_custom
      : [...(cfg.defaultInstructions ?? [])];

    instructionsSource.forEach((inst) => {
      bodyChildren.push(
        mkPara([mkRun(inst, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    // "מסקנות הבדיקה:" heading (bold, cfg.headingSize)
    bodyChildren.push(
      mkPara(
        [mkRun('מסקנות הבדיקה:', { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Conclusion (custom overrides standard)
    let conclusionText;
    if (data.conclusion_custom && String(data.conclusion_custom).trim()) {
      conclusionText = String(data.conclusion_custom).trim();
    } else if (data.has_defects) {
      conclusionText = cfg.conclusionDefects;
    } else {
      conclusionText = cfg.conclusionOk;
    }

    conclusionText.split('\n').forEach((line) => {
      bodyChildren.push(
        mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));

  } else if (cfg.layout === 'freeform') {
    // ── 'freeform' layout: group7 (מסמך כללי) ────────────────────────────────
    // User provides all content as free text in conclusion_custom (newline = paragraph)
    // notes_custom used for an optional first section with heading

    const freeformNotes = Array.isArray(data.notes_custom) ? data.notes_custom.filter(s => s.trim()) : [];
    freeformNotes.forEach((line, idx) => {
      bodyChildren.push(
        mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY })
      );
    });

    const bodyText = data.conclusion_custom && String(data.conclusion_custom).trim()
      ? String(data.conclusion_custom).trim()
      : '';

    bodyText.split('\n').forEach((line, idx) => {
      bodyChildren.push(
        mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY })
      );
    });

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));

  } else if (cfg.layout === 'opinion') {
    // ── 'opinion' layout: group6, group8 ─────────────────────────────────────

    // Intro paragraph
    const introText = typeof cfg.introTemplate === 'function'
      ? cfg.introTemplate(effectiveData)
      : '';
    const fullIntro = data.intro_extra ? introText + '\n' + data.intro_extra : introText;

    fullIntro.split('\n').forEach((line, idx) => {
      bodyChildren.push(
        mkPara(
          [mkRun(line, { size: cfg.bodySize, bold: cfg.introBold })],
          { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY }
        )
      );
    });

    // "נתונים כלליים וממצאים:" heading
    bodyChildren.push(
      mkPara(
        [mkRun(cfg.findingsSectionHeading, { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Findings as bullet paragraphs
    const findingsSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
      ? data.notes_custom
      : cfg.defaultFindings;

    findingsSource.forEach((item) => {
      const text = String(item).replace(/^[•\-]\s*/, '');
      bodyChildren.push(
        mkPara([mkRun(`• ${text}`, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    // "הערות ומסקנות:" heading
    bodyChildren.push(
      mkPara(
        [mkRun(cfg.conclusionSectionHeading, { size: cfg.headingSize, bold: true })],
        { spacing: SP_SECTION }
      )
    );

    // Conclusions as numbered paragraphs
    const conclusionsRaw = data.conclusion_custom && String(data.conclusion_custom).trim()
      ? String(data.conclusion_custom).trim()
      : '';

    if (conclusionsRaw) {
      conclusionsRaw.split('\n').forEach((line, i) => {
        bodyChildren.push(
          mkPara([mkRun(`${i + 1}. ${line}`, { size: cfg.bodySize })], { spacing: SP_BODY })
        );
      });
    }

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));

  } else {
    // ── 'gap-survey' layout: group2 (סקר פערי בטיחות) ─────────────────────

    // Intro paragraphs (no heading — straight paragraphs)
    const introParts = cfg.introTemplate(effectiveData);
    const effectiveParts = data.intro_extra
      ? [...introParts, { text: data.intro_extra, bold: false }]
      : introParts;

    effectiveParts.forEach((part, idx) => {
      bodyChildren.push(
        mkPara([mkRun(part.text, { size: cfg.bodySize, bold: part.bold })], { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY })
      );
    });

    // Priority legend (הערות section)
    bodyChildren.push(
      mkPara([mkRun('הערות:', { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION })
    );
    const notesSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
      ? data.notes_custom
      : [...cfg.defaultNotes];
    notesSource.forEach((note) => {
      bodyChildren.push(
        mkPara([mkRun(note, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    // Authorization line + validity line
    bodyChildren.push(
      mkPara([mkRun(cfg.authorizationLine, { size: cfg.bodySize })], { spacing: SP_BODY })
    );
    bodyChildren.push(
      mkPara([mkRun(cfg.validityLine, { size: cfg.bodySize })], { spacing: SP_BODY })
    );

    // "נתונים וממצאים:" heading + main table
    bodyChildren.push(
      mkPara([mkRun('נתונים וממצאים:', { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION })
    );
    if (mainTable) bodyChildren.push(mainTable);

    // Approvals checklist table (fixed template rows — empty, user fills in Word)
    if (cfg.approvalsTableHeaders) {
      bodyChildren.push(
        mkPara([mkRun('ריכוז בדיקות בטיחות:', { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION })
      );
      const emptyApprovalsRows = Array.from({ length: 20 }, () =>
        Array(cfg.approvalsTableHeaders.length).fill('')
      );
      bodyChildren.push(
        mkTable(cfg.approvalsTableHeaders, emptyApprovalsRows, cfg.approvalsColWidths, {
          hdrSize: cfg.tableHdrSize,
          dataSize: cfg.tableDataSize,
        })
      );
    }

    // Conclusion heading + fixed lines
    bodyChildren.push(
      mkPara([mkRun(cfg.conclusionHeading, { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION })
    );
    const conclusionLines = data.conclusion_custom && String(data.conclusion_custom).trim()
      ? String(data.conclusion_custom).trim().split('\n')
      : cfg.conclusionLines;
    conclusionLines.forEach((line) => {
      bodyChildren.push(
        mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: SP_BODY })
      );
    });

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));
  }

  // ── 9. Defects table (if applicable) ─────────────────────────────────────
  if (
    data.has_defects &&
    cfg.hasDefectsTable &&
    cfg.defectsColumns &&
    Array.isArray(data.defects_rows) &&
    data.defects_rows.length > 0
  ) {
    const defectsTitlePara = mkPara(
      [mkRun('טבלת הליקויים', { size: cfg.headingSize, bold: true })],
      { pageBreak: true }
    );
    const defectsTable = mkTable(
      cfg.defectsColumns,
      data.defects_rows,
      cfg.defectsColWidths,
      { hdrSize: cfg.tableHdrSize, dataSize: cfg.tableDataSize }
    );
    bodyChildren.push(defectsTitlePara, defectsTable, mkPara([mkRun('')]));
  }

  // ── 10. Photos section ────────────────────────────────────────────────────
  const photos = Array.isArray(data.photos) ? data.photos : [];

  if (photos.length > 0) {
    const photosTitlePara = mkPara(
      [mkRun('נספח תמונות:', { size: cfg.headingSize, bold: true })],
      { pageBreak: true }
    );

    // Load all photos (skip failures)
    const loadedPhotos = [];
    for (let i = 0; i < photos.length; i++) {
      const photo = photos[i];
      try {
        const dataUrl = photo.data;
        if (!dataUrl) continue;

        const b64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
        const imgData = base64ToUint8Array(b64);

        const { width: natW, height: natH } = await getImageDimensions(dataUrl);
        const isPortrait = natH > natW;

        // Compute display dimensions in cm
        let dispW, dispH;
        if (isPortrait) {
          dispH = 8; // cm absolute height
          dispW = natW > 0 ? (natW / natH) * dispH : 6;
        } else {
          dispW = 8; // cm absolute width
          dispH = natH > 0 ? (natH / natW) * dispW : 5;
        }

        const caption = photo.caption
          ? `תמונה ${i + 1} - ${photo.caption}`
          : `תמונה ${i + 1}`;

        loadedPhotos.push({ imgData, dispW, dispH, caption, index: i + 1 });
      } catch (_) {
        // skip broken photos
      }
    }

    // Build 2-per-row table
    const photoTableRows = [];
    for (let r = 0; r < loadedPhotos.length; r += 2) {
      const left  = loadedPhotos[r];
      const right = loadedPhotos[r + 1] ?? null;

      const makePhotoCell = (p) => {
        // No bidirectional on image paragraphs
        const imgPara = new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: p.imgData,
              transformation: {
                width:  cmPx(p.dispW),
                height: cmPx(p.dispH),
              },
            }),
          ],
        });

        const capPara = mkPara(
          [mkRun(p.caption, { size: 9 })],
          { alignment: AlignmentType.CENTER }
        );

        return new TableCell({
          width: { size: cm(8.5), type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          borders: NO_BORDER,
          children: [capPara, imgPara],
        });
      };

      const emptyCell = () =>
        new TableCell({
          width: { size: cm(8.5), type: WidthType.DXA },
          borders: NO_BORDER,
          children: [mkPara([mkRun('')])],
        });

      photoTableRows.push(
        new TableRow({
          children: [
            makePhotoCell(left),
            right ? makePhotoCell(right) : emptyCell(),
          ],
        })
      );
    }

    const photoTable = new Table({
      visuallyRightToLeft: true,
      layout: TableLayoutType.FIXED,
      rows: photoTableRows,
    });

    bodyChildren.push(photosTitlePara, photoTable);
  }

  // ── Build Document ────────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: {
          // Set RTL + right-align as document-wide defaults (docDefaults in styles.xml)
          paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT },
          run: { font: { name: FONT } },
        },
      },
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          quickFormat: true,
          paragraph: {
            bidirectional: true,
            alignment: AlignmentType.RIGHT,
          },
          run: { font: { name: FONT } },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: mm(210), height: mm(297) },
            margin: {
              top:    mm(31.70),
              bottom: mm(12.51),
              left:   mm(20.00),
              right:  mm(24.98),
              header: mm(3.00),
              footer: mm(1.99),
            },
          },
          // Note: bidi is injected via JSZip post-processing below (docx library ignores it here)
        },
        headers: { default: docHeader },
        footers: { default: docFooter },
        children: bodyChildren,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);

  // Post-process: inject <w:bidi/> into <w:sectPr> so Word treats the section as RTL.
  // The docx library silently ignores bidi:true in SectionProperties, so we patch the XML directly.
  try {
    const zip = await JSZip.loadAsync(blob);
    const docXml = await zip.file('word/document.xml').async('string');
    // Insert <w:bidi/> just before the closing </w:sectPr> tag (once — the last sectPr is the section)
    const patched = docXml.replace(/<\/w:sectPr>/, '<w:bidi/></w:sectPr>');
    if (patched !== docXml) {
      zip.file('word/document.xml', patched);
      return await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });
    }
  } catch (_) {
    // If post-processing fails, return the original blob
  }

  return blob;
}
