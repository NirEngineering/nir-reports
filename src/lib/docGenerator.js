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
  LevelFormat,
  VerticalMergeType,
} from 'docx';

// ── Font constant ─────────────────────────────────────────────────────────────
const FONT = 'Arial';

// Exact pixel sizes derived from sample Word document (914400 EMU = 1 inch = 96px)
// Verified against the header/footer <wp:extent> in two independent, current
// (2026) real Drive documents — identical in both, so treated as the standard.
const HEADER_W = 337, HEADER_H = 133;   // 89.06mm × 35.18mm
const FOOTER_W = 569, FOOTER_H = 39;    // 150.50mm × 10.28mm

// ── Doc-type configuration ────────────────────────────────────────────────────
// Per-type font sizes and content extracted from 48 original Hebrew Word documents
const DOC_TYPES = {
  // group1: אלמנטים תלויים — 'simple' layout, first-person "ערכתי"
  group1: {
    name: 'אלמנטים תלויים',
    layout: 'simple',
    titleSize: 16,
    bodySize: 10,
    headingSize: 10,
    tableHdrSize: 10,
    tableDataSize: 9,
    titleSuffix: false,  // subject does NOT include "– {location}"
    introBold: false,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ערכתי סיור בדיקה ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את יציבות האלמנטים והמתקנים התלויים.`,
    tableColumns: ['מיקום', 'האלמנט/המתקן', 'נתונים וממצאים', 'הערות'],
    colWidths: [3.2, 4.1, 5.4, 2.7],
    defaultNotes: [
      "על כל שינוי קונסטרוקטיבי ועיוותים באופן חיבור/תליות האלמנטים (סדקים, עיוותים, שקיעות, ניתוקים, קורוזיה וכד') – לדווח לח''מ מיד.",
      'אין להעמיס עומסים על האלמנטים שנבדקו שאינם מיועדים לכך.',
      'הבדיקה הינה ויזואלית ונכונה ליום הבדיקה.',
    ],
    validityDays: 365,
    conclusionOk: 'האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו מספר ליקויים והערות שיש לתקנם ולטפלם. שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: false,
  },

  // group2: סקר פערי בטיחות — 'gap-survey' layout
  group2: {
    name: 'סקר פערי בטיחות',
    layout: 'gap-survey',
    titleSize: 15,
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
    colWidths: [0.9, 2.6, 2.1, 3.0, 4.7, 2.1],
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
    // Presence of this flag (not its content — see mkApprovalsTable) triggers
    // rendering the ריכוז בדיקות בטיחות table for this type.
    approvalsTableHeaders: true,
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
    colWidths: [3.2, 3.2, 2.7, 2.7, 3.6],
    defaultNotes: [
      "ממצאי סקר זה הם כפי שהועברו לח''מ ע''י בעלי התפקידים באתר ומציגים מצב קיים ביום הסיור בלבד, מזמין העבודה אחראי לביצוע תיקון הליקויים שנמצאו בפרק זמן שהוגדר, לא תשמע טענה כנגד הח''מ בגין ליקויים שהצביע עליהם במסגרת סקר זה ושאינם תוקנו בתוך מסגרת הזמן שנקבעה.",
      "מצורפת טבלת ליקויים בהמשך המסמך הכוללת את מיקום הליקוי, פירוט הממצא, דרישות/הנחיות לטיפול ותמונות הליקוי.",
      'מקרא קדימות לטיפול בליקויים:',
      "קדימות 1 – ''ליקוי חמור'' בהגדרתו - ליקוי/מפגע המחייב הסרתו/תיקונו וטיפולו המיידי, לאישור ובדיקה חוזרת.",
      "קדימות 2 – ''ליקוי בינוני'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 3 חודשים מתאריך דו''ח זה.",
      "קדימות 3 – ''ליקוי קל'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 6 חודשים מתאריך דו''ח זה.",
    ],
    validityDays: 1825,
    defaultInstructions: [
      "על כל שינוי קונסטרוקטיבי בתקרות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה, כפף, שבר, סדק, תזוזות וכו' – יש לזמן לבדיקה חוזרת וטיפול מתאים עפ''י הממצאים.",
      'אין לטפס, להיתלות ו/או להעמיס עומסים על התקרות שנבדקו.',
    ],
    conclusionOk: 'התקרות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו ליקויים בתקרות התותב. יש לטפל בליקויים בהתאם לטבלת הליקויים המצורפת. שאר התקרות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: true,
    defectsColumns: ['מיקום', 'ממצאים וליקויים', 'הדרישה', 'תמונות', 'קדימות ליקוי'],
    defectsColWidths: [3.2, 5.0, 3.6, 1.8, 1.8],
    // Fixed legal/technical appendix, verified present (word-for-word) in a
    // real ceiling-survey report from August 2026 — appears in every תקרות
    // תותב report, unlike סככות which has no such appendix.
    appendixTitle: "נספח דרישות לתכן והתקנה של תקרות תותב פריקות לפי ת''י 5103",
    appendixIntro: "להלן רשימת סעיפי דרישת ת''י 5103 (דרישות לתכן ולהתקנה של תקרות תותב פריקות לא נושאות)",
    appendixClauses: [
      { label: "סעיף 5.3 (ת''י 5103 חלק 1):", lines: [
        "המתלים יהיו מבין האפשריות האלה: מוט פלדה מגלוון בקוטר 3.5 מ''מ לפחות, סרט פלדה מגלוון ברוחב 18 מ''מ לפחות ובעובי 0.8 מ''מ לפחות, מוטות הברגה מפלדה מגלוונים בקוטר 6 מ''מ לפחות.",
      ] },
      { label: "סעיף 5.4 (ת''י 5103 חלק 1):", lines: [
        "המרחק בין המתלים לאורך המנשא הראשי לא יהיה גדול מ-1200 מ''מ.",
        "המרחק בין המתלה הראשון (במנשא הראשי) לבין הקיר או הקצה לא יהיה גדול מ-400 מ''מ.",
      ] },
      { label: "סעיף 5.5 (ת''י 5103 חלק 1):", lines: [
        "מקבעים עליונים יעמדו בעומס שליפה או גזירה (בהתאם לאופן התקנתם) השווה למשקל שעליהם לשאת כפול 4 אך לא פחות מ-500 ניוטון למקבעים העליונים ולא קטן מ-300 ניוטון למתלים. (500 ניוטון = 50 ק''ג).",
        "בשימוש בבורג ומיתר יעוגן המתלה ע''י שימוש בדסקית בעובי 1 מ''מ לפחות ובקוטר 15 מ''מ לפחות מתחת לראש הבורג.",
      ] },
      { label: "סעיף 5.9 (ת''י 5103 חלק 1):", lines: [
        "עובי הדופן המינימלי של פרופילי אלומיניום ופרופילי אומגה יהיה 1 מ''מ.",
      ] },
      { label: "סעיף 5.21 (ת''י 5103 חלק 1):", lines: [
        "מגשים שמשולבים בהם אביזרים כגון גופי תאורה או התקני קצה של מערכות איורור יחוברו בצורה מכנית לפרופילי ההיקף.",
      ] },
      { label: "סעיף 7.2.1 (ת''י 5103 חלק 1):", lines: [
        "גופי תאורה יחוברו באמצעות מתלים נפרדים לתקרה המבנית או לקונס' הנושאת.",
      ] },
      { label: "סעיף 7.3.1 (ת''י 5103 חלק 1):", lines: [
        "כל האביזרים שמעל תקרת התותב ו/או אלה המשולבים בה יחוברו באמצעות מתלים נפרדים לתקרה המבנית.",
      ] },
      { label: "סעיף 6.3 (ת''י 5103 חלק 3):", lines: [
        "לא יעשה שימוש במסמרות לקיבוע החלק העליון או התחתון של המתלה.",
        "המתלים יהיו אנכיים או אנכיים בקירוב ולא ילחצו על חומרי בידוד המכסים מובלים או צינורות.",
        "בשום מקרה אין לעשות חיתוך בסרט הפלדה לצורך השחלתו למנשא הראשי.",
      ] },
      { label: "נספח א- א5 (ת''י 5103 חלק 3):", lines: [
        "קיבוע בקורות עץ יעשה רק בתוך הדופן הצדדית של הקורה בגובה 50 מ''מ מחלקה התחתון. במידה ואין אפשרות, ניתן להתחבר לדופן תחתונה באמצעות ברגי עץ בעומק 38 מ''מ לפחות.",
      ] },
      { label: "נספח א- א7 (ת''י 5103 חלק 3):", lines: [
        "בקיבוע לתקרת צלעות יאותרו צלעות הבטון והחיבור אליהם יעשה במיתדים וברגים בעומק 35 מ''מ לפחות. (כנ''ל לגבי העומק בתקרת בטון מקשית).",
      ] },
    ],
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
    introBold: false,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ביקרתי ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את יציבות האלמנטים והמבנים.`,
    tableColumns: ['האלמנט/המבנה הנבדק', 'נתונים ופירוט', 'הערות', 'תקין/לא תקין', 'קדימות ליקוי'],
    colWidths: [4.1, 5.0, 2.7, 1.8, 1.8],
    defaultNotes: [
      "על כל שינוי קונסטרוקטיבי ועיוותים באופן חיבור/תליות האלמנטים (סדקים, עיוותים, שקיעות, ניתוקים, קורוזיה וכד') – לדווח לח''מ מיד.",
      'אין להעמיס עומסים על האלמנטים שנבדקו שאינם מיועדים לכך.',
      'הבדיקה הינה ויזואלית ונכונה ליום הבדיקה.',
    ],
    conclusionOk: 'האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו מספר ליקויים והערות שיש לתקנם ולטפלם. שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: false,
  },

  // group5: סקקות — 'survey' layout, canopy-specific
  group5: {
    name: 'סקקות',
    layout: 'survey',
    titleSize: 15,
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
    colWidths: [1.4, 2.2, 2.6, 3.6, 2.2, 2.2, 1.2],
    defaultNotes: [
      'מקרא קדימות לטיפול בליקויים:',
      "קדימות 1 – ''ליקוי חמור'' בהגדרתו - ליקוי/מפגע המחייב הסרתו/תיקונו וטיפולו המיידי, לאישור ובדיקה חוזרת.",
      "קדימות 2 – ''ליקוי בינוני'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 3 חודשים מתאריך דו''ח זה.",
      "קדימות 3 – ''ליקוי קל'' בהגדרתו - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 6 חודשים מתאריך דו''ח זה.",
    ],
    validityDays: 365,
    defaultInstructions: [
      'על כל שינוי קונסטרוקטיבי בסככות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה וכד\', יש לדווח לח\'\'מ מיד.',
      'אין לטפס, להיתלות, לפרוש שלטים/מפרשי רוח ו/או להעמיס עומסים על הסככות שנבדקו.',
      'הסככות נבדקו מבדיקה ויזואלית לתקינות ושלמות כללית.',
    ],
    conclusionOk: 'הסככות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'נמצאו ליקויים בסככות. יש לטפל בליקויים בהתאם לטבלת הליקויים המצורפת. שאר הסככות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה.',
    hasDefectsTable: true,
    defectsColumns: ['מספר סככה', 'מיקום', 'ממצאי ליקויים ודרישות', 'תמונות הליקוי', 'קדימות'],
    defectsColWidths: [1.8, 2.7, 5.5, 3.6, 1.8],
  },
  // group6: חוות דעת הנדסיות — 'opinion' layout, free-text findings + numbered conclusions
  group6: {
    name: 'חוות דעת הנדסיות',
    layout: 'opinion',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    titleSuffix: true,
    introBold: false,
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
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    titleSuffix: false,
    hasDefectsTable: false,
  },

  // group8: אישור מבנים ארעיים — 'opinion' layout, temporary structure approval
  group8: {
    name: 'אישור מבנים ארעיים',
    layout: 'opinion',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    titleSuffix: true,
    introBold: false,
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

  // group9: אישור ביצוע חיזוק — 'opinion' layout, execution-of-strengthening-works approval
  group9: {
    name: 'אישור ביצוע חיזוק',
    layout: 'opinion',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    titleSuffix: false,
    subjectLocationLine: true,
    introBold: false,
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} ביקרתי ב${d.location}${d.address ? ', ' + d.address : ''} ובדקתי את ביצוע עבודות החיזוק המפורטות לעיל.`,
    findingsSectionHeading: 'העבודות כללו:',
    plainFindings: true,
    conclusionSectionHeading: 'הערות:',
    defaultFindings: [],
    defaultConclusions: [
      'האישור מתייחס לביצוע עבודות המפורטות במסמך זה בלבד.',
      'אין לבצע שינויים קונסטרוקטיבים, אין להחסיר ו/או להוסיף רכיבים.',
      "על כל שינוי קונסטרוקטיבי ועיוותים כלשהם (סדקים, עיוותים, שקיעות, ניתוקים, חלודה, אלמנטים רופפים, חוסרים/תוספות וכדומה) יש לדווח על כך לבדיקה חוזרת וטיפול מתאים עפ''י הממצאים.",
      "תוקף המסמך בהתאם לת''י 1525 לחמש שנים מיום כתיבתו ו/או עד ביצוע שינויים מיבניים או בתכולה המצוינת לעיל – המוקדם מביניהם.",
    ],
    hasDefectsTable: false,
  },

  // group10: אישור רשת הגנה — 'opinion' layout, falling-object safety-net stability approval
  group10: {
    name: 'אישור רשת הגנה',
    layout: 'opinion',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    titleSuffix: false,
    subjectLocationLine: true,
    introBold: false,
    introTemplate: () =>
      'נבדקה התקנת רשת הגנה ובלימת נפילת חפצים באתר שבנדון, שבוצעה על ידי צוות המתקינים.',
    findingsSectionHeading: null,
    plainFindings: true,
    conclusionSectionHeading: 'הערות:',
    defaultFindings: [],
    defaultConclusions: [
      'יש לפנות במיידי חפצ/ים שנפל/ו על חלקי ורכיבי מערכת ההגנה.',
      'במקרה הצורך יש לבצע חיזוקים ומתיחות בהתאם כנדרש.',
      "על כל שינוי במערכת ההגנה והאחיזה יש לזמן לביקורת נוספת וטיפול מתאים עפ''י הממצאים.",
      'אין להעמיס עומסים ולבצע שינויים תוספות/הפחתות על מערכת ההגנה שבוצעה.',
      'תוקף המסמך הינו לשנה מיום כתיבתו.',
    ],
    hasDefectsTable: false,
  },

  // group11: אישור בטיחות לקייטנה — 'checklist' layout, fixed 15-item safety checklist
  group11: {
    name: 'אישור בטיחות לקייטנה',
    layout: 'checklist',
    titleSize: 15,
    bodySize: 9,
    headingSize: 9,
    tableHdrSize: 9,
    tableDataSize: 8,
    titleSuffix: false,
    introTemplate: (d) => [
      `הנני מאשר כי ערכתי מבדק בטיחות בשטח הקייטנה במבנה ${d.location}${d.address ? ', ' + d.address : ''}, הכולל את תשתית המבנה, מצורפת רשימת עזר לבדיקת בטיחות וטבלת ריכוז אישורים נדרשים וליקויים שנתגלו שיש לטפלם לפני תחילת הקייטנה.`,
      'אישור הבטיחות מותנה בהצגת כלל האישורים הרלוונטיים הנדרשים המפורטים בסוף המסמך.',
    ],
    checklistHeading: 'רשימת עזר לבדיקת הבטיחות:',
    checklistColumns: ["מס'", 'נושא הבדיקה', 'קיים/חסר/בוצע/לא בוצע', 'הערות'],
    checklistColWidths: [0.9, 8.8, 3.0, 2.7],
    defectsHeading: 'טבלת ליקויים לטיפול',
    defectsColumns: ['', 'תחום הבדיקה', 'סעיף ברשימת המבדק', 'הדרישה', 'הממצא, מהותו ומיקומו', 'קדימות הליקוי'],
    defectsColWidths: [0.9, 2.6, 2.1, 3.0, 4.7, 2.1],
    approvalsHeading: 'ריכוז בדיקות בטיחות:',
  },
};

// ── Fixed 15-item safety checklist (group11, קייטנה) — identical wording in
// every real Drive sample found; only the status/הערות per item vary per site.
const CAMP_CHECKLIST_ITEMS = [
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

// The "ריכוז בדיקות בטיחות" table — external certifications/inspections the
// client must be able to present. Domain/frequency/checking-body text is
// fixed boilerplate verified verbatim against real Drive samples; only the
// הוצג/לא הוצג status column is filled in per visit (left blank by default,
// matching most rows in the real samples). Shared by group2 (סקר פערי
// בטיחות) and group11 (אישור בטיחות לקייטנה), which use the identical table.
const SAFETY_APPROVALS_ROWS = [
  { domain: 'יציבות ותקינות המבנים (לרבות מבנים יבילים)', freq: 'במקרה ונצפו כשלי יציבות דוגמת שקיעות וסדקים.\nמבנה יביל – אחת לחמש שנים', checker: 'מהנדס מבנים' },
  { domain: 'יציבות ותקינות סככות', freq: 'בהקמה ואחת לחמש שנים.\nלפי הצורך (שקיעות, קורוזיה, סדקים)', checker: 'מהנדס/הנדסאי מבנים' },
  { domain: 'יציבות עמודי תאורה', freq: 'בהקמה ואחת לחמש שנים.\nלפי הצורך (שקיעות, קורוזיה, סדקים)', checker: 'מהנדס/הנדסאי מבנים' },
  { domain: 'תקרות תלויות', freq: 'בהקמה ואחת לחמש שנים.\nלפי הצורך (שקיעות, קורוזיה, סדקים)', checker: 'מהנדס/הנדסאי מבנים' },
  { domain: 'מנשאים תלויים למזגנים', freq: 'בהקמה ואחת לחמש שנים.\nלפי הצורך (שקיעות, קורוזיה, סדקים, ריקבון)', checker: 'מהנדס/הנדסאי מבנים' },
  { domain: 'תחנת הסעה ומסופים להסעות\nסידורי בטיחות בתחנת איסוף והורדת ילדים ובמסופי הסעה הצמודים למוסד חינוכי', freq: 'בהקמה ולאחר שינוי', checker: 'מהנדס תנועה' },
  { domain: 'מתקני משחקים', freq: 'בהתאם לדרישות תקן מתקני משחק 1498', checker: 'הצגת אישור לתחזוקת מתקני המשחקים בתו תקן ממכון התקנים לפי ת"י 1498' },
  { domain: 'מתקני כושר בשטחי חוץ', freq: 'בהתאם לדרישות תקן מתקני כושר 1497', checker: 'אישור מעבדה - התאמה לתקן ישראלי 1497' },
  { domain: 'וילונות חלוקה באולמות', freq: 'בהתאם לדרישות תקן וילונות חלוקה 5517', checker: 'אישור מעבדה – התאמה לתקן 5517' },
  { domain: 'מתקני סל וספורט \nבמגרשים ובאולמות.', freq: 'בהקמה ואחת לשנה', checker: 'אישור מעבדה – התאמה לתקן 5515' },
  { domain: 'מוצג', freq: 'בהקמה ולפי הצורך', checker: 'הנדסאי מבנים,\nבמוצג המשלב חשמל אישור בודק חשמל' },
  { domain: 'חשמל', domainGroup: true, sub: 'מערכת חשמל', freq: 'בתום חמש שנים מקבלת טופס 4 ובתדירות של אחת לחמש שנים.', checker: 'חשמלאי בודק מתאים' },
  { domain: 'חשמל', domainGroup: true, sub: 'ציוד ומכשירי חשמל (ויזואלית בלבד)', freq: 'שנתית', checker: 'חשמלאי מוסמך' },
  { domain: 'מערכת גז', freq: 'בהקמה ואחת לחמש שנים', checker: 'התאמה לתקן 158 – טכנאי גז סוג 2' },
  { domain: 'ציוד וכלים טעוני בדיקה: \n(מעליות, מתקני הרמה, אבזרי הרמה, מתקני לחץ, דודי קיטור, קולטי קיטור, אוטוקלבים)', freq: 'על פי פקודת הבטיחות בעבודה-1970: למעליות כל 6 חודשים\nלמכונת ההרמה כל 14 חודשים\nלאבזרי ההרמה כל 6 חודשים\nלמתקני הלחץ כל 26 חודשים\nלדודי הקיטור ולקולטי קיטור כל 26 חודשים', checker: 'אישור בדיקה בתוקף מבודק מוסמך (כהגדרתו בפקודת הבטיחות בעבודה לכל סוג של מתקן בהתאמה) \n\nמהנדס מבנים' },
  { domain: 'מדידת קרינה אלקטרומגנטית (מדידת קרינה בתחום תדרי הרדיו) בתחום תדרי הרדיו RF באזורי שהייה.\nקרינה סלולרית\nאינטרנט אלחוטי', freq: 'בהתקנה ובמידה ויש שינוי ברשת ובהיקף הציוד ומכשירי הקצה.', checker: 'בודק ובעל ציוד המוסמך ע"י המשרד להגנת הסביבה.\nביצוע על פי הנחיות המשרד להגנת הסביבה מספר 09-04-01' },
  { domain: 'מדידת שדות מגנטיים וחשמליים בתחום תדר רשת החשמל ELF באזורי שהייה\nשנאי חברת חשמל\nקווי חשמל תת קרקעיים\nארונות חשמל\nמחשבים\nציוד חשמלי', freq: 'בהתקנה ובמידה ויש שינוי ברשת ובהיקף הציוד ומכשירי הקצה.', checker: 'בודק ובעל ציוד המוסמך ע"י המשרד להגנת הסביבה.\nביצוע על פי הנחיות המשרד להגנת הסביבה מספר 09-04-01' },
  { domain: 'יציבות עצים וענפים\n(בתחום בית הספר ובסמוך לגדרות מבחוץ)', freq: 'שנתית', checker: 'אגרונום או גוזם עצים מוסמך' },
  { domain: 'מוכנות ותקינות ציוד כיבוי אש', freq: 'שנתית', checker: 'חברה המאושרת לביקורת ותחזוקה לציוד גילוי וכיבוי אש' },
];


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

function mkRun(text, opts = {}) {
  const {
    size   = 8.5,
    bold   = false,
    color  = undefined,
    underline = false,
    italic = false,
  } = opts;

  return new TextRun({
    text,
    font: FONT,
    size: size * 2,
    bold,
    color,
    italics: italic,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
    // Every run in the real documents carries <w:rtl/> (run-level RTL) in
    // addition to the paragraph's own bidi flag — without it, an embedded
    // Latin/digit sequence (dates, numbers) at the end of a Hebrew sentence
    // can throw off the ordering of trailing punctuation.
    rightToLeft: true,
  });
}

/**
 * Create a RTL paragraph. bidirectional:true sets paragraph direction to RTL.
 * Default is AlignmentType.LEFT, not RIGHT — counterintuitive, but confirmed
 * against a real generated document opened in Word: for a bidirectional
 * paragraph, docx's "right" justification renders on the PHYSICAL LEFT of
 * the page and "left" renders on the physical right.
 */
function mkPara(children = [], opts = {}) {
  const {
    alignment = AlignmentType.LEFT,
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

// The list numbers under "מסקנות והערות:" / "מסקנות:" used to be a plain
// "1. " string glued onto the Hebrew line — a real Word numbered list is
// what actually keeps the digit on the right and the sentence flowing
// right-to-left instead of the two fighting over which side is "first".
const NUMBERING_REFERENCE = 'auto-numbered-list';
function mkNumberedPara(text, opts = {}) {
  const { size = 8.5, spacing = undefined } = opts;
  return new Paragraph({
    children: [mkRun(text, { size })],
    alignment: AlignmentType.LEFT,
    spacing,
    bidirectional: true,
    numbering: { reference: NUMBERING_REFERENCE, level: 0 },
  });
}

// Shared by the 'simple' and 'survey' layouts: an explicit expiry date
// (computed from the visit date, not just "valid for N years" prose) for
// the types that carry one — 1 year for hanging-element/canopy inspections,
// 5 years for ceiling surveys.
function mkValidityLine(cfg, data, spacing) {
  if (!cfg.validityDays) return null;
  const validUntil = addDays(data.inspection_date, cfg.validityDays);
  if (!validUntil) return null;
  const years = cfg.validityDays >= 1825 ? 'חמש שנים' : 'שנה';
  return mkPara([mkRun(`תוקף הדו''ח הינו ל${years} מיום הבדיקה ועד לתאריך ${validUntil} בכפוף למסקנות הבדיקה.`, { size: cfg.bodySize })], { spacing });
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
  const { headerBg = 'EAF1DD', hdrSize, dataSize, noMerge = false } = opts;
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
  const dataRows = rows.map((row, rowIdx) => {
    const cells = row.map((cellText, colIdx) => {
      const text = String(cellText ?? '');
      const isNotOk   = text === 'לא תקין'          || text.startsWith('לא תקין');
      const isWatchOk = text === 'תקין - דורש מעקב' || text.startsWith('תקין - דורש מעקב');

      let color;
      let bold = false;
      if (isNotOk)        { color = 'C00000'; bold = true; }
      else if (isWatchOk) { color = 'C55A11'; bold = true; }

      // Consecutive rows repeating the same value in a column (typically
      // מיקום/סוג האלמנט when one location has several findings) merge into
      // one spanning cell instead of repeating the text on every row. Fixed
      // checklists (e.g. group11) disable this via noMerge — there, a
      // repeated status value like "קיים" across unrelated rows must NOT
      // visually merge them together.
      const prevText = rowIdx > 0 ? String(rows[rowIdx - 1][colIdx] ?? '') : null;
      const nextText = rowIdx < rows.length - 1 ? String(rows[rowIdx + 1][colIdx] ?? '') : null;
      const continuesFromAbove = !noMerge && text !== '' && text === prevText;
      const continuesBelow     = !noMerge && text !== '' && text === nextText;
      const verticalMerge = continuesFromAbove
        ? VerticalMergeType.CONTINUE
        : (continuesBelow ? VerticalMergeType.RESTART : undefined);

      return new TableCell({
        width:         { size: colWidths[colIdx] ?? colWidths[colWidths.length - 1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading:       { type: ShadingType.CLEAR, fill: 'FFFFFF' },
        borders:       THIN_BORDER,
        verticalMerge,
        children: [mkPara(
          [mkRun(continuesFromAbove ? '' : text, { size: dataSize ?? 8.5, bold, color })],
          { alignment: AlignmentType.CENTER }
        )],
      });
    });

    return new TableRow({ children: cells });
  });

  const totalWidth = colWidths.reduce((sum, w) => sum + w, 0);
  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.CENTER,
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// One or more stacked paragraphs for a table cell whose text contains
// embedded newlines (SAFETY_APPROVALS_ROWS' תדירות/checker text is often
// multi-line) — a single mkPara/mkRun pair can't represent a line break.
function mkCellLines(text, opts = {}) {
  const { size = 8, bold = false, alignment = AlignmentType.CENTER } = opts;
  const lines = String(text ?? '').split('\n');
  return lines.map((line) => mkPara([mkRun(line, { size, bold })], { alignment, spacing: { after: 0 } }));
}

// The "ריכוז בדיקות בטיחות" table (see SAFETY_APPROVALS_ROWS). Built by hand
// rather than through mkTable(): it needs a 2-way column split (domain-group
// label vs. specific sub-item) only for the two "חשמל" rows, via columnSpan/
// verticalMerge — not something the generic per-column merge logic supports.
// statusByIndex optionally supplies the הוצג/לא הוצג value per row (0-based);
// left blank by default, matching most rows in the real Drive samples.
function mkApprovalsTable(statusByIndex = {}) {
  const w = [0.8, 1.8, 2.4, 2.6, 4.7, 3.1].map((v) => cm(v));
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({ width: { size: w[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'EAF1DD' }, borders: THIN_BORDER, children: mkCellLines("מס'", { size: 8.5, bold: true }) }),
      new TableCell({ width: { size: w[1] + w[2], type: WidthType.DXA }, columnSpan: 2, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'EAF1DD' }, borders: THIN_BORDER, children: mkCellLines('תחום הבדיקה', { size: 8.5, bold: true }) }),
      new TableCell({ width: { size: w[3], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'EAF1DD' }, borders: THIN_BORDER, children: mkCellLines('תדירות', { size: 8.5, bold: true }) }),
      new TableCell({ width: { size: w[4], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'EAF1DD' }, borders: THIN_BORDER, children: mkCellLines('הגוף המקצועי הבודק והמאשר', { size: 8.5, bold: true }) }),
      new TableCell({ width: { size: w[5], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, fill: 'EAF1DD' }, borders: THIN_BORDER, children: mkCellLines('הוצג/לא הוצג', { size: 8.5, bold: true }) }),
    ],
  });

  const dataRows = SAFETY_APPROVALS_ROWS.map((row, i) => {
    const isElectric = !!row.domainGroup;
    const prevSameGroup = isElectric && SAFETY_APPROVALS_ROWS[i - 1]?.domainGroup && SAFETY_APPROVALS_ROWS[i - 1].domain === row.domain;

    const noCell = new TableCell({ width: { size: w[0], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines('') });

    const cells = [noCell];
    if (isElectric) {
      cells.push(new TableCell({
        width: { size: w[1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        borders: THIN_BORDER,
        verticalMerge: prevSameGroup ? VerticalMergeType.CONTINUE : VerticalMergeType.RESTART,
        children: prevSameGroup ? [] : mkCellLines(row.domain),
      }));
      cells.push(new TableCell({ width: { size: w[2], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines(row.sub, { alignment: AlignmentType.LEFT }) }));
    } else {
      cells.push(new TableCell({ width: { size: w[1] + w[2], type: WidthType.DXA }, columnSpan: 2, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines(row.domain, { alignment: AlignmentType.LEFT }) }));
    }
    cells.push(new TableCell({ width: { size: w[3], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines(row.freq, { alignment: AlignmentType.LEFT }) }));
    cells.push(new TableCell({ width: { size: w[4], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines(row.checker, { alignment: AlignmentType.LEFT }) }));
    cells.push(new TableCell({ width: { size: w[5], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: THIN_BORDER, children: mkCellLines(statusByIndex[i] ?? '') }));

    return new TableRow({ children: cells });
  });

  const totalWidth = w.reduce((sum, v) => sum + v, 0);
  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.CENTER,
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: w,
    rows: [headerRow, ...dataRows],
  });
}

// group11 (קייטנה) signature/fill-in line — unlike every other type's plain
// mkSignatureBlock(), the real documents use a single packed line with
// underscore-filled blanks for the inspector's identity, followed by a
// plain (non-bold) label row underneath. חתימה is always left blank.
function mkCampSignatureBlock(dateStr) {
  const fill = (val) => `____${val}____`;
  const line1 = `${fill('ניר')}          ${fill('בן דוד')}         ${fill('70382')}         ${fill(dateStr)}             __________________`;
  const line2 = '                שם פרטי                  שם משפחה                   מספר תעודה                             תאריך                                        חתימה';
  return [
    mkPara([mkRun(line1, { size: 9 })], { alignment: AlignmentType.CENTER, spacing: { before: 480, after: 0 } }),
    mkPara([mkRun(line2, { size: 8 })], { alignment: AlignmentType.CENTER, spacing: { after: 0 } }),
  ];
}

// Plain physically-left paragraphs, not the 2-column table this used to be —
// a table's position under RTL bidi was too unreliable to land precisely in
// the page's bottom-left corner. AlignmentType.RIGHT is not a typo: per the
// convention established above, that's what renders on the physical left.
function mkSignatureBlock() {
  return [
    new Paragraph({
      children: [mkRun('ניר בן דוד', { size: 9, bold: true })],
      alignment: AlignmentType.RIGHT,
      spacing: { before: 240, after: 0 },
      bidirectional: true,
    }),
    new Paragraph({
      children: [mkRun('מהנדס מבנים B.sc', { size: 9 })],
      alignment: AlignmentType.RIGHT,
      spacing: { after: 0 },
      bidirectional: true,
    }),
    new Paragraph({
      children: [mkRun('מ.ר 28566561', { size: 9 })],
      alignment: AlignmentType.RIGHT,
      spacing: { after: 0 },
      bidirectional: true,
    }),
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

  // Use client name as fallback when location is empty. inspection_date
  // feeds directly into the intro sentence template below — formatDate() is
  // a no-op on an already-formatted string, so this is safe whether the
  // caller supplied strict ISO or DD.MM.YYYY.
  const effectiveData = {
    ...data,
    location: (data.location || '').trim() || (data.client || '').trim(),
    inspection_date: formatDate(data.inspection_date),
  };

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

  // Pulled past the page margin (1418 twip / 2.5cm) so it sits 3mm (170 twip)
  // from the physical right edge, per the real client doc's own page-number
  // placement, instead of sitting flush with the body's right margin.
  const pageNumPara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,
    spacing: { after: 0 },
    indent: { left: -1248, right: -1248 },
    children: [
      new TextRun({ text: 'עמוד ', font: FONT, size: 7 * 2 }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 7 * 2 }),
      new TextRun({ text: ' מתוך ', font: FONT, size: 7 * 2 }),
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

  // ── 4. Date paragraph — left side of page
  const dateStr  = formatDate(data.date);
  const datePara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    spacing: { after: 0 },
    children: [mkRun(dateStr, { size: 8 })],
  });

  // ── 5. Client block ───────────────────────────────────────────────────────
  const toLabel    = mkPara([mkRun('לכבוד',                 { size: 8 })],                        { spacing: { after: 0 } });
  const clientPara = mkPara([mkRun(data.client ?? '',        { size: 8, underline: true })],       { spacing: { after: 0 } });
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

  // A few types (אישור ביצוע, רשת הגנה) put the project/site name on its own
  // centered line right under the subject, instead of appended to it with "–".
  const subjectLocationPara = cfg.subjectLocationLine
    ? mkPara([mkRun(effectiveData.location, { size: cfg.titleSize, bold: true })], { alignment: AlignmentType.CENTER, spacing: { after: 0 } })
    : null;

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
    ...(subjectLocationPara ? [subjectLocationPara] : []),
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
      bodyChildren.push(mkNumberedPara(note, { size: cfg.bodySize, spacing: SP_BODY }));
    });

    // Validity line — computed expiry date, per type (group1/group3/group5)
    const validityLine = mkValidityLine(cfg, data, SP_BODY);
    if (validityLine) bodyChildren.push(validityLine);

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

    const surveyValidityLine = mkValidityLine(cfg, data, SP_BODY);
    if (surveyValidityLine) bodyChildren.push(surveyValidityLine);

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
    // Some types (רשת הגנה) have no separate findings heading at all — the
    // technical description just flows on as plain paragraphs after the intro.
    if (cfg.findingsSectionHeading) {
      bodyChildren.push(
        mkPara(
          [mkRun(cfg.findingsSectionHeading, { size: cfg.headingSize, bold: true })],
          { spacing: SP_SECTION }
        )
      );
    }

    // Findings as bullet paragraphs
    const findingsSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
      ? data.notes_custom
      : cfg.defaultFindings;

    findingsSource.forEach((item) => {
      const text = String(item).replace(/^[•\-]\s*/, '');
      const prefix = cfg.plainFindings ? '' : '• ';
      bodyChildren.push(
        mkPara([mkRun(`${prefix}${text}`, { size: cfg.bodySize })], { spacing: SP_BODY })
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
    const conclusionLines = conclusionsRaw ? conclusionsRaw.split('\n') : (cfg.defaultConclusions || []);
    conclusionLines.forEach((line) => {
      bodyChildren.push(mkNumberedPara(line, { size: cfg.bodySize, spacing: SP_BODY }));
    });

    bodyChildren.push(...mkSignatureBlock(cfg.bodySize));

  } else if (cfg.layout === 'checklist') {
    // ── 'checklist' layout: group11 (אישור בטיחות לקייטנה) ────────────────
    // Fixed 15-item checklist + optional ליקויים/ריכוז בדיקות tables, closing
    // with the underscore fill-in line instead of the usual signature block.
    const introParts = typeof cfg.introTemplate === 'function' ? cfg.introTemplate(effectiveData) : [];
    const effectiveIntroParts = data.intro_extra ? [...introParts, data.intro_extra] : introParts;
    effectiveIntroParts.forEach((line, idx) => {
      bodyChildren.push(mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: idx === 0 ? { ...SP_BODY, before: 360 } : SP_BODY }));
    });

    bodyChildren.push(mkPara([mkRun(cfg.checklistHeading, { size: cfg.bodySize })], { spacing: SP_SECTION }));
    const checklistAnswers = Array.isArray(data.checklist_items) ? data.checklist_items : [];
    const checklistRows = CAMP_CHECKLIST_ITEMS.map((item, i) => [
      String(i + 1), item, checklistAnswers[i]?.status || '', checklistAnswers[i]?.note || '',
    ]);
    bodyChildren.push(mkTable(cfg.checklistColumns, checklistRows, cfg.checklistColWidths, { hdrSize: cfg.tableHdrSize, dataSize: cfg.tableDataSize, noMerge: true }));

    bodyChildren.push(mkPara([mkRun(cfg.defectsHeading, { size: cfg.bodySize })], { spacing: SP_SECTION }));
    const campDefectsRows = Array.isArray(data.defects_rows) && data.defects_rows.length > 0
      ? data.defects_rows
      : [['--', '--', '--', '--', '--', '--']];
    bodyChildren.push(mkTable(cfg.defectsColumns, campDefectsRows, cfg.defectsColWidths, { hdrSize: cfg.tableHdrSize, dataSize: cfg.tableDataSize, noMerge: true }));

    bodyChildren.push(mkPara([mkRun(cfg.approvalsHeading, { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION }));
    bodyChildren.push(mkApprovalsTable(data.approvals_status || {}));

    bodyChildren.push(...mkCampSignatureBlock(effectiveData.inspection_date || formatDate(data.date)));

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

    // Approvals checklist table — fixed real content, see SAFETY_APPROVALS_ROWS
    if (cfg.approvalsTableHeaders) {
      bodyChildren.push(
        mkPara([mkRun('ריכוז בדיקות בטיחות:', { size: cfg.headingSize, bold: true })], { spacing: SP_SECTION })
      );
      bodyChildren.push(mkApprovalsTable(data.approvals_status || {}));
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
        // Cap at 7.2cm so the image fits inside the 7.7cm-wide photo cell with margin
        if (isPortrait) {
          dispH = 7.2; // cm absolute height
          dispW = natW > 0 ? (natW / natH) * dispH : 5.4;
        } else {
          dispW = 7.2; // cm absolute width
          dispH = natH > 0 ? (natH / natW) * dispW : 4.5;
        }

        // Short by design — which finding a photo illustrates is already
        // cross-referenced from the table ("ראה תמונה N"), so the appendix
        // itself just needs the number, not a repeated description.
        const caption = `תמונה ${i + 1}`;

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
          [mkRun(p.caption, { size: 8 })],
          { alignment: AlignmentType.CENTER }
        );

        return new TableCell({
          width: { size: cm(7.7), type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          borders: NO_BORDER,
          children: [capPara, imgPara],
        });
      };

      const emptyCell = () =>
        new TableCell({
          width: { size: cm(7.7), type: WidthType.DXA },
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

    const photoColWidths = [cm(7.7), cm(7.7)];
    const photoTable = new Table({
      visuallyRightToLeft: true,
      layout: TableLayoutType.FIXED,
      width: { size: photoColWidths[0] + photoColWidths[1], type: WidthType.DXA },
      columnWidths: photoColWidths,
      rows: photoTableRows,
    });

    bodyChildren.push(photosTitlePara, photoTable);
  }

  // Fixed standard-reference appendix (currently only group3 / תקרות תותב) — appears last, after photos.
  if (cfg.appendixClauses) {
    bodyChildren.push(mkPara([mkRun(cfg.appendixTitle, { size: cfg.headingSize, bold: true })], { pageBreak: true }));
    bodyChildren.push(mkPara([mkRun(cfg.appendixIntro, { size: cfg.bodySize })], { spacing: SP_BODY }));
    cfg.appendixClauses.forEach((clause) => {
      bodyChildren.push(mkPara([mkRun(clause.label, { size: cfg.bodySize, bold: true })], { spacing: SP_SECTION }));
      clause.lines.forEach((line) => {
        bodyChildren.push(mkPara([mkRun(line, { size: cfg.bodySize })], { spacing: SP_BODY }));
      });
    });
  }

  // ── Build Document ────────────────────────────────────────────────────────
  const doc = new Document({
    numbering: {
      config: [{
        reference: NUMBERING_REFERENCE,
        levels: [{
          level: 0,
          format: LevelFormat.DECIMAL,
          text: '%1.',
          // The real template's own numPr uses lvlJc=left/ind-left here,
          // but that rendered with the digit on the physical LEFT of the
          // sentence — confirmed wrong against the actual generated report.
          // Mirrored to the opposite side, matching the physical-right
          // digit placement that was asked for.
          alignment: AlignmentType.RIGHT,
          style: { paragraph: { indent: { right: 340, hanging: 340 } } },
        }],
      }],
    },
    styles: {
      default: {
        document: {
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
            alignment: AlignmentType.LEFT,
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
            // Exact twip (DXA) values read straight from <w:pgMar> of a genuine
            // field report from Drive (13.8.2026, מרכז דניאל לחתירה) — symmetric
            // left/right (2.5cm each), unlike the earlier asymmetric values which
            // came from a self-generated sample and drifted from the real standard.
            // footer: 4mm (227 twip) from the physical bottom edge, per the
            // field engineer's own measurement against the real Word doc.
            margin: {
              top:    2552,
              bottom: 567,
              left:   1418,
              right:  1418,
              header: 142,
              footer: 227,
            },
          },
          bidi: true,  // RTL section — all paragraphs inherit right-to-left direction
        },
        headers: { default: docHeader },
        footers: { default: docFooter },
        children: bodyChildren,
      },
    ],
  });

  return await Packer.toBlob(doc);
}
