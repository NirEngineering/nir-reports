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
  SimpleField,
  UnderlineType,
  convertMillimetersToTwip as mm,
  TableLayoutType,
  VerticalMergeType,
  LevelFormat,
} from 'docx';

const FN_NUMBERING = [
  {
    reference: 'fn-bullet',
    levels: [{
      level: 0,
      format: LevelFormat.BULLET,
      text: '•',
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { right: mm(12), hanging: mm(6) } } },
    }],
  },
  {
    reference: 'fn-numbered',
    levels: [{
      level: 0,
      format: LevelFormat.DECIMAL,
      text: '%1.',
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { right: mm(12), hanging: mm(6) } } },
    }],
  },
];

// ── Constants (match V1 doc_generator.py exactly) ────────────────────────────
const FONT        = 'Arial';
const FS_DATE     = 8;
const FS_CLIENT   = 8;
const FS_TITLE    = 15;
const FS_INTRO    = 9;
const FS_HEADING  = 9;
const FS_BODY     = 8;
const FS_TABLE    = 9;
const FS_PHOTO    = 9;
const FS_PAGE_NUM = 7;

const LS_BODY  = 1.5;
const LS_INTRO = 1.15;
const lsTwips = (m) => Math.round(240 * m);
const pt = (v) => v * 20;

// ── Notes ────────────────────────────────────────────────────────────────────
const NOTES_HANGING = [
  'על כל שינוי קונסטרוקטיבי ועיוותים באופן חיבור/תליות האלמנטים (סדקים, עיוותים, שקיעות, ניתוקים, חלודה, אלמנטים רופפים, חוסרים/תוספות וכדומה) יש לדווח ולזמן לבדיקה חוזרת.',
  'אין להעמיס עומסים על האלמנטים שנבדקו שאינם מיועדים לכך.',
  'הבדיקה הינה ויזואלית ונכונה ליום הבדיקה.',
];

const NOTES_CEILING = [
  'על כל שינוי קונסטרוקטיבי בתקרות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה, כפף, שבר, סדק, תזוזות וכו\' - יש לזמן לבדיקה חוזרת.',
  'אין לטפס, להיתלות ו/או להעמיס עומסים על התקרות שנבדקו.',
  'תוקף הדו"ח הינו לחמש שנים מיום הבדיקה בכפוף למסקנות הבדיקה.',
];

const NOTES_SURVEY = [
  'על כל שינוי קונסטרוקטיבי, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה, כפף, שבר, סדק, תזוזות וכו\' - יש לזמן לבדיקה חוזרת.',
  'ממצאי סקר זה הם כפי שהועברו לח"מ ע"י בעלי התפקידים באתר ומציגים מצב קיים ביום הסיור בלבד.',
  'מצורפת טבלת ליקויים בהמשך המסמך הכוללת את מיקום הליקוי, פירוט הממצא, דרישות/הנחיות לטיפול ותמונות.',
  'מקרא קדימות לטיפול בליקויים:',
  'קדימות 1 – "ליקוי חמור" - ליקוי/מפגע המחייב הסרתו/תיקונו וטיפולו המיידי.',
  'קדימות 2 – "ליקוי בינוני" - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 3 חודשים מתאריך דו"ח זה.',
  'קדימות 3 – "ליקוי קל" - ליקוי/מפגע המחייב טיפול בתכנית עבודה עד 6 חודשים מתאריך דו"ח זה.',
];

const NOTES_BUILDINGS = [
  'על כל שינוי קונסטרוקטיבי במבנים ובמתקנים, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור/תליות האלמנטים שנבדקו, קורוזיה, כפף, שבר, סדק, תזוזות וכו\' – יש לזמן לבדיקה חוזרת וטיפול מתאים עפ"י הממצאים.',
  'אין לטפס, להיתלות ו/או להעמיס עומסים על האלמנטים והמתקנים שנבדקו.',
  'תוקף הדו"ח הינו לשנה מיום הבדיקה בכפוף למסקנות הבדיקה.',
  'מקרא קדימות לטיפול בליקויים:',
  'קדימות 1 – ליקוי/מפגע המחייב הסרתו/תיקונו המיידי.',
  'קדימות 2 – ליקוי/מפגע המחייב טיפול בתכנית עבודה ועד 3 חודשים מיום דו"ח זה.',
];

const NOTES_SHADES = [
  'על כל שינוי קונסטרוקטיבי בסככות ורכיביהן, הוספה ו/או הפחתת רכיבים, עיוותים באופן חיבור, קורוזיה, כפף, שבר, סדק, תזוזות וכו\' - יש לזמן לבדיקה חוזרת.',
  'אין לטפס, להיתלות, לפרוש שלטים/מפרשי רוח ו/או להעמיס עומסים על הסככות שנבדקו.',
  'תוקף הדו"ח הינו לשנה מיום הבדיקה בכפוף למסקנות הבדיקה.',
];

// ── Doc-type configuration ────────────────────────────────────────────────────
const DOC_TYPES = {
  group1: {
    logoFile: 'letterhead_main.jpg',
    tableColumns: ['מיקום', 'האלמנט/המתקן', 'נתונים וממצאים', 'הערות'],
    colWidths: [3.5, 4.0, 7.0, 2.5],
    defectsColumns: null,
    defectsColWidths: null,
    sectionTitle: 'נתונים וממצאים:',
    notesTitle: 'הערות והנחיות:',
    defaultNotes: NOTES_HANGING,
    hasValidityNote: true,
    hasDefectsTable: false,
    conclusionOk:      'האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'בחלק מהאלמנטים שנבדקו נמצאו ליקויים – ראה פירוט בטבלה.',
    introTemplate: (d) => {
      const loc = d.address ? `${d.location}, ${d.address}` : d.location;
      return `בתאריך ${d.inspection_date} ערכתי סיור בדיקה ב${loc} ובדקתי את יציבות האלמנטים והמתקנים התלויים במתחם.`;
    },
  },
  group3: {
    logoFile: 'letterhead_main.jpg',
    tableColumns: ['מיקום/חדר', 'סוג התקרה', 'תקין/לא תקין', 'קדימות ליקויים', 'הערות'],
    colWidths: [3.5, 4.5, 2.5, 2.5, 4.0],
    defectsColumns: ["מס'", 'מיקום', 'ממצאים וליקויים', 'טיפול נדרש', 'תמונות', 'קדימות'],
    defectsColWidths: [1.0, 2.5, 4.5, 4.5, 2.0, 2.0],
    sectionTitle: 'הקדמה:',
    notesTitle: 'הנחיות:',
    defaultNotes: NOTES_CEILING,
    hasValidityNote: false,
    hasDefectsTable: true,
    appendixTitles: [
      'נספח דרישות לתכן והתקנה של תקרות תותב פריקות לפי ת"י 5103',
      'נספח דרישות לתכן והתקנה של תקרות תותב לא פריקות לפי ת"י 1924',
    ],
    conclusionOk:      'התקרות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה ומיום הבדיקה.',
    conclusionDefects: 'נמצאו ליקויים בתקרות – ראה פירוט בנספח טבלת הליקויים בהמשך.',
    introTemplate: (d) =>
      `סקר זה מתייחס למצב קיים של התשתיות במבנה (תשתית תקרות התותב שנבדקו) והצגת הפערים הקיימים בין המצב בשטח לבין הדרישות הרלוונטיות.\nהסקר בוצע לבקשת ${d.client} ב${d.location} בתאריך ${d.inspection_date}.`,
  },
  group4: {
    logoFile: 'letterhead_scouts.png',
    tableColumns: ['האלמנט/המבנה הנבדק', 'נתונים ופירוט', 'הערות', 'תקין/לא תקין'],
    colWidths: [4.0, 6.0, 4.5, 2.5],
    defectsColumns: ["מס'", 'מיקום', 'ממצאי הסיור', 'טיפול נדרש', 'תמונות', 'קדימות'],
    defectsColWidths: [1.0, 2.5, 4.5, 4.5, 2.0, 2.0],
    sectionTitle: 'נתונים וממצאים:',
    notesTitle: 'הערות והנחיות:',
    defaultNotes: NOTES_SURVEY,
    hasValidityNote: false,
    hasDefectsTable: true,
    conclusionOk:      'שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'יש לטפל בליקויים הקיימים עפ"י ההערות בטבלה.',
    introTemplate: (d) => {
      const loc = d.address ? `${d.location}, ${d.address}` : d.location;
      return `בתאריך ${d.inspection_date} ביקרתי ב${loc} ובדקתי את יציבות ותקינות מספר אלמנטים ומתקנים במתחם.`;
    },
  },
  group6: {
    logoFile: 'letterhead_main.jpg',
    tableColumns: ['האלמנט/המבנה הנבדק', 'מיקום', 'סוג / האלמנט', 'נתונים וממצאים', 'תקין/לא תקין', 'קדימות'],
    colWidths: [1247, 891, 992, 2300, 2694, 835, 692],
    defectsColumns: null,
    defectsColWidths: null,
    sectionTitle: 'נתונים וממצאים:',
    notesTitle: 'הערות והנחיות:',
    defaultNotes: NOTES_BUILDINGS,
    hasValidityNote: false,
    hasDefectsTable: false,
    conclusionOk:      'שאר האלמנטים שנבדקו נמצאו יציבים ובטוחים לשימוש נכון ליום הבדיקה.',
    conclusionDefects: 'יש לטפל בליקויים שנמצאו וצויינו במסמך.',
    conclusionExtra:   'נדרשת בדיקה חוזרת לאחר תיקון הליקויים לאישור / נדרש עדכון ביצוע לאישור.',
    pageMargin: { top: 2410, right: 1800, bottom: 568, left: 1260, header: 142, footer: 0 },
    introTemplate: (d) => {
      const loc = d.address ? `${d.location}, ${d.address}` : d.location;
      return `בתאריך ${d.inspection_date} ביקרתי ב${loc} ובדקתי את יציבות ותקינות מספר אלמנטים ומתקנים שונים במתחם.`;
    },
  },
  group5: {
    logoFile: 'letterhead_main.jpg',
    tableColumns: ["מס'", 'מיקום', 'סוג הסככה', "מידות (מ') ונתונים", 'תקין/לא תקין', 'קדימות ליקויים', 'תמונה'],
    colWidths: [1.0, 2.5, 3.0, 4.5, 2.5, 2.0, 1.5],
    defectsColumns: ['מספר סככה', 'מיקום', 'ממצאי ליקויים ודרישות', 'תמונות הליקוי', 'קדימות'],
    defectsColWidths: [2.0, 2.5, 6.5, 4.0, 2.0],
    sectionTitle: 'הקדמה:',
    notesTitle: 'הנחיות:',
    defaultNotes: NOTES_SHADES,
    hasValidityNote: false,
    hasDefectsTable: true,
    conclusionOk:      'הסככות שנבדקו נמצאו יציבות ובטוחות לשימוש נכון ליום הבדיקה ומיום הבדיקה.',
    conclusionDefects: 'במספר סככות נמצאו ליקויים – ראה פירוט הליקויים בנספח טבלת הליקויים המצורפת.',
    introTemplate: (d) =>
      `סקר זה מתייחס למצב קיים של תשתיות הסככות והצגת הפערים הקיימים בין המצב בשטח לבין הדרישות הרלוונטיות.\nהסקר בוצע לבקשת ${d.client} ב${d.location} בתאריך ${d.inspection_date}.`,
  },
};

// ── Helpers ──────────────────────────────────────────────────────────────────

const cm   = (v) => Math.round(v * 567);
const cmPx = (v) => Math.round(v * 37.795275591);

function formatDate(s) {
  if (!s) return '';
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${parseInt(m[3])}.${parseInt(m[2])}.${m[1]}` : s;
}

function addDays(s, days) {
  if (!s) return '';
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  d.setDate(d.getDate() + days);
  return `${d.getDate()}.${d.getMonth() + 1}.${d.getFullYear()}`;
}

async function fetchLogo(url, fallbackAspect = 4.0) {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const rawBuf = await res.arrayBuffer();
    const ext = url.split('.').pop().toLowerCase();
    const mimeType = ext === 'png' ? 'image/png' : 'image/jpeg';
    const type = ext === 'png' ? 'png' : 'jpg';

    const blob = new Blob([rawBuf], { type: mimeType });
    const blobUrl = URL.createObjectURL(blob);

    const { buf, aspectRatio } = await new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const w = img.naturalWidth || 800;
        const h = img.naturalHeight || Math.round(800 / fallbackAspect);
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        canvas.toBlob(
          (outBlob) => {
            outBlob.arrayBuffer().then((outBuf) =>
              resolve({ buf: outBuf, aspectRatio: w / h })
            );
          },
          mimeType,
          0.95,
        );
      };
      img.onerror = () => resolve({ buf: rawBuf, aspectRatio: fallbackAspect });
      img.src = blobUrl;
    });

    URL.revokeObjectURL(blobUrl);
    return { buf, type, aspectRatio };
  } catch (_) { return null; }
}

function b64ToArr(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

function imgDims(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 4, h: 3 });
    img.src = dataUrl;
  });
}

// ── Core paragraph builder ────────────────────────────────────────────────────

function mkRun(text, opts = {}) {
  const { size = FS_BODY, bold = false, color, underline = false, rtl = true } = opts;
  // In RTL Word paragraphs, trailing punctuation can float to the wrong side.
  // Appending U+200F (Right-to-Left Mark) after punctuation anchors it correctly.
  const t = (rtl && text) ? text.replace(/([:\?!,\.;])\s*$/, '$1\u200F') : text;
  return new TextRun({
    text: t,
    font: FONT,
    size: size * 2,
    bold,
    color,
    rtl,
    ...(rtl ? { language: { bidi: 'he-IL' } } : {}),
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

/** Build an RTL paragraph. */
function mkPara(runs, { alignment = AlignmentType.LEFT, spacing } = {}) {
  return new Paragraph({ children: runs, alignment, bidirectional: true, spacing });
}

// Border specs
const NO_BORDER = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};
const THIN_BORDER = {
  top:    { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  left:   { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  right:  { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
};

// ── Table builder ─────────────────────────────────────────────────────────────
function mkTable(headers, rows, colWidthsCm, mergeCol = -1, headerBg = 'EAF1DD') {
  const colW = colWidthsCm.map((w) => cm(w));
  const CELL_SPACING = { before: pt(2), after: pt(2) };

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      new TableCell({
        width: { size: colW[i], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.CLEAR, fill: headerBg },
        borders: THIN_BORDER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          bidirectional: true,

          spacing: CELL_SPACING,
          children: [mkRun(h, { size: FS_TABLE, bold: true })],
        })],
      })
    ),
  });

  let mergeMarks = rows.map(() => 'restart');
  if (mergeCol >= 0 && rows.length >= 2) {
    const effective = [];
    let prev = '';
    for (const row of rows) {
      const val = String(row[mergeCol] ?? '').trim();
      effective.push(val || prev);
      if (val) prev = val;
    }
    mergeMarks = effective.map((v, i) =>
      i > 0 && v === effective[i - 1] ? 'continue' : 'restart'
    );
  }

  const RIGHT_COL_KEYWORDS = ['נתונים וממצאים', 'נתונים ופירוט'];
  const isRightCol = (h) => RIGHT_COL_KEYWORDS.some((kw) => h?.includes(kw));

  const dataRows = rows.map((row, rIdx) => {
    const bg = 'FFFFFF';

    const cells = row.map((cellText, cIdx) => {
      const text       = String(cellText ?? '');
      const isNotOk   = text === 'לא תקין' || text.startsWith('לא תקין');
      const isWatch   = text === 'תקין - דורש מעקב' || text === 'דורש בדיקה חוזרת';
      const isPrio1   = text === '1' && cIdx < headers.length && headers[cIdx].includes('קדימות');
      let color, bold = false;
      if (isNotOk || isPrio1) { color = 'C00000'; bold = true; }
      if (isWatch)            { color = 'C55A11'; bold = true; }

      const cellAlign = isRightCol(headers[cIdx]) ? AlignmentType.LEFT : AlignmentType.CENTER;

      if (cIdx === mergeCol) {
        const isCont = mergeMarks[rIdx] === 'continue';
        return new TableCell({
          width: { size: colW[cIdx] ?? colW[colW.length - 1], type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.CLEAR, fill: bg },
          borders: THIN_BORDER,
          verticalMerge: isCont ? VerticalMergeType.CONTINUE : VerticalMergeType.RESTART,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: true,
            spacing: CELL_SPACING,
            children: isCont ? [] : [mkRun(text, { size: FS_TABLE, bold, color })],
          })],
        });
      }

      return new TableCell({
        width: { size: colW[cIdx] ?? colW[colW.length - 1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.CLEAR, fill: bg },
        borders: THIN_BORDER,
        children: [new Paragraph({
          alignment: cellAlign,
          bidirectional: true,
          spacing: CELL_SPACING,
          children: [mkRun(text, { size: FS_TABLE, bold, color })],
        })],
      });
    });

    return new TableRow({ children: cells });
  });

  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    rows: [headerRow, ...dataRows],
  });
}

// ── Survey table (group6) ─────────────────────────────────────────────────────
function mkSurveyTable(buildingData, buildingStatus, buildingPriority, rows) {
  const COL = [1247, 891, 992, 2300, 2694, 835, 692];
  const HEADER_BG = 'C5E0B3';
  const SUBHDR_BG = 'DBDBDB';
  const DATA_BG   = 'F7CAAC';

  const CELL_SPACING = { before: pt(2), after: pt(2) };

  const mkCell = (text, {
    width, span = 1, bg = DATA_BG, bold = false, color,
    align = AlignmentType.CENTER, vMerge = null, fontSize = FS_TABLE,
  } = {}) => new TableCell({
    width: { size: width, type: WidthType.DXA },
    ...(span > 1 ? { columnSpan: span } : {}),
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: bg },
    borders: THIN_BORDER,
    ...(vMerge === 'continue' ? { verticalMerge: VerticalMergeType.CONTINUE } : {}),
    ...(vMerge === 'restart' ? { verticalMerge: VerticalMergeType.RESTART } : {}),
    children: [new Paragraph({
      alignment: align,
      bidirectional: true,
      spacing: CELL_SPACING,
      children: vMerge === 'continue' ? [] : [mkRun(text, { size: fontSize, bold, color })],
    })],
  });

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      mkCell('האלמנט/המבנה הנבדק', { width: COL[0],                    bg: HEADER_BG, bold: true }),
      mkCell('נתונים ופירוט',       { width: COL[1]+COL[2]+COL[3], span: 3, bg: HEADER_BG, bold: true }),
      mkCell('הערות',               { width: COL[4],                    bg: HEADER_BG, bold: true }),
      mkCell('תקין/לא תקין',        { width: COL[5],                    bg: HEADER_BG, bold: true }),
      mkCell('קדימות ליקוי',        { width: COL[6],                    bg: HEADER_BG, bold: true }),
    ],
  });

  const bldText   = String(buildingData || '');
  const bldStatus = String(buildingStatus || 'לא תקין');
  const bldPrio   = String(buildingPriority || '2');
  const bldNotOk  = bldStatus.startsWith('לא תקין');
  const buildingRow = new TableRow({
    children: [
      mkCell('יציבות מבנה', { width: COL[0],                    bg: DATA_BG, bold: true }),
      mkCell(bldText,        { width: COL[1]+COL[2]+COL[3], span: 3, bg: DATA_BG, align: AlignmentType.LEFT }),
      mkCell('',             { width: COL[4],                    bg: DATA_BG }),
      mkCell(bldStatus,      { width: COL[5], bg: DATA_BG, color: bldNotOk ? 'C00000' : undefined, bold: bldNotOk }),
      mkCell(bldPrio,        { width: COL[6],                    bg: DATA_BG }),
    ],
  });

  const subHdrRow = new TableRow({
    children: [
      mkCell('האלמנט/המבנה הנבדק', { width: COL[0],            bg: SUBHDR_BG, bold: true }),
      mkCell('מיקום',               { width: COL[1],            bg: SUBHDR_BG, bold: true }),
      mkCell('סוג / האלמנט',        { width: COL[2],            bg: SUBHDR_BG, bold: true }),
      mkCell('נתונים וממצאים',      { width: COL[3]+COL[4], span: 2, bg: SUBHDR_BG, bold: true }),
      mkCell('תקין/לא תקין',        { width: COL[5]+COL[6], span: 2, bg: SUBHDR_BG, bold: true }),
    ],
  });

  let prevElement = null;
  const dataRows = (Array.isArray(rows) ? rows : []).map((row) => {
    const [element, location, type, dataText, status, priority] = row;
    const elemStr   = String(element  ?? '');
    const locStr    = String(location ?? '');
    const typeStr   = String(type     ?? '');
    const dataStr   = String(dataText ?? '');
    const statusStr = String(status   ?? '');
    const prioStr   = String(priority ?? '');

    const isNotOk = statusStr.startsWith('לא תקין');
    const isPrio1 = prioStr === '1';
    const statusColor = (isNotOk || isPrio1) ? 'C00000' : undefined;
    const statusBold  = isNotOk || isPrio1;

    const vMerge = (elemStr && elemStr === prevElement) ? 'continue' : 'restart';
    if (elemStr) prevElement = elemStr;

    return new TableRow({ children: [
      mkCell(elemStr,  { width: COL[0],            bg: DATA_BG, bold: true, vMerge }),
      mkCell(locStr,   { width: COL[1],            bg: DATA_BG }),
      mkCell(typeStr,  { width: COL[2],            bg: DATA_BG }),
      mkCell(dataStr,  { width: COL[3]+COL[4], span: 2, bg: DATA_BG, align: AlignmentType.LEFT }),
      mkCell(statusStr,{ width: COL[5],            bg: DATA_BG, color: statusColor, bold: statusBold }),
      mkCell(prioStr,  { width: COL[6],            bg: DATA_BG, color: isPrio1 ? 'C00000' : undefined, bold: isPrio1 }),
    ]});
  });

  return new Table({
    visuallyRightToLeft: true,
    layout: TableLayoutType.FIXED,
    rows: [headerRow, buildingRow, subHdrRow, ...dataRows],
  });
}

// ── Main export ──────────────────────────────────────────────────────────────
export async function generateDocument(data) {
  const cfg  = DOC_TYPES[data.doc_type] ?? DOC_TYPES.group4;
  const base = import.meta.env.BASE_URL;

  const [headerLogo, footerLogo] = await Promise.all([
    fetchLogo(base + cfg.logoFile, 5.0),
    fetchLogo(base + 'footer_logo.png', 5.0),
  ]);

  function mkPageField(instr) {
    const sf = new SimpleField(instr);
    sf.root.push(new TextRun({ text: '1', font: FONT, size: FS_PAGE_NUM * 2, rtl: true }));
    return sf;
  }

  // ── Header: page number "עמוד X מתוך Y" — LTR+RIGHT avoids bidi digit-grouping ──
  const headerParas = [
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: false,
      spacing: { before: 0, after: pt(2) },
      children: [
        new TextRun({ text: 'עמוד ', font: FONT, size: FS_PAGE_NUM * 2 }),
        mkPageField('PAGE'),
        new TextRun({ text: ' מתוך ', font: FONT, size: FS_PAGE_NUM * 2 }),
        mkPageField('NUMPAGES'),
      ],
    }),
  ];

  if (headerLogo) {
    const logoW = 9.5;
    const logoH = logoW / headerLogo.aspectRatio;
    headerParas.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,

      spacing: { before: 0, after: 0 },
      children: [new ImageRun({
        data: headerLogo.buf,
        transformation: { width: cmPx(logoW), height: cmPx(logoH) },
        type: headerLogo.type,
      })],
    }));
  }

  // ── Footer: logo centered (14cm) ──────────────────────────────────────────
  const footerParas = [];
  if (footerLogo) {
    const fW = 14;
    const fH = fW / footerLogo.aspectRatio;
    footerParas.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,

      spacing: { before: 0, after: 0 },
      children: [new ImageRun({
        data: footerLogo.buf,
        transformation: { width: cmPx(fW), height: cmPx(fH) },
        type: footerLogo.type,
      })],
    }));
  } else {
    footerParas.push(new Paragraph({ children: [] }));
  }

  // ── Date — LEFT aligned, LTR, explicitly non-bidi so Word shows it on the left ──
  const datePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: false,
    spacing: { before: 0, after: 0 },
    children: [mkRun(formatDate(data.date), { size: FS_DATE, rtl: false })],
  });

  // ── Client block ──────────────────────────────────────────────────────────
  const clientBlock = [
    mkPara([mkRun('לכבוד', { size: FS_CLIENT })], { spacing: { before: 0, after: 0 } }),
    ...(data.client ? [mkPara([mkRun(data.client, { size: FS_CLIENT })], { spacing: { before: 0, after: 0 } })] : []),
    mkPara([mkRun(data.organization ?? '', { size: FS_CLIENT, underline: true })], {
      spacing: { before: 0, after: pt(6) },
    }),
  ];

  // ── Subject ───────────────────────────────────────────────────────────────
  const subjectPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    bidirectional: true,

    spacing: { before: 0, after: pt(14) },
    children: [mkRun(`הנדון: ${data.subject ?? ''}`, { size: FS_TITLE, bold: true, underline: true })],
  });

  // ── Intro ─────────────────────────────────────────────────────────────────
  const introRaw = data.intro_extra
    ? cfg.introTemplate(data) + '\n' + data.intro_extra
    : cfg.introTemplate(data);

  const introParas = introRaw.split('\n')
    .filter(l => l.trim())
    .map(l => new Paragraph({
      alignment: AlignmentType.LEFT,
      bidirectional: true,

      spacing: { before: 0, after: pt(3), line: lsTwips(LS_INTRO), lineRule: 'auto' },
      children: [mkRun(l.trim(), { size: FS_INTRO })],
    }));

  const gapAfterIntro = new Paragraph({
    spacing: { before: 0, after: pt(4) },
    children: [],
  });

  // ── Section title ─────────────────────────────────────────────────────────
  const sectionTitlePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,

    spacing: { before: 0, after: pt(4), line: lsTwips(LS_BODY), lineRule: 'auto' },
    children: [mkRun(cfg.sectionTitle, { size: FS_HEADING, bold: true })],
  });

  // ── Main table ────────────────────────────────────────────────────────────
  const tableRows = Array.isArray(data.table_rows) ? data.table_rows : [];
  let mainTable = null;

  if (data.doc_type === 'group6') {
    mainTable = mkSurveyTable(
      data.building_stability_data   ?? '',
      data.building_stability_status ?? 'לא תקין',
      data.building_stability_priority ?? '2',
      tableRows,
    );
  } else if (tableRows.length > 0) {
    const mergeColIdx = data.doc_type === 'group5' ? -1 : 0;
    mainTable = mkTable(cfg.tableColumns, tableRows, cfg.colWidths, mergeColIdx);
  }

  // ── Notes ─────────────────────────────────────────────────────────────────
  const notesTitlePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,

    spacing: { before: pt(4), after: pt(2), line: lsTwips(LS_BODY), lineRule: 'auto' },
    children: [mkRun(cfg.notesTitle, { size: FS_HEADING, bold: true })],
  });

  let notesSource;
  if (Array.isArray(data.notes_custom) && data.notes_custom.length > 0) {
    notesSource = data.notes_custom;
  } else {
    notesSource = [...cfg.defaultNotes];
    if (cfg.hasValidityNote) {
      const expiry = addDays(data.inspection_date_raw || data.inspection_date, 365);
      notesSource.push(`תוקף האישור לשנה מיום הבדיקה ועד לתאריך ${expiry} בכפוף למסקנות.`);
    }
  }

  const notesParas = notesSource
    .filter(n => n && String(n).trim())
    .map((n, i) => new Paragraph({
      alignment: AlignmentType.LEFT,
      bidirectional: true,
      spacing: { before: 0, after: pt(3), line: lsTwips(LS_BODY), lineRule: 'auto' },
      children: [mkRun(`\u200F${i + 1}. ${String(n).trim()}`, { size: FS_BODY })],
    }));

  // ── Conclusions ───────────────────────────────────────────────────────────
  const conclusionsTitlePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,

    spacing: { before: pt(4), after: pt(2), line: lsTwips(LS_BODY), lineRule: 'auto' },
    children: [mkRun('מסקנות הבדיקה:', { size: FS_HEADING, bold: true })],
  });

  let conclusionLines;
  if (data.conclusion_custom && String(data.conclusion_custom).trim()) {
    conclusionLines = String(data.conclusion_custom).split('\n').filter(l => l.trim());
  } else if (data.has_defects) {
    conclusionLines = [cfg.conclusionDefects, cfg.conclusionExtra, cfg.conclusionOk].filter(Boolean);
  } else {
    conclusionLines = [cfg.conclusionOk];
  }

  const conclusionParas = conclusionLines.map((l, i) => new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,
    spacing: { before: 0, after: pt(3), line: lsTwips(LS_BODY), lineRule: 'auto' },
    children: [mkRun(`\u200F${i + 1}. ${l.trim()}`, { size: FS_BODY })],
  }));

  // ── Defects appendix ──────────────────────────────────────────────────────
  const defectsSection = [];
  if (
    data.has_defects && cfg.hasDefectsTable && cfg.defectsColumns &&
    Array.isArray(data.defects_rows) && data.defects_rows.length > 0
  ) {
    defectsSection.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: true,

        pageBreakBefore: true,
        spacing: { before: 0, after: pt(6) },
        children: [mkRun('נספח טבלת ליקויים:', { size: FS_HEADING, bold: true })],
      }),
      mkTable(cfg.defectsColumns, data.defects_rows, cfg.defectsColWidths, -1),
    );
  }

  // ── Photos appendix ───────────────────────────────────────────────────────
  const photosSection = [];
  const photos = Array.isArray(data.photos) ? data.photos : [];

  if (photos.length > 0) {
    const loaded = [];
    for (let i = 0; i < photos.length; i++) {
      const ph = photos[i];
      try {
        const dataUrl = ph.data;
        if (!dataUrl) continue;
        const b64    = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
        const imgArr = b64ToArr(b64);
        const { w: natW, h: natH } = await imgDims(dataUrl);
        const isPortrait = natH > natW;

        let dispW, dispH;
        if (isPortrait) {
          dispH = 8;   dispW = natW > 0 ? (natW / natH) * dispH : 6;
        } else {
          dispW = 7.5; dispH = natH > 0 ? (natH / natW) * dispW : 5.625;
        }

        const mimeM = dataUrl.match(/^data:image\/(jpeg|jpg|png|gif|webp);/);
        const imgType = mimeM?.[1] === 'png' ? 'png' : 'jpg';
        const caption = ph.caption ? `תמונה ${i + 1} - ${ph.caption}` : `תמונה ${i + 1}`;
        loaded.push({ imgArr, dispW, dispH, imgType, caption });
      } catch (_) { /* skip broken */ }
    }

    if (loaded.length > 0) {
      photosSection.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: true,

        pageBreakBefore: true,
        spacing: { before: 0, after: pt(6) },
        children: [mkRun('נספח תמונות:', { size: FS_HEADING, bold: true })],
      }));

      for (let i = 0; i < loaded.length; i += 2) {
        const left  = loaded[i];
        const right = loaded[i + 1] ?? null;

        const capCell = (p) => new TableCell({
          width: { size: cm(8.5), type: WidthType.DXA },
          borders: NO_BORDER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: true,

            spacing: { before: pt(4), after: pt(2) },
            children: [mkRun(p.caption, { size: FS_PHOTO })],
          })],
        });
        const imgCell = (p) => new TableCell({
          width: { size: cm(8.5), type: WidthType.DXA },
          verticalAlign: VerticalAlign.TOP,
          borders: NO_BORDER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: true,

            spacing: { before: 0, after: pt(4) },
            children: [new ImageRun({
              data: p.imgArr,
              transformation: { width: cmPx(p.dispW), height: cmPx(p.dispH) },
              type: p.imgType,
            })],
          })],
        });
        const emptyCell = () => new TableCell({
          width: { size: cm(8.5), type: WidthType.DXA },
          borders: NO_BORDER,
          children: [new Paragraph({ children: [] })],
        });

        const pairTable = new Table({
          layout: TableLayoutType.FIXED,
          rows: [
            new TableRow({ children: [capCell(left), right ? capCell(right) : emptyCell()] }),
            new TableRow({ children: [imgCell(left), right ? imgCell(right) : emptyCell()] }),
          ],
        });
        photosSection.push(pairTable);
        photosSection.push(new Paragraph({

          spacing: { before: 0, after: pt(6) },
          children: [],
        }));
      }
    }
  }

  // ── Group3 appendix titles ────────────────────────────────────────────────
  const appendixSection = [];
  if (data.doc_type === 'group3' && cfg.appendixTitles) {
    appendixSection.push(
      new Paragraph({ spacing: { before: pt(6), after: 0 }, children: [] }),
    );
    for (const title of cfg.appendixTitles) {
      if (title) {
        appendixSection.push(new Paragraph({
          alignment: AlignmentType.LEFT,
          bidirectional: true,

          spacing: { before: 0, after: pt(3) },
          children: [mkRun(title, { size: FS_BODY, bold: true })],
        }));
      }
    }
  }

  // ── Assemble document body ────────────────────────────────────────────────
  const bodyChildren = [
    datePara,
    ...clientBlock,
    subjectPara,
    ...introParas,
    gapAfterIntro,
    sectionTitlePara,
  ];

  if (mainTable) {
    bodyChildren.push(mainTable);
    bodyChildren.push(new Paragraph({
      spacing: { before: 0, after: pt(6) },
      children: [],
    }));
  }

  bodyChildren.push(
    notesTitlePara,
    ...notesParas,
    conclusionsTitlePara,
    ...conclusionParas,
    ...defectsSection,
    ...photosSection,
    ...appendixSection,
  );

  // ── Build Document ────────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: {
          run:       { font: { name: FONT }, size: FS_BODY * 2 },
          paragraph: { alignment: AlignmentType.LEFT },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: mm(210), height: mm(297) },
          margin: cfg.pageMargin ?? {
            top:    mm(42.5),
            bottom: mm(7.5),
            left:   mm(20),
            right:  mm(20),
            header: mm(3),
            footer: mm(0),
          },
        },
      },
      headers: { default: new Header({ children: headerParas }) },
      footers: { default: new Footer({ children: footerParas }) },
      children: bodyChildren,
    }],
  });

  return Packer.toBlob(doc);
}

// ── Field-notes document generator ──────────────────────────────────────────

function parseHtmlToDocxChildren(html, { textSize = FS_BODY, lineSpacing = null } = {}) {
  const container = document.createElement('div');
  container.innerHTML = html || '';
  const paras = [];
  const sp = lineSpacing ? { line: lsTwips(lineSpacing), lineRule: 'auto' } : {};
  const hebrewRe = /[\u0590-\u05FF]/;
  const isRtl = t => hebrewRe.test(t);

  // bidi alignment flip: editor dir=rtl means justifyRight → visual-left in OOXML
  function alignFromStyle(el) {
    const ta = el.style?.textAlign || '';
    if (ta === 'left')   return AlignmentType.RIGHT;
    if (ta === 'center') return AlignmentType.CENTER;
    if (ta === 'right')  return AlignmentType.LEFT;
    return AlignmentType.LEFT;
  }

  // Extract inline runs only — do NOT recurse into block children
  function runsFromLeaf(node) {
    const runs = [];
    node.childNodes.forEach(child => {
      if (child.nodeType === Node.TEXT_NODE) {
        const t = child.textContent;
        if (t) runs.push(mkRun(t, { size: textSize, rtl: isRtl(t) }));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const tag = child.tagName.toUpperCase();
        if (tag === 'BR') { runs.push(new TextRun({ break: 1 })); return; }
        const inner = child.textContent || '';
        if (!inner.trim()) return;
        const bold = tag === 'B' || tag === 'STRONG' || child.style?.fontWeight === 'bold';
        const underline = tag === 'U';
        runs.push(mkRun(inner, { size: textSize, bold, underline, rtl: isRtl(inner) }));
      }
    });
    if (runs.length === 0) runs.push(mkRun('', { size: textSize }));
    return runs;
  }

  const BLOCK_TAGS = new Set(['DIV','P','H1','H2','H3','H4','H5','H6','BLOCKQUOTE','OL','UL','LI']);

  function processNode(node) {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = (node.textContent || '').trim();
      if (t) paras.push(mkPara([mkRun(t, { size: textSize, rtl: isRtl(t) })],
        { spacing: { before: 0, after: 0, ...sp } }));
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = node.tagName.toUpperCase();

    if (tag === 'BR') {
      paras.push(mkPara([mkRun('', { size: textSize })], { spacing: { before: 0, after: 0, ...sp } }));
      return;
    }

    if (tag === 'OL' || tag === 'UL') {
      let counter = 0;
      node.querySelectorAll(':scope > li').forEach(li => {
        counter++;
        const prefix = tag === 'UL'
          ? mkRun('• ', { size: textSize })
          : mkRun(`\u200F${counter}. `, { size: textSize });
        paras.push(mkPara([prefix, ...runsFromLeaf(li)], {
          spacing: { before: 0, after: 0, ...sp },
        }));
      });
      return;
    }

    // Block container (div/p/etc.) — recurse if it has block children, else treat as leaf
    const hasBlockChild = [...node.childNodes].some(
      c => c.nodeType === Node.ELEMENT_NODE && BLOCK_TAGS.has(c.tagName?.toUpperCase())
    );
    if (hasBlockChild) {
      node.childNodes.forEach(processNode);
    } else {
      const align = alignFromStyle(node);
      paras.push(mkPara(runsFromLeaf(node), { alignment: align, spacing: { before: 0, after: 0, ...sp } }));
    }
  }

  container.childNodes.forEach(processNode);
  return paras.length ? paras : [mkPara([mkRun('', { size: textSize })])];
}

export async function generateFieldNotesDocument(data) {
  const base = import.meta.env.BASE_URL;
  const [headerLogo, footerLogo] = await Promise.all([
    fetchLogo(base + 'letterhead_main.jpg', 5.0),
    fetchLogo(base + 'footer_logo.png', 5.0),
  ]);

  function mkPageField(instr) {
    const sf = new SimpleField(instr);
    sf.root.push(new TextRun({ text: '1', font: FONT, size: FS_PAGE_NUM * 2, rtl: true }));
    return sf;
  }

  const headerParas = [
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: false,
      spacing: { before: 0, after: pt(2) },
      children: [
        new TextRun({ text: 'עמוד ', font: FONT, size: FS_PAGE_NUM * 2 }),
        mkPageField('PAGE'),
        new TextRun({ text: ' מתוך ', font: FONT, size: FS_PAGE_NUM * 2 }),
        mkPageField('NUMPAGES'),
      ],
    }),
  ];
  if (headerLogo) {
    const logoW = 9.5;
    const logoH = logoW / headerLogo.aspectRatio;
    headerParas.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 0, after: 0 },
      children: [new ImageRun({
        data: headerLogo.buf,
        transformation: { width: cmPx(logoW), height: cmPx(logoH) },
        type: headerLogo.type,
      })],
    }));
  }

  const footerParas = [];
  if (footerLogo) {
    const fW = 14;
    const fH = fW / footerLogo.aspectRatio;
    footerParas.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 0, after: 0 },
      children: [new ImageRun({
        data: footerLogo.buf,
        transformation: { width: cmPx(fW), height: cmPx(fH) },
        type: footerLogo.type,
      })],
    }));
  } else {
    footerParas.push(new Paragraph({ children: [] }));
  }

  const datePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: false,
    spacing: { before: 0, after: 0 },
    children: [mkRun(formatDate(data.date), { size: FS_DATE, rtl: false })],
  });

  const leKavodPara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,
    spacing: { before: 0, after: 0, line: lsTwips(1.15), lineRule: 'auto' },
    children: [mkRun('לכבוד', { size: 8 })],
  });

  const clientPara = new Paragraph({
    alignment: AlignmentType.LEFT,
    bidirectional: true,
    spacing: { before: 0, after: 0, line: lsTwips(1.15), lineRule: 'auto' },
    children: [mkRun(data.client || '', { size: 8, underline: true })],
  });

  const titlePara = new Paragraph({
    alignment: AlignmentType.CENTER,
    bidirectional: true,
    spacing: { before: pt(18), after: pt(10) },
    children: [mkRun(data.title || 'יומן שטח', { size: FS_TITLE, bold: true, underline: true })],
  });

  const metaParas = [];
  if (data.location && data.location !== data.client) {
    metaParas.push(mkPara([mkRun(data.location, { size: FS_CLIENT })], { spacing: { before: 0, after: pt(8) } }));
  }

  const contentParas = parseHtmlToDocxChildren(data.content, { textSize: 9, lineSpacing: 1.15 });

  // Photos
  const photosSection = [];
  const photos = Array.isArray(data.photos) ? data.photos : [];
  if (photos.length > 0) {
    const loaded = [];
    for (let i = 0; i < photos.length; i++) {
      const ph = photos[i];
      try {
        const dataUrl = ph.data;
        if (!dataUrl) continue;
        const b64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
        const imgArr = b64ToArr(b64);
        const { w: natW, h: natH } = await imgDims(dataUrl);
        const isPortrait = natH > natW;
        let dispW, dispH;
        if (isPortrait) { dispH = 8; dispW = natW > 0 ? (natW / natH) * dispH : 6; }
        else { dispW = 8.4; dispH = natH > 0 ? (natH / natW) * dispW : 6.3; }
        const mimeM = dataUrl.match(/^data:image\/(jpeg|jpg|png|gif|webp);/);
        const imgType = mimeM?.[1] === 'png' ? 'png' : 'jpg';
        const num = i + 1;
        const label = ph.caption ? `תמונה ${num} - ${ph.caption}` : `תמונה ${num}`;
        loaded.push({ imgArr, dispW, dispH, imgType, label });
      } catch (_) {}
    }
    if (loaded.length > 0) {
      photosSection.push(new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: true,
        pageBreakBefore: true,
        spacing: { before: 0, after: pt(6) },
        children: [mkRun('תמונות:', { size: FS_HEADING, bold: true })],
      }));
      // Each cell: image on top, caption below — keeps them inseparable across pages.
      const photoCell = (p) => new TableCell({
        width: { size: cm(8.5), type: WidthType.DXA },
        verticalAlign: VerticalAlign.TOP,
        borders: NO_BORDER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: true,
            spacing: { before: pt(4), after: pt(2) },
            children: [mkRun(p.label, { size: FS_PHOTO })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            bidirectional: true,
            spacing: { before: 0, after: pt(6) },
            children: [new ImageRun({
              data: p.imgArr,
              transformation: { width: cmPx(p.dispW), height: cmPx(p.dispH) },
              type: p.imgType,
            })],
          }),
        ],
      });
      const emptyCell = () => new TableCell({
        width: { size: cm(8.5), type: WidthType.DXA },
        borders: NO_BORDER,
        children: [new Paragraph({ children: [] })],
      });
      for (let i = 0; i < loaded.length; i += 2) {
        const left = loaded[i];
        const right = loaded[i + 1] ?? null;
        photosSection.push(new Table({
          layout: TableLayoutType.FIXED,
          rows: [new TableRow({ cantSplit: true, children: [photoCell(left), right ? photoCell(right) : emptyCell()] })],
        }));
      }
    }
  }

  const doc = new Document({
    numbering: { config: FN_NUMBERING },
    styles: {
      default: {
        document: {
          run:       { font: { name: FONT }, size: FS_BODY * 2 },
          paragraph: { alignment: AlignmentType.LEFT },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: mm(210), height: mm(297) },
          margin: { top: mm(42.5), bottom: mm(7.5), left: mm(20), right: mm(20), header: mm(3), footer: mm(0) },
        },
      },
      headers: { default: new Header({ children: headerParas }) },
      footers: { default: new Footer({ children: footerParas }) },
      children: [datePara, leKavodPara, clientPara, titlePara, ...metaParas, ...contentParas, ...photosSection],
    }],
  });

  return Packer.toBlob(doc);
}
