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

// ── Font / size constants ────────────────────────────────────────────────────
const FONT       = 'Arial';
const FS_DATE    = 8;
const FS_CLIENT  = 8;
const FS_TITLE   = 15;
const FS_INTRO   = 9;
const FS_HEADING = 9;
const FS_BODY    = 8;
const FS_TABLE   = 9;
const FS_PHOTO_NUM = 9;
const FS_PHOTO_CAP = 9;

// ── Doc-type configuration ───────────────────────────────────────────────────
const DOC_TYPES = {
  group1: {
    name: 'אלמנטים תלויים',
    tableColumns: ['מיקום', 'האלמנט/המתקן', 'נתונים וממצאים', 'הערות'],
    colWidths: [3.5, 4.5, 6.0, 3.0],
    sectionTitle: 'נתונים וממצאים:',
    notesTitle: 'הערות והנחיות:',
    defaultNotes: [
      'אין לתלות עומסים נוספים מעל האלמנטים הנבדקים ללא אישור מהנדס.',
      'כל פגם או שינוי שיתגלה יש לדווח מיידית למהנדס האחראי.',
      'יש לבצע בדיקות תחזוקה שוטפות אחת לשלושה חודשים.',
    ],
    conclusionOk: 'האלמנטים הנבדקים תקינים ובטוחים לשימוש.',
    conclusionDefects: 'נמצאו ליקויים הדורשים טיפול מיידי כמפורט בטבלה.',
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} בוצע סיור בדיקה ב${d.location}${d.address ? ', ' + d.address : ''}.\nהבדיקה בוצעה על ידי מהנדס מוסמך מטעם ניר הנדסה.\nלהלן ממצאי הבדיקה:`,
    validityMonths: 12,
    hasDefectsTable: false,
  },
  group3: {
    name: 'תקרות תותב',
    tableColumns: ['מיקום/חדר', 'סוג התקרה', 'תקין/לא תקין', 'קדימות ליקויים', 'הערות'],
    colWidths: [3.5, 3.5, 3.0, 3.0, 4.0],
    sectionTitle: 'הקדמה:',
    notesTitle: 'הנחיות:',
    defaultNotes: [
      'יש לתקן את הליקויים שצוינו בהתאם לסדר הקדימות.',
      'בדיקה חוזרת תבוצע לאחר ביצוע התיקונים.',
      'אין להוסיף עומסים על תקרות התותב ללא אישור מהנדס.',
    ],
    conclusionOk: 'תקרות התותב תקינות ובטוחות לשימוש.',
    conclusionDefects: 'נמצאו ליקויים בתקרות התותב הדורשים תיקון דחוף.',
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} בוצע סיור בדיקת תקרות תותב ב${d.location}.\nלהלן ממצאי הבדיקה:`,
    hasDefectsTable: true,
    defectsColumns: ["מס'", 'מיקום', 'ממצאים וליקויים', 'טיפול נדרש', 'תמונות', 'קדימות'],
    defectsColWidths: [1.5, 3.0, 5.0, 4.0, 2.0, 2.0],
  },
  group4: {
    name: 'סקר שנתי',
    tableColumns: ['האלמנט/המבנה הנבדק', 'נתונים ופירוט', 'הערות', 'תקין/לא תקין'],
    colWidths: [4.5, 6.0, 4.0, 2.5],
    sectionTitle: 'נתונים וממצאים:',
    notesTitle: 'הערות והנחיות:',
    defaultNotes: [
      'הסקר בוצע על ידי מהנדס מוסמך.',
      'יש לטפל בממצאים בהתאם לסדר הקדימות שנקבע.',
      'בדיקה חוזרת מומלצת בתום ביצוע התיקונים.',
    ],
    conclusionOk: 'המצב הכללי תקין.',
    conclusionDefects: 'נמצאו ליקויים הדורשים טיפול.',
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} בוצע סקר שנתי ב${d.location}.\nלהלן ממצאי הסקר:`,
    hasDefectsTable: true,
    defectsColumns: ["מס'", 'מיקום', 'ממצאי הסיור', 'טיפול נדרש', 'תמונות', 'קדימות ליקוי'],
    defectsColWidths: [1.5, 3.0, 5.0, 4.0, 2.0, 2.0],
  },
  group5: {
    name: 'סככות',
    tableColumns: ["מס'", 'מיקום', 'סוג הסככה', "מידות (מ') ונתונים", 'תקין/לא תקין', 'קדימות ליקויים', 'תמונה'],
    colWidths: [1.5, 2.5, 3.0, 4.0, 2.5, 2.5, 1.5],
    sectionTitle: 'הקדמה:',
    notesTitle: 'הנחיות:',
    defaultNotes: [
      'יש לוודא שאין עומסים על גג הסככה מעבר לתכן.',
      'בדיקה עונתית מומלצת לפני עונת הגשמים.',
    ],
    conclusionOk: 'הסככות תקינות ובטוחות לשימוש.',
    conclusionDefects: 'נמצאו ליקויים בסככות הדורשים טיפול.',
    introTemplate: (d) =>
      `בתאריך ${d.inspection_date} בוצע סיור בדיקת סככות ב${d.location}.\nלהלן ממצאי הבדיקה:`,
    hasDefectsTable: true,
    defectsColumns: ['מספר סככה', 'מיקום', 'ממצאי ליקויים ודרישות', 'תמונות הליקוי', 'קדימות'],
    defectsColWidths: [2.0, 3.0, 6.0, 4.0, 2.0],
  },
};

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Convert cm to twips (docx unit). 1 cm = 567 twips */
const cm = (v) => Math.round(v * 567);

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
    size   = FS_BODY,
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
    rtl,
    italics: italic,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

/**
 * Create a Paragraph with bidirectional RTL support.
 * opts: { alignment, spacing, indent, pageBreak }
 */
function mkPara(children = [], opts = {}) {
  const {
    alignment = AlignmentType.RIGHT,
    spacing   = undefined,
    pageBreak = false,
  } = opts;

  return new Paragraph({
    children,
    bidirectional: true,
    alignment,
    spacing,
    pageBreakBefore: pageBreak,
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
 * headers: string[]
 * rows: string[][] (each cell may be a string)
 * colWidthsCm: number[] (cm per column)
 * headerBg: hex color string without '#'
 */
function mkTable(headers, rows, colWidthsCm, headerBg = 'EAF1DD') {
  const colWidths = colWidthsCm.map((w) => cm(w));

  // Header row
  const headerCells = headers.map((h, i) =>
    new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, fill: headerBg },
      borders: THIN_BORDER,
      children: [
        mkPara([mkRun(h, { size: FS_TABLE, bold: true })], {
          alignment: AlignmentType.CENTER,
        }),
      ],
    })
  );

  const headerRow = new TableRow({
    children: headerCells,
    tableHeader: true,
  });

  // Data rows
  const dataRows = rows.map((row, rowIdx) => {
    const bg = rowIdx % 2 === 0 ? 'FFFFFF' : 'F9FBFF';

    const cells = row.map((cellText, colIdx) => {
      const text = String(cellText ?? '');
      const isNotOk =
        text === 'לא תקין' || text.startsWith('לא תקין');
      const isWatchOk =
        text === 'תקין - דורש מעקב' || text.startsWith('תקין - דורש מעקב');

      let color;
      let bold = false;
      if (isNotOk) {
        color = 'C00000';
        bold  = true;
      } else if (isWatchOk) {
        color = 'C55A11';
        bold  = true;
      }

      return new TableCell({
        width: { size: colWidths[colIdx] ?? colWidths[colWidths.length - 1], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.CLEAR, fill: bg },
        borders: THIN_BORDER,
        children: [
          mkPara([mkRun(text, { size: FS_TABLE, bold, color })], {
            alignment: AlignmentType.CENTER,
          }),
        ],
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

// ── Main export ──────────────────────────────────────────────────────────────

export async function generateDocument(data) {
  const cfg = DOC_TYPES[data.doc_type] ?? DOC_TYPES.group4;

  // ── Logo ─────────────────────────────────────────────────────────────────
  // Logo path must account for Vite's base path (/nir-reports/ in production)
  let logoBuffer = null;
  const logoPaths = [
    `${import.meta.env.BASE_URL}logo.png`,
    `${import.meta.env.BASE_URL}logo.svg`,
    '/logo.png',
  ];
  for (const path of logoPaths) {
    try {
      const res = await fetch(path);
      if (res.ok) {
        // Convert to Uint8Array – ArrayBuffer is not supported by all docx versions
        logoBuffer = new Uint8Array(await res.arrayBuffer());
        break;
      }
    } catch (_) {
      // try next path
    }
  }

  // ── Header ───────────────────────────────────────────────────────────────
  const headerChildren = [];

  if (logoBuffer) {
    headerChildren.push(
      mkPara(
        [
          new ImageRun({
            data: logoBuffer,
            transformation: { width: cm(4), height: cm(2.5) },
          }),
        ],
        { alignment: AlignmentType.RIGHT }
      )
    );
  }

  // Page number line: "עמוד PAGE מתוך NUMPAGES"
  // Use TextRun.children with PageNumber enum – the correct docx v8 API
  headerChildren.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: true,
      children: [
        new TextRun({ text: ' ', font: FONT, size: 7 * 2 }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 7 * 2 }),
        new TextRun({ text: ' מתוך ', font: FONT, size: 7 * 2, rtl: true }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 7 * 2 }),
        new TextRun({ text: 'עמוד ', font: FONT, size: 7 * 2, rtl: true }),
      ],
    })
  );

  const docHeader = new Header({ children: headerChildren });

  // ── Date paragraph (left-aligned, not RTL) ────────────────────────────────
  const dateStr = formatDate(data.date);
  const datePara = new Paragraph({
    alignment: AlignmentType.LEFT,
    children: [mkRun(dateStr, { size: FS_DATE, rtl: false })],
  });

  // ── Client block ─────────────────────────────────────────────────────────
  const toLabel  = mkPara([mkRun('לכבוד', { size: FS_CLIENT })]);
  const clientPara = mkPara([mkRun(data.client ?? '', { size: FS_CLIENT, bold: false })]);
  const orgPara    = mkPara([mkRun(data.organization ?? '', { size: FS_CLIENT, underline: true })]);
  const gapPara    = mkPara([mkRun('')]);

  // ── Subject ───────────────────────────────────────────────────────────────
  const subjectPara = mkPara(
    [mkRun(`הנדון: ${data.subject ?? ''}`, { size: FS_TITLE, bold: true, underline: true })],
    { alignment: AlignmentType.CENTER }
  );

  // ── Intro text ────────────────────────────────────────────────────────────
  const introText = data.intro_extra
    ? cfg.introTemplate(data) + '\n' + data.intro_extra
    : cfg.introTemplate(data);

  const introParas = introText.split('\n').map((line) =>
    mkPara([mkRun(line, { size: FS_INTRO })], {
      spacing: { line: Math.round(276 * 1.15), lineRule: 'AUTO' }, // 276 = single line in twips
    })
  );

  const gapPara2 = mkPara([mkRun('')]);

  // ── Section title ─────────────────────────────────────────────────────────
  const sectionTitlePara = mkPara([mkRun(cfg.sectionTitle, { size: FS_HEADING, bold: true })]);

  // ── Main data table ───────────────────────────────────────────────────────
  const tableRows = Array.isArray(data.table_rows) ? data.table_rows : [];
  let mainTable = null;
  if (tableRows.length > 0) {
    mainTable = mkTable(cfg.tableColumns, tableRows, cfg.colWidths);
  }

  const gapPara3 = mkPara([mkRun('')]);

  // ── Notes ─────────────────────────────────────────────────────────────────
  const notesTitlePara = mkPara([mkRun(cfg.notesTitle, { size: FS_HEADING, bold: true })]);

  const notesSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
    ? data.notes_custom
    : [...cfg.defaultNotes];

  // For group1: add validity note
  if (data.doc_type === 'group1' && cfg.validityMonths) {
    const validUntil = addDays(data.inspection_date, 365);
    notesSource.push(
      `תוקף האישור לשנה מיום הבדיקה ועד לתאריך ${validUntil} בכפוף למסקנות.`
    );
  }

  const notesParas = notesSource.map((note, i) =>
    mkPara([mkRun(`${i + 1}. ${note}`, { size: FS_BODY })], {})
  );

  // ── Conclusions ───────────────────────────────────────────────────────────
  const conclusionsTitlePara = mkPara([mkRun('מסקנות הבדיקה:', { size: FS_HEADING, bold: true })]);

  let conclusionLines;
  if (data.conclusion_custom && String(data.conclusion_custom).trim()) {
    conclusionLines = String(data.conclusion_custom).split('\n');
  } else if (data.has_defects) {
    conclusionLines = [cfg.conclusionDefects];
  } else {
    conclusionLines = [cfg.conclusionOk];
  }

  const conclusionParas = conclusionLines.map((line, i) =>
    mkPara([mkRun(`${i + 1}. ${line}`, { size: FS_BODY })], {})
  );

  // ── Defects table (appendix) ──────────────────────────────────────────────
  let defectsSection = [];
  if (
    data.has_defects &&
    cfg.hasDefectsTable &&
    cfg.defectsColumns &&
    Array.isArray(data.defects_rows) &&
    data.defects_rows.length > 0
  ) {
    const defectsTitlePara = mkPara(
      [mkRun('נספח טבלת ליקויים:', { size: FS_HEADING, bold: true })],
      { pageBreak: true }
    );
    const defectsTable = mkTable(cfg.defectsColumns, data.defects_rows, cfg.defectsColWidths);
    defectsSection = [defectsTitlePara, defectsTable, mkPara([mkRun('')])];
  }

  // ── Photos appendix ───────────────────────────────────────────────────────
  let photosSection = [];
  const photos = Array.isArray(data.photos) ? data.photos : [];

  if (photos.length > 0) {
    const photosTitlePara = mkPara(
      [mkRun('נספח תמונות:', { size: FS_HEADING, bold: true })],
      { pageBreak: defectsSection.length === 0 } // page break only if no defects section added it
    );

    // Actually always page break before photos appendix
    const photosTitleParaBreak = mkPara(
      [mkRun('נספח תמונות:', { size: FS_HEADING, bold: true })],
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
          dispH = 8; // cm
          dispW = natW > 0 ? (natW / natH) * dispH : 6;
        } else {
          dispW = 7.5; // cm
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
        const imgPara = mkPara([
          new ImageRun({
            data: p.imgData,
            transformation: {
              width:  cm(p.dispW),
              height: cm(p.dispH),
            },
          }),
        ], { alignment: AlignmentType.CENTER });

        const capPara = mkPara(
          [mkRun(p.caption, { size: FS_PHOTO_CAP })],
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

    photosSection = [photosTitleParaBreak, photoTable];
  }

  // ── Assemble document body ────────────────────────────────────────────────
  const bodyChildren = [
    datePara,
    toLabel,
    clientPara,
    orgPara,
    gapPara,
    subjectPara,
    mkPara([mkRun('')]),
    ...introParas,
    gapPara2,
    sectionTitlePara,
  ];

  if (mainTable) {
    bodyChildren.push(mainTable);
  }

  bodyChildren.push(
    gapPara3,
    notesTitlePara,
    ...notesParas,
    mkPara([mkRun('')]),
    conclusionsTitlePara,
    ...conclusionParas,
    ...defectsSection,
    ...photosSection,
  );

  // ── Build Document ────────────────────────────────────────────────────────
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width:  mm(210),   // A4 width
              height: mm(297),   // A4 height
            },
            margin: {
              top:    mm(42),    // 4.2 cm — room for logo header
              bottom: mm(8),
              left:   mm(20),
              right:  mm(20),
            },
          },
        },
        headers: { default: docHeader },
        children: bodyChildren,
      },
    ],
  });

  return await Packer.toBlob(doc);
}
