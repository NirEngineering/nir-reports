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

// ── Font / size constants (matched to original Word documents) ───────────────
const FONT         = 'Arial';
const FS_DATE      = 8;    // date line (top right)
const FS_CLIENT    = 8;    // "לכבוד", client name, org
const FS_TITLE     = 15;   // הנדון — bold, centered
const FS_INTRO     = 8.5;  // intro / body paragraphs
const FS_HEADING   = 9;    // section headings (bold)
const FS_BODY      = 8.5;  // general body text
const FS_TABLE_H   = 9;    // table header row
const FS_TABLE_D   = 8.5;  // table data cells
const FS_PHOTO_CAP = 9;    // photo captions

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
 * Create a paragraph. RTL direction is inherited from the section (bidi: true in sectPr).
 * Do NOT set bidirectional on individual paragraphs — it causes image mirroring.
 * opts: { alignment, spacing, pageBreak }
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
 * - Data cells: always RIGHT (matches original document style)
 */
function mkTable(headers, rows, colWidthsCm, headerBg = 'EAF1DD') {
  const colWidths = colWidthsCm.map((w) => cm(w));

  // Header row – always centered
  const headerCells = headers.map((h, i) =>
    new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.CLEAR, fill: headerBg },
      borders: THIN_BORDER,
      children: [mkPara([mkRun(h, { size: FS_TABLE_H, bold: true })], { alignment: AlignmentType.CENTER })],
    })
  );

  const headerRow = new TableRow({ children: headerCells, tableHeader: true });

  // Data rows – all white (no alternating tint)
  const dataRows = rows.map((row) => {
    const bg = 'FFFFFF';

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
        shading:       { type: ShadingType.CLEAR, fill: bg },
        borders:       THIN_BORDER,
        children: [mkPara([mkRun(text, { size: FS_TABLE_D, bold, color })], { alignment: AlignmentType.CENTER })],
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

  // Use client name as fallback when location is empty
  const effectiveData = (data.location || '').trim()
    ? data
    : { ...data, location: (data.client || '').trim() };

  // ── Logo images (extracted from original Word documents) ─────────────────
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

  // ── Header: page-number + header logo (8.91 × 3.52 cm) ──────────────────
  // Dimensions from original documents. No bidirectional on image paragraphs.
  const HEADER_W = cmPx(8.91), HEADER_H = cmPx(3.52);
  const FOOTER_W = cmPx(15.05), FOOTER_H = cmPx(1.03);

  const pageNumPara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0 },
    children: [
      new TextRun({ text: 'עמוד ', font: FONT, size: 7 * 2, rtl: true }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 7 * 2 }),
      new TextRun({ text: ' מתוך ', font: FONT, size: 7 * 2, rtl: true }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 7 * 2 }),
    ],
  });

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

  // ── Footer: footer strip image (15.05 × 1.03 cm) ─────────────────────────
  const footerLogoPara = footerLogoBuffer
    ? new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [new ImageRun({ data: footerLogoBuffer, transformation: { width: FOOTER_W, height: FOOTER_H } })],
      })
    : new Paragraph({ children: [] });

  const docFooter = new Footer({ children: [footerLogoPara] });

  // Spacing matched to original documents: 1.5× line spacing, no paragraph gaps
  const SP_BODY    = { line: 360, lineRule: 'auto', after: 0 };
  const SP_SECTION = { line: 360, lineRule: 'auto', before: 120, after: 0 };

  // ── Date paragraph (right-aligned, matching original documents) ──────────
  const dateStr  = formatDate(data.date);
  const datePara = new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0 },
    children: [mkRun(dateStr, { size: FS_DATE })],
  });

  // ── Client block ─────────────────────────────────────────────────────────
  const toLabel    = mkPara([mkRun('לכבוד',                   { size: FS_CLIENT })],         { spacing: { after: 0 } });
  const clientPara = mkPara([mkRun(data.client ?? '',          { size: FS_CLIENT })],         { spacing: { after: 0 } });
  const orgPara    = mkPara([mkRun(data.organization ?? '',    { size: FS_CLIENT, underline: true })], { spacing: { after: 0 } });

  // ── Subject ───────────────────────────────────────────────────────────────
  const subjectPara = mkPara(
    [mkRun(`הנדון: ${data.subject ?? ''}`, { size: FS_TITLE, bold: true, underline: true })],
    { alignment: AlignmentType.CENTER, spacing: { before: 40, after: 160 } }
  );

  // ── Intro text ────────────────────────────────────────────────────────────
  const introText = data.intro_extra
    ? cfg.introTemplate(effectiveData) + '\n' + data.intro_extra
    : cfg.introTemplate(effectiveData);

  const introParas = introText.split('\n').map((line) =>
    mkPara([mkRun(line, { size: FS_INTRO })], { spacing: SP_BODY })
  );

  // ── Section title ─────────────────────────────────────────────────────────
  const sectionTitlePara = mkPara(
    [mkRun(cfg.sectionTitle, { size: FS_HEADING, bold: true })],
    { spacing: SP_SECTION }
  );

  // ── Main data table ───────────────────────────────────────────────────────
  const tableRows = Array.isArray(data.table_rows) ? data.table_rows : [];
  let mainTable = null;
  if (tableRows.length > 0) {
    mainTable = mkTable(cfg.tableColumns, tableRows, cfg.colWidths);
  }

  // ── Notes ─────────────────────────────────────────────────────────────────
  const notesTitlePara = mkPara(
    [mkRun(cfg.notesTitle, { size: FS_HEADING, bold: true })],
    { spacing: SP_SECTION }
  );

  const notesSource = Array.isArray(data.notes_custom) && data.notes_custom.length > 0
    ? data.notes_custom
    : [...cfg.defaultNotes];

  // For group1: add validity note
  if (data.doc_type === 'group1' && cfg.validityMonths) {
    const validUntil = addDays(data.inspection_date, 365);
    notesSource.push(`תוקף האישור לשנה מיום הבדיקה ועד לתאריך ${validUntil} בכפוף למסקנות.`);
  }

  const notesParas = notesSource.map((note, i) =>
    mkPara([mkRun(`${i + 1}. ${note}`, { size: FS_BODY })], { spacing: SP_BODY })
  );

  // ── Conclusions ───────────────────────────────────────────────────────────
  const conclusionsTitlePara = mkPara(
    [mkRun('מסקנות הבדיקה:', { size: FS_HEADING, bold: true })],
    { spacing: SP_SECTION }
  );

  let conclusionLines;
  if (data.conclusion_custom && String(data.conclusion_custom).trim()) {
    conclusionLines = String(data.conclusion_custom).split('\n');
  } else if (data.has_defects) {
    conclusionLines = [cfg.conclusionDefects];
  } else {
    conclusionLines = [cfg.conclusionOk];
  }

  const conclusionParas = conclusionLines.map((line, i) =>
    mkPara([mkRun(`${i + 1}. ${line}`, { size: FS_BODY })], { spacing: SP_BODY })
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
        const imgPara = mkPara([
          new ImageRun({
            data: p.imgData,
            transformation: {
              width:  cmPx(p.dispW),
              height: cmPx(p.dispH),
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
    subjectPara,
    ...introParas,
    sectionTitlePara,
  ];

  if (mainTable) {
    bodyChildren.push(mainTable);
  }

  bodyChildren.push(
    notesTitlePara,
    ...notesParas,
    conclusionsTitlePara,
    ...conclusionParas,
    ...defectsSection,
    ...photosSection,
  );

  // ── Build Document ────────────────────────────────────────────────────────
  const doc = new Document({
    // Set document-level RTL default so Word renders Hebrew correctly in all views
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
              // Exact values from original Word documents
              top:    mm(40),    // body starts 40mm from top (header image fits in header area)
              bottom: mm(10),    // body ends 10mm from bottom
              left:   mm(22),    // physical left margin
              right:  mm(32),    // physical right margin (wider, Hebrew text starts here)
              header: mm(2.5),   // header content starts 2.5mm from top edge
              footer: mm(0),     // footer content at bottom edge
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
