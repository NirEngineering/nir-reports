import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Header,
  Footer,
  AlignmentType,
  ImageRun,
  convertMillimetersToTwip as mm,
  UnderlineType,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from 'docx';

const FONT = 'Arial';

const HEADER_W = 359, HEADER_H = 142;
const FOOTER_W = 568, FOOTER_H = 39;

const SP = { line: 360, lineRule: 'auto', after: 0 };

const NO_BORDER = { style: 'none', size: 0, color: 'FFFFFF' };
const ALL_NO_BORDER = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER, insideH: NO_BORDER, insideV: NO_BORDER };

function mkRun(text, { size = 9, bold = false, underline = false } = {}) {
  return new TextRun({
    text: String(text ?? ''),
    font: FONT,
    size: size * 2,
    bold,
    rtl: true,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function mkPara(children = [], { alignment = AlignmentType.RIGHT, spacing } = {}) {
  return new Paragraph({ children, alignment, spacing, bidirectional: true });
}

function emptyPara() {
  return mkPara([mkRun('')], { spacing: SP });
}

function b64ToUint8(b64) {
  const raw = b64.replace(/^data:[^,]+,/, '');
  const bin = atob(raw);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

async function fetchBuf(paths) {
  for (const url of paths) {
    try {
      const r = await fetch(url);
      if (r.ok) return new Uint8Array(await r.arrayBuffer());
    } catch { /* try next */ }
  }
  return new Uint8Array(0);
}

// ── Main export ───────────────────────────────────────────────────────────────
export async function generateEventApproval(data) {
  const base = import.meta.env.BASE_URL || '/';

  const [headerBuf, footerBuf] = await Promise.all([
    fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
    fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
  ]);

  const structures = typeof data.structures === 'string'
    ? data.structures.split('\n').filter(s => s.trim())
    : [];

  // Keep ~28 rows of space in the structures section
  const padLines = Math.max(0, 28 - structures.length);

  // ── Header ────────────────────────────────────────────────────────────────
  const headerSection = new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: headerBuf, transformation: { width: HEADER_W, height: HEADER_H } })],
      }),
    ],
  });

  // ── Footer ────────────────────────────────────────────────────────────────
  const footerSection = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })],
      }),
    ],
  });

  // ── לכבוד / תאריך on same visual line via borderless 2-cell table ──────────
  const toDateTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: ALL_NO_BORDER,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 65, type: WidthType.PERCENTAGE },
            borders: ALL_NO_BORDER,
            children: [mkPara([mkRun(`לכבוד: ${data.to || ''}`, { size: 8 })], { spacing: SP })],
          }),
          new TableCell({
            width: { size: 35, type: WidthType.PERCENTAGE },
            borders: ALL_NO_BORDER,
            children: [new Paragraph({
              alignment: AlignmentType.LEFT,
              bidirectional: true,
              spacing: SP,
              children: [mkRun(`תאריך: ${data.date || ''}`, { size: 8 })],
            })],
          }),
        ],
      }),
    ],
  });

  // ── Body ──────────────────────────────────────────────────────────────────
  const body = [

    toDateTable,
    emptyPara(),

    // הנדון – 11pt BOLD UNDERLINED CENTER
    new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: SP,
      children: [
        mkRun('הנדון : אישור מבנים ומתקנים ארעיים/קבועים', { size: 11, bold: true, underline: true }),
      ],
    }),

    emptyPara(),

    // פרטי העסק ובעל העסק – 9pt UNDERLINED (not bold)
    mkPara([mkRun('פרטי העסק ובעל העסק', { size: 9, underline: true })], { spacing: SP }),

    emptyPara(),
    emptyPara(),

    // כתובת + שם
    mkPara([
      mkRun(`כתובת העסק: ${data.address || ''}`, { size: 9 }),
      mkRun('          ', { size: 9 }),
      mkRun(`שם בעל העסק/המזמין: ${data.owner || ''}`, { size: 9 }),
    ], { spacing: SP }),

    // ת.זהות + טלפון
    mkPara([
      mkRun(`מספר ת.זהות: ${data.id_num || ''}`, { size: 9 }),
      mkRun('          ', { size: 9 }),
      mkRun(`טלפון סלולארי: ${data.phone || ''}`, { size: 9 }),
    ], { spacing: SP }),

    emptyPara(),

    // להלן המתקנים
    mkPara([mkRun('להלן המתקנים הארעיים/הקבועים שנבדקו:', { size: 9 })], { spacing: SP }),

    // Each structure – CENTER 14pt
    ...structures.map(s =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        bidirectional: true,
        spacing: SP,
        children: [mkRun(s, { size: 14 })],
      })
    ),

    // Padding empty rows
    ...Array.from({ length: padLines }, () => emptyPara()),

    // המתקנים לעיל נבדקו – 9pt (NOT bold)
    mkPara([mkRun('המתקנים לעיל נבדקו בהיבטים הבאים:', { size: 9 })], { spacing: SP }),

    ...['תקינות ויציבות המבנה',
        'יציבות הקרקע/התשתית עליה מונח המבנה',
        'חיבורים',
        'העמסות',
    ].map(t => mkPara([mkRun(`◦  ${t}`, { size: 9 })], { spacing: SP })),

    emptyPara(),

    // הערות: – 9pt (NOT bold)
    mkPara([mkRun('הערות:', { size: 9 })], { spacing: SP }),

    ...['האישור תקף למבנים/מתקנים שצויינו במסמך זה בלבד.',
        "אין לבצע שינויים במבנים/מתקנים - כל שינוי מבני ו/או הפחתות/תוספות למבנים, ללא ידיעת הח''מ, תגרור לביטול אישור זה.",
        'האישור אינו מתייחס לתכנון המבנים ועיגונם אלא את לכשירותם ותקינותם בלבד.',
        'אישור זה מתייחס לזמן הפעילות בלבד ואינו כולל זמני פירוק והרכבה.',
    ].map(t => mkPara([mkRun(`•  ${t}`, { size: 9 })], { spacing: SP })),

    emptyPara(),
    emptyPara(),

    // הערות נוספות – 9pt BOLD
    mkPara([
      mkRun('הערות נוספות: ', { size: 9, bold: true }),
      mkRun(data.notes || '', { size: 9 }),
    ], { spacing: SP }),

    emptyPara(),
    emptyPara(),
    emptyPara(),

    // תוקף האישור – 10pt BOLD UNDERLINED
    mkPara([
      mkRun('תוקף האישור : ', { size: 10, bold: true, underline: true }),
      mkRun(data.validity || '', { size: 10, bold: true }),
    ], { spacing: SP }),
  ];

  // ── Photos section (2 per row) ────────────────────────────────────────────
  const photos = Array.isArray(data.photos) ? data.photos : [];
  const photoItems = [];
  for (let i = 0; i < photos.length; i += 2) {
    const pair = photos.slice(i, i + 2);
    const emptyCell = () => new TableCell({
      children: [new Paragraph({ children: [] })],
      width: { size: 50, type: WidthType.PERCENTAGE },
      borders: ALL_NO_BORDER,
    });
    const imgCells = pair.map(ph => new TableCell({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: b64ToUint8(ph.data), transformation: { width: 255, height: 191 } })],
      })],
      width: { size: 50, type: WidthType.PERCENTAGE },
      margins: { top: 40, bottom: 40, left: 40, right: 40 },
      borders: ALL_NO_BORDER,
    }));
    const capCells = pair.map(ph => new TableCell({
      children: [mkPara([mkRun(ph.caption || '', { size: 8 })], { alignment: AlignmentType.CENTER, spacing: SP })],
      width: { size: 50, type: WidthType.PERCENTAGE },
      borders: ALL_NO_BORDER,
    }));
    if (pair.length < 2) { imgCells.push(emptyCell()); capCells.push(emptyCell()); }
    photoItems.push(
      new Table({
        alignment: AlignmentType.CENTER,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: ALL_NO_BORDER,
        rows: [
          new TableRow({ children: imgCells }),
          new TableRow({ children: capCells }),
        ],
      }),
      emptyPara()
    );
  }

  if (photos.length > 0) {
    body.push(
      emptyPara(),
      mkPara([mkRun('תמונות:', { size: 9, bold: true })], { spacing: { ...SP, before: 360 } }),
      ...photoItems
    );
  }

  const doc = new Document({
    styles: {
      default: { document: { run: { font: { name: FONT } } } },
      paragraphStyles: [{
        id: 'Normal', name: 'Normal', quickFormat: true,
        paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT },
        run: { font: { name: FONT } },
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: mm(210), height: mm(297) },
          margin: {
            top:    mm(40),
            bottom: mm(2.5),
            left:   mm(15),
            right:  mm(15),
            header: mm(0),
            footer: mm(4.6),
          },
        },
        bidi: true,
      },
      headers: { default: headerSection },
      footers: { default: footerSection },
      children: body,
    }],
  });

  return Packer.toBlob(doc);
}
