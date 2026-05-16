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

// Pixels at 96 DPI — used for ImageRun.transformation
const cmPx = (v) => Math.round(v / 2.54 * 96);

// Exact pixel sizes derived from sample Word document (914400 EMU = 1 inch = 96px)
const HEADER_W = 359, HEADER_H = 142;   // 95.12mm × 37.57mm
const FOOTER_W = 568, FOOTER_H = 39;    // 150.50mm × 10.28mm
const STAMP_W  = cmPx(2.349);   // 23.49 mm
const STAMP_H  = cmPx(1.313);   // 13.13 mm

// Line spacing constants
const SP = { line: 360, lineRule: 'auto', after: 0 };

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

  const [headerBuf, footerBuf, stampBuf] = await Promise.all([
    fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
    fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
    fetchBuf([`${base}stamp.png`,       '/nir-reports/stamp.png',       '/stamp.png']),
  ]);

  const structures = typeof data.structures === 'string'
    ? data.structures.split('\n').filter(s => s.trim())
    : [];

  // Pad empty lines so the aspects/conditions section stays below mid-page
  const padLines = Math.max(0, 10 - structures.length);

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

  // ── Stamp paragraph ───────────────────────────────────────────────────────
  const stampPara = stampBuf.length > 100
    ? new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: SP,
        children: [new ImageRun({ data: stampBuf, transformation: { width: STAMP_W, height: STAMP_H } })],
      })
    : mkPara([mkRun('')], { spacing: SP });

  // ── Body paragraphs ───────────────────────────────────────────────────────
  const body = [

    // לכבוד — right-aligned; תאריך — left-aligned (separate paragraphs)
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      bidirectional: true,
      spacing: SP,
      children: [mkRun(`לכבוד: ${data.to || ''}`, { size: 8 })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      bidirectional: true,
      spacing: SP,
      children: [mkRun(`תאריך: ${data.date || ''}`, { size: 10 })],
    }),

    mkPara([mkRun('')], { spacing: SP }),

    // הנדון — centered, bold/underline
    new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 480, after: 0 },
      children: [
        mkRun('הנדון  :  ', { size: 16, bold: true }),
        mkRun('אישור מבנים ומתקנים ארעיים/קבועים', { size: 16, bold: true, underline: true }),
      ],
    }),

    mkPara([mkRun('')], { spacing: SP }),

    // פרטי העסק ובעל העסק (section header)
    mkPara([mkRun('פרטי העסק ובעל העסק', { size: 10, bold: true })], { spacing: { ...SP, before: 360 } }),

    // כתובת + שם
    mkPara([
      mkRun(`כתובת העסק: ${data.address || ''}`, { size: 9 }),
      mkRun(' '.repeat(12), { size: 9 }),
      mkRun(`שם בעל העסק/המזמין: ${data.owner || ''}`, { size: 9 }),
    ], { spacing: SP }),

    // ת.זהות + טלפון
    mkPara([
      mkRun(`מספר ת.זהות: ${data.id_num || ''}`, { size: 9 }),
      mkRun(' '.repeat(12), { size: 9 }),
      mkRun(`טלפון סלולארי: ${data.phone || ''}`, { size: 9 }),
    ], { spacing: SP }),

    mkPara([mkRun('')], { spacing: SP }),

    // section label: structures
    mkPara([mkRun('להלן המתקנים הארעיים/הקבועים שנבדקו:', { size: 9 })], { spacing: SP }),

    // Each structure on its own line
    ...structures.map((s, i) =>
      mkPara([mkRun(`${i + 1}. ${s}`, { size: 9 })], { spacing: SP })
    ),

    // Padding empty lines to keep spacing consistent
    ...Array.from({ length: padLines }, () => mkPara([mkRun('')], { spacing: SP })),

    mkPara([mkRun('')], { spacing: SP }),
    mkPara([mkRun('')], { spacing: SP }),

    // Inspected aspects
    mkPara([mkRun('המתקנים לעיל נבדקו בהיבטים הבאים:', { size: 10, bold: true })], { spacing: { ...SP, before: 360 } }),

    ...['תקינות ויציבות המבנה',
        'יציבות הקרקע/התשתית עליה מונח המבנה',
        'חיבורים',
        'העמסות',
    ].map(t => mkPara([mkRun(`◦  ${t}`, { size: 9 })], { spacing: SP })),

    mkPara([mkRun('')], { spacing: SP }),

    // Legal conditions
    mkPara([mkRun('הערות:', { size: 10, bold: true })], { spacing: { ...SP, before: 360 } }),

    ...['האישור תקף למבנים/מתקנים שצויינו במסמך זה בלבד.',
        "אין לבצע שינויים במבנים/מתקנים - כל שינוי מבני ו/או הפחתות/תוספות למבנים, ללא ידיעת הח''מ, תגרור לביטול אישור זה.",
        'האישור אינו מתייחס לתכנון המבנים ועיגונם אלא את לכשירותם ותקינותם בלבד.',
        'אישור זה מתייחס לזמן הפעילות בלבד ואינו כולל זמני פירוק והרכבה.',
    ].map(t => mkPara([mkRun(`•  ${t}`, { size: 9 })], { spacing: SP })),

    mkPara([mkRun('')], { spacing: SP }),
    mkPara([mkRun('')], { spacing: SP }),

    // Additional notes
    mkPara([
      mkRun('הערות נוספות:  ', { size: 9, bold: true }),
      mkRun(data.notes || '', { size: 9 }),
    ], { spacing: SP }),

    mkPara([mkRun('')], { spacing: SP }),
    mkPara([mkRun('')], { spacing: SP }),

    // Stamp
    stampPara,

    mkPara([mkRun('')], { spacing: SP }),

    // Validity / signature line
    mkPara([
      mkRun('תוקף האישור :   ', { size: 9, bold: true }),
      mkRun(data.validity || '', { size: 9, bold: true }),
      mkRun(' '.repeat(50), { size: 9 }),
      mkRun('חותמת וחתימה: _______________', { size: 9, bold: true }),
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
    });
    const imgCells = pair.map(ph => new TableCell({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({ data: b64ToUint8(ph.data), transformation: { width: 255, height: 191 } })],
      })],
      width: { size: 50, type: WidthType.PERCENTAGE },
      margins: { top: 40, bottom: 40, left: 40, right: 40 },
    }));
    const capCells = pair.map(ph => new TableCell({
      children: [mkPara([mkRun(ph.caption || '', { size: 9 })], { alignment: AlignmentType.CENTER, spacing: SP })],
      width: { size: 50, type: WidthType.PERCENTAGE },
    }));
    if (pair.length < 2) { imgCells.push(emptyCell()); capCells.push(emptyCell()); }
    photoItems.push(
      new Table({
        alignment: AlignmentType.CENTER,
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: imgCells }),
          new TableRow({ children: capCells }),
        ],
      }),
      mkPara([mkRun('')], { spacing: SP })
    );
  }

  if (photos.length > 0) {
    body.push(
      mkPara([mkRun('')], { spacing: SP }),
      mkPara([mkRun('תמונות:', { size: 10, bold: true })], { spacing: { ...SP, before: 360 } }),
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
            top:    mm(31.70),
            bottom: mm(12.51),
            left:   mm(20.00),
            right:  mm(24.98),
            header: mm(3.00),
            footer: mm(1.99),
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
