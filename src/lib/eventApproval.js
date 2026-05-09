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
} from 'docx';

const FONT = 'Arial';

// Pixels at 96 DPI — used for ImageRun.transformation
const cmPx = (v) => Math.round(v / 2.54 * 96);

// Image display sizes from original document (exact measurements)
const HEADER_W = cmPx(10.595);  // 105.95 mm
const HEADER_H = cmPx(4.178);   // 41.78 mm
const FOOTER_W = cmPx(13.5);    // 135.00 mm
const FOOTER_H = cmPx(0.995);   // 9.95 mm
const STAMP_W  = cmPx(2.349);   // 23.49 mm
const STAMP_H  = cmPx(1.313);   // 13.13 mm

// Line spacing constants
const SP     = { line: 276, lineRule: 'auto', after: 0 };
const SP_DBL = { line: 480, lineRule: 'auto', after: 0 };
const SP_15  = { line: 360, lineRule: 'auto', after: 0 };

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
  return new Paragraph({ children, alignment, spacing });
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
        alignment: AlignmentType.LEFT,
        children: [new ImageRun({ data: footerBuf, transformation: { width: FOOTER_W, height: FOOTER_H } })],
      }),
    ],
  });

  // ── Stamp paragraph ───────────────────────────────────────────────────────
  const stampPara = stampBuf.length > 100
    ? new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: SP_15,
        children: [new ImageRun({ data: stampBuf, transformation: { width: STAMP_W, height: STAMP_H } })],
      })
    : mkPara([mkRun('')], { spacing: SP_15 });

  // ── Body paragraphs ───────────────────────────────────────────────────────
  const body = [

    // Row 1: לכבוד ... תאריך (8pt, right-aligned, on same line via spaces)
    mkPara([
      mkRun(`לכבוד: ${data.to || ''}`, { size: 8 }),
      mkRun(' '.repeat(55), { size: 8 }),
      mkRun(`תאריך: ${data.date || ''}`, { size: 8 }),
    ], { spacing: SP }),

    mkPara([mkRun('')], { spacing: SP }),

    // הנדון — centered, bold/underline
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: SP,
      children: [
        mkRun('הנדון  :  ', { size: 11.5, bold: true }),
        mkRun('אישור מבנים ומתקנים ארעיים/קבועים', { size: 11.5, bold: true, underline: true }),
      ],
    }),

    mkPara([mkRun('')], { spacing: SP }),

    // פרטי העסק ובעל העסק (underlined section header)
    mkPara([mkRun('פרטי העסק ובעל העסק ', { size: 9, underline: true })], { spacing: SP }),

    // כתובת + שם (double spacing)
    mkPara([
      mkRun(`כתובת העסק: ${data.address || ''}`, { size: 9 }),
      mkRun(' '.repeat(12), { size: 9 }),
      mkRun(`שם בעל העסק/המזמין: ${data.owner || ''}`, { size: 9 }),
    ], { spacing: SP_DBL }),

    // ת.זהות + טלפון (double spacing)
    mkPara([
      mkRun(`מספר ת.זהות: ${data.id_num || ''}`, { size: 9 }),
      mkRun(' '.repeat(12), { size: 9 }),
      mkRun(`טלפון סלולארי: ${data.phone || ''}`, { size: 9 }),
    ], { spacing: SP_DBL }),

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
    mkPara([mkRun('המתקנים לעיל נבדקו בהיבטים הבאים:', { size: 9 })], { spacing: SP_15 }),

    ...['תקינות ויציבות המבנה',
        'יציבות הקרקע/התשתית עליה מונח המבנה',
        'חיבורים',
        'העמסות',
    ].map(t => mkPara([mkRun(`◦  ${t}`, { size: 9 })], { spacing: SP })),

    mkPara([mkRun('')], { spacing: SP }),

    // Legal conditions
    mkPara([mkRun('הערות:', { size: 9 })], { spacing: SP_15 }),

    ...['האישור תקף למבנים/מתקנים שצויינו במסמך זה בלבד.',
        "אין לבצע שינויים במבנים/מתקנים - כל שינוי מבני ו/או הפחתות/תוספות למבנים, ללא ידיעת הח''מ, תגרור לביטול אישור זה.",
        'האישור אינו מתייחס לתכנון המבנים ועיגונם אלא את לכשירותם ותקינותם בלבד.',
        'אישור זה מתייחס לזמן הפעילות בלבד ואינו כולל זמני פירוק והרכבה.',
    ].map(t => mkPara([mkRun(`•  ${t}`, { size: 9 })], { spacing: SP })),

    mkPara([mkRun('')], { spacing: SP }),
    mkPara([mkRun('')], { spacing: SP_15 }),

    // Additional notes
    mkPara([
      mkRun('הערות נוספות:  ', { size: 9, bold: true }),
      mkRun(data.notes || '', { size: 9 }),
    ], { spacing: SP_15 }),

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

  const doc = new Document({
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
            footer: mm(4.55),
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
