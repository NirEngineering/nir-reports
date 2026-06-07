import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, ImageRun, AlignmentType, UnderlineType, WidthType,
  ShadingType, convertMillimetersToTwip as mm, BorderStyle,
  TableBorders, VerticalAlign, PageNumber,
} from 'docx';

const FONT = 'Arial';
const DEFAULT_SIZE = 20; // half-points = 10pt

async function fetchBuf(paths) {
  for (const url of paths) {
    try { const r = await fetch(url); if (r.ok) return new Uint8Array(await r.arrayBuffer()); } catch {}
  }
  return new Uint8Array(0);
}

// ── Text alignment (no section bidi – LEFT/RIGHT are physical) ────────────────
function tiptapAlign(align) {
  switch (align) {
    case 'left':    return AlignmentType.LEFT;
    case 'center':  return AlignmentType.CENTER;
    case 'justify': return AlignmentType.BOTH;
    default:        return AlignmentType.RIGHT; // 'right' and default (Hebrew RTL)
  }
}

// ── Single text run ───────────────────────────────────────────────────────────
function convertTextNode(node) {
  if (node.type !== 'text') return null;
  const marks = node.marks || [];
  const bold      = marks.some(m => m.type === 'bold');
  const italic    = marks.some(m => m.type === 'italic');
  const underline = marks.some(m => m.type === 'underline');
  const strike    = marks.some(m => m.type === 'strike');
  const ts        = marks.find(m => m.type === 'textStyle')?.attrs || {};
  const rawColor  = ts.color || '';
  const color     = rawColor.startsWith('#') ? rawColor.slice(1) : (rawColor || undefined);
  const fontFam   = ts.fontFamily || FONT;
  const fsPt      = ts.fontSize ? parseInt(ts.fontSize) : 0;

  return new TextRun({
    text:      node.text || '',
    font:      fontFam,
    size:      fsPt ? fsPt * 2 : DEFAULT_SIZE,
    bold,
    italics:   italic,
    strike,
    underline: underline ? { type: UnderlineType.SINGLE } : undefined,
    color,
    rtl:       true,
  });
}

// ── Paragraph / heading ───────────────────────────────────────────────────────
function convertInlineContent(inlineNodes) {
  const runs = [];
  for (const n of (inlineNodes || [])) {
    if (n.type === 'text') {
      const r = convertTextNode(n);
      if (r) runs.push(r);
    } else if (n.type === 'hardBreak') {
      runs.push(new TextRun({ break: 1 }));
    }
  }
  return runs;
}

function convertParagraph(node) {
  const align = tiptapAlign(node.attrs?.textAlign);
  const runs = convertInlineContent(node.content);
  return new Paragraph({
    alignment: align,
    bidirectional: true,
    spacing: { after: 80 },
    children: runs.length ? runs : [new TextRun({ text: '' })],
  });
}

function convertHeading(node) {
  const level = node.attrs?.level || 1;
  const align = tiptapAlign(node.attrs?.textAlign);
  const sizes = [36, 30, 26, 24, 22, 20];
  const size  = sizes[level - 1] || 20;
  const forcedRuns = (node.content || []).map(n => {
    if (n.type !== 'text') return null;
    const marks = n.marks || [];
    const ts = marks.find(m => m.type === 'textStyle')?.attrs || {};
    return new TextRun({
      text: n.text || '', font: ts.fontFamily || FONT,
      size, bold: true, rtl: true,
      italics: marks.some(m => m.type === 'italic'),
      underline: marks.some(m => m.type === 'underline') ? { type: UnderlineType.SINGLE } : undefined,
    });
  }).filter(Boolean);

  return new Paragraph({
    alignment: align, bidirectional: true,
    spacing: { before: 200, after: 100 },
    children: forcedRuns.length ? forcedRuns : [new TextRun({ text: '', size, bold: true })],
  });
}

// ── Lists ────────────────────────────────────────────────────────────────────
function convertListItem(itemNode, listType, index = 1) {
  const paras = [];
  for (const child of (itemNode.content || [])) {
    if (child.type === 'paragraph') {
      const align = tiptapAlign(child.attrs?.textAlign);
      const prefix = listType === 'bullet' ? '•  ' : `${index}.  `;
      const runs = [
        new TextRun({ text: prefix, font: FONT, size: DEFAULT_SIZE, rtl: true }),
        ...convertInlineContent(child.content),
      ];
      paras.push(new Paragraph({
        alignment: align, bidirectional: true,
        spacing: { after: 40 },
        indent: { right: mm(6) },
        children: runs,
      }));
    }
  }
  return paras;
}

// ── Table ─────────────────────────────────────────────────────────────────────
const NO_BORDER   = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const CELL_BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };

function convertTableCell(node) {
  const bgRaw  = node.attrs?.backgroundColor || node.attrs?.background || '';
  const bg     = bgRaw.startsWith('#') ? bgRaw.slice(1) : bgRaw || undefined;
  const colspan = node.attrs?.colspan || 1;
  const rowspan = node.attrs?.rowspan || 1;
  const children = convertNodes(node.content || []);

  return new TableCell({
    children: children.length ? children : [new Paragraph({ children: [] })],
    columnSpan: colspan,
    rowSpan:    rowspan,
    verticalAlign: VerticalAlign.CENTER,
    shading: bg ? { fill: bg, type: ShadingType.SOLID, color: bg } : undefined,
    margins: { top: mm(1.5), bottom: mm(1.5), left: mm(2), right: mm(2) },
    borders: {
      top: CELL_BORDER, bottom: CELL_BORDER,
      left: CELL_BORDER, right:  CELL_BORDER,
    },
  });
}

function convertTable(node) {
  const rows = (node.content || []).map(rowNode =>
    new TableRow({
      children: (rowNode.content || []).map(cellNode => convertTableCell(cellNode)),
    })
  );
  const colCount = Math.max(...(node.content || []).map(r => (r.content || []).length), 1);
  return new Table({
    width:  { size: 100, type: WidthType.PERCENTAGE },
    layout: 'fixed',
    rows,
    columnWidths: Array(colCount).fill(Math.floor(9072 / colCount)),
    borders: {
      top: CELL_BORDER, bottom: CELL_BORDER,
      left: CELL_BORDER, right: CELL_BORDER,
      insideH: CELL_BORDER, insideV: CELL_BORDER,
    },
  });
}

// ── Image ─────────────────────────────────────────────────────────────────────
function b64ToUint8Array(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function convertImage(node) {
  const src = node.attrs?.src || '';
  if (!src.startsWith('data:')) return null;
  const [, b64] = src.split(',');
  if (!b64) return null;
  const data = b64ToUint8Array(b64);
  const MAX_W = 620; // ≈ 163mm at 96 DPI (within content area)
  const origW  = node.attrs?.width  || MAX_W;
  const origH  = node.attrs?.height || Math.round(origW * 0.6);
  const scale  = Math.min(1, MAX_W / origW);
  return new ImageRun({
    data,
    transformation: { width: Math.round(origW * scale), height: Math.round(origH * scale) },
  });
}

// ── Node dispatcher ───────────────────────────────────────────────────────────
function convertNodes(nodes = []) {
  const out = [];
  for (const node of nodes) {
    switch (node.type) {
      case 'paragraph':   out.push(convertParagraph(node)); break;
      case 'heading':     out.push(convertHeading(node)); break;
      case 'bulletList':
        out.push(...(node.content || []).flatMap(li => convertListItem(li, 'bullet')));
        break;
      case 'orderedList':
        out.push(...(node.content || []).flatMap((li, i) => convertListItem(li, 'ordered', i + 1)));
        break;
      case 'table': out.push(convertTable(node)); break;
      case 'horizontalRule':
        out.push(new Paragraph({
          alignment: AlignmentType.CENTER, bidirectional: true,
          children: [new TextRun({ text: '─'.repeat(50), font: FONT, size: 16 })],
        }));
        break;
      case 'image': {
        const run = convertImage(node);
        if (run) out.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          bidirectional: true,
          spacing: { after: 80 },
          children: [run],
        }));
        break;
      }
      default: break;
    }
  }
  return out;
}

// ── Main export ───────────────────────────────────────────────────────────────
export async function tiptapToDocx(jsonDoc, { title = 'מסמך' } = {}) {
  const base = import.meta.env.BASE_URL || '/';
  const [headerBuf, footerBuf] = await Promise.all([
    fetchBuf([`${base}header-logo.jpg`, '/nir-reports/header-logo.jpg', '/header-logo.jpg']),
    fetchBuf([`${base}footer-logo.png`, '/nir-reports/footer-logo.png', '/footer-logo.png']),
  ]);

  // Header: page number (top-right, ~3mm from top) then logo centered (~5mm from top)
  const headerSection = new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18 }),
          new TextRun({ text: '/', font: FONT, size: 18 }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18 }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: headerBuf.length > 0
          ? [new ImageRun({ data: headerBuf, transformation: { width: 376, height: 149 } })]
          : [new TextRun({ text: '' })],
      }),
    ],
  });

  const footerSection = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: footerBuf.length > 0
        ? [new ImageRun({ data: footerBuf, transformation: { width: 432, height: 30 } })]
        : [new TextRun({ text: '' })],
    })],
  });

  const body = convertNodes(jsonDoc?.content || []);
  if (!body.length) body.push(new Paragraph({ children: [] }));

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
          size:   { width: mm(210), height: mm(297) },
          margin: {
            top:    mm(50),    // leaves room for page-num (~3mm) + logo (~39mm) + gap
            bottom: mm(18),    // leaves room for footer logo (~8mm) + gap above 5mm margin
            left:   mm(17.92),
            right:  mm(17.46),
            header: mm(2),     // header starts at 2mm → page number at ~2mm, logo top at ~5mm
            footer: mm(5),     // footer bottom edge at 5mm from page bottom
          },
        },
        // No section-level bidi: paragraph-level bidirectional keeps RTL text direction
        // while LEFT/RIGHT alignment values mean physical left/right (no inversion)
      },
      headers: { default: headerSection },
      footers: { default: footerSection },
      children: body,
    }],
  });

  return Packer.toBlob(doc);
}
