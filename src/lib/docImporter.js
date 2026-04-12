/**
 * docImporter.js
 * Browser-side .docx import / edit / re-export for the Hebrew RTL engineering report app.
 *
 * Exports:
 *   importDocx(file)            → Promise<ImportResult>
 *   exportDocx(importResult, edits) → Promise<Blob>
 */

import JSZip from 'jszip';

// ---------------------------------------------------------------------------
// Namespace helpers
// ---------------------------------------------------------------------------
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Return a live HTMLCollection of all w:<tag> descendants of `el`. */
const wEl = (el, tag) => el.getElementsByTagNameNS(W_NS, tag);

/** Extract all text inside a w:p / w:tc element. */
function textOf(el) {
  return Array.from(wEl(el, 't'))
    .map((t) => t.textContent)
    .join('');
}

/** Return first child element with local name `tag` under `el` in W_NS. */
function wFirst(el, tag) {
  return el.getElementsByTagNameNS(W_NS, tag)[0] || null;
}

// ---------------------------------------------------------------------------
// importDocx
// ---------------------------------------------------------------------------

/**
 * Import a .docx File, parse its content, and return a structured ImportResult.
 *
 * @param {File} file
 * @returns {Promise<ImportResult>}
 */
export async function importDocx(file) {
  // ── 1. Open ZIP ────────────────────────────────────────────────────────────
  let zip;
  try {
    zip = await JSZip.loadAsync(file);
  } catch (err) {
    throw new Error(`Failed to open .docx file as ZIP: ${err.message}`);
  }

  // ── 2. Read document.xml ──────────────────────────────────────────────────
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) throw new Error('word/document.xml not found in .docx');
  const originalXml = await docXmlFile.async('string');

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(originalXml, 'application/xml');

  const parseError = xmlDoc.querySelector('parsererror');
  if (parseError) throw new Error(`XML parse error: ${parseError.textContent}`);

  // ── 3. Parse paragraphs ───────────────────────────────────────────────────
  const wParagraphs = Array.from(wEl(xmlDoc, 'p'));
  const paragraphs = wParagraphs.map((p, idx) => {
    const text = textOf(p);

    // Bold: any w:rPr/w:b element present with no w:b val="false"
    const boldEls = Array.from(wEl(p, 'b'));
    const isBold = boldEls.some((b) => {
      const val = b.getAttributeNS(W_NS, 'val');
      return val !== 'false' && val !== '0';
    });

    // Center alignment
    const jcEls = Array.from(wEl(p, 'jc'));
    const isCenter = jcEls.some((jc) => {
      const val = jc.getAttributeNS(W_NS, 'val');
      return val === 'center' || val === 'both';
    });

    // Font size (half-points → pt)
    const szEl = wFirst(p, 'sz');
    const halfPt = szEl ? parseInt(szEl.getAttributeNS(W_NS, 'val') || '0', 10) : 0;
    const fontSize = halfPt ? halfPt / 2 : 0;

    // Heading heuristic
    const isHeading = isBold && text.trim().length > 0 && text.trim().length < 60;

    return {
      id: `p_${idx}`,
      text,
      isBold,
      isCenter,
      fontSize,
      isHeading,
      rawXml: p.outerHTML || new XMLSerializer().serializeToString(p),
    };
  });

  // ── 4. Parse tables ───────────────────────────────────────────────────────
  const wTables = Array.from(wEl(xmlDoc, 'tbl'));
  const tables = wTables.map((tbl, tIdx) => {
    const wRows = Array.from(wEl(tbl, 'tr'));
    const allRows = wRows.map((row) => {
      const cells = Array.from(wEl(row, 'tc'));
      // Deduplicate adjacent cells that have identical text (merged cell artefact)
      const cellTexts = [];
      let prev = null;
      for (const cell of cells) {
        const ct = textOf(cell).trim();
        if (ct !== prev) cellTexts.push(ct);
        prev = ct;
      }
      return cellTexts;
    });

    const headers = allRows[0] || [];
    const rows = allRows.slice(1);

    return {
      id: `t_${tIdx}`,
      headers,
      rows,
      rawXml: tbl.outerHTML || new XMLSerializer().serializeToString(tbl),
    };
  });

  // ── 5. Extract images ─────────────────────────────────────────────────────
  const images = await extractImages(zip);

  // ── 6. Smart field extraction ─────────────────────────────────────────────
  const extracted = extractFields(paragraphs, tables);

  return {
    _zip: zip,
    _originalXml: originalXml,
    paragraphs,
    tables,
    images,
    extracted,
  };
}

// ---------------------------------------------------------------------------
// exportDocx
// ---------------------------------------------------------------------------

/**
 * Apply text edits to the original XML and re-package the .docx.
 *
 * @param {ImportResult} importResult   - result from importDocx()
 * @param {{ paragraphs: Record<string,string> }} edits
 * @returns {Promise<Blob>}
 */
export async function exportDocx(importResult, edits) {
  const { _zip, _originalXml } = importResult;
  if (!_zip || !_originalXml) throw new Error('Invalid importResult – missing _zip or _originalXml');

  const editMap = (edits && edits.paragraphs) ? edits.paragraphs : {};

  // ── 1. Parse original XML ─────────────────────────────────────────────────
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(_originalXml, 'application/xml');

  const parseError = xmlDoc.querySelector('parsererror');
  if (parseError) throw new Error(`XML parse error: ${parseError.textContent}`);

  // ── 2. Apply edits ────────────────────────────────────────────────────────
  const wParagraphs = Array.from(wEl(xmlDoc, 'p'));

  for (const [id, newText] of Object.entries(editMap)) {
    // id format: "p_N"
    const match = id.match(/^p_(\d+)$/);
    if (!match) continue;
    const idx = parseInt(match[1], 10);
    if (idx < 0 || idx >= wParagraphs.length) continue;

    const p = wParagraphs[idx];
    const runs = Array.from(wEl(p, 'r'));

    if (runs.length === 0) {
      // No runs – create one
      const r = xmlDoc.createElementNS(W_NS, 'w:r');
      const t = xmlDoc.createElementNS(W_NS, 'w:t');
      t.setAttribute('xml:space', 'preserve');
      t.textContent = newText;
      r.appendChild(t);
      p.appendChild(r);
      continue;
    }

    // Keep first run's formatting, set its text, remove extra runs
    const firstRun = runs[0];
    const firstT = wFirst(firstRun, 't') || (() => {
      const t = xmlDoc.createElementNS(W_NS, 'w:t');
      firstRun.appendChild(t);
      return t;
    })();

    firstT.setAttribute('xml:space', 'preserve');
    firstT.textContent = newText;

    // Remove all w:t in first run except the one we just set
    Array.from(wEl(firstRun, 't')).forEach((t) => {
      if (t !== firstT) t.parentNode.removeChild(t);
    });

    // Remove subsequent runs
    for (let i = 1; i < runs.length; i++) {
      runs[i].parentNode.removeChild(runs[i]);
    }
  }

  // ── 3. Serialize back ─────────────────────────────────────────────────────
  const serializer = new XMLSerializer();
  const newXml = serializer.serializeToString(xmlDoc);

  // ── 4. Update ZIP and generate blob ───────────────────────────────────────
  _zip.file('word/document.xml', newXml);

  const blob = await _zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return blob;
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Extract images from the ZIP and return array of base64 data-URLs (max 20).
 */
async function extractImages(zip) {
  const images = [];

  // Read relationships to find image files
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!relsFile) return images;

  let relsXml;
  try {
    relsXml = await relsFile.async('string');
  } catch {
    return images;
  }

  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  const rels = Array.from(relsDoc.getElementsByTagName('Relationship'));

  for (const rel of rels) {
    if (images.length >= 20) break;

    const type = rel.getAttribute('Type') || '';
    if (!type.endsWith('/image')) continue;

    const target = rel.getAttribute('Target') || '';
    // Target is relative to word/, e.g. "media/image1.png"
    const zipPath = `word/${target.replace(/^\//, '')}`;
    const imgFile = zip.file(zipPath);
    if (!imgFile) continue;

    try {
      const base64 = await imgFile.async('base64');
      const ext = (zipPath.split('.').pop() || 'png').toLowerCase();
      const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
        : ext === 'gif' ? 'image/gif'
        : ext === 'bmp' ? 'image/bmp'
        : ext === 'webp' ? 'image/webp'
        : 'image/png';
      images.push(`data:${mime};base64,${base64}`);
    } catch {
      // skip unreadable image
    }
  }

  return images;
}

/**
 * Smart field extraction from parsed paragraphs and tables.
 */
function extractFields(paragraphs, tables) {
  const texts = paragraphs.map((p) => p.text.trim());

  // ── Dates ─────────────────────────────────────────────────────────────────
  const dateRe = /(\d{1,2})[./](\d{1,2})[./](\d{4})/;
  let date = '';
  let inspection_date = '';
  let dateCount = 0;
  for (let i = 0; i < Math.min(texts.length, 30); i++) {
    const m = dateRe.exec(texts[i]);
    if (m) {
      const iso = `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
      if (dateCount === 0) date = iso;
      else if (dateCount === 1) inspection_date = iso;
      dateCount++;
      if (dateCount >= 2) break;
    }
  }

  // ── לכבוד → client / organization ────────────────────────────────────────
  let client = '';
  let organization = '';
  for (let i = 0; i < texts.length - 2; i++) {
    if (texts[i].includes('לכבוד')) {
      client = texts[i + 1] || '';
      organization = texts[i + 2] || '';
      break;
    }
  }

  // ── הנדון → subject ───────────────────────────────────────────────────────
  let subject = '';
  for (let i = 0; i < texts.length; i++) {
    if (texts[i].startsWith('הנדון')) {
      subject = texts[i].replace(/^הנדון\s*[:：]?\s*/, '').trim();
      if (!subject && i + 1 < texts.length) subject = texts[i + 1];
      break;
    }
  }

  // ── location (מיקום / כתובת) ──────────────────────────────────────────────
  let location = '';
  for (let i = 0; i < texts.length; i++) {
    if (/מיקום|כתובת|נכס/.test(texts[i])) {
      const inline = texts[i].replace(/^(מיקום|כתובת|נכס)\s*[:：]?\s*/, '').trim();
      location = inline || texts[i + 1] || '';
      break;
    }
  }

  // ── Section text blocks ───────────────────────────────────────────────────
  const introLines = [];
  const notesLines = [];
  const conclusionLines = [];

  let mode = null; // 'intro' | 'notes' | 'conclusion' | null

  for (let i = 0; i < texts.length; i++) {
    const t = texts[i];

    if (/נתונים|הקדמה|ממצאים/.test(t) && paragraphs[i].isHeading) {
      mode = 'intro';
      continue;
    }
    if (/הערות|הנחיות/.test(t) && paragraphs[i].isHeading) {
      mode = 'notes';
      continue;
    }
    if (/מסקנות/.test(t) && paragraphs[i].isHeading) {
      mode = 'conclusion';
      continue;
    }
    // Stop accumulating when hitting a new heading-like paragraph
    if (paragraphs[i].isHeading && t.length > 0 && mode !== null) {
      mode = null;
    }

    if (!t) continue; // skip blank lines inside blocks

    if (mode === 'intro') introLines.push(t);
    else if (mode === 'notes') notesLines.push(t);
    else if (mode === 'conclusion') conclusionLines.push(t);
  }

  // ── Table rows ────────────────────────────────────────────────────────────
  let table_rows = [];
  let defects_rows = [];

  for (const tbl of tables) {
    const headerStr = tbl.headers.join(' ');
    if (!defects_rows.length && /ליקוי|ליקויים/.test(headerStr)) {
      defects_rows = tbl.rows;
    } else if (!table_rows.length) {
      table_rows = tbl.rows;
    }
  }

  // ── doc_type ──────────────────────────────────────────────────────────────
  let doc_type = 'group1';
  if (/תקרות/.test(subject)) doc_type = 'group3';
  else if (/סככ/.test(subject)) doc_type = 'group5';
  else if (/שנתי|סקר/.test(subject)) doc_type = 'group4';

  return {
    doc_type,
    subject,
    client,
    organization,
    location,
    date,
    inspection_date,
    intro_text: introLines.join('\n'),
    notes_text: notesLines.join('\n'),
    conclusion_text: conclusionLines.join('\n'),
    table_rows,
    defects_rows,
  };
}
