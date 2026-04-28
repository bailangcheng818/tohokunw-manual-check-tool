/**
 * docx-generator.js
 * Core template engine: JSON → .docx comparison table
 * Matches the 新旧比較表 A3 landscape format with Word comments
 *
 * Fixes v1.1:
 *  - A3 landscape: docx-js swaps w/h internally, so pass portrait dims (16838x23811)
 *  - Color/formatting: resolveColor now correctly passes hex to TextRun
 *  - Comments: robust anchor search with partial match fallback
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, PageOrientation, BorderStyle, WidthType, ShadingType,
  VerticalAlign, UnderlineType,
} = require('docx');

// ─── Page Layout ──────────────────────────────────────────────────────────────
// docx-js BUG: swaps width/height when orientation=LANDSCAPE
// So pass portrait values (16838 x 23811) → docx-js outputs w:w=23811 w:h=16838 ✓
const PAGE_W_INPUT = 16838;   // passed to docx-js (gets swapped to become w:w=23811)
const PAGE_H_INPUT = 23811;   // passed to docx-js (gets swapped to become w:h=16838)

// Actual content dimensions (after swap, real page is 23811 wide)
const CONTENT_W  = 23811 - 700 - 700;  // 22411 DXA usable width
const COL_OLD    = 10100;
const COL_NEW    = 10100;
const COL_NOTE   = CONTENT_W - COL_OLD - COL_NEW;  // 2211

// ─── Typography ───────────────────────────────────────────────────────────────
const FONT       = 'MS Mincho';
const SZ         = 19;        // 9.5pt
const RED        = 'FF0000';
const BLUE_UNDER = '4472C4';  // Word comment anchor blue

// ─── Border helpers ───────────────────────────────────────────────────────────
const mkBorder = (color = '888888', size = 4) =>
  ({ style: BorderStyle.SINGLE, size, color });
const BORDERS = {
  top: mkBorder(), bottom: mkBorder(), left: mkBorder(), right: mkBorder(),
};
const SUB_BORDERS = {
  top: mkBorder('AAAAAA', 2), bottom: mkBorder('AAAAAA', 2),
  left: mkBorder('AAAAAA', 2), right: mkBorder('AAAAAA', 2),
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Resolve color string to hex. Returns undefined for black (default).
 * FIX: was previously returning undefined for valid hex strings.
 */
function resolveColor(c) {
  if (!c || c === 'black') return undefined;
  if (c === 'red')         return RED;
  // Raw 6-char hex (e.g. "FF0000")
  if (/^[0-9A-Fa-f]{6}$/.test(c)) return c.toUpperCase();
  return undefined;
}

/**
 * Build a TextRun with full formatting support.
 */
function mkRun(text, opts = {}) {
  const cfg = {
    text,
    font: FONT,
    size: opts.sz || SZ,
    bold: opts.bold || false,
  };
  const color = resolveColor(opts.color);
  if (color)          cfg.color     = color;
  if (opts.underline) cfg.underline = { type: UnderlineType.SINGLE, color: BLUE_UNDER };
  if (opts.strike)    cfg.strike    = true;
  return new TextRun(cfg);
}

function mkPara(runs, align) {
  if (!Array.isArray(runs)) runs = [runs];
  return new Paragraph({
    alignment: align || AlignmentType.JUSTIFY,
    spacing:   { before: 0, after: 40, line: 260, lineRule: 'auto' },
    children:  runs,
  });
}

function emptyPara() {
  return new Paragraph({
    children: [new TextRun({ text: '', size: SZ, font: FONT })],
    spacing:  { before: 0, after: 40 },
  });
}

function mainCell(children, colW, shadeColor) {
  return new TableCell({
    borders:       BORDERS,
    verticalAlign: VerticalAlign.TOP,
    width:         { size: colW, type: WidthType.DXA },
    margins:       { top: 100, bottom: 100, left: 150, right: 150 },
    shading: shadeColor
      ? { fill: shadeColor, type: ShadingType.CLEAR, color: 'auto' }
      : undefined,
    children: Array.isArray(children) ? children : [children],
  });
}

function subCell(paragraphs, w) {
  return new TableCell({
    borders:       SUB_BORDERS,
    verticalAlign: VerticalAlign.TOP,
    width:         { size: w, type: WidthType.DXA },
    margins:       { top: 60, bottom: 60, left: 80, right: 80 },
    children:      Array.isArray(paragraphs) ? paragraphs : [paragraphs],
  });
}

function headerRow() {
  return new TableRow({
    children: [
      mainCell(mkPara([mkRun('旧',   { bold: true, sz: SZ + 1 })], AlignmentType.CENTER), COL_OLD,  'D9D9D9'),
      mainCell(mkPara([mkRun('新',   { bold: true, sz: SZ + 1 })], AlignmentType.CENTER), COL_NEW,  'D9D9D9'),
      mainCell(mkPara([mkRun('備考', { bold: true, sz: SZ + 1 })], AlignmentType.CENTER), COL_NOTE, 'D9D9D9'),
    ],
  });
}

// ─── Paragraph spec → Paragraph ───────────────────────────────────────────────
/**
 * Convert one paragraph spec from the JSON input into a docx Paragraph.
 *
 * Spec:
 *   { text, bold?, color?, underline?, align?, indent?, sz? }
 *   { segments: [{text, bold?, color?, underline?}], align?, indent?, sz? }
 */
function specToPara(spec) {
  if (!spec) return emptyPara();

  const sz    = spec.sz || SZ;
  const align = ({
    left:    AlignmentType.LEFT,
    center:  AlignmentType.CENTER,
    justify: AlignmentType.JUSTIFY,
    right:   AlignmentType.RIGHT,
  })[spec.align] || AlignmentType.JUSTIFY;

  let runs;
  if (spec.segments && spec.segments.length > 0) {
    runs = spec.segments.map(seg =>
      mkRun(seg.text || '', {
        bold:      seg.bold      || false,
        color:     seg.color,
        underline: seg.underline || false,
        sz,
      })
    );
  } else {
    runs = [mkRun(spec.text || '', {
      bold:      spec.bold      || false,
      color:     spec.color,
      underline: spec.underline || false,
      sz,
    })];
  }

  return new Paragraph({
    alignment: align,
    spacing:   { before: 0, after: 40, line: 260, lineRule: 'auto' },
    indent:    spec.indent ? { left: spec.indent } : undefined,
    children:  runs,
  });
}

function specsToParagraphs(specs) {
  if (!specs || specs.length === 0) return [emptyPara()];
  return specs.map(specToPara);
}

// ─── Inner history tables (render inside comparison column, width = COL_OLD) ──
// type(700) + date(1500) + reason(6700) + note(1200) = 10100
const HIST_INNER_COLS = [700, 1500, 6700, 1200];

const HIST_BORDERS = {
  top: mkBorder(), bottom: mkBorder(), left: mkBorder(), right: mkBorder(),
};

function histCell(children, w, shade) {
  return new TableCell({
    borders:       HIST_BORDERS,
    verticalAlign: VerticalAlign.TOP,
    width:         { size: w, type: WidthType.DXA },
    margins:       { top: 50, bottom: 50, left: 80, right: 80 },
    shading: shade ? { fill: shade, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    children: Array.isArray(children) ? children : [children],
  });
}

function histCellPara(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing:   { before: 0, after: 30, line: 240, lineRule: 'auto' },
    children:  [mkRun(text, { bold: opts.bold, color: opts.color })],
  });
}

function buildInnerEstablishedTable(historyData) {
  const headerRow = new TableRow({
    children: [
      histCell(histCellPara('',             { bold: true }),                                            HIST_INNER_COLS[0], 'D9D9D9'),
      histCell(histCellPara('制定・廃止年月日', { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[1], 'D9D9D9'),
      histCell(histCellPara('主な理由',       { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[2], 'D9D9D9'),
      histCell(histCellPara('備考',           { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[3], 'D9D9D9'),
    ],
  });

  const rows = [headerRow];
  for (const e of (historyData.established || [])) {
    rows.push(new TableRow({
      children: [
        histCell(histCellPara('制定', { align: AlignmentType.CENTER, color: e.color }), HIST_INNER_COLS[0]),
        histCell(histCellPara(e.date   || '', { color: e.color }), HIST_INNER_COLS[1]),
        histCell(histCellPara(e.reason || '', { color: e.color }), HIST_INNER_COLS[2]),
        histCell(histCellPara(e.note   || '', { color: e.color }), HIST_INNER_COLS[3]),
      ],
    }));
  }
  // Always add 廃止 row
  rows.push(new TableRow({
    children: [
      histCell(histCellPara('廃止', { align: AlignmentType.CENTER }), HIST_INNER_COLS[0]),
      ...HIST_INNER_COLS.slice(1).map(w => histCell(histCellPara(''), w)),
    ],
  }));

  return new Table({
    width:        { size: HIST_INNER_COLS.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: HIST_INNER_COLS,
    rows,
  });
}

function buildInnerRevisedTable(historyData) {
  const REVISED_MIN_ROWS = 8;

  const headerRow = new TableRow({
    children: [
      histCell(histCellPara('番号',    { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[0], 'D9D9D9'),
      histCell(histCellPara('改正年月日', { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[1], 'D9D9D9'),
      histCell(histCellPara('主な理由', { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[2], 'D9D9D9'),
      histCell(histCellPara('備考',    { bold: true, align: AlignmentType.CENTER }), HIST_INNER_COLS[3], 'D9D9D9'),
    ],
  });

  const rows = [headerRow];
  for (const e of (historyData.revised || [])) {
    rows.push(new TableRow({
      children: [
        histCell(histCellPara(e.number || '', { align: AlignmentType.CENTER, color: e.color }), HIST_INNER_COLS[0]),
        histCell(histCellPara(e.date   || '', { color: e.color }), HIST_INNER_COLS[1]),
        histCell(histCellPara(e.reason || '', { color: e.color }), HIST_INNER_COLS[2]),
        histCell(histCellPara(e.note   || '', { color: e.color }), HIST_INNER_COLS[3]),
      ],
    }));
  }
  const needed = Math.max(0, REVISED_MIN_ROWS - (historyData.revised || []).length);
  for (let i = 0; i < needed; i++) {
    rows.push(new TableRow({ children: HIST_INNER_COLS.map(w => histCell(histCellPara(''), w)) }));
  }

  return new Table({
    width:        { size: HIST_INNER_COLS.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: HIST_INNER_COLS,
    rows,
  });
}

function buildHistoryCellContent(historyData) {
  const title = (historyData && historyData.doc_title) || '';
  const data  = historyData || {};
  return [
    emptyPara(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 160, line: 260, lineRule: 'auto' },
      children:  [new TextRun({ text: '制定改廃経歴表', font: FONT, size: 26, bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:   { before: 0, after: 100 },
      border:    { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA', space: 1 } },
      children:  [mkRun('仕様書名：', { bold: true }), mkRun(title, { underline: true })],
    }),
    emptyPara(),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:   { before: 60, after: 50 },
      children:  [mkRun('制定・廃止', { bold: true })],
    }),
    buildInnerEstablishedTable(data),
    emptyPara(),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:   { before: 60, after: 50 },
      children:  [mkRun('改正', { bold: true })],
    }),
    buildInnerRevisedTable(data),
  ];
}

// ─── Section table ────────────────────────────────────────────────────────────
function buildSectionTable(section) {
  let oldContent, newContent;

  if (section.old_history || section.new_history) {
    // History mode: render nested 制定改廃経歴表 tables in each column
    oldContent = buildHistoryCellContent(section.old_history);
    newContent = buildHistoryCellContent(section.new_history);
  } else {
    oldContent = specsToParagraphs(section.old_paragraphs);
    newContent = specsToParagraphs(section.new_paragraphs);
  }

  const noteParas = (section.notes && section.notes.length > 0)
    ? section.notes.map(n => mkPara([mkRun(n, { color: 'red' })]))
    : [emptyPara()];

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [COL_OLD, COL_NEW, COL_NOTE],
    rows: [
      headerRow(),
      new TableRow({
        children: [
          mainCell(oldContent, COL_OLD),
          mainCell(newContent, COL_NEW),
          mainCell(noteParas,  COL_NOTE),
        ],
      }),
    ],
  });
}

// ─── XML helpers (comment injection) ─────────────────────────────────────────

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildCommentsXml(comments) {
  if (!comments || comments.length === 0) return null;
  const date = '2025-01-01T00:00:00Z';
  const els = comments.map((c, i) => `
  <w:comment w:id="${i}" w:author="Editor" w:date="${date}" w:initials="E">
    <w:p>
      <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>
      <w:r>
        <w:rPr><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
        <w:t>${escapeXml(c.text)}</w:t>
      </w:r>
    </w:p>
  </w:comment>`).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  mc:Ignorable="w14 w15">
${els}
</w:comments>`;
}

/**
 * Inject comment markers around anchor text in document XML.
 * FIX: tries both escaped and unescaped variants, and searches within <w:t> tags.
 */
function injectCommentMarkers(docXml, comments) {
  if (!comments || comments.length === 0) return docXml;

  let xml = docXml;

  for (const { anchor, id } of comments) {
    if (!anchor) continue;

    // Try to find the text in <w:t>...</w:t>
    // Anchors may appear escaped or unescaped in the XML
    const candidates = [
      escapeXml(anchor),
      anchor,
    ];

    let found = false;
    for (const needle of candidates) {
      // Look for the text inside a <w:t> tag
      const tOpen  = `>${needle}<`;
      const tIdx   = xml.indexOf(tOpen);
      if (tIdx === -1) continue;

      // Walk back to find the opening <w:r> of this run
      const rStart = xml.lastIndexOf('<w:r>', tIdx);
      if (rStart === -1) continue;

      // Find the closing </w:r>
      const rEnd = xml.indexOf('</w:r>', tIdx);
      if (rEnd === -1) continue;
      const rEndFull = rEnd + '</w:r>'.length;

      const run = xml.slice(rStart, rEndFull);
      const wrapped =
        `<w:commentRangeStart w:id="${id}"/>` +
        run +
        `<w:commentRangeEnd w:id="${id}"/>` +
        `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
        `<w:commentReference w:id="${id}"/></w:r>`;

      xml    = xml.slice(0, rStart) + wrapped + xml.slice(rEndFull);
      found  = true;
      break;
    }

    if (!found) {
      console.warn(`[docx-mcp] Comment anchor not found in XML: "${anchor}"`);
    }
  }

  return xml;
}

// ─── Comment injection via JSZip ──────────────────────────────────────────────
async function injectComments(buffer, comments) {
  let JSZip;
  try { JSZip = require('jszip'); }
  catch {
    console.warn('[docx-mcp] jszip not available – comments skipped');
    return buffer;
  }

  const zip = await JSZip.loadAsync(buffer);

  // Patch document.xml
  let docXml = await zip.file('word/document.xml').async('string');
  docXml = injectCommentMarkers(docXml, comments);
  zip.file('word/document.xml', docXml);

  // Write comments.xml
  zip.file('word/comments.xml', buildCommentsXml(comments));

  // Patch _rels
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  if (!relsXml.includes('relationships/comments')) {
    relsXml = relsXml.replace('</Relationships>',
      `  <Relationship Id="rId_cmt" ` +
      `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" ` +
      `Target="comments.xml"/>\n</Relationships>`);
    zip.file('word/_rels/document.xml.rels', relsXml);
  }

  // Patch [Content_Types].xml
  let typesXml = await zip.file('[Content_Types].xml').async('string');
  if (!typesXml.includes('comments+xml')) {
    typesXml = typesXml.replace('</Types>',
      `  <Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml" ` +
      `PartName="/word/comments.xml"/>\n</Types>`);
    zip.file('[Content_Types].xml', typesXml);
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ─── Main export ──────────────────────────────────────────────────────────────
async function generateComparisonDoc(spec, outputPath) {
  // Build body
  const bodyChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 200 },
      children:  [new TextRun({ text: spec.doc_title || '新旧比較表', font: FONT, size: 28, bold: true })],
    }),
  ];

  for (const section of (spec.sections || [])) {
    const statusLabel = { new: '　新規追加', unchanged: '　変更なし' }[section.status] || '　変更あり';
    bodyChildren.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:   { before: 0, after: 80 },
      children:  [new TextRun({ text: section.title + statusLabel, font: FONT, size: SZ, bold: true })],
    }));
    bodyChildren.push(buildSectionTable(section));
    bodyChildren.push(emptyPara());
  }

  // Assemble doc — NOTE: swap width/height to work around docx-js landscape bug
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: SZ } } } },
    sections: [{
      properties: {
        page: {
          size: {
            width:       PAGE_W_INPUT,   // 16838 → docx-js outputs w:w=23811 ✓
            height:      PAGE_H_INPUT,   // 23811 → docx-js outputs w:h=16838 ✓
            orientation: PageOrientation.LANDSCAPE,
          },
          margin: { top: 850, bottom: 500, left: 700, right: 700 },
        },
      },
      children: bodyChildren,
    }],
  });

  let buffer = await Packer.toBuffer(doc);

  // Collect all comments
  const allComments = [];
  for (const section of (spec.sections || [])) {
    for (const c of (section.comments || [])) {
      allComments.push({ ...c, id: allComments.length });
    }
  }
  if (allComments.length > 0) {
    buffer = await injectComments(buffer, allComments);
  }

  const result = { buffer, base64: buffer.toString('base64') };

  if (outputPath) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buffer);
    result.path = outputPath;
  }

  return result;
}

// ─── Edit-based comparison doc (新旧比較表 from Dify edit format) ─────────────

const { readScheme, readOriginalBuffer } = require('./file-store');

const SECTION_HEADING_TYPES = new Set(['numbered_heading', 'bracket_heading']);

// ─── XML helpers for comparison table ────────────────────────────────────────

const TC_BORDER = `w:val="single" w:sz="4" w:color="888888" w:space="0"`;
const TC_BORDERS_XML = `<w:tcBorders>
  <w:top ${TC_BORDER}/><w:left ${TC_BORDER}/>
  <w:bottom ${TC_BORDER}/><w:right ${TC_BORDER}/>
</w:tcBorders>`;

/**
 * Find the next occurrence of <tagName> or <tagName (space) starting at pos.
 * Returns -1 if not found. Does not match longer tag names (e.g. <w:pPr>).
 */
function findNextOpeningTag(text, pos, tagName) {
  const n1 = `<${tagName}>`;
  const n2 = `<${tagName} `;
  let i = pos;
  while (i < text.length) {
    const i1 = text.indexOf(n1, i);
    const i2 = text.indexOf(n2, i);
    if (i1 === -1 && i2 === -1) return -1;
    if (i1 === -1) return i2;
    if (i2 === -1) return i1;
    return Math.min(i1, i2);
  }
  return -1;
}

/**
 * Extract a balanced <tagName>...</tagName> block starting at `start`.
 * Handles nested occurrences of the same tag.
 * Returns { xml, endPos } where endPos points to just after the closing tag.
 */
function extractBalancedXmlTag(text, start, tagName) {
  const openTag  = `<${tagName}`;
  const closeTag = `</${tagName}>`;
  let depth = 1;
  let pos   = text.indexOf('>', start) + 1; // skip past the opening tag's >
  while (pos < text.length && depth > 0) {
    const nextOpen  = findNextOpeningTag(text, pos, tagName);
    const nextClose = text.indexOf(closeTag, pos);
    if (nextClose === -1) return { xml: null, endPos: text.length };
    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++;
      pos = text.indexOf('>', nextOpen) + 1;
    } else {
      depth--;
      pos = nextClose + closeTag.length;
    }
  }
  if (depth !== 0) return { xml: null, endPos: text.length };
  return { xml: text.slice(start, pos), endPos: pos };
}

/**
 * Extract top-level body children (paragraphs and tables) in document order.
 * Paragraphs are extracted the same way as the ingest regex (non-greedy: stop
 * at the first </w:p>), so their flatIdx aligns with scheme.json xml_index values.
 * Tables are extracted as balanced <w:tbl>…</w:tbl> blocks verbatim.
 *
 * Returns [{ type: 'p', flatIdx, xml } | { type: 'tbl', xml }]
 */
function extractTopLevelBodyChildren(docXml) {
  const bodyStartIdx = docXml.indexOf('<w:body>');
  const bodyEndIdx   = docXml.lastIndexOf('</w:body>');
  if (bodyStartIdx === -1) return [];

  // Count paragraphs before <w:body> to establish flatIdx offset.
  // Standard Word documents have none, but be safe.
  const preBody  = docXml.slice(0, bodyStartIdx);
  let flatIdx    = (preBody.match(/<w:p[ >]/g) || []).length;

  const body     = docXml.slice(bodyStartIdx + '<w:body>'.length, bodyEndIdx);
  const children = [];
  let pos        = 0;

  while (pos < body.length) {
    const pPos   = findNextOpeningTag(body, pos, 'w:p');
    const tblPos = findNextOpeningTag(body, pos, 'w:tbl');

    if (pPos === -1 && tblPos === -1) break;

    const isTable = tblPos !== -1 && (pPos === -1 || tblPos < pPos);
    const start   = isTable ? tblPos : pPos;

    if (!isTable) {
      // Same extraction as the ingest regex: stop at the first </w:p>
      const endIdx = body.indexOf('</w:p>', start);
      if (endIdx === -1) break;
      const xml = body.slice(start, endIdx + 6);
      children.push({ type: 'p', flatIdx, xml });
      flatIdx++;             // one flat-para entry per non-greedy paragraph match
      pos = endIdx + 6;
    } else {
      const { xml, endPos } = extractBalancedXmlTag(body, start, 'w:tbl');
      if (!xml) break;
      // Count how many flat-para entries the table occupies (same regex logic)
      const tblParaCount = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).length;
      children.push({ type: 'tbl', xml, flatIdxStart: flatIdx, flatIdxEnd: flatIdx + tblParaCount });
      flatIdx += tblParaCount;
      pos = endPos;
    }
  }

  return children;
}

/**
 * Group top-level body children into sections keyed by heading paragraphs.
 * schemeParaMap: Map<xml_index, schemePara> — used to detect heading types
 * and to provide text/edit fallback for paragraphs.
 *
 * Returns [{ heading: bodyChild|null, items: bodyChild[] }]
 */
function groupBodyChildrenIntoSections(bodyChildren, schemeParaMap) {
  const sections = [];
  let current = { heading: null, items: [] };

  for (const child of bodyChildren) {
    if (child.type === 'p') {
      const schemePara = schemeParaMap.get(child.flatIdx);
      const isHeading  = schemePara && SECTION_HEADING_TYPES.has(schemePara.type);
      const enriched   = { ...child, schemePara: schemePara || null };
      if (isHeading) {
        if (current.heading !== null || current.items.length > 0) sections.push(current);
        current = { heading: enriched, items: [] };
      } else {
        current.items.push(enriched);
      }
    } else {
      current.items.push(child); // tables go straight into the current section
    }
  }

  if (current.heading !== null || current.items.length > 0) sections.push(current);
  return sections;
}

/**
 * Add <w:strike/> to every <w:rPr> block in a paragraph XML string.
 * Handles both pPr/rPr (paragraph mark) and r/rPr (run text).
 * Also adds a bare <w:rPr><w:strike/></w:rPr> for runs that have none.
 */
function addStrikethroughToParaXml(paraXml) {
  // Step 1: add to existing <w:rPr> blocks
  let result = paraXml.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/g,
    (m, open, inner, close) => inner.includes('<w:strike') ? m : `${open}${inner}<w:strike/>${close}`
  );
  // Step 2: for bare <w:r...> immediately followed by <w:t (no rPr present)
  result = result.replace(/(<w:r(?:\s[^>]*)?>)\s*(<w:t)/g,
    (m, openR, openT) => `${openR}<w:rPr><w:strike/></w:rPr>${openT}`
  );
  return result;
}

/** Extract <w:pPr>…</w:pPr> from a paragraph XML string, or '' if absent. */
function extractPPrXml(paraXml) {
  if (!paraXml) return '';
  const m = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : '';
}

/**
 * Take a <w:pPr> block and ensure its <w:rPr> child has a red color marker
 * (for the paragraph mark, so Word renders the paragraph mark in red too).
 * If pPrXml is empty, returns a minimal red-pPr block.
 */
function injectRedIntoPPrRpr(pPrXml) {
  if (!pPrXml) return '<w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>';
  if (pPrXml.includes('<w:rPr>')) {
    return pPrXml.replace(/(<w:rPr>)([\s\S]*?)(<\/w:rPr>)/, (m, open, inner, close) => {
      const noColor = inner.replace(/<w:color[^/]*\/>/g, '');
      return `${open}${noColor}<w:color w:val="FF0000"/>${close}`;
    });
  }
  return pPrXml.replace('</w:pPr>', '<w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>');
}

/**
 * Build a red-text paragraph for the 新 column of changed rows.
 * origParaXml: optional — when provided, <w:pPr> (indentation, spacing, tab stops)
 * is copied from the original paragraph so layout is preserved.
 */
function buildRedTextParaXml(text, origParaXml) {
  const pPr = injectRedIntoPPrRpr(extractPPrXml(origParaXml));
  const t   = escapeXml(text);
  return `<w:p>${pPr}` +
    `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}"/>` +
    `<w:color w:val="FF0000"/><w:sz w:val="${SZ}"/><w:szCs w:val="${SZ}"/></w:rPr>` +
    `<w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
}

/** Fallback: plain paragraph XML when the original XML is unavailable. */
function buildPlainParaXml(text) {
  return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

/**
 * Fallback for paragraphs that are "unsafe" (drawing-containing, truncated XML).
 * Recovers <w:pPr> from the truncated rawXml (properties precede drawing content)
 * and pairs it with the scheme text — preserving indentation/spacing even when
 * the drawing itself is lost.
 */
function buildFallbackParaXml(rawXml, text) {
  const pPr = extractPPrXml(rawXml);
  const t   = escapeXml(text || '');
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
}

/**
 * Strip <w:drawing> and <mc:AlternateContent> blocks from paragraph XML.
 * These contain nested <w:p> inside <w:txbxContent>, which would produce
 * unclosed tags when the paragraph is pasted into a table cell.
 */
function stripDrawingsFromParaXml(xml) {
  const patterns = [
    { open: '<w:drawing', close: '</w:drawing>' },
    { open: '<mc:AlternateContent', close: '</mc:AlternateContent>' },
  ];
  let result = xml;
  for (const { open, close } of patterns) {
    let out = '';
    let pos = 0;
    while (pos < result.length) {
      const s = result.indexOf(open, pos);
      if (s === -1) { out += result.slice(pos); break; }
      out += result.slice(pos, s);
      let depth = 1, search = s + open.length;
      while (depth > 0 && search < result.length) {
        const no = result.indexOf(open, search);
        const nc = result.indexOf(close, search);
        if (nc === -1) { search = result.length; break; }
        if (no !== -1 && no < nc) { depth++; search = no + open.length; }
        else { depth--; search = nc + close.length; }
      }
      pos = search;
    }
    result = out;
  }
  return result;
}

/**
 * Rebuild table XML replacing specific paragraph slots (identified by their
 * sequential index in the non-greedy para scan) with patched versions.
 * patchMap: Map<localParaIdx, newParaXml>
 */
function patchTableParas(tblXml, patchMap) {
  if (!patchMap || patchMap.size === 0) return tblXml;
  let result   = '';
  let pos      = 0;
  let localIdx = 0;
  while (pos < tblXml.length) {
    const pStart = findNextOpeningTag(tblXml, pos, 'w:p');
    if (pStart === -1) { result += tblXml.slice(pos); break; }
    result += tblXml.slice(pos, pStart);
    const endIdx = tblXml.indexOf('</w:p>', pStart);
    if (endIdx === -1) { result += tblXml.slice(pStart); break; }
    const paraEnd = endIdx + 6;
    result += patchMap.has(localIdx) ? patchMap.get(localIdx) : tblXml.slice(pStart, paraEnd);
    localIdx++;
    pos = paraEnd;
  }
  return result;
}

/**
 * Build a <w:tc> from an array of items.
 * Each item may be:
 *   - a string (paragraph XML) → drawings are stripped before insertion
 *   - { type: 'p', xml }       → same as string
 *   - { type: 'tbl', xml }     → nested table, included verbatim (no drawing strip)
 */
function buildTcXml(items, colWidth, shadeHex) {
  const shade = shadeHex
    ? `<w:shd w:val="clear" w:color="auto" w:fill="${shadeHex}"/>`
    : '';
  const content = (items || []).map(item => {
    const xml  = typeof item === 'string' ? item : item.xml;
    const type = typeof item === 'string' ? 'p'  : item.type;
    return type === 'tbl' ? xml : stripDrawingsFromParaXml(xml);
  }).join('') || '<w:p><w:pPr/></w:p>';
  return `<w:tc><w:tcPr>` +
    `<w:tcW w:w="${colWidth}" w:type="dxa"/>` +
    TC_BORDERS_XML +
    shade +
    `<w:tcMar>` +
    `<w:top w:w="100" w:type="dxa"/><w:left w:w="150" w:type="dxa"/>` +
    `<w:bottom w:w="100" w:type="dxa"/><w:right w:w="150" w:type="dxa"/>` +
    `</w:tcMar></w:tcPr>${content}</w:tc>`;
}

/** Build a header cell with centered bold text and grey shading. */
function buildHeaderTcXml(title, colWidth) {
  const t   = escapeXml(title);
  const sz  = SZ + 1;
  const para = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}"/>` +
    `<w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
    `<w:t>${t}</w:t></w:r></w:p>`;
  return buildTcXml([para], colWidth, 'D9D9D9');
}

/** Build the repeating header row XML (<w:tblHeader/>). */
function buildHeaderRowXml(oldTitle, newTitle) {
  return `<w:tr><w:trPr><w:tblHeader/></w:trPr>` +
    buildHeaderTcXml(oldTitle, COL_OLD) +
    buildHeaderTcXml(newTitle, COL_NEW) +
    buildHeaderTcXml('備考', COL_NOTE) +
    `</w:tr>`;
}

/**
 * Return true if the extracted paragraph XML has balanced drawing/textbox tags.
 * Paragraphs containing drawings are extracted non-greedily (up to the first
 * inner </w:p>), leaving outer tags unclosed. Such XML must not be placed
 * verbatim into a table cell.
 */
function isParaXmlSafe(xml) {
  if (!xml) return false;
  const pairs = [
    ['<mc:AlternateContent', '</mc:AlternateContent>'],
    ['<w:drawing',           '</w:drawing>'          ],
    ['<w:txbxContent',       '</w:txbxContent>'      ],
    ['<wps:txbx',            '</wps:txbx>'           ],
    ['<v:textbox',           '</v:textbox>'          ],
  ];
  for (const [open, close] of pairs) {
    const opens  = (xml.split(open).length  - 1);
    const closes = (xml.split(close).length - 1);
    if (opens !== closes) return false;
  }
  return true;
}

/**
 * Build one comparison data row for a section.
 * section.heading: body paragraph child (or null)
 * section.items:   array of body children — each is {type:'p', flatIdx, xml, schemePara}
 *                  or {type:'tbl', xml}
 * allParaXmls:     full flat array of all <w:p> XML strings from document (xml_index aligned)
 */
function buildSectionDataRowXml(section, allParaXmls, editsByXmlIndex, editsByOldText) {
  const allItems = section.heading ? [section.heading, ...section.items] : section.items;
  const oldItems  = [];
  const newItems  = [];
  const bikoItems = [];

  for (const item of allItems) {
    if (item.type === 'tbl') {
      const { xml, flatIdxStart, flatIdxEnd } = item;
      // Collect edits whose xml_index falls inside this table's paragraph range
      const tableEdits = new Map(); // localIdx → edit
      if (flatIdxStart != null && flatIdxEnd != null) {
        for (const [xmlIdx, edit] of editsByXmlIndex) {
          if (xmlIdx >= flatIdxStart && xmlIdx < flatIdxEnd) {
            tableEdits.set(xmlIdx - flatIdxStart, edit);
          }
        }
      }
      if (tableEdits.size === 0) {
        oldItems.push(item);
        newItems.push(item);
      } else {
        const oldPatch = new Map();
        const newPatch = new Map();
        for (const [localIdx, edit] of tableEdits) {
          const globalIdx    = flatIdxStart + localIdx;
          const rawXml       = allParaXmls[globalIdx];
          const origXml      = isParaXmlSafe(rawXml) ? rawXml : null;
          const fallbackText = edit.old_text || '';
          oldPatch.set(localIdx, origXml
            ? addStrikethroughToParaXml(origXml)
            : buildFallbackParaXml(rawXml, fallbackText));
          newPatch.set(localIdx, buildRedTextParaXml(edit.new_text || '', rawXml));
          if (edit.rationale) bikoItems.push(buildPlainParaXml(edit.rationale));
        }
        oldItems.push({ type: 'tbl', xml: patchTableParas(xml, oldPatch) });
        newItems.push({ type: 'tbl', xml: patchTableParas(xml, newPatch) });
      }
      continue;
    }

    // Paragraph item
    const flatIdx    = item.flatIdx;
    const schemePara = item.schemePara;
    const edit       = editsByXmlIndex.get(flatIdx)
      || (schemePara?.text ? editsByOldText.get(schemePara.text.trim()) : undefined);
    const rawXml  = allParaXmls[flatIdx];
    const origXml = isParaXmlSafe(rawXml) ? rawXml : null;
    const fallbackText = schemePara?.text || '';

    if (edit) {
      // 旧 column: original XML with strikethrough, or pPr-preserving fallback
      oldItems.push(origXml
        ? addStrikethroughToParaXml(origXml)
        : buildFallbackParaXml(rawXml, fallbackText));
      // 新 column: red text, inheriting pPr from rawXml (Fix D)
      newItems.push(buildRedTextParaXml(edit.new_text || '', rawXml));
      if (edit.rationale) bikoItems.push(buildPlainParaXml(edit.rationale));
    } else {
      // Unchanged: use original XML or pPr-preserving fallback (Fix C)
      const xml = origXml || buildFallbackParaXml(rawXml, fallbackText);
      oldItems.push(xml);
      newItems.push(xml);
    }
  }

  return `<w:tr>` +
    buildTcXml(oldItems,  COL_OLD) +
    buildTcXml(newItems,  COL_NEW) +
    buildTcXml(bikoItems, COL_NOTE) +
    `</w:tr>`;
}

/** Assemble the full <w:tbl> XML for the comparison table. */
function buildFullComparisonTableXml(sections, allParaXmls, editsByXmlIndex, editsByOldText, filename) {
  const insideBorder = `<w:insideH ${TC_BORDER}/><w:insideV ${TC_BORDER}/>`;
  const tblPr = `<w:tblPr>` +
    `<w:tblW w:w="${CONTENT_W}" w:type="dxa"/>` +
    `<w:tblBorders><w:top ${TC_BORDER}/><w:left ${TC_BORDER}/>` +
    `<w:bottom ${TC_BORDER}/><w:right ${TC_BORDER}/>${insideBorder}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/>` +
    `</w:tblPr>`;
  const tblGrid = `<w:tblGrid>` +
    `<w:gridCol w:w="${COL_OLD}"/><w:gridCol w:w="${COL_NEW}"/><w:gridCol w:w="${COL_NOTE}"/>` +
    `</w:tblGrid>`;
  const header   = buildHeaderRowXml(`${filename}（旧）`, `${filename}（新）`);
  const dataRows = sections
    .map(s => buildSectionDataRowXml(s, allParaXmls, editsByXmlIndex, editsByOldText))
    .join('');
  return `<w:tbl>${tblPr}${tblGrid}${header}${dataRows}</w:tbl>`;
}

// ─── Main export ──────────────────────────────────────────────────────────────

/** A3 landscape sectPr — applied to the whole comparison document. */
const SECT_PR_A3_LANDSCAPE =
  `<w:sectPr>` +
  `<w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/>` +
  `<w:pgMar w:top="850" w:right="700" w:bottom="500" w:left="700" ` +
  `w:header="0" w:footer="0" w:gutter="0"/>` +
  `</w:sectPr>`;

/** Build cover page XML (pure XML strings, no docx-js). */
function buildCoverPageXml(filename, date) {
  const f = escapeXml(filename);
  const d = escapeXml(date);
  const mkPara = (text, jc, bold, sz) => {
    const jcXml = jc !== 'left' ? `<w:jc w:val="${jc}"/>` : '';
    const bXml  = bold ? '<w:b/>' : '';
    return `<w:p><w:pPr>${jcXml}</w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}"/>` +
      `${bXml}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
      `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  };
  return [
    '<w:p/>',
    mkPara(f + '（改正）', 'center', true, SZ),
    mkPara('新旧比較表',    'center', true, 32),
    mkPara(`作成日：${d}`,  'left',   false, SZ),
    '<w:p/>',
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
  ].join('');
}

/**
 * Generate a 新旧比較表 DOCX from a Dify workflow output payload.
 *
 * Strategy:
 *  1. Load original DOCX as the base ZIP (preserves all metadata, theme,
 *     styles, numbering, relationships — nothing is missing).
 *  2. Extract all <w:p> XML strings from the original document body.
 *  3. Build comparison table XML with original paragraphs verbatim (unchanged)
 *     or modified (strikethrough/red) for edited paragraphs.
 *  4. Replace <w:body> in original document.xml with: cover page + table + A3 sectPr.
 *  5. Write modified ZIP as output.
 *
 * @param {object} payload - { filename, date, amendment_edits (JSON string), ... }
 * @param {string} [outputPath]
 */
async function generateEditComparisonDoc(payload, outputPath) {
  let JSZip;
  try { JSZip = require('jszip'); } catch { throw new Error('jszip not available'); }

  // 1. Parse edits
  const rawEdits = typeof payload.amendment_edits === 'string'
    ? JSON.parse(payload.amendment_edits)
    : (payload.amendment_edits || []);

  const ref_id   = rawEdits[0]?.ref_id;
  const filename = payload.filename || '新旧比較表';
  const date     = payload.date     || '';

  // 2. Read scheme (for section structure) and original DOCX (base ZIP)
  const scheme     = readScheme(ref_id);
  const paragraphs = scheme.paragraphs || [];
  const origBuffer = readOriginalBuffer(ref_id);
  const origZip    = await JSZip.loadAsync(origBuffer);
  const origDocXml = await origZip.file('word/document.xml').async('string');

  // 3. Fix A — build flat para array using same regex as ingest.js (aligns xml_index)
  const allParaXmls = origDocXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];

  // 4. Fix B — extract top-level body children (paragraphs + tables in document order)
  const bodyChildren = extractTopLevelBodyChildren(origDocXml);

  // 5. Build edit lookup maps
  const editsByXmlIndex = new Map();
  const editsByOldText  = new Map();
  for (const e of rawEdits) {
    if (e.xml_index != null) editsByXmlIndex.set(e.xml_index, e);
    if (e.old_text)          editsByOldText.set(e.old_text.trim(), e);
  }

  // 6. Build scheme paragraph map for heading detection and text fallback
  const schemeParaMap = new Map(paragraphs.map(p => [p.xml_index, p]));

  // 7. Group body children into sections, then build comparison table XML
  const sections = groupBodyChildrenIntoSections(bodyChildren, schemeParaMap);
  const tableXml = buildFullComparisonTableXml(
    sections, allParaXmls, editsByXmlIndex, editsByOldText, filename
  );

  // 8. Build new body: cover page + table + A3 landscape sectPr
  const coverXml = buildCoverPageXml(filename, date);
  const newBody  = `<w:body>${coverXml}${tableXml}<w:p/>${SECT_PR_A3_LANDSCAPE}</w:body>`;

  // 9. Replace only the <w:body>...</w:body> in original document.xml,
  //    preserving the root element with all its namespace declarations.
  const bodyStart = origDocXml.indexOf('<w:body');
  const afterBody = origDocXml.lastIndexOf('</w:body>') + '</w:body>'.length;
  const newDocXml = origDocXml.slice(0, bodyStart) + newBody + origDocXml.slice(afterBody);

  origZip.file('word/document.xml', newDocXml);

  const buffer = await origZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

  const result = { buffer, base64: buffer.toString('base64') };
  if (outputPath) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buffer);
    result.path = outputPath;
  }
  return result;
}

module.exports = {
  generateComparisonDoc,
  injectComments,
  generateEditComparisonDoc,
  addStrikethroughToParaXml,
  buildRedTextParaXml,
  isParaXmlSafe,
  buildFallbackParaXml,
};
