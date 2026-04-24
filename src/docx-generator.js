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

module.exports = {
  generateComparisonDoc,
  injectComments,
};
