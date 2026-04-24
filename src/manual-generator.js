/**
 * manual-generator.js
 * Core template engine: JSON → .docx A4 纵向 仕様書マニュアル
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, UnderlineType, Header, Footer, PageNumber,
  PageBorderDisplay, PageBorderOffsetFrom,
} = require('docx');

const { injectComments } = require('./docx-generator');

// ─── Page Layout (A4 Portrait) ────────────────────────────────────────────────
const PAGE_W    = 11906;
const PAGE_H    = 16838;
const MARGIN_TOP = 1200;   // slightly more top margin to leave room for header
const MARGIN_BTM = 1000;
const MARGIN_LR  = 1100;
const CONTENT_W  = PAGE_W - MARGIN_LR * 2;  // 9706 DXA

// ─── Typography ───────────────────────────────────────────────────────────────
const FONT       = 'MS Mincho';
const SZ         = 19;       // 9.5pt
const SZ_HDR     = 18;       // header/footer smaller
const RED        = 'FF0000';
const BLUE_UNDER = '4472C4';

// ─── History table column widths ──────────────────────────────────────────────
// type(800) + date(1800) + reason(5906) + note(1200) = 9706
const HIST_COLS = [800, 1800, 5906, 1200];

// ─── Border helpers ───────────────────────────────────────────────────────────
const mkBorder = (color = '888888', size = 4) =>
  ({ style: BorderStyle.SINGLE, size, color });
const BORDERS = {
  top: mkBorder(), bottom: mkBorder(), left: mkBorder(), right: mkBorder(),
};
const THIN_BORDERS = {
  top:    mkBorder('AAAAAA', 2), bottom: mkBorder('AAAAAA', 2),
  left:   mkBorder('AAAAAA', 2), right:  mkBorder('AAAAAA', 2),
};

// ─── Core helpers ─────────────────────────────────────────────────────────────

function resolveColor(c) {
  if (!c || c === 'black') return undefined;
  if (c === 'red')         return RED;
  if (/^[0-9A-Fa-f]{6}$/.test(c)) return c.toUpperCase();
  return undefined;
}

function mkRun(text, opts = {}) {
  const cfg = {
    text,
    font:  FONT,
    size:  opts.sz || SZ,
    bold:  opts.bold  || false,
  };
  const color = resolveColor(opts.color);
  if (color)          cfg.color     = color;
  if (opts.underline) cfg.underline = { type: UnderlineType.SINGLE, color: BLUE_UNDER };
  return new TextRun(cfg);
}

function mkPara(runs, align, spacingAfter) {
  if (!Array.isArray(runs)) runs = [runs];
  return new Paragraph({
    alignment: align || AlignmentType.JUSTIFY,
    spacing:   { before: 0, after: spacingAfter !== undefined ? spacingAfter : 40, line: 260, lineRule: 'auto' },
    children:  runs,
  });
}

function emptyPara(sz) {
  return new Paragraph({
    children: [new TextRun({ text: '', size: sz || SZ, font: FONT })],
    spacing:  { before: 0, after: 40 },
  });
}

function spacerParas(n) {
  return Array.from({ length: n }, () => emptyPara());
}

// ─── Paragraph spec → docx Paragraph ─────────────────────────────────────────
function specToPara(spec) {
  if (!spec) return emptyPara();

  const sz    = spec.sz || SZ;
  const align = ({
    left:    AlignmentType.LEFT,
    center:  AlignmentType.CENTER,
    justify: AlignmentType.JUSTIFY,
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

// ─── Table cell helpers ───────────────────────────────────────────────────────
function mkCell(children, colW, opts = {}) {
  return new TableCell({
    borders:       opts.thin ? THIN_BORDERS : BORDERS,
    verticalAlign: VerticalAlign.TOP,
    width:         { size: colW, type: WidthType.DXA },
    margins:       { top: 80, bottom: 80, left: 120, right: 120 },
    shading: opts.shade
      ? { fill: opts.shade, type: ShadingType.CLEAR, color: 'auto' }
      : undefined,
    children: Array.isArray(children) ? children : [children],
  });
}

function mkCellPara(text, opts = {}) {
  return mkPara(
    [mkRun(text, { bold: opts.bold, sz: opts.sz || SZ, color: opts.color })],
    opts.align || AlignmentType.LEFT,
    40,
  );
}

// ─── Page header & footer ─────────────────────────────────────────────────────
function makeHeader(title) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border:    { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
        spacing:   { before: 0, after: 120 },
        children:  [new TextRun({ text: title, font: FONT, size: SZ_HDR, bold: true })],
      }),
    ],
  });
}

function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing:   { before: 60, after: 0 },
        children:  [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_HDR })],
      }),
    ],
  });
}

// ─── 表紙ブロック ──────────────────────────────────────────────────────────────
function buildCoverBlock(spec) {
  const paras = [];

  // Subtitle — right-aligned
  if (spec.doc_subtitle) {
    paras.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing:   { before: 0, after: 80 },
      children:  [mkRun(spec.doc_subtitle, { sz: SZ })],
    }));
  }

  // Vertical space before title (~1/4 of page)
  paras.push(...spacerParas(7));

  // Document title — large, bold, centered
  paras.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 240 },
    children:  [mkRun(spec.doc_title || '仕様書', { sz: 32, bold: true })],
  }));

  // Revision / date — red, centered
  if (spec.revision) {
    paras.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 120 },
      children:  [mkRun(spec.revision, { color: 'red', sz: SZ })],
    }));
  }

  // Vertical space before company
  paras.push(...spacerParas(7));

  // Company
  if (spec.company) {
    paras.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 120 },
      children:  [mkRun(spec.company, { sz: SZ })],
    }));
  }

  // Department
  if (spec.department) {
    paras.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 0 },
      children:  [mkRun(spec.department, { sz: SZ })],
    }));
  }

  return paras;
}

// ─── 制定改廃経歴 ─────────────────────────────────────────────────────────────
function buildEstablishedTable(history) {
  const headerRow = new TableRow({
    children: [
      mkCell(mkCellPara('', { bold: true }), HIST_COLS[0], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('制定・廃止年月日', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[1], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('主な理由', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[2], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('備考', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[3], { shade: 'D9D9D9' }),
    ],
  });

  const dataRows = [];
  for (const e of (history.established || [])) {
    dataRows.push(new TableRow({
      children: [
        mkCell(mkCellPara('制定', { align: AlignmentType.CENTER }), HIST_COLS[0]),
        mkCell(mkCellPara(e.date || ''), HIST_COLS[1]),
        mkCell(mkCellPara(e.reason || ''), HIST_COLS[2]),
        mkCell(mkCellPara(e.note || ''), HIST_COLS[3]),
      ],
    }));
  }
  // Always add a 廃止 row
  dataRows.push(new TableRow({
    children: [
      mkCell(mkCellPara('廃止', { align: AlignmentType.CENTER }), HIST_COLS[0]),
      mkCell(mkCellPara(''), HIST_COLS[1]),
      mkCell(mkCellPara(''), HIST_COLS[2]),
      mkCell(mkCellPara(''), HIST_COLS[3]),
    ],
  }));

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: HIST_COLS,
    rows:         [headerRow, ...dataRows],
  });
}

function buildRevisedTable(history) {
  const REVISED_MIN_ROWS = 10;  // minimum total data rows (empty rows pad the rest)

  const headerRow = new TableRow({
    children: [
      mkCell(mkCellPara('番号', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[0], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('改正年月日', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[1], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('主な理由', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[2], { shade: 'D9D9D9' }),
      mkCell(mkCellPara('備考', { bold: true, align: AlignmentType.CENTER }), HIST_COLS[3], { shade: 'D9D9D9' }),
    ],
  });

  const dataRows = [];
  for (const e of (history.revised || [])) {
    dataRows.push(new TableRow({
      children: [
        mkCell(mkCellPara(e.number || '', { align: AlignmentType.CENTER }), HIST_COLS[0]),
        mkCell(mkCellPara(e.date || ''), HIST_COLS[1]),
        mkCell(mkCellPara(e.reason || ''), HIST_COLS[2]),
        mkCell(mkCellPara(e.note || ''), HIST_COLS[3]),
      ],
    }));
  }

  // Pad with empty rows
  const needed = Math.max(0, REVISED_MIN_ROWS - dataRows.length);
  for (let i = 0; i < needed; i++) {
    dataRows.push(new TableRow({
      children: HIST_COLS.map(w => mkCell(mkCellPara(''), w)),
    }));
  }

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: HIST_COLS,
    rows:         [headerRow, ...dataRows],
  });
}

function buildHistoryBlock(spec) {
  const paras = [];

  // Page break — new page for history
  paras.push(new Paragraph({
    pageBreakBefore: true,
    spacing:         { before: 0, after: 0 },
    children:        [],
  }));

  // Large title
  paras.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing:   { before: 0, after: 240 },
    children:  [mkRun('制定改廃経歴表', { sz: 32, bold: true })],
  }));

  // 仕様書名 line
  paras.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing:   { before: 0, after: 160 },
    border:    { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA', space: 1 } },
    children:  [
      mkRun('仕様書名：', { bold: true }),
      mkRun(spec.doc_title || '', { underline: true }),
    ],
  }));

  // 制定・廃止 section
  paras.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing:   { before: 160, after: 60 },
    children:  [mkRun('制定・廃止', { bold: true })],
  }));
  paras.push(buildEstablishedTable(spec.history));
  paras.push(emptyPara());

  // 改正 section
  paras.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing:   { before: 160, after: 60 },
    children:  [mkRun('改正', { bold: true })],
  }));
  paras.push(buildRevisedTable(spec.history));

  return paras;
}

// ─── Body sections ────────────────────────────────────────────────────────────
function buildBodySections(sections) {
  const paras = [];
  let first = true;

  for (const section of sections) {
    if (first) {
      // First section starts on new page
      if (section.heading) {
        paras.push(new Paragraph({
          pageBreakBefore: true,
          alignment:       AlignmentType.LEFT,
          spacing:         { before: 0, after: 80, line: 260, lineRule: 'auto' },
          children:        [mkRun(section.heading, { bold: true })],
        }));
      } else {
        paras.push(new Paragraph({
          pageBreakBefore: true,
          children:        [],
        }));
      }
      first = false;
    } else {
      if (section.heading) {
        paras.push(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing:   { before: 180, after: 80, line: 260, lineRule: 'auto' },
          children:  [mkRun(section.heading, { bold: true })],
        }));
      }
    }

    for (const paraSpec of (section.paragraphs || [])) {
      paras.push(specToPara(paraSpec));
    }
  }

  return paras;
}

// ─── Main export ──────────────────────────────────────────────────────────────
async function generateManualDoc(spec, outputPath) {
  const bodyChildren = [];

  // 1. 表紙
  bodyChildren.push(...buildCoverBlock(spec));

  // 2. 制定改廃経歴表 (if provided)
  if (spec.history) {
    bodyChildren.push(...buildHistoryBlock(spec));
  }

  // 3. 本文 sections
  if (spec.sections && spec.sections.length > 0) {
    bodyChildren.push(...buildBodySections(spec.sections));
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: SZ } } } },
    sections: [{
      headers: { default: makeHeader(spec.doc_title || '仕様書') },
      footers: { default: makeFooter() },
      properties: {
        page: {
          size:   { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN_TOP, bottom: MARGIN_BTM, left: MARGIN_LR, right: MARGIN_LR, header: 500, footer: 500 },
          borders: {
            pageBorders: {
              display:    PageBorderDisplay.ALL_PAGES,
              offsetFrom: PageBorderOffsetFrom.TEXT,
            },
            pageBorderTop:    { style: BorderStyle.SINGLE, size: 12, color: '000000', space: 24 },
            pageBorderRight:  { style: BorderStyle.SINGLE, size: 12, color: '000000', space: 24 },
            pageBorderBottom: { style: BorderStyle.SINGLE, size: 12, color: '000000', space: 24 },
            pageBorderLeft:   { style: BorderStyle.SINGLE, size: 12, color: '000000', space: 24 },
          },
        },
      },
      children: bodyChildren,
    }],
  });

  let buffer = await Packer.toBuffer(doc);

  // Collect comments from sections (if any)
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

module.exports = { generateManualDoc };
