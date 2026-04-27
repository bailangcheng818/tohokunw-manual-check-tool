'use strict';

const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const { createStore, writeScheme, writeContent, writeImage, writeImageMeta } = require('./file-store');
const { xmlDecode, extractRuns, classifyParagraph } = require('./template-analyzer');

// Vertex AI is loaded lazily to avoid startup errors when not configured
let _vertexAI = null;
function getVertexAI() {
  if (!_vertexAI) _vertexAI = require('./vertex-ai');
  return _vertexAI;
}

// ---------------------------------------------------------------------------
// In-flight deduplication
// Dify retries on timeout — if the same file URL set is already being processed,
// attach to the existing promise instead of starting a new job.
// ---------------------------------------------------------------------------
const _inFlightJobs = new Map(); // key → Promise

function makeJobKey(files) {
  return (files || [])
    .map(f => (f.url || '').replace(/[?&](timestamp|nonce|sign)=[^&]*/g, ''))
    .sort()
    .join('|');
}

// ---------------------------------------------------------------------------
// URL construction
// ---------------------------------------------------------------------------

function buildFileUrl(fileUrl, difyBaseUrl) {
  if (!fileUrl) throw new Error('Missing file URL');
  // Strip wrapping { } that Dify sometimes leaves around variable values
  const cleaned = fileUrl.replace(/^\{/, '').replace(/\}$/, '').trim();
  if (/^https?:\/\//i.test(cleaned)) return cleaned;
  const base = (difyBaseUrl || '').replace(/\/$/, '');
  const rel = cleaned.startsWith('/') ? cleaned : `/${cleaned}`;
  return base + rel;
}

// ---------------------------------------------------------------------------
// File download
// ---------------------------------------------------------------------------

async function downloadFile(url, difyApiKey) {
  console.log(`  [ingest] downloading: ${url}`);
  const response = await fetch(url, {
    headers: difyApiKey ? { Authorization: difyApiKey } : {},
  });
  console.log(`  [ingest] download status: ${response.status}`);
  if (!response.ok) {
    throw new Error(`Download failed: ${response.status} ${response.statusText} — ${url}`);
  }
  const arrayBuffer = await response.arrayBuffer();
  const buf = Buffer.from(arrayBuffer);
  console.log(`  [ingest] downloaded ${buf.length} bytes`);
  return buf;
}

// ---------------------------------------------------------------------------
// File → key mapping  (manual / bessi)
// ---------------------------------------------------------------------------

function mapFilesToKeys(files) {
  const result = { manual: -1, bessi: -1 };
  for (const [i, file] of (files || []).entries()) {
    const name = (file.name || '').toLowerCase();
    const ext = path.extname(name);
    if (result.manual === -1 && (name.includes('manual') || name.includes('仕様') || ext === '.docx')) {
      result.manual = i;
    } else if (result.bessi === -1 && (name.includes('bessi') || name.includes('別紙') || ext === '.xls' || ext === '.xlsx')) {
      result.bessi = i;
    }
  }
  // Fallback: index-based
  if (result.manual === -1) result.manual = 0;
  if (result.bessi === -1 && (files || []).length > 1) result.bessi = 1;
  return result;
}

// ---------------------------------------------------------------------------
// Markdown table helper
// ---------------------------------------------------------------------------

function rowsToMarkdownTable(rows) {
  if (!rows || rows.length === 0) return '';
  const lines = [];
  for (const [i, row] of rows.entries()) {
    const cells = (Array.isArray(row) ? row : []).map(cell => {
      const s = String(cell === null || cell === undefined ? '' : cell).replace(/\|/g, '｜');
      return s.length > 200 ? s.slice(0, 200) + '…' : s;
    });
    lines.push('| ' + cells.join(' | ') + ' |');
    if (i === 0) {
      lines.push('| ' + cells.map(() => '---').join(' | ') + ' |');
    }
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Legacy .doc processing  (Word 97-2003 binary format)
// ---------------------------------------------------------------------------

async function processDocFile(buffer, ref_id) {
  const WordExtractor = require('word-extractor');
  const extractor = new WordExtractor();

  // word-extractor expects a file path, so write to a temp path first
  const os = require('os');
  const tmpPath = path.join(os.tmpdir(), `ingest_${ref_id}.doc`);
  require('fs').writeFileSync(tmpPath, buffer);

  let content = '';
  try {
    const doc = await extractor.extract(tmpPath);
    const body = doc.getBody() || '';
    const footnotes = doc.getFootnotes() || '';
    const annotations = doc.getAnnotations() || '';
    content = [body, footnotes, annotations].filter(Boolean).join('\n\n').trim();
  } finally {
    try { require('fs').unlinkSync(tmpPath); } catch {}
  }

  const scheme = {
    file_type: 'doc',
    note: 'Legacy .doc format — text extracted only, no structure or images',
    char_count: content.length,
  };

  writeScheme(ref_id, scheme);
  writeContent(ref_id, content);

  return { scheme, content, images_summary: [] };
}

// ---------------------------------------------------------------------------
// Excel processing  (.xls via SheetJS, .xlsx via exceljs)
// ---------------------------------------------------------------------------

function processWithSheetJS(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheets = {};
  const contentParts = [];

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const headers = rows[0] ? rows[0].map(String) : [];
    sheets[sheetName] = {
      row_count: rows.length,
      col_count: rows[0] ? rows[0].length : 0,
      headers,
    };
    contentParts.push(`## ${sheetName}`);
    contentParts.push(rowsToMarkdownTable(rows));
  }

  return {
    scheme: { file_type: 'xls', sheet_names: workbook.SheetNames, sheets },
    content: contentParts.join('\n\n'),
    images_summary: [],
  };
}

async function processWithExcelJS(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheets = {};
  const contentParts = [];

  wb.eachSheet(ws => {
    const rows = [];
    ws.eachRow({ includeEmpty: false }, row => {
      rows.push(row.values.slice(1));
    });
    const headers = rows[0] ? rows[0].map(v => String(v === null || v === undefined ? '' : v)) : [];
    sheets[ws.name] = {
      row_count: ws.rowCount,
      col_count: ws.columnCount,
      headers,
    };
    contentParts.push(`## ${ws.name}`);
    contentParts.push(rowsToMarkdownTable(rows));
  });

  const sheetNames = Object.keys(sheets);
  return {
    scheme: { file_type: 'xlsx', sheet_names: sheetNames, sheets },
    content: contentParts.join('\n\n'),
    images_summary: [],
  };
}

async function processExcelFile(buffer, ext, ref_id) {
  const result = ext === 'xls'
    ? processWithSheetJS(buffer)
    : await processWithExcelJS(buffer);

  writeScheme(ref_id, result.scheme);
  writeContent(ref_id, result.content);
  return result;
}

// ---------------------------------------------------------------------------
// 指令一: DOCX content type detection (drawing / image / table / text)
// ---------------------------------------------------------------------------

/**
 * Scan document XML and classify each paragraph and table by content type.
 *
 * Types:
 *   "drawing"  — paragraph contains <w:drawing><wp:anchor> or <mc:AlternateContent>
 *                (SmartArt, charts, shapes — not extractable as raw image bytes)
 *   "image"    — paragraph contains <w:drawing><wp:inline> (raster image embed)
 *   "table"    — top-level <w:tbl> element
 *   "text"     — plain paragraph with text
 *
 * @param {string} docXml
 * @returns {{ items: Array, summary: object }}
 */
/**
 * @param {string} docXml
 * @returns {{
 *   items: Array,
 *   summary: { has_drawing: bool, has_table: bool, has_image: bool, drawing_count: number },
 *   drawingParas: Array<{ para_index: number, para_total: number, context_text: string }>
 * }}
 */
function detectDocxContentTypes(docXml) {
  const items = [];
  let drawingCount = 0;
  let imageCount = 0;

  const paragraphMatches = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  const paraTotal = paragraphMatches.length;

  // Pre-extract plain text for each paragraph (for contextText building)
  const paraTexts = paragraphMatches.map(pXml =>
    [...pXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)]
      .map(m => xmlDecode(m[1]))
      .join('')
  );

  const drawingParas = [];

  for (const [i, pXml] of paragraphMatches.entries()) {
    const hasText = /<w:t[ >]/.test(pXml);
    const hasAlternate = pXml.includes('<mc:AlternateContent');
    const hasDrawingTag = pXml.includes('<w:drawing');
    const hasVmlPict = pXml.includes('<w:pict') &&
      (pXml.includes('<v:shape') || pXml.includes('<v:imagedata'));
    const hasObjectTag = pXml.includes('<w:object');

    let type = 'text';
    if (hasAlternate || hasVmlPict || hasObjectTag) {
      type = 'drawing';
      drawingCount++;
    } else if (hasDrawingTag) {
      if (pXml.includes('<wp:inline') && !pXml.includes('<wp:anchor')) {
        type = 'image';
        imageCount++;
      } else {
        type = 'drawing';
        drawingCount++;
      }
    }

    items.push({ type, page_hint: i, has_text: hasText });

    if (type === 'drawing') {
      // Collect ±200 chars of surrounding plain text as context
      const before = paraTexts.slice(Math.max(0, i - 5), i).join(' ').slice(-200);
      const after = paraTexts.slice(i + 1, i + 6).join(' ').slice(0, 200);
      const context_text = [before, after].filter(Boolean).join('\n');
      drawingParas.push({ para_index: i, para_total: paraTotal, context_text });
    }
  }

  const tableMatches = docXml.match(/<w:tbl[ >][\s\S]*?<\/w:tbl>/g) || [];
  for (let i = 0; i < tableMatches.length; i++) {
    items.push({ type: 'table', page_hint: -1, has_text: true });
  }

  const summary = {
    has_drawing: drawingCount > 0,
    has_table: tableMatches.length > 0,
    has_image: imageCount > 0,
    drawing_count: drawingCount,
  };

  return { items, summary, drawingParas };
}

// ---------------------------------------------------------------------------
// 指令二: Convert DOCX pages to PNG images via LibreOffice + pdftoppm
// ---------------------------------------------------------------------------

/**
 * Convert all pages of a DOCX to PNG images.
 * Requires: `soffice` (LibreOffice) and `pdftoppm` (poppler-utils) on PATH.
 *
 * Install on Debian/Ubuntu:
 *   apt-get install -y libreoffice poppler-utils
 *
 * @param {string} docxPath  - absolute path to .docx file
 * @param {string} outputDir - directory to write images into
 * @returns {Promise<string[]>} sorted list of absolute PNG file paths
 */
async function convertDocxPagesToImages(docxPath, outputDir) {
  const { execFile } = require('child_process');
  const { promisify } = require('util');
  const execFileAsync = promisify(execFile);

  // Check tool availability
  for (const tool of ['soffice', 'pdftoppm']) {
    try {
      await execFileAsync('which', [tool]);
    } catch {
      throw new Error(`Required tool not found: ${tool}. Install with: apt-get install -y libreoffice poppler-utils`);
    }
  }

  // Step 1: DOCX → PDF
  await execFileAsync('soffice', [
    '--headless',
    '--convert-to', 'pdf',
    '--outdir', outputDir,
    docxPath,
  ]);

  const pdfName = path.basename(docxPath, '.docx') + '.pdf';
  const pdfPath = path.join(outputDir, pdfName);

  if (!fs.existsSync(pdfPath)) {
    throw new Error(`LibreOffice did not produce a PDF at: ${pdfPath}`);
  }

  // Step 2: PDF → PNG (150 dpi)
  const imgPrefix = path.join(outputDir, 'page');
  await execFileAsync('pdftoppm', ['-r', '150', '-png', pdfPath, imgPrefix]);

  // Collect generated PNG files
  const pngFiles = fs.readdirSync(outputDir)
    .filter(f => f.startsWith('page') && f.endsWith('.png'))
    .sort()
    .map(f => path.join(outputDir, f));

  return pngFiles;
}

// ---------------------------------------------------------------------------
// DOCX image extraction
// ---------------------------------------------------------------------------

async function extractDocxImages(zip) {
  const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string') || '';
  const imageRels = [...relsXml.matchAll(/Type="[^"]*\/image"[^>]*Target="([^"]+)"/g)]
    .map(m => m[1]);

  const images = [];
  for (const target of imageRels) {
    const zipPath = target.startsWith('media/')
      ? `word/${target}`
      : target.startsWith('word/') ? target : `word/media/${path.basename(target)}`;
    const file = zip.file(zipPath);
    if (!file) continue;
    const buffer = await file.async('nodebuffer');
    const rawExt = path.extname(target).toLowerCase().replace('.', '') || 'png';
    // Skip Windows metafiles — Gemini Vision can't read them
    if (rawExt === 'emf' || rawExt === 'wmf') continue;
    const ext = rawExt === 'jpeg' ? 'jpg' : rawExt;
    images.push({ buffer, ext, target });
  }
  return images;
}

// ---------------------------------------------------------------------------
// Header / footer extraction
// ---------------------------------------------------------------------------

async function extractHeadersFooters(zip) {
  const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string') || '';
  const relPattern = /Type="[^"]*\/(header|footer)"[^/]*Target="([^"]+)"/g;
  const parts = [];
  let m;
  while ((m = relPattern.exec(relsXml)) !== null) {
    parts.push({ role: m[1], target: m[2] });
  }

  const headers = [];
  const footers = [];
  for (const { role, target } of parts) {
    const zipPath = target.startsWith('word/') ? target : `word/${target}`;
    const xml = await zip.file(zipPath)?.async('string');
    if (!xml) continue;
    const paragraphs = [];
    for (const pXml of (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [])) {
      const runs = extractRuns(pXml);
      const text = runs.map(r => r.text).join('');
      if (!text.trim()) continue;
      const para_id = `p_${String(paragraphs.length + 1).padStart(3, '0')}`;
      paragraphs.push({ para_id, xml_index: paragraphs.length, type: 'body', text, runs });
    }
    const entry = { part: target.replace(/^word\//, '').replace(/\.xml$/, ''), paragraphs };
    (role === 'header' ? headers : footers).push(entry);
  }
  return { headers, footers };
}

// ---------------------------------------------------------------------------
// Text box extraction
// ---------------------------------------------------------------------------

function extractTextboxes(docXml) {
  const textboxes = [];
  for (const [idx, txXml] of (docXml.match(/<w:txbx[\s\S]*?<\/w:txbx>/g) || []).entries()) {
    const paragraphs = [];
    for (const pXml of (txXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [])) {
      const runs = extractRuns(pXml);
      const text = runs.map(r => r.text).join('');
      if (text.trim()) paragraphs.push({ text, runs });
    }
    if (paragraphs.length > 0) textboxes.push({ index: idx, paragraphs });
  }
  return textboxes;
}

// ---------------------------------------------------------------------------
// DOCX processing
// ---------------------------------------------------------------------------

async function processDocxFile(buffer, ref_id) {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml').async('string');

  // --- Paragraphs ---
  const paragraphMatches = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  const paragraphs = [];
  const contentLines = [];

  for (const [xmlIndex, pXml] of paragraphMatches.entries()) {
    const runs = extractRuns(pXml);
    const text = runs.map(r => r.text).join('');
    if (!text.trim()) continue;

    const boldMatch = pXml.match(/<w:sz[^>]+w:val="(\d+)"/);
    const size = boldMatch ? Number(boldMatch[1]) : null;
    const alignMatch = pXml.match(/<w:jc[^>]+w:val="([^"]+)"/);
    const align = alignMatch ? alignMatch[1] : 'left';

    const type = classifyParagraph({ text, runs, size, align });
    const para_id = `p_${String(paragraphs.length + 1).padStart(3, '0')}`;
    paragraphs.push({ para_id, xml_index: xmlIndex, type, text, runs, size, align });

    if (type === 'title' || type === 'numbered_heading' || type === 'bracket_heading') {
      contentLines.push(`\n### ${text}`);
    } else {
      contentLines.push(text);
    }
  }

  // --- Tables ---
  const tableMatches = docXml.match(/<w:tbl[ >][\s\S]*?<\/w:tbl>/g) || [];
  const tables = [];

  for (const [ti, tblXml] of tableMatches.entries()) {
    const rowMatches = tblXml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    const rows = rowMatches.map(trXml => {
      const cellMatches = trXml.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || [];
      return cellMatches.map(tcXml => {
        const runs = extractRuns(tcXml);
        return runs.map(r => r.text).join('');
      });
    });
    tables.push({ table_index: ti, rows, row_count: rows.length, col_count: rows[0]?.length || 0 });

    contentLines.push(`\n#### 表${ti + 1}`);
    contentLines.push(rowsToMarkdownTable(rows));
  }

  // --- Images ---
  const images = await extractDocxImages(zip);
  const images_summary = [];
  const vertexAI = getVertexAI();

  for (const [idx, img] of images.entries()) {
    const imgIndex = idx + 1;
    writeImage(ref_id, imgIndex, img.buffer, img.ext);

    const mimeType = img.ext === 'jpg' ? 'image/jpeg' : `image/${img.ext}`;
    let meta = { label: '（未解析）', summary: '' };
    try {
      meta = await vertexAI.labelImage({ imageBuffer: img.buffer, mimeType });
    } catch (err) {
      meta = { label: 'ラベル取得失敗', summary: err.message };
    }
    writeImageMeta(ref_id, imgIndex, meta);
    images_summary.push({ ref: `img_${String(imgIndex).padStart(3, '0')}.${img.ext}`, ...meta });
  }

  // --- 指令一: Drawing detection ---
  const { summary: contentSummary, drawingParas } = detectDocxContentTypes(docXml);
  console.log(`  [ingest] content types: drawing=${contentSummary.drawing_count}, has_image=${contentSummary.has_image}, has_table=${contentSummary.has_table}`);

  // --- 指令三〜四: Drawing → page images → LLM analysis (drawing pages only) ---
  let drawing_detected = false;
  let drawing_preview = [];

  if (contentSummary.has_drawing) {
    drawing_detected = true;
    const storeDir = require('./file-store').getStoreDir(ref_id);
    const drawingDir = path.join(storeDir, 'drawing_pages');
    fs.mkdirSync(drawingDir, { recursive: true });

    try {
      const originalDocxPath = path.join(storeDir, 'original.docx');
      console.log(`  [ingest] drawing detected (${contentSummary.drawing_count} shapes) — converting to page images...`);

      const pagePngs = await convertDocxPagesToImages(originalDocxPath, drawingDir);
      const totalPages = pagePngs.length;
      console.log(`  [ingest] converted ${totalPages} pages to PNG`);

      // Build a Map: 1-based page number → context_text
      // Estimation: page ≈ round(para_index / (para_total-1) * (totalPages-1)) + 1
      const pageContextMap = new Map();
      for (const dp of drawingParas) {
        const p = Math.min(
          totalPages,
          Math.max(1, Math.round((dp.para_index / Math.max(dp.para_total - 1, 1)) * (totalPages - 1)) + 1)
        );
        const existing = pageContextMap.get(p) || '';
        pageContextMap.set(p, existing ? `${existing}\n${dp.context_text}` : dp.context_text);
      }

      console.log(`  [ingest] drawing pages to analyze: [${[...pageContextMap.keys()].sort((a, b) => a - b).join(', ')}] / ${totalPages}`);

      for (const [pi, pngPath] of pagePngs.entries()) {
        const pageNum = pi + 1;
        if (!pageContextMap.has(pageNum)) continue; // skip non-drawing pages

        const contextText = pageContextMap.get(pageNum) || '';
        const imageBuffer = fs.readFileSync(pngPath);
        const SIZE_THRESHOLD = 150 * 1024; // 150 KB — text-only pages at 150 DPI are typically smaller
        if (imageBuffer.length < SIZE_THRESHOLD) {
          console.log(`  [ingest] page ${pageNum} skipped (${Math.round(imageBuffer.length / 1024)}KB < 150KB threshold)`);
          continue;
        }
        let analysis = { label: '（未解析）', summary: '', figure_type: 'other', key_elements: [], mermaid: '' };
        try {
          analysis = await vertexAI.labelImage({ imageBuffer, mimeType: 'image/png', contextText });
          console.log(`  [ingest] page ${pageNum} analyzed: ${analysis.label}`);
        } catch (err) {
          console.warn(`  [ingest] page ${pageNum} analysis failed: ${err.message}`);
          analysis.summary = err.message;
        }

        // Save drawing page as a formal image entry (continuous index after raster images)
        const imgIndex = images.length + drawing_preview.length + 1;
        writeImage(ref_id, imgIndex, imageBuffer, 'png');
        writeImageMeta(ref_id, imgIndex, { ...analysis, source: 'drawing_page', page: pageNum });
        const imgRef = `img_${String(imgIndex).padStart(3, '0')}.png`;
        images_summary.push({ ref: imgRef, source: 'drawing_page', page: pageNum, ...analysis });

        drawing_preview.push({
          page: pageNum,
          png_path: pngPath,
          img_ref: imgRef,
          ...analysis,
        });
      }
    } catch (err) {
      console.warn(`  [ingest] drawing-to-image conversion failed: ${err.message}`);
      drawing_preview = [{ page: 0, label: '変換失敗', summary: err.message, figure_type: 'other', key_elements: [], mermaid: '' }];
    }
  }

  const { headers, footers } = await extractHeadersFooters(zip);
  const textboxes = extractTextboxes(docXml);

  const scheme = {
    file_type: 'docx',
    paragraph_count: paragraphs.length,
    heading_count: paragraphs.filter(p => p.type.includes('heading') || p.type === 'title').length,
    table_count: tables.length,
    image_count: images.length + drawing_preview.length,
    drawing_count: contentSummary.drawing_count,
    paragraphs,
    tables,
    drawing_pages: drawing_preview,
    headers,
    footers,
    textboxes,
  };

  writeScheme(ref_id, scheme);
  writeContent(ref_id, contentLines.join('\n'));

  return {
    scheme,
    content: contentLines.join('\n'),
    images_summary,
    drawing_detected,
    drawing_preview,
  };
}

// ---------------------------------------------------------------------------
// Main handler
// ---------------------------------------------------------------------------

async function handleIngest(body) {
  const { files, dify_base_url: difyBaseUrl, dify_api_key: difyApiKey } = body || {};

  if (!files || !Array.isArray(files) || files.length === 0) {
    const err = new Error('Missing or empty files array');
    err.statusCode = 400;
    throw err;
  }

  // Deduplication: if the same file set is already being processed, wait for it
  const jobKey = makeJobKey(files);
  if (_inFlightJobs.has(jobKey)) {
    console.log(`  [ingest] duplicate request — attaching to existing job (key=${jobKey.slice(0, 60)}...)`);
    return _inFlightJobs.get(jobKey);
  }

  const job = _doIngest({ files, difyBaseUrl, difyApiKey });
  _inFlightJobs.set(jobKey, job);
  job.finally(() => _inFlightJobs.delete(jobKey));
  return job;
}

async function _doIngest({ files, difyBaseUrl, difyApiKey }) {

  const keyMap = mapFilesToKeys(files);
  const result = {};

  for (const [key, fileIndex] of Object.entries(keyMap)) {
    if (fileIndex < 0 || fileIndex >= files.length) continue;

    const file = files[fileIndex];
    const filename = (file.name || `file_${fileIndex}`).replace(/\}$/, '').trim();
    const ext = path.extname(filename).toLowerCase().replace('.', '') || 'bin';

    console.log(`  [ingest] key=${key} file="${filename}" ext=${ext}`);

    const fullUrl = buildFileUrl(file.url, difyBaseUrl);
    console.log(`  [ingest] full url: ${fullUrl}`);

    const buffer = await downloadFile(fullUrl, difyApiKey);

    const { ref_id, originalPath } = createStore(ext);
    fs.writeFileSync(originalPath, buffer);
    console.log(`  [ingest] ref_id=${ref_id} saved original to ${originalPath}`);

    let processed;
    if (ext === 'docx') {
      console.log(`  [ingest] processing DOCX...`);
      processed = await processDocxFile(buffer, ref_id);
      console.log(`  [ingest] DOCX done: ${processed.scheme.paragraph_count} paras, ${processed.scheme.table_count} tables, ${processed.scheme.image_count} images`);
    } else if (ext === 'doc') {
      console.log(`  [ingest] processing DOC (legacy Word)...`);
      processed = await processDocFile(buffer, ref_id);
      console.log(`  [ingest] DOC done: ${processed.content.length} chars`);
    } else if (ext === 'xls' || ext === 'xlsx') {
      console.log(`  [ingest] processing Excel (${ext})...`);
      processed = await processExcelFile(buffer, ext, ref_id);
      console.log(`  [ingest] Excel done: sheets=${processed.scheme.sheet_names?.join(',')}`);
    } else {
      console.log(`  [ingest] unsupported ext "${ext}", storing raw only`);
      processed = { scheme: { file_type: ext }, content: '', images_summary: [] };
      writeScheme(ref_id, processed.scheme);
      writeContent(ref_id, '');
    }

    result[key] = {
      ref_id,
      content: processed.content,
      scheme: processed.scheme,
      images_summary: processed.images_summary,
      drawing_detected: processed.drawing_detected || false,
      drawing_preview: processed.drawing_preview || [],
    };
  }

  return result;
}

module.exports = { handleIngest, processDocxFile, processDocFile, processExcelFile };
