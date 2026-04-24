'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { FILE_STORE_DIR, MANUAL_DATABASE_DIR } = require('./config');
const { createStore, writeContent, writeScheme, getManifestEntry, readImagesMeta } = require('./file-store');
const { getHeaders } = require('./excel-writer');

const WORD_EXTS = new Set(['.docx', '.doc']);
const EXCEL_EXTS = new Set(['.xlsx', '.xls']);
const DEFAULT_EXTS = new Set(['.docx', '.doc', '.xlsx', '.xls']);

// Session-scoped map: folder_name (NFC) → ref_id (backed by file-store on disk)
const folderRefMap = new Map();

function fileType(ext) {
  if (WORD_EXTS.has(ext)) return 'word';
  if (EXCEL_EXTS.has(ext)) return 'excel';
  return 'other';
}

// macOS HFS+/APFS returns filenames in NFD from readdirSync, but callers
// (Dify, browser) typically send NFC. Normalize everything to NFC so
// comparisons are consistent regardless of Unicode normalization form.
function nfc(str) {
  return typeof str === 'string' ? str.normalize('NFC') : str;
}

function validateFolderName(name) {
  if (!name || typeof name !== 'string') {
    throw Object.assign(new Error('folder_name is required'), { statusCode: 400 });
  }
  if (name.includes('/') || name.includes('\\') || name.includes('..')) {
    throw Object.assign(new Error('folder_name must not contain path separators or ".."'), { statusCode: 400 });
  }
}

function parseExtensions(extensions) {
  if (!extensions) return DEFAULT_EXTS;
  return new Set(extensions.split(',').map(e => e.trim().toLowerCase()));
}

/**
 * List manual folders under MANUAL_DATABASE_DIR (or a subfolder).
 * Each subdirectory is treated as one "manual unit".
 */
async function listManualFolders({ folder, extensions } = {}) {
  const extFilter = parseExtensions(extensions);
  const baseDir = folder ? path.join(MANUAL_DATABASE_DIR, nfc(folder)) : MANUAL_DATABASE_DIR;

  if (!fs.existsSync(baseDir)) {
    return { folder: baseDir, manuals: [], count: 0 };
  }

  const entries = fs.readdirSync(baseDir, { withFileTypes: true });
  const manuals = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const folderName = nfc(entry.name);
    const folderPath = path.join(baseDir, entry.name);
    const files = fs.readdirSync(folderPath, { withFileTypes: true }).filter(f => f.isFile());
    const filtered = files.filter(f => extFilter.has(path.extname(f.name).toLowerCase()));
    if (filtered.length === 0) continue;

    // Primary doc: file sharing the folder name → first .docx → first file
    // Compare NFC-normalised stems to handle macOS NFD filenames correctly.
    let primary = filtered.find(f => nfc(path.basename(f.name, path.extname(f.name))) === folderName);
    if (!primary) primary = filtered.find(f => WORD_EXTS.has(path.extname(f.name).toLowerCase()));
    if (!primary) primary = filtered[0];

    const primaryExt = path.extname(primary.name).toLowerCase();
    const primaryStat = fs.statSync(path.join(folderPath, primary.name));

    const attachments = filtered
      .filter(f => f.name !== primary.name)
      .map(f => {
        const ext = path.extname(f.name).toLowerCase();
        const stat = fs.statSync(path.join(folderPath, f.name));
        return { name: nfc(f.name), type: fileType(ext), ext, size_kb: Math.round(stat.size / 1024) };
      });

    const manifestEntry = getManifestEntry(folderName);
    const entry = {
      name: folderName,
      primary_doc: nfc(primary.name),
      type: fileType(primaryExt),
      ext: primaryExt,
      size_kb: Math.round(primaryStat.size / 1024),
      modified: primaryStat.mtime.toISOString(),
      attachments,
      ref_id: manifestEntry?.ref_id || folderRefMap.get(folderName) || null,
      images_analyzed: manifestEntry?.images_analyzed || false,
    };
    if (manifestEntry?.summary) entry.summary = manifestEntry.summary;
    if (manifestEntry?.effective_date) entry.effective_date = manifestEntry.effective_date;
    if (manifestEntry?.key_topics) entry.key_topics = manifestEntry.key_topics;
    if (manifestEntry?.document_type) entry.document_type = manifestEntry.document_type;
    if (manifestEntry?.attachment_summaries?.length) entry.attachment_summaries = manifestEntry.attachment_summaries;
    manuals.push(entry);
  }

  return { folder: baseDir, manuals, count: manuals.length };
}

/**
 * Extract plain text and structure from a DOCX buffer.
 * Lightweight: no AI calls, no LibreOffice dependency.
 */
async function extractDocxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')?.async('text');
  if (!docXml) return { text: '', paragraphs: 0, tables: 0 };

  // Split on paragraph boundaries and extract w:t text within each paragraph
  const paraChunks = docXml.split(/<w:p[ >\/]/);
  const lines = paraChunks
    .map(p => {
      const matches = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
      return matches.map(m => m.replace(/<[^>]+>/g, '')).join('');
    })
    .filter(Boolean);

  const tableCount = (docXml.match(/<w:tbl[ >]/g) || []).length;

  return { text: lines.join('\n'), paragraphs: lines.length, tables: tableCount };
}

/**
 * Read a manual folder: extract primary DOCX text, list attachment metadata.
 * Stores the primary file in file-store.js so ref_id works with /generate/from-edit.
 */
async function readManualFolder({ folder_name, mode = 'full' } = {}) {
  validateFolderName(folder_name);
  const normalizedName = nfc(folder_name);

  const folderPath = path.join(MANUAL_DATABASE_DIR, normalizedName);
  if (!fs.existsSync(folderPath)) {
    throw Object.assign(new Error(`Folder not found: ${normalizedName}`), { statusCode: 404 });
  }

  const files = fs.readdirSync(folderPath, { withFileTypes: true }).filter(f => f.isFile());

  // Find primary document — NFC-compare stem vs folder name
  let primaryFile = files.find(f => {
    const ext = path.extname(f.name).toLowerCase();
    return WORD_EXTS.has(ext) && nfc(path.basename(f.name, ext)) === normalizedName;
  });
  if (!primaryFile) primaryFile = files.find(f => WORD_EXTS.has(path.extname(f.name).toLowerCase()));

  if (!primaryFile) {
    throw Object.assign(
      new Error(`No Word document found in folder: ${normalizedName}`),
      { statusCode: 422 },
    );
  }

  const primaryPath = path.join(folderPath, primaryFile.name);
  const primaryExt = path.extname(primaryFile.name).toLowerCase();
  const primaryMtime = fs.statSync(primaryPath).mtime.toISOString();

  // Check manifest for pre-ingested cache (full image analysis)
  const manifestEntry = getManifestEntry(normalizedName);
  const cacheHit = manifestEntry && manifestEntry.primary_mtime === primaryMtime;

  let ref_id;
  let primaryResult;
  let images_analyzed = false;
  let images_stale = false;
  let images_summary = [];

  if (cacheHit && mode !== 'schema') {
    // Serve from pre-ingest cache — no re-processing needed
    ref_id = manifestEntry.ref_id;
    folderRefMap.set(normalizedName, ref_id);
    const contentPath = path.join(FILE_STORE_DIR, ref_id, 'content.txt');
    const content = fs.existsSync(contentPath) ? fs.readFileSync(contentPath, 'utf8') : null;
    images_summary = readImagesMeta(ref_id);
    images_analyzed = manifestEntry.images_analyzed || false;
    primaryResult = {
      file: nfc(primaryFile.name),
      type: 'word',
      content,
      scheme: null,
    };
  } else {
    // Lightweight path — text extraction only
    if (manifestEntry && !cacheHit) images_stale = true;

    const buffer = fs.readFileSync(primaryPath);

    // Reuse session ref_id or create new store
    ref_id = folderRefMap.get(normalizedName);
    if (!ref_id) {
      const store = createStore(primaryExt.replace('.', ''));
      ref_id = store.ref_id;
      fs.writeFileSync(store.originalPath, buffer);
      folderRefMap.set(normalizedName, ref_id);
    }

    if (primaryExt === '.docx' && mode !== 'schema') {
      const extracted = await extractDocxText(buffer);
      const scheme = {
        file_name: nfc(primaryFile.name),
        file_type: 'docx',
        paragraphs: extracted.paragraphs,
        tables: extracted.tables,
      };
      writeContent(ref_id, extracted.text);
      writeScheme(ref_id, scheme);
      primaryResult = { file: nfc(primaryFile.name), type: 'word', content: extracted.text, scheme };
    } else {
      // Excel primary or schema-only mode
      const headers = EXCEL_EXTS.has(primaryExt) ? await getHeaders(primaryPath) : [];
      const scheme = { file_name: nfc(primaryFile.name), file_type: primaryExt.replace('.', ''), headers };
      writeScheme(ref_id, scheme);
      primaryResult = { file: nfc(primaryFile.name), type: fileType(primaryExt), content: null, scheme };
    }
  }

  // Build attachment list with headers for Excel files
  const attachments = await Promise.all(
    files
      .filter(f => f.name !== primaryFile.name && DEFAULT_EXTS.has(path.extname(f.name).toLowerCase()))
      .map(async f => {
        const fPath = path.join(folderPath, f.name);
        const ext = path.extname(f.name).toLowerCase();
        const stat = fs.statSync(fPath);
        const info = { name: nfc(f.name), type: fileType(ext), ext, size_kb: Math.round(stat.size / 1024) };
        if (EXCEL_EXTS.has(ext)) {
          try { info.headers = await getHeaders(fPath); } catch { info.headers = []; }
        }
        return info;
      }),
  );

  const result = { folder_name: normalizedName, ref_id, primary: primaryResult, attachments, images_analyzed };
  if (images_analyzed) result.images_summary = images_summary;
  if (images_stale) result.images_stale = true;
  return result;
}

function getFolderRef(folderName) {
  return folderRefMap.get(nfc(folderName)) || null;
}

module.exports = { listManualFolders, readManualFolder, getFolderRef };
