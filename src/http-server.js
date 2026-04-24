#!/usr/bin/env node
'use strict';

const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const {
  getAssetDefinition,
  getAssetSchemaDescription,
  listAssetTypes,
  validateAssetSpec,
} = require('./asset-registry');
const {
  excelSchemaDescription,
  handleAppendEditRecord,
  handleAppendRow,
  handleFindRow,
  handleReadExcelHeaders,
  handleUpdateCell,
  handleUpdateRow,
} = require('./excel-tools');
const { HOST, MANUAL_DATABASE_DIR, OUTPUT_DIR, PORT, SERVICE_NAME, VERSION, safeFilename } = require('./config');
const { listManualFolders, readManualFolder } = require('./file-discovery');
const { handleIngest, processDocxFile, processDocFile, processExcelFile } = require('./ingest');
const { createStore, getManifestEntry, setManifestEntry } = require('./file-store');
const { summarizeDocument } = require('./vertex-ai');
const { applyEdits } = require('./edit-applier');

// ── Logger ────────────────────────────────────────────────────────────────────
const C = {
  reset: '\x1b[0m', bold: '\x1b[1m', dim: '\x1b[2m',
  red: '\x1b[31m', green: '\x1b[32m', yellow: '\x1b[33m',
  blue: '\x1b[34m', magenta: '\x1b[35m', cyan: '\x1b[36m', gray: '\x1b[90m',
};

const METHOD_COLOR = { GET: C.blue, POST: C.green, PUT: C.yellow, DELETE: C.red, PATCH: C.magenta };

// Route group tag for quick scanning
const ROUTE_TAG = [
  [/^\/(list-files|read-file)/, 'files '],
  [/^\/(excel\/|append-row|update-row|update-cell|edit-record|find-row|schema\/excel)/, 'excel '],
  [/^\/(generate|ingest)/, 'doc   '],
  [/^\/schema/, 'schema'],
  [/^\/health/, 'health'],
];

function ts() {
  return new Date().toLocaleTimeString('ja-JP', { hour12: false });
}

function tag(path) {
  for (const [re, label] of ROUTE_TAG) {
    if (re.test(path)) return C.dim + '[' + label + ']' + C.reset;
  }
  return C.dim + '[other ]' + C.reset;
}

function colorMethod(m) {
  return (METHOD_COLOR[m] || C.gray) + C.bold + m.padEnd(6) + C.reset;
}

function colorStatus(code) {
  const c = code >= 500 ? C.red : code >= 400 ? C.yellow : code >= 300 ? C.cyan : C.green;
  return c + C.bold + code + C.reset;
}

function bodyLine(body) {
  if (!body || typeof body !== 'object' || !Object.keys(body).length) return '';
  return Object.entries(body).map(([k, v]) => {
    if (v === null || v === undefined) return C.dim + k + C.reset;
    if (typeof v === 'object') return `${C.cyan}${k}${C.reset}:{${Object.keys(v).join(',')}}`;
    const s = String(v);
    return `${C.cyan}${k}${C.reset}:${s.length > 28 ? s.slice(0, 28) + '…' : s}`;
  }).join('  ');
}

function responseSummary(body) {
  if (!body || typeof body !== 'object') return '';
  const picks = [];
  if (body.ref_id)          picks.push(`ref_id:${String(body.ref_id).slice(0,8)}…`);
  if (body.count != null)   picks.push(`count:${body.count}`);
  if (body.appended_rows)   picks.push(`appended:${body.appended_rows}`);
  if (body.updated_columns) picks.push(`updated:[${body.updated_columns.join(',')}]`);
  if (body.found != null)   picks.push(`found:${body.found}`);
  if (body.folder_name)     picks.push(`folder:${body.folder_name}`);
  if (!body.success && body.error) picks.push(`${C.red}${body.error}${C.reset}`);
  return picks.length ? '  ' + C.dim + picks.join('  ') + C.reset : '';
}

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// ── Request / Response logger ─────────────────────────────────────────────────
app.use((req, res, next) => {
  const start = Date.now();

  // Intercept res.json to capture the response body for summary logging
  let resBody;
  const origJson = res.json.bind(res);
  res.json = (body) => { resBody = body; return origJson(body); };

  const body = bodyLine(req.body);
  process.stdout.write(
    `\n${C.gray}${ts()}${C.reset}  ${tag(req.path)}  ${colorMethod(req.method)} ${C.bold}${req.path}${C.reset}` +
    (body ? `\n         ${C.dim}body${C.reset}  ${body}` : '') + '\n',
  );

  res.on('finish', () => {
    const ms = Date.now() - start;
    const msStr = ms > 2000 ? C.yellow + ms + 'ms' + C.reset : C.dim + ms + 'ms' + C.reset;
    const summary = responseSummary(resBody);
    process.stdout.write(`         ${colorStatus(res.statusCode)}  ${msStr}${summary}\n`);
  });

  next();
});

fs.mkdirSync(OUTPUT_DIR, { recursive: true });

app.use('/files', express.static(OUTPUT_DIR, {
  setHeaders: (res, filePath) => {
    const filename = path.basename(filePath);
    if (filename.endsWith('.docx')) {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } else if (filename.endsWith('.xlsx')) {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
  },
}));

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    service: SERVICE_NAME,
    version: VERSION,
    port: PORT,
    output_dir: OUTPUT_DIR,
    excel_output_dir: OUTPUT_DIR,
    excel_output_dir_set: !!process.env.OUTPUT_DIR,
    manual_database_dir: MANUAL_DATABASE_DIR,
    manual_database_dir_set: !!process.env.MANUAL_DATABASE_DIR,
    endpoints: [
      'GET /health', 'GET /schema', 'GET /assets',
      'GET /list-files', 'POST /read-file',
      'POST /ingest', 'POST /generate/from-edit',
      'GET /schema/excel/:file', 'POST /excel/append-row', 'POST /excel/update-cell',
      'POST /excel/update-row', 'GET /find-row/:file',
      'POST /append-row', 'POST /update-cell', 'POST /update-row',
    ],
  });
});

app.get('/schema', (req, res) => {
  res.json({
    service: SERVICE_NAME,
    description: 'Unified document generation and Excel manual-check support API.',
    document_endpoints: {
      'GET /assets': 'List supported document asset types.',
      'GET /schema/assets/:asset_type': 'Get schema guidance for a document asset type.',
      'POST /generate/:asset_type': 'Generate a document asset.',
      'POST /generate/:asset_type/download': 'Generate and stream a document asset.',
    },
    ingest_endpoints: {
      'POST /ingest': 'Ingest DOCX/XLS files from Dify. Returns ref_id, content, scheme, images_summary per file.',
      'POST /generate/from-edit': 'Apply edit diff JSON to a stored original DOCX. Returns edited DOCX as base64 + download_url.',
    },
    excel_endpoints: {
      'GET /schema/excel': 'Get Excel write/update guidance.',
      'GET /schema/excel/:file': 'Read workbook headers.',
      'POST /excel/append-row': 'Append row(s) to Excel.',
      'POST /excel/update-cell': 'Update cell(s) in Excel.',
      'POST /excel/edit-record': 'Append 改正管理表 record.',
    },
    compatibility_endpoints: {
      'GET /schema/comparison': 'Alias for /schema/assets/comparison_doc.',
      'GET /schema/manual': 'Alias for /schema/assets/manual_doc.',
      'POST /generate': 'Alias for /generate/comparison_doc.',
      'POST /generate/manual': 'Alias for /generate/manual_doc.',
      'POST /append-row': 'Alias for /excel/append-row.',
      'POST /update-cell': 'Alias for /excel/update-cell.',
      'POST /edit-record': 'Alias for /excel/edit-record.',
    },
  });
});

app.get('/assets', (req, res) => {
  res.json({ assets: listAssetTypes() });
});

app.get('/schema/assets/:asset_type', (req, res) => {
  const schema = getAssetSchemaDescription(req.params.asset_type);
  if (!schema) return res.status(404).json({ error: `Unknown asset_type: ${req.params.asset_type}` });
  res.json(schema);
});

app.get('/schema/comparison', (req, res) => {
  res.json(getAssetSchemaDescription('comparison_doc'));
});

app.get('/schema/manual', (req, res) => {
  res.json(getAssetSchemaDescription('manual_doc'));
});

app.get('/schema/excel', (req, res) => {
  res.json(excelSchemaDescription());
});

app.get('/schema/excel/:file(*)', async (req, res) => {
  const result = await mcpToHttp(handleReadExcelHeaders({
    file_path: decodeURIComponent(req.params.file),
    sheet: req.query.sheet || undefined,
  }));
  sendToolResult(res, result);
});

app.post('/validate/:asset_type', (req, res) => {
  const validation = validateAssetSpec(req.params.asset_type, req.body.spec);
  res.status(validation.success ? 200 : 400).json(validation);
});

app.post('/preview/:asset_type', (req, res) => {
  const validation = validateAssetSpec(req.params.asset_type, req.body.spec);
  if (!validation.success) return res.status(400).json(validation);
  const asset = getAssetDefinition(req.params.asset_type);
  res.type('text/plain').send(asset.preview(validation.data));
});

app.post('/generate/from-edit', async (req, res) => {
  try {
    const result = await applyEdits(req.body);
    res.json({ success: true, ...result });
  } catch (err) {
    const status = err.statusCode || 500;
    res.status(status).json({ success: false, error: err.message });
  }
});

app.post('/generate/:asset_type', async (req, res) => {
  const result = await generateAsset(req.params.asset_type, req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);
  res.json(result);
});

app.post('/generate/:asset_type/download', async (req, res) => {
  const result = await generateAssetBuffer(req.params.asset_type, req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"; filename*=UTF-8''${encodeURIComponent(result.filename)}`);
  res.send(result.buffer);
});

app.post('/generate', async (req, res) => {
  const result = await generateAsset('comparison_doc', req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);
  res.json(result);
});

app.post('/generate/download', async (req, res) => {
  const result = await generateAssetBuffer('comparison_doc', req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"; filename*=UTF-8''${encodeURIComponent(result.filename)}`);
  res.send(result.buffer);
});

app.post('/generate/manual', async (req, res) => {
  const result = await generateAsset('manual_doc', req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);
  res.json(result);
});

app.post('/generate/manual/download', async (req, res) => {
  const result = await generateAssetBuffer('manual_doc', req.body || {});
  if (!result.success) return res.status(result.status || 500).json(result);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"; filename*=UTF-8''${encodeURIComponent(result.filename)}`);
  res.send(result.buffer);
});

app.post('/excel/append-row', async (req, res) => sendToolResult(res, await mcpToHttp(handleAppendRow(req.body))));
app.post('/excel/update-cell', async (req, res) => sendToolResult(res, await mcpToHttp(handleUpdateCell(req.body))));
app.post('/excel/edit-record', async (req, res) => sendToolResult(res, await mcpToHttp(handleAppendEditRecord(req.body))));

app.post('/append-row', async (req, res) => sendToolResult(res, await mcpToHttp(handleAppendRow(req.body))));
app.post('/update-cell', async (req, res) => sendToolResult(res, await mcpToHttp(handleUpdateCell(req.body))));
app.post('/edit-record', async (req, res) => sendToolResult(res, await mcpToHttp(handleAppendEditRecord(req.body))));

app.post('/excel/update-row', async (req, res) => sendToolResult(res, await mcpToHttp(handleUpdateRow(req.body))));
app.post('/update-row', async (req, res) => sendToolResult(res, await mcpToHttp(handleUpdateRow(req.body))));

app.get('/find-row/:file(*)', async (req, res) => {
  const result = await mcpToHttp(handleFindRow({
    file_path: decodeURIComponent(req.params.file),
    sheet: req.query.sheet || undefined,
    column: req.query.column,
    value: req.query.value,
    header_row: req.query.header_row ? Number(req.query.header_row) : undefined,
  }));
  sendToolResult(res, result);
});

app.post('/ingest', async (req, res) => {
  try {
    const result = await handleIngest(req.body);
    res.json({ success: true, ...result });
  } catch (err) {
    const status = err.statusCode || 500;
    console.error(`  [ingest] ERROR: ${err.message}`);
    if (err.cause) console.error(`  [ingest] CAUSE: ${err.cause}`);
    res.status(status).json({ success: false, error: err.message, stack: err.stack });
  }
});

app.get('/list-files', async (req, res) => {
  try {
    const result = await listManualFolders({
      folder: req.query.folder || undefined,
      extensions: req.query.extensions || undefined,
    });
    res.json({ success: true, ...result });
  } catch (err) {
    res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }
});

app.post('/read-file', async (req, res) => {
  try {
    const result = await readManualFolder({
      folder_name: req.body.folder_name,
      mode: req.body.mode || 'full',
    });
    res.json({ success: true, ...result });
  } catch (err) {
    res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }
});

// ---------------------------------------------------------------------------
// Pre-ingest: shared core logic (used by HTTP endpoint and startup prompt)
// ---------------------------------------------------------------------------

const _PREINGEST_WORD_EXTS = new Set(['.docx', '.doc']);
const _PREINGEST_EXCEL_EXTS = new Set(['.xlsx', '.xls']);

async function runPreIngestFolder(folderName) {
  const normalizedName = folderName.normalize('NFC');
  const folderPath = path.join(MANUAL_DATABASE_DIR, normalizedName);
  if (!fs.existsSync(folderPath)) {
    throw Object.assign(new Error(`Folder not found: ${normalizedName}`), { statusCode: 404 });
  }

  const files = fs.readdirSync(folderPath, { withFileTypes: true }).filter(f => f.isFile());
  let primaryFile = files.find(f => {
    const ext = path.extname(f.name).toLowerCase();
    return _PREINGEST_WORD_EXTS.has(ext) && f.name.normalize('NFC').slice(0, -ext.length) === normalizedName;
  });
  if (!primaryFile) primaryFile = files.find(f => _PREINGEST_WORD_EXTS.has(path.extname(f.name).toLowerCase()));
  if (!primaryFile) {
    throw Object.assign(new Error(`No Word document found in folder: ${normalizedName}`), { statusCode: 422 });
  }

  const primaryPath = path.join(folderPath, primaryFile.name);
  const primaryExt = path.extname(primaryFile.name).toLowerCase();
  const primaryMtime = fs.statSync(primaryPath).mtime.toISOString();

  const existing = getManifestEntry(normalizedName);
  if (existing && existing.primary_mtime === primaryMtime) {
    return { ...existing, folder_name: normalizedName, cached: true };
  }

  console.log(`  [pre-ingest] "${normalizedName}" — processing primary doc...`);
  const buffer = fs.readFileSync(primaryPath);
  const { ref_id, originalPath } = createStore(primaryExt.replace('.', ''));
  fs.writeFileSync(originalPath, buffer);

  let processed;
  if (primaryExt === '.docx') {
    processed = await processDocxFile(buffer, ref_id);
  } else {
    processed = await processDocFile(buffer, ref_id);
  }
  const image_count = (processed.images_summary || []).length;

  // Summarize document content via Vertex AI (gracefully degrades if not configured)
  let docSummary = { summary: '', key_topics: [], effective_date: null, document_type: null };
  if (processed.content) {
    try {
      console.log(`  [pre-ingest] summarizing "${normalizedName}"...`);
      docSummary = await summarizeDocument({ text: processed.content, fileName: primaryFile.name });
    } catch (err) {
      console.warn(`  [pre-ingest] summarization failed: ${err.message}`);
    }
  }

  // Process Excel attachments
  const attachment_summaries = [];
  for (const attFile of files) {
    const attExt = path.extname(attFile.name).toLowerCase();
    if (attFile.name === primaryFile.name || !_PREINGEST_EXCEL_EXTS.has(attExt)) continue;
    const attPath = path.join(folderPath, attFile.name);
    try {
      const attBuffer = fs.readFileSync(attPath);
      const attExtNoDot = attExt.replace('.', '');
      const { ref_id: attRefId, originalPath: attOrigPath } = createStore(attExtNoDot);
      fs.writeFileSync(attOrigPath, attBuffer);
      const attProcessed = await processExcelFile(attBuffer, attExtNoDot, attRefId);
      attachment_summaries.push({
        name: attFile.name.normalize('NFC'),
        type: 'excel',
        ref_id: attRefId,
        sheet_names: attProcessed.scheme.sheet_names || [],
        sheets: attProcessed.scheme.sheets || {},
      });
      console.log(`  [pre-ingest] attachment "${attFile.name}" done`);
    } catch (err) {
      console.warn(`  [pre-ingest] attachment "${attFile.name}" failed: ${err.message}`);
      attachment_summaries.push({ name: attFile.name.normalize('NFC'), type: 'excel', error: err.message });
    }
  }

  const entry = {
    ref_id,
    primary_mtime: primaryMtime,
    ingested_at: new Date().toISOString(),
    images_analyzed: true,
    image_count,
    ...docSummary,
    attachment_summaries,
  };
  setManifestEntry(normalizedName, entry);
  return { ...entry, folder_name: normalizedName, cached: false };
}

// ---------------------------------------------------------------------------
// Startup: check for un-ingested folders and optionally run pre-ingest
// ---------------------------------------------------------------------------

async function checkPendingPreIngest() {
  if (!fs.existsSync(MANUAL_DATABASE_DIR)) return;

  const entries = fs.readdirSync(MANUAL_DATABASE_DIR, { withFileTypes: true }).filter(e => e.isDirectory());
  if (entries.length === 0) return;

  const pending = [];
  for (const entry of entries) {
    const folderName = entry.name.normalize('NFC');
    const folderPath = path.join(MANUAL_DATABASE_DIR, folderName);
    const primaryFile = fs.readdirSync(folderPath, { withFileTypes: true })
      .filter(f => f.isFile())
      .find(f => _PREINGEST_WORD_EXTS.has(path.extname(f.name).toLowerCase()));
    if (!primaryFile) continue;
    const mtime = fs.statSync(path.join(folderPath, primaryFile.name)).mtime.toISOString();
    const cached = getManifestEntry(folderName);
    if (!cached || cached.primary_mtime !== mtime) pending.push(folderName);
  }

  if (pending.length === 0) {
    console.log('[pre-ingest] All folders are up to date.\n');
    return;
  }

  console.log(`\n[pre-ingest] ${pending.length} folder(s) not yet ingested or stale:`);
  pending.forEach(f => console.log(`  • ${f}`));

  if (!process.stdin.isTTY) {
    console.log('[pre-ingest] Non-interactive mode — skipping. POST /pre-ingest-folder to ingest manually.\n');
    return;
  }

  const { createInterface } = require('readline');
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  rl.question('\nRun pre-ingest now? (y/N) ', async answer => {
    rl.close();
    if (answer.trim().toLowerCase() !== 'y') {
      console.log('[pre-ingest] Skipped. Use POST /pre-ingest-folder to ingest individual folders.\n');
      return;
    }
    console.log('');
    for (const folderName of pending) {
      process.stdout.write(`[pre-ingest] "${folderName}"... `);
      try {
        const r = await runPreIngestFolder(folderName);
        console.log(r.cached ? 'already cached' : `done (${r.image_count} images)`);
      } catch (err) {
        console.log(`FAILED: ${err.message}`);
      }
    }
    console.log('\n[pre-ingest] Complete.\n');
  });
}

app.post('/pre-ingest-folder', async (req, res) => {
  try {
    const { folder_name } = req.body || {};
    if (!folder_name || typeof folder_name !== 'string') {
      return res.status(400).json({ success: false, error: 'folder_name is required' });
    }
    if (folder_name.includes('/') || folder_name.includes('\\') || folder_name.includes('..')) {
      return res.status(400).json({ success: false, error: 'folder_name must not contain path separators or ".."' });
    }
    const result = await runPreIngestFolder(folder_name);
    res.json({ success: true, ...result });
  } catch (err) {
    res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }
});

app.get('/ingest-status', (req, res) => {
  const { FILE_STORE_DIR: storeDir } = require('./config');
  const manifestPath = path.join(storeDir, 'manifest.json');
  const status = fs.existsSync(manifestPath)
    ? JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
    : {};
  res.json({ success: true, status });
});

// Must be registered AFTER all specific /schema/* routes to avoid shadowing them.
app.get('/schema/:file(*)', async (req, res) => {
  const result = await mcpToHttp(handleReadExcelHeaders({
    file_path: decodeURIComponent(req.params.file),
    sheet: req.query.sheet || undefined,
    header_row: req.query.header_row ? Number(req.query.header_row) : undefined,
  }));
  sendToolResult(res, result);
});

async function generateAsset(assetType, body) {
  if (!body.spec) return { success: false, status: 400, error: 'Missing spec' };

  const validation = validateAssetSpec(assetType, body.spec);
  if (!validation.success) return { success: false, status: 400, ...validation };

  const asset = getAssetDefinition(assetType);
  if (!asset) return { success: false, status: 404, error: `Unknown asset_type: ${assetType}` };

  const filenameBase = safeFilename(body.output_filename || asset.defaultFileName(validation.data), assetType);
  const outputPath = path.join(OUTPUT_DIR, `${filenameBase}.docx`);

  try {
    const result = await asset.generate(validation.data, outputPath);
    const filename = path.basename(result.path);
    const response = {
      success: true,
      asset_type: assetType,
      path: result.path,
      filename,
      download_url: `http://host.docker.internal:${PORT}/files/${encodeURIComponent(filename)}`,
      size_kb: Math.round(result.buffer.length / 1024),
    };
    if (body.return_base64 === true) response.base64 = result.base64;
    return response;
  } catch (error) {
    return { success: false, status: 500, error: error.message, stack: error.stack };
  }
}

async function generateAssetBuffer(assetType, body) {
  if (!body.spec) return { success: false, status: 400, error: 'Missing spec' };

  const validation = validateAssetSpec(assetType, body.spec);
  if (!validation.success) return { success: false, status: 400, ...validation };

  const asset = getAssetDefinition(assetType);
  if (!asset) return { success: false, status: 404, error: `Unknown asset_type: ${assetType}` };

  try {
    const result = await asset.generate(validation.data);
    const filenameBase = safeFilename(body.filename || body.output_filename || asset.defaultFileName(validation.data), assetType);
    return {
      success: true,
      filename: `${filenameBase}.docx`,
      buffer: result.buffer,
    };
  } catch (error) {
    return { success: false, status: 500, error: error.message, stack: error.stack };
  }
}

async function mcpToHttp(resultPromise) {
  const result = await resultPromise;
  const text = result.content?.[0]?.text || '';
  let data = text;
  try {
    data = JSON.parse(text);
  } catch (error) {
    data = { message: text };
  }
  return { success: !result.isError, status: result.isError ? 400 : 200, data };
}

function sendToolResult(res, result) {
  res.status(result.status).json(result.success ? { success: true, ...result.data } : { success: false, ...result.data });
}

app.listen(PORT, HOST, () => {
  const displayHost = HOST === '0.0.0.0' ? 'localhost' : HOST;
  const url = `http://${displayHost}:${PORT}`;
  const W = 62;
  const line  = (s = '')  => '│  ' + s + ' '.repeat(Math.max(0, W - 4 - s.replace(/\x1b\[[^m]*m/g, '').length)) + '│';
  const divider = '├' + '─'.repeat(W - 2) + '┤';

  const label  = (l, v, warn = false) =>
    line(`${C.dim}${l.padEnd(14)}${C.reset}${warn ? C.yellow : C.cyan}${v}${warn ? '  ⚠ not set in .env' : ''}${C.reset}`);

  console.log('\n' + '┌' + '─'.repeat(W - 2) + '┐');
  console.log(line(`${C.bold}${C.green}${SERVICE_NAME}${C.reset}  ${C.dim}v${VERSION}${C.reset}`));
  console.log(divider);
  console.log(label('URL', url));
  console.log(label('Output dir', OUTPUT_DIR, !process.env.OUTPUT_DIR));
  console.log(label('Manual DB', MANUAL_DATABASE_DIR, !process.env.MANUAL_DATABASE_DIR));
  console.log(divider);
  console.log(line(`${C.dim}Document${C.reset}   POST /generate/:type  POST /ingest  POST /generate/from-edit`));
  console.log(line(`${C.dim}Excel${C.reset}      POST /append-row  /update-row  GET /find-row/:file`));
  console.log(line(`${C.dim}Files${C.reset}      GET /list-files  POST /read-file  POST /pre-ingest-folder`));
  console.log(line(`${C.dim}Schema${C.reset}     GET /schema/:file  GET /schema/excel/:file`));
  console.log('└' + '─'.repeat(W - 2) + '┘\n');
  setImmediate(() => checkPendingPreIngest());
});
