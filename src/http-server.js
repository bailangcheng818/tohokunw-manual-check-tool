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
  handleReadExcelHeaders,
  handleUpdateCell,
} = require('./excel-tools');
const { HOST, OUTPUT_DIR, PORT, SERVICE_NAME, VERSION, safeFilename } = require('./config');
const { handleIngest } = require('./ingest');
const { applyEdits } = require('./edit-applier');

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// Request logger
app.use((req, res, next) => {
  const start = Date.now();
  const bodyPreview = req.body && Object.keys(req.body).length
    ? JSON.stringify(req.body).slice(0, 300)
    : '(empty)';
  console.log(`\n→ ${req.method} ${req.path}`);
  if (req.method !== 'GET') console.log(`  body: ${bodyPreview}`);
  res.on('finish', () => {
    const ms = Date.now() - start;
    console.log(`← ${res.statusCode} ${req.method} ${req.path} (${ms}ms)`);
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
  console.log(`[${SERVICE_NAME} HTTP] Running on http://${displayHost}:${PORT}`);
  console.log(`[${SERVICE_NAME} HTTP] Output dir: ${OUTPUT_DIR}`);
  console.log(`[${SERVICE_NAME} HTTP] GET /health, GET /schema, GET /assets`);
});
