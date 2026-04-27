'use strict';

const path = require('path');
const os = require('os');

const SERVICE_NAME = 'tohokunw-manual-check-tool';
const VERSION = '1.0.0';

const OUTPUT_DIR = process.env.OUTPUT_DIR
  || path.join(os.homedir(), 'Desktop', 'tohokunw-manual-check-output');

const FILE_STORE_DIR = process.env.FILE_STORE_DIR
  || path.join(OUTPUT_DIR, 'file_store');

// Separate read-only source directory for manual folders (file discovery).
// Generated documents (OUTPUT_DIR) and manual source files are intentionally kept apart.
const MANUAL_DATABASE_DIR = process.env.MANUAL_DATABASE_DIR
  || path.join(os.homedir(), 'Desktop', 'tohokunw-manual-database');

const PORT = parseInt(process.env.PORT || '3456', 10);
const HOST = process.env.HOST || '0.0.0.0';
const PUBLIC_URL = (process.env.PUBLIC_URL || `http://localhost:${PORT}`).replace(/\/$/, '');

function resolveOutputPath(filePath) {
  if (!filePath) return filePath;
  return path.isAbsolute(filePath) ? filePath : path.join(OUTPUT_DIR, filePath);
}

// Resolve a relative path against MANUAL_DATABASE_DIR.
// Used by Excel tools so files like 気づき管理表.xlsx land in the manual database.
function resolveManualPath(filePath) {
  if (!filePath) return filePath;
  return path.isAbsolute(filePath) ? filePath : path.join(MANUAL_DATABASE_DIR, filePath);
}

function safeFilename(name, fallback = 'document') {
  return String(name || fallback)
    .replace(/[\\/:*?"<>|　]/g, '_')
    .substring(0, 80);
}

module.exports = {
  FILE_STORE_DIR,
  HOST,
  MANUAL_DATABASE_DIR,
  OUTPUT_DIR,
  PORT,
  PUBLIC_URL,
  SERVICE_NAME,
  VERSION,
  resolveManualPath,
  resolveOutputPath,
  safeFilename,
};
