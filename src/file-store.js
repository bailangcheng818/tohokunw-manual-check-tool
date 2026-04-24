'use strict';

const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const { FILE_STORE_DIR } = require('./config');

const MANIFEST_PATH = path.join(FILE_STORE_DIR, 'manifest.json');

function _readManifest() {
  if (!fs.existsSync(MANIFEST_PATH)) return {};
  try { return JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8')); } catch { return {}; }
}

function getManifestEntry(folderName) {
  return _readManifest()[folderName] || null;
}

function setManifestEntry(folderName, entry) {
  fs.mkdirSync(FILE_STORE_DIR, { recursive: true });
  const manifest = _readManifest();
  manifest[folderName] = entry;
  fs.writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2), 'utf8');
}

function readImagesMeta(ref_id) {
  const imagesDir = path.join(FILE_STORE_DIR, ref_id, 'images');
  if (!fs.existsSync(imagesDir)) return [];
  const allFiles = fs.readdirSync(imagesDir);
  return allFiles
    .filter(f => f.endsWith('_meta.json'))
    .sort()
    .map(f => {
      try {
        const meta = JSON.parse(fs.readFileSync(path.join(imagesDir, f), 'utf8'));
        const base = f.replace('_meta.json', '');
        const imgFile = allFiles.find(x => x.startsWith(base + '.') && !x.endsWith('_meta.json'));
        const ref = imgFile || base + '.png';
        return { ref, ...meta };
      } catch { return null; }
    })
    .filter(Boolean);
}

function _storeDir(ref_id) {
  return path.join(FILE_STORE_DIR, ref_id);
}

function _imagesDir(ref_id) {
  return path.join(_storeDir(ref_id), 'images');
}

function _imgPad(index) {
  return String(index).padStart(3, '0');
}

/**
 * Create a new file store entry.
 * @param {string} ext - file extension without dot, e.g. 'docx' or 'xls'
 * @returns {{ ref_id: string, storeDir: string, originalPath: string }}
 */
function createStore(ext) {
  const ref_id = uuidv4();
  const storeDir = _storeDir(ref_id);
  fs.mkdirSync(storeDir, { recursive: true });
  fs.mkdirSync(_imagesDir(ref_id), { recursive: true });
  const originalPath = path.join(storeDir, `original.${ext}`);
  return { ref_id, storeDir, originalPath };
}

/**
 * Resolve the store directory for an existing ref_id.
 * Throws if it does not exist.
 * @param {string} ref_id
 * @returns {string} storeDir
 */
function getStoreDir(ref_id) {
  const storeDir = _storeDir(ref_id);
  if (!fs.existsSync(storeDir)) {
    const err = new Error(`ref_id not found: ${ref_id}`);
    err.statusCode = 404;
    throw err;
  }
  return storeDir;
}

function writeScheme(ref_id, obj) {
  const storeDir = _storeDir(ref_id);
  fs.writeFileSync(path.join(storeDir, 'scheme.json'), JSON.stringify(obj, null, 2), 'utf8');
}

function writeContent(ref_id, text) {
  const storeDir = _storeDir(ref_id);
  fs.writeFileSync(path.join(storeDir, 'content.txt'), text, 'utf8');
}

/**
 * Write an image buffer to images/img_NNN.{ext}.
 * @param {string} ref_id
 * @param {number} index - 1-based index
 * @param {Buffer} buffer
 * @param {string} ext - e.g. 'png', 'jpg'
 * @returns {string} absolute file path
 */
function writeImage(ref_id, index, buffer, ext) {
  const imgPath = path.join(_imagesDir(ref_id), `img_${_imgPad(index)}.${ext}`);
  fs.writeFileSync(imgPath, buffer);
  return imgPath;
}

/**
 * Write image metadata JSON to images/img_NNN_meta.json.
 * @param {string} ref_id
 * @param {number} index - 1-based index
 * @param {object} meta - { label, summary, ... }
 */
function writeImageMeta(ref_id, index, meta) {
  const metaPath = path.join(_imagesDir(ref_id), `img_${_imgPad(index)}_meta.json`);
  fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2), 'utf8');
}

/**
 * Read the original file buffer for a given ref_id.
 * @param {string} ref_id
 * @returns {Buffer}
 */
function readOriginalBuffer(ref_id) {
  const storeDir = getStoreDir(ref_id);
  const files = fs.readdirSync(storeDir).filter(f => f.startsWith('original.'));
  if (files.length === 0) {
    throw new Error(`No original file found for ref_id: ${ref_id}`);
  }
  return fs.readFileSync(path.join(storeDir, files[0]));
}

module.exports = {
  createStore,
  getStoreDir,
  writeScheme,
  writeContent,
  writeImage,
  writeImageMeta,
  readOriginalBuffer,
  getManifestEntry,
  setManifestEntry,
  readImagesMeta,
};
