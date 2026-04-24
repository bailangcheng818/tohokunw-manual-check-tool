'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const { readOriginalBuffer, getStoreDir } = require('./file-store');
const { OUTPUT_DIR, PUBLIC_URL, safeFilename } = require('./config');

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

function xmlDecode(value) {
  return String(value || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function xmlEncode(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Extract concatenated plain text from a paragraph or cell XML block.
 */
function extractFullText(xml) {
  const parts = [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map(m => xmlDecode(m[1]));
  return parts.join('');
}

// ---------------------------------------------------------------------------
// Paragraph edit  (type: 'paragraph')
// ---------------------------------------------------------------------------

/**
 * Find the first <w:rPr>...</w:rPr> block within a run, or empty string.
 */
function extractFirstRpr(runXml) {
  const m = runXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

/**
 * Replace paragraph text using "keep first run's formatting" strategy.
 * Returns the modified paragraph XML, or null if old_text doesn't match.
 */
function applyParagraphEdit(paraXml, oldText, newText) {
  const fullText = extractFullText(paraXml);
  if (fullText.trim() !== oldText.trim()) return null;

  // Find all runs
  const runMatches = paraXml.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
  if (runMatches.length === 0) return null;

  const firstRun = runMatches[0];
  const rPr = extractFirstRpr(firstRun);

  // Build replacement: single run with original rPr and new text
  const replacementRun =
    `<w:r>${rPr}<w:t xml:space="preserve">${xmlEncode(newText)}</w:t></w:r>`;

  // Replace all runs in the paragraph with just the replacement run
  // We replace from the first run start to the last run end
  const firstRunStart = paraXml.indexOf(runMatches[0]);
  const lastRun = runMatches[runMatches.length - 1];
  const lastRunEnd = paraXml.lastIndexOf(lastRun) + lastRun.length;

  return paraXml.slice(0, firstRunStart) + replacementRun + paraXml.slice(lastRunEnd);
}

/**
 * Apply all paragraph-type edits to the document XML string.
 */
function applyParagraphEdits(docXml, edits) {
  const paraEdits = edits.filter(e => e.type === 'paragraph');
  if (paraEdits.length === 0) return docXml;

  // Split into paragraph blocks preserving surrounding content
  // We process by finding and replacing individual <w:p> blocks
  for (const edit of paraEdits) {
    const paraBlocks = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
    let replaced = false;

    for (const block of paraBlocks) {
      const newBlock = applyParagraphEdit(block, edit.old_text, edit.new_text);
      if (newBlock !== null) {
        // Replace first matching occurrence only
        docXml = docXml.replace(block, newBlock);
        replaced = true;
        break;
      }
    }

    if (!replaced) {
      console.warn(`[edit-applier] No matching paragraph for old_text: "${edit.old_text.slice(0, 60)}"`);
    }
  }

  return docXml;
}

// ---------------------------------------------------------------------------
// Table cell edit  (type: 'table_cell')
// ---------------------------------------------------------------------------

/**
 * Extract top-level <w:tbl> blocks (skip nested tables).
 * Returns array of { start, end, xml } positions in the docXml string.
 */
function extractTopLevelTables(docXml) {
  const tables = [];
  let i = 0;
  while (i < docXml.length) {
    const start = docXml.indexOf('<w:tbl', i);
    if (start === -1) break;

    // Find matching </w:tbl> by counting nesting
    let depth = 0;
    let pos = start;
    while (pos < docXml.length) {
      const openIdx = docXml.indexOf('<w:tbl', pos + 1);
      const closeIdx = docXml.indexOf('</w:tbl>', pos);
      if (closeIdx === -1) break;

      if (openIdx !== -1 && openIdx < closeIdx) {
        depth++;
        pos = openIdx;
      } else {
        if (depth === 0) {
          const end = closeIdx + '</w:tbl>'.length;
          tables.push({ start, end, xml: docXml.slice(start, end) });
          i = end;
          break;
        }
        depth--;
        pos = closeIdx + 1;
      }
    }
    if (pos >= docXml.length) break;
  }
  return tables;
}

/**
 * Replace a cell's text content within a table XML string.
 * row and col are 0-based.
 */
function applyTableCellEdit(tblXml, row, col, newText) {
  const rowMatches = tblXml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
  if (row >= rowMatches.length) return null;

  const targetRow = rowMatches[row];
  const cellMatches = targetRow.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || [];
  if (col >= cellMatches.length) return null;

  const targetCell = cellMatches[col];

  // Replace text in the cell: keep first run's rPr, replace <w:t>
  const runMatches = targetCell.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
  let newCell;
  if (runMatches.length > 0) {
    const firstRun = runMatches[0];
    const rPr = extractFirstRpr(firstRun);
    const replacementRun =
      `<w:r>${rPr}<w:t xml:space="preserve">${xmlEncode(newText)}</w:t></w:r>`;
    const firstRunStart = targetCell.indexOf(runMatches[0]);
    const lastRun = runMatches[runMatches.length - 1];
    const lastRunEnd = targetCell.lastIndexOf(lastRun) + lastRun.length;
    newCell = targetCell.slice(0, firstRunStart) + replacementRun + targetCell.slice(lastRunEnd);
  } else {
    // No runs — inject a minimal run
    const insertAt = targetCell.lastIndexOf('</w:tc>');
    newCell =
      targetCell.slice(0, insertAt) +
      `<w:p><w:r><w:t xml:space="preserve">${xmlEncode(newText)}</w:t></w:r></w:p>` +
      targetCell.slice(insertAt);
  }

  const newRow = targetRow.replace(targetCell, newCell);
  return tblXml.replace(targetRow, newRow);
}

/**
 * Apply all table_cell-type edits to the document XML string.
 */
function applyTableCellEdits(docXml, edits) {
  const cellEdits = edits.filter(e => e.type === 'table_cell');
  if (cellEdits.length === 0) return docXml;

  for (const edit of cellEdits) {
    const tables = extractTopLevelTables(docXml);
    const { table_index, row, col, new_text } = edit;

    if (table_index >= tables.length) {
      console.warn(`[edit-applier] table_index ${table_index} out of range (${tables.length} tables)`);
      continue;
    }

    const table = tables[table_index];
    const newTblXml = applyTableCellEdit(table.xml, row, col, new_text);
    if (newTblXml === null) {
      console.warn(`[edit-applier] Could not find cell [${row},${col}] in table ${table_index}`);
      continue;
    }

    docXml = docXml.slice(0, table.start) + newTblXml + docXml.slice(table.end);
  }

  return docXml;
}

// ---------------------------------------------------------------------------
// Main: applyEdits
// ---------------------------------------------------------------------------

/**
 * Apply edit diff to a stored original DOCX and return the new file.
 *
 * @param {object} params
 * @param {string} params.ref_id
 * @param {Array}  params.edits
 * @param {string} [params.output_filename]
 * @returns {Promise<{path: string, filename: string, download_url: string, size_kb: number, base64: string}>}
 */
async function applyEdits({ ref_id, edits, output_filename, return_base64 = false } = {}) {
  if (!ref_id) {
    const err = new Error('Missing ref_id');
    err.statusCode = 400;
    throw err;
  }
  if (!Array.isArray(edits) || edits.length === 0) {
    const err = new Error('Missing or empty edits array');
    err.statusCode = 400;
    throw err;
  }

  // Confirm the store exists (throws 404 if not)
  getStoreDir(ref_id);

  const originalBuffer = readOriginalBuffer(ref_id);
  const zip = await JSZip.loadAsync(originalBuffer);

  let docXml = await zip.file('word/document.xml').async('string');

  // Apply edits
  docXml = applyParagraphEdits(docXml, edits);
  docXml = applyTableCellEdits(docXml, edits);

  zip.file('word/document.xml', docXml);

  const outputBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

  const filenameBase = safeFilename(output_filename || `edited_${ref_id.slice(0, 8)}`);
  const filename = `${filenameBase}.docx`;
  const outputPath = path.join(OUTPUT_DIR, filename);
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  fs.writeFileSync(outputPath, outputBuffer);

  const result = {
    path: outputPath,
    filename,
    download_url: `${PUBLIC_URL}/files/${encodeURIComponent(filename)}`,
    size_kb: Math.round(outputBuffer.length / 1024),
  };
  if (return_base64) result.base64 = outputBuffer.toString('base64');
  return result;
}

module.exports = { applyEdits };
