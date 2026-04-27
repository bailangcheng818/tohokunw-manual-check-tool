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

// ---------------------------------------------------------------------------
// Run-level edit  (type: 'paragraph_runs')
// ---------------------------------------------------------------------------

/**
 * Build a <w:r> XML element from a run spec.
 * baseRpr: raw <w:rPr>...</w:rPr> from the first original run (font/size passthrough).
 * Spec properties (bold, underline, color) override anything in baseRpr.
 */
function buildRunXml(run, baseRpr = '') {
  // Strip formatting tags from baseRpr so only font/size/language carry over
  const cleanBase = baseRpr
    .replace(/<w:b(?:\s[^>]*)?\s*\/>/g, '')
    .replace(/<w:u[^>]*\/>/g, '')
    .replace(/<w:color[^>]*\/>/g, '');

  const props = [];
  if (run.bold)      props.push('<w:b/>');
  if (run.underline) props.push('<w:u w:val="single"/>');
  if (run.color)     props.push(`<w:color w:val="${xmlEncode(run.color.replace('#', ''))}"/>`);

  const rPrContent = cleanBase + props.join('');
  const rPr = rPrContent ? `<w:rPr>${rPrContent}</w:rPr>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${xmlEncode(run.text)}</w:t></w:r>`;
}

/**
 * Replace all runs in a paragraph with runs built from runsSpec[].
 * Preserves <w:pPr> and content outside the run region.
 * Returns modified paragraph XML, or null if no runs found.
 */
function applyRunsEdit(paraXml, runsSpec) {
  const runMatches = paraXml.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
  if (runMatches.length === 0) return null;

  const baseRpr = extractFirstRpr(runMatches[0]);
  const newRunsXml = runsSpec
    .filter(r => r.text)
    .map(r => buildRunXml(r, baseRpr))
    .join('');

  const firstRunStart = paraXml.indexOf(runMatches[0]);
  const lastRun = runMatches[runMatches.length - 1];
  const lastRunEnd = paraXml.lastIndexOf(lastRun) + lastRun.length;
  return paraXml.slice(0, firstRunStart) + newRunsXml + paraXml.slice(lastRunEnd);
}

/**
 * Apply all paragraph_runs-type edits to the document XML string.
 */
function applyRunsEdits(docXml, edits) {
  const runsEdits = edits.filter(e => e.type === 'paragraph_runs');
  if (runsEdits.length === 0) return docXml;

  for (const edit of runsEdits) {
    const paraBlocks = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
    let replaced = false;

    if (edit.xml_index != null) {
      const block = paraBlocks[edit.xml_index];
      if (block) {
        const newBlock = applyRunsEdit(block, edit.runs);
        if (newBlock !== null) {
          docXml = docXml.replace(block, newBlock);
          replaced = true;
        }
      }
    }

    if (!replaced && edit.old_text) {
      for (const block of paraBlocks) {
        if (extractFullText(block).trim() === edit.old_text.trim()) {
          const newBlock = applyRunsEdit(block, edit.runs);
          if (newBlock !== null) {
            docXml = docXml.replace(block, newBlock);
            replaced = true;
            break;
          }
        }
      }
    }

    if (!replaced) {
      console.warn(`[edit-applier] paragraph_runs: no match para_id=${edit.para_id || '?'}`);
    }
  }
  return docXml;
}

// ---------------------------------------------------------------------------
// Paragraph edit  (type: 'paragraph')
// ---------------------------------------------------------------------------

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
 * Supports xml_index-based direct addressing (primary) with old_text fallback.
 */
function applyParagraphEdits(docXml, edits) {
  const paraEdits = edits.filter(e => e.type === 'paragraph');
  if (paraEdits.length === 0) return docXml;

  for (const edit of paraEdits) {
    const paraBlocks = docXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
    let replaced = false;

    // Primary: xml_index による直接アドレス指定
    if (edit.xml_index != null) {
      const block = paraBlocks[edit.xml_index];
      if (block) {
        if (edit.old_text && extractFullText(block).trim() !== edit.old_text.trim()) {
          console.warn(`[edit-applier] para_id=${edit.para_id || edit.xml_index}: text mismatch (expected "${edit.old_text.slice(0, 40)}", found "${extractFullText(block).slice(0, 40)}")`);
        }
        const newBlock = applyParagraphEdit(block, extractFullText(block), edit.new_text);
        if (newBlock !== null) {
          docXml = docXml.replace(block, newBlock);
          replaced = true;
        }
      }
    }

    // Fallback: 従来の old_text マッチ（後方互換）
    if (!replaced && edit.old_text) {
      for (const block of paraBlocks) {
        const newBlock = applyParagraphEdit(block, edit.old_text, edit.new_text);
        if (newBlock !== null) {
          docXml = docXml.replace(block, newBlock);
          replaced = true;
          break;
        }
      }
    }

    if (!replaced) {
      console.warn(`[edit-applier] No matching paragraph for para_id=${edit.para_id || '?'} old_text="${(edit.old_text || '').slice(0, 60)}"`);
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
  docXml = applyRunsEdits(docXml, edits);
  docXml = applyTableCellEdits(docXml, edits);

  zip.file('word/document.xml', docXml);

  // Apply header / footer edits
  const hfEdits = edits.filter(e => e.type === 'header_paragraph' || e.type === 'footer_paragraph');
  for (const edit of hfEdits) {
    const zipPath = `word/${edit.part}.xml`;
    let partXml = await zip.file(zipPath)?.async('string');
    if (!partXml) {
      console.warn(`[edit-applier] header/footer part not found: ${zipPath}`);
      continue;
    }
    const mockEdits = [{ type: 'paragraph', xml_index: edit.xml_index, old_text: edit.old_text, new_text: edit.new_text, para_id: edit.para_id }];
    partXml = applyParagraphEdits(partXml, mockEdits);
    zip.file(zipPath, partXml);
  }

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
