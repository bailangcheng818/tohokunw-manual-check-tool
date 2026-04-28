'use strict';

const JSZip = require('jszip');
const fs    = require('fs');
const path  = require('path');

const { readOriginalBuffer, getStoreDir } = require('./file-store');
const { OUTPUT_DIR, PUBLIC_URL, safeFilename } = require('./config');
const {
  injectComments,
  addStrikethroughToParaXml,
  buildRedTextParaXml,
  isParaXmlSafe,
  buildFallbackParaXml,
} = require('./docx-generator');

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

function normalizeForMatch(text) {
  return String(text || '')
    .split('\n')
    .map(line => line.replace(/^#{1,6}\s+/, ''))
    .join('\n')
    .trim();
}

function extractFullText(xml) {
  const parts = [...String(xml || '').matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map(m => xmlDecode(m[1]));
  return parts.join('');
}

function extractFirstRpr(runXml) {
  const m = runXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

function getParagraphMatches(docXml) {
  return [...docXml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)]
    .map(m => ({ xml: m[0], start: m.index, end: m.index + m[0].length }));
}

function applyParagraphEdit(paraXml, oldText, newText) {
  const fullText = extractFullText(paraXml);
  if (oldText && fullText.trim() !== oldText.trim()) return null;

  const runMatches = paraXml.match(/<w:r[ >][\s\S]*?<\/w:r>/g) || [];
  if (runMatches.length === 0) return null;

  const rPr = extractFirstRpr(runMatches[0]);
  const replacementRun =
    `<w:r>${rPr}<w:t xml:space="preserve">${xmlEncode(newText)}</w:t></w:r>`;

  const firstRunStart = paraXml.indexOf(runMatches[0]);
  const lastRun       = runMatches[runMatches.length - 1];
  const lastRunEnd    = paraXml.lastIndexOf(lastRun) + lastRun.length;

  return paraXml.slice(0, firstRunStart) + replacementRun + paraXml.slice(lastRunEnd);
}

function applyParagraphEdits(docXml, edits, pendingComments) {
  const paraEdits = edits.filter(e => e.type === 'paragraph');
  if (paraEdits.length === 0) return docXml;

  // Process highest xml_index first so each old+new insertion doesn't shift
  // the indices of lower-positioned edits that are processed afterwards.
  paraEdits.sort((a, b) => {
    if (a.xml_index == null && b.xml_index == null) return 0;
    if (a.xml_index == null) return 1;
    if (b.xml_index == null) return -1;
    return b.xml_index - a.xml_index;
  });

  for (const edit of paraEdits) {
    const paraBlocks = getParagraphMatches(docXml);
    let replaced = false;

    if (edit.xml_index != null) {
      const blockInfo = paraBlocks[edit.xml_index];
      if (blockInfo) {
        const block = blockInfo.xml;
        const textOk = !edit.old_text ||
          normalizeForMatch(extractFullText(block)) === normalizeForMatch(edit.old_text);
        if (!textOk) {
          // xml_index points to the wrong paragraph — fall through to text matching
          console.warn(`[comparison-test] para_id=${edit.para_id || edit.xml_index}: text mismatch at xml_index ${edit.xml_index}, falling back to text search (expected "${edit.old_text.slice(0, 40)}", found "${extractFullText(block).slice(0, 40)}")`);
        } else {
          const oldPara = isParaXmlSafe(block)
            ? addStrikethroughToParaXml(block)
            : buildFallbackParaXml(block, edit.old_text || extractFullText(block));
          const newPara = buildRedTextParaXml(edit.new_text || '', block);
          docXml = docXml.slice(0, blockInfo.start) + oldPara + newPara + docXml.slice(blockInfo.end);
          if (edit.rationale) pendingComments.push({ text: edit.rationale, anchor: edit.new_text || '' });
          replaced = true;
        }
      }
    }

    if (!replaced && edit.old_text) {
      const normalizedOld = normalizeForMatch(edit.old_text);

      // single-paragraph match with normalization
      for (const { xml: block } of paraBlocks) {
        if (normalizeForMatch(extractFullText(block)) !== normalizedOld) continue;
        const oldPara = isParaXmlSafe(block)
          ? addStrikethroughToParaXml(block)
          : buildFallbackParaXml(block, edit.old_text || extractFullText(block));
        const newPara = buildRedTextParaXml(edit.new_text || '', block);
        docXml = docXml.replace(block, oldPara + newPara);
        if (edit.rationale) pendingComments.push({ text: edit.rationale, anchor: edit.new_text || '' });
        replaced = true;
        break;
      }

      // multi-paragraph match: strikethrough all matched paragraphs, insert new text after last
      if (!replaced) {
        const oldLines = normalizedOld.split('\n').filter(l => l.trim());
        if (oldLines.length >= 2) {
          for (let si = 0; si < paraBlocks.length; si++) {
            let li = 0;
            let ei = si - 1;
            for (let pi = si; pi < paraBlocks.length && li < oldLines.length; pi++) {
              const pt = normalizeForMatch(extractFullText(paraBlocks[pi].xml)).trim();
              if (!pt) continue;
              if (pt !== oldLines[li]) break;
              li++;
              ei = pi;
            }
            if (li === oldLines.length) {
              const firstInfo = paraBlocks[si];
              const lastInfo  = paraBlocks[ei];
              const oldParasXml = paraBlocks.slice(si, ei + 1)
                .map(({ xml }) => isParaXmlSafe(xml)
                  ? addStrikethroughToParaXml(xml)
                  : buildFallbackParaXml(xml, extractFullText(xml))
                ).join('');
              const newPara = buildRedTextParaXml(edit.new_text || '', firstInfo.xml);
              docXml = docXml.slice(0, firstInfo.start) + oldParasXml + newPara + docXml.slice(lastInfo.end);
              if (edit.rationale) pendingComments.push({ text: edit.rationale, anchor: edit.new_text || '' });
              replaced = true;
              break;
            }
          }
        }
      }
    }

    if (!replaced) {
      console.warn(`[comparison-test] No matching paragraph for para_id=${edit.para_id || '?'} old_text="${(edit.old_text || '').slice(0, 60)}"`);
    }
  }

  return docXml;
}

function applySimpleParagraphEdits(docXml, edits) {
  for (const edit of edits) {
    const paraBlocks = getParagraphMatches(docXml);
    let replaced = false;

    if (edit.xml_index != null && paraBlocks[edit.xml_index]) {
      const blockInfo = paraBlocks[edit.xml_index];
      const block = blockInfo.xml;
      const newBlock = applyParagraphEdit(block, extractFullText(block), edit.new_text || '');
      if (newBlock) {
        docXml = docXml.slice(0, blockInfo.start) + newBlock + docXml.slice(blockInfo.end);
        replaced = true;
      }
    }

    if (!replaced && edit.old_text) {
      for (const { xml: block } of paraBlocks) {
        const newBlock = applyParagraphEdit(block, edit.old_text, edit.new_text || '');
        if (newBlock) {
          docXml = docXml.replace(block, newBlock);
          break;
        }
      }
    }
  }
  return docXml;
}

function extractTopLevelTables(docXml) {
  const tables = [];
  let i = 0;
  while (i < docXml.length) {
    const start = docXml.indexOf('<w:tbl', i);
    if (start === -1) break;

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

function buildMarkedCellContent(targetCell, newText) {
  const paraMatches = targetCell.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  const basePara = paraMatches[0] || '';
  const oldParas = paraMatches.length > 0
    ? paraMatches.map(p => (
      isParaXmlSafe(p)
        ? addStrikethroughToParaXml(p)
        : buildFallbackParaXml(p, extractFullText(p))
    )).join('')
    : buildFallbackParaXml('', extractFullText(targetCell));
  return oldParas + buildRedTextParaXml(newText || '', basePara);
}

function applyTableCellEdit(tblXml, row, col, newText) {
  const rowMatches = tblXml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
  if (row == null || col == null || row >= rowMatches.length) return null;

  const targetRow = rowMatches[row];
  const cellMatches = targetRow.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || [];
  if (col >= cellMatches.length) return null;

  const targetCell = cellMatches[col];
  const paraMatches = targetCell.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  let newCell;

  if (paraMatches.length > 0) {
    const firstParaStart = targetCell.indexOf(paraMatches[0]);
    const lastPara       = paraMatches[paraMatches.length - 1];
    const lastParaEnd    = targetCell.lastIndexOf(lastPara) + lastPara.length;
    newCell = targetCell.slice(0, firstParaStart)
      + buildMarkedCellContent(targetCell, newText)
      + targetCell.slice(lastParaEnd);
  } else {
    const insertAt = targetCell.lastIndexOf('</w:tc>');
    if (insertAt === -1) return null;
    newCell = targetCell.slice(0, insertAt)
      + buildMarkedCellContent(targetCell, newText)
      + targetCell.slice(insertAt);
  }

  const newRow = targetRow.replace(targetCell, newCell);
  return tblXml.replace(targetRow, newRow);
}

function applyTableCellEdits(docXml, edits, pendingComments) {
  const cellEdits = edits.filter(e => e.type === 'table_cell');
  if (cellEdits.length === 0) return docXml;

  for (const edit of cellEdits) {
    const tables = extractTopLevelTables(docXml);
    if (edit.table_index == null || edit.table_index >= tables.length) {
      console.warn(`[comparison-test] table_index ${edit.table_index} out of range (${tables.length} tables)`);
      continue;
    }

    const table = tables[edit.table_index];
    const newTblXml = applyTableCellEdit(table.xml, edit.row, edit.col, edit.new_text || '');
    if (!newTblXml) {
      console.warn(`[comparison-test] Could not find cell [${edit.row},${edit.col}] in table ${edit.table_index}`);
      continue;
    }
    docXml = docXml.slice(0, table.start) + newTblXml + docXml.slice(table.end);
    if (edit.rationale) pendingComments.push({ text: edit.rationale, anchor: edit.new_text || '' });
  }

  return docXml;
}

async function applyComparisonTestEdits({ ref_id, edits, output_filename, return_base64 = false } = {}) {
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

  getStoreDir(ref_id);

  const originalBuffer = readOriginalBuffer(ref_id);
  const zip = await JSZip.loadAsync(originalBuffer);

  let docXml = await zip.file('word/document.xml').async('string');
  const pendingComments = [];

  docXml = applyParagraphEdits(docXml, edits, pendingComments);
  docXml = applyTableCellEdits(docXml, edits, pendingComments);

  // table_row_append: clone last row, fill cells with red text to show addition
  for (const edit of edits.filter(e => e.type === 'table_row_append')) {
    const tables = extractTopLevelTables(docXml);
    const idx = edit.table_index ?? 0;
    if (idx >= tables.length) {
      console.warn(`[comparison-test] table_row_append: table_index ${idx} out of range`);
      continue;
    }
    const cellTexts = String(edit.new_text || '').split('|').map(s => s.trim());
    const table = tables[idx];
    const rows = table.xml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    if (rows.length === 0) continue;
    const templateRow = rows[rows.length - 1];
    const cells = templateRow.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || [];
    let newRowXml = templateRow;
    cells.forEach((cell, i) => {
      const text = cellTexts[i] ?? '';
      const paras = cell.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
      const basePara = paras[0] || '';
      const redPara = buildRedTextParaXml(text, basePara);
      const firstStart = basePara ? cell.indexOf(basePara) : cell.lastIndexOf('</w:tc>');
      const lastEnd = paras.length > 0
        ? cell.lastIndexOf(paras[paras.length - 1]) + paras[paras.length - 1].length
        : firstStart;
      const newCell = cell.slice(0, firstStart) + redPara + cell.slice(lastEnd);
      newRowXml = newRowXml.replace(cell, newCell);
    });
    const newTblXml = table.xml.replace('</w:tbl>', newRowXml + '</w:tbl>');
    docXml = docXml.slice(0, table.start) + newTblXml + docXml.slice(table.end);
    if (edit.rationale) pendingComments.push({ text: edit.rationale, anchor: edit.new_text || '' });
    console.log(`[comparison-test] table_row_append: appended red row to table ${idx}`);
  }

  zip.file('word/document.xml', docXml);

  const hfEdits = edits.filter(e => e.type === 'header_paragraph' || e.type === 'footer_paragraph');
  for (const edit of hfEdits) {
    const zipPath = `word/${edit.part}.xml`;
    let partXml = await zip.file(zipPath)?.async('string');
    if (!partXml) {
      console.warn(`[comparison-test] header/footer part not found: ${zipPath}`);
      continue;
    }
    partXml = applySimpleParagraphEdits(partXml, [edit]);
    zip.file(zipPath, partXml);
  }

  let outputBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  if (pendingComments.length > 0) {
    outputBuffer = await injectComments(outputBuffer, pendingComments.map((c, i) => ({ ...c, id: i })));
  }

  const filenameBase = safeFilename(output_filename || `comparison_test_${ref_id.slice(0, 8)}`);
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

module.exports = { applyComparisonTestEdits };
