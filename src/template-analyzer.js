'use strict';

const JSZip = require('jszip');

function xmlDecode(value) {
  return String(value || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function stripTags(value) {
  return xmlDecode(String(value || '').replace(/<[^>]+>/g, ''));
}

function normalizeAssetType(name) {
  return String(name || 'generated_asset')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'generated_asset';
}

function detectAlignment(paragraphXml) {
  const match = paragraphXml.match(/<w:jc[^>]+w:val="([^"]+)"/);
  const value = match?.[1];
  if (value === 'both') return 'justify';
  if (value === 'center') return 'center';
  if (value === 'right') return 'right';
  return 'left';
}

function detectSize(paragraphXml) {
  const sizes = [...paragraphXml.matchAll(/<w:sz[^>]+w:val="(\d+)"/g)].map((m) => Number(m[1]));
  if (sizes.length === 0) return null;
  return Math.max(...sizes);
}

function extractRuns(paragraphXml) {
  const runs = [];
  const runMatches = paragraphXml.match(/<w:r[\s\S]*?<\/w:r>/g) || [];

  for (const runXml of runMatches) {
    const textParts = [...runXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => {
      const decoded = xmlDecode(m[1]);
      // Strip entity-decoded OOXML namespace tags that leaked into text content
      return decoded.replace(/<\/?[a-zA-Z][a-zA-Z0-9]*:[^>]{0,200}>/g, '');
    });
    const text = textParts.join('');
    if (!text) continue;

    const colorMatch = runXml.match(/<w:color[^>]+w:val="([^"]+)"/);
    runs.push({
      text,
      bold: /<w:b(?:\/>| )/.test(runXml),
      underline: /<w:u(?:\/>| )/.test(runXml),
      color: colorMatch?.[1] ? colorMatch[1].toUpperCase() : undefined,
    });
  }

  return runs;
}

function classifyParagraph(paragraph) {
  const text = paragraph.text.trim();
  if (!text) return 'empty';
  if (paragraph.align === 'center' && (paragraph.size || 0) >= 28) return 'title';
  if (paragraph.align === 'center') return 'centered_line';
  if (/^[0-9０-９一二三四五六七八九十]+\s*[.．、]/.test(text)) return 'numbered_heading';
  if (/^【.+】$/.test(text)) return 'bracket_heading';
  if (paragraph.runs.some((run) => run.bold)) return 'emphasized_line';
  return 'body';
}

function suggestFieldKey(text, index) {
  const normalized = text
    .replace(/[【】()[\]（）]/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5ぁ-んァ-ヶ]+/g, '_')
    .replace(/^_+|_+$/g, '');

  return normalized || `field_${index + 1}`;
}

function buildBlueprint(paragraphs, formatName) {
  const blocks = paragraphs
    .filter((paragraph) => paragraph.kind !== 'empty')
    .slice(0, 80)
    .map((paragraph, index) => ({
      id: `block_${index + 1}`,
      role: paragraph.kind,
      field_key: suggestFieldKey(paragraph.text, index),
      source_text: paragraph.text,
      formatting: {
        align: paragraph.align,
        size: paragraph.size,
        bold: paragraph.runs.some((run) => run.bold),
        underline: paragraph.runs.some((run) => run.underline),
      },
      example: {
        text: paragraph.text,
      },
    }));

  return {
    format_name: formatName,
    asset_type: normalizeAssetType(formatName),
    schema_style: 'block_document',
    recommended_top_level_fields: [
      { name: 'doc_title', type: 'string', required: false },
      { name: 'blocks', type: 'array', required: true },
    ],
    blocks,
  };
}

async function analyzeDocxTemplate({ buffer, formatName = 'Generated Format' }) {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file('word/document.xml')?.async('string');

  if (!documentXml) {
    throw new Error('word/document.xml was not found in the docx file');
  }

  const paragraphXmlList = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const paragraphs = paragraphXmlList.map((paragraphXml, index) => {
    const runs = extractRuns(paragraphXml);
    const text = runs.map((run) => run.text).join('') || stripTags(paragraphXml);
    const align = detectAlignment(paragraphXml);
    const size = detectSize(paragraphXml);

    const paragraph = {
      index,
      text,
      align,
      size,
      runs,
    };

    return {
      ...paragraph,
      kind: classifyParagraph(paragraph),
    };
  });

  const nonEmptyParagraphs = paragraphs.filter((paragraph) => paragraph.text.trim().length > 0);
  const headings = nonEmptyParagraphs.filter((paragraph) => paragraph.kind !== 'body' && paragraph.kind !== 'centered_line');
  const blueprint = buildBlueprint(nonEmptyParagraphs, formatName);

  return {
    format_name: formatName,
    asset_type: blueprint.asset_type,
    paragraph_count: nonEmptyParagraphs.length,
    headings: headings.slice(0, 20).map((paragraph) => ({
      index: paragraph.index,
      text: paragraph.text,
      kind: paragraph.kind,
    })),
    paragraphs: nonEmptyParagraphs.slice(0, 120),
    suggested_asset_blueprint: blueprint,
  };
}

async function analyzeDocxTemplateBase64(base64Docx, formatName) {
  const buffer = Buffer.from(base64Docx, 'base64');
  return analyzeDocxTemplate({ buffer, formatName });
}

module.exports = {
  analyzeDocxTemplate,
  analyzeDocxTemplateBase64,
  classifyParagraph,
  extractRuns,
  normalizeAssetType,
  xmlDecode,
};
