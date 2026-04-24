'use strict';

const { z } = require('zod');

const { generateComparisonDoc } = require('./docx-generator');
const { generateManualDoc } = require('./manual-generator');
const {
  DocSpecSchema,
  ManualDocSchema,
  SCHEMA_DESCRIPTION,
} = require('./schema');

function previewComparisonSpec(spec = {}) {
  const lines = [`Preview: ${spec.doc_title || '(no title)'}`, ''];

  for (const section of (spec.sections || [])) {
    lines.push(`=== ${section.title} [${section.status}] ===`);
    lines.push('[old]');
    for (const p of (section.old_paragraphs || [])) {
      lines.push(`- ${p.text || (p.segments || []).map((s) => s.text).join('')}`);
    }
    lines.push('[new]');
    for (const p of (section.new_paragraphs || [])) {
      lines.push(`- ${p.text || (p.segments || []).map((s) => s.text).join('')}`);
    }
    if (section.notes?.length) {
      lines.push('[notes]');
      for (const note of section.notes) lines.push(`- ${note}`);
    }
    if (section.comments?.length) {
      lines.push('[comments]');
      for (const comment of section.comments) lines.push(`- "${comment.anchor}" => ${comment.text}`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

function previewManualSpec(spec = {}) {
  const lines = [`Preview: ${spec.doc_title || '(no title)'}`];
  if (spec.doc_subtitle) lines.push(`Subtitle: ${spec.doc_subtitle}`);
  if (spec.revision) lines.push(`Revision: ${spec.revision}`);
  if (spec.company || spec.department) lines.push(`Org: ${[spec.company, spec.department].filter(Boolean).join(' / ')}`);
  lines.push('');

  for (const section of (spec.sections || [])) {
    lines.push(`=== ${section.heading || section.id} ===`);
    for (const p of (section.paragraphs || [])) {
      lines.push(`- ${p.text || (p.segments || []).map((s) => s.text).join('')}`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

const ASSET_TYPES = {
  comparison_doc: {
    asset_type: 'comparison_doc',
    title: 'Comparison Doc',
    description: 'A3 landscape before/after comparison document with old/new/note columns.',
    schema: DocSpecSchema,
    schemaDescription: SCHEMA_DESCRIPTION,
    defaultFileName: (spec) => spec.doc_title || 'comparison',
    preview: previewComparisonSpec,
    generate: generateComparisonDoc,
  },
  manual_doc: {
    asset_type: 'manual_doc',
    title: 'Manual Doc',
    description: 'A4 portrait manual/specification document with cover, history, and body sections.',
    schema: ManualDocSchema,
    schemaDescription: {
      description: 'Manual/specification Word document',
      required_fields: ['doc_title', 'sections'],
      schema: {
        doc_title: 'string - document title',
        doc_subtitle: 'string? - subtitle',
        revision: 'string? - revision label',
        company: 'string? - company name',
        department: 'string? - department name',
        history: '{ established?: HistoryEntry[], revised?: RevisionEntry[] }?',
        sections: [
          {
            id: 'string - unique section ID',
            heading: 'string? - section heading',
            paragraphs: ['ParagraphSchema[] - text or formatted segments'],
          },
        ],
      },
      example: {
        doc_title: '通信関係請負工事共通仕様書',
        doc_subtitle: '仕様書',
        revision: '第3回改正',
        sections: [
          {
            id: 'overview',
            heading: '1. 総則',
            paragraphs: [{ text: '本仕様書は、通信関係請負工事に適用する。' }],
          },
        ],
      },
    },
    defaultFileName: (spec) => spec.doc_title || 'manual',
    preview: previewManualSpec,
    generate: generateManualDoc,
  },
};

const AssetTypeSchema = z.enum(Object.keys(ASSET_TYPES));

function listAssetTypes() {
  return Object.values(ASSET_TYPES).map((asset) => ({
    asset_type: asset.asset_type,
    title: asset.title,
    description: asset.description,
  }));
}

function getAssetDefinition(assetType) {
  return ASSET_TYPES[assetType] || null;
}

function getAssetSchemaDescription(assetType) {
  const asset = getAssetDefinition(assetType);
  return asset ? asset.schemaDescription : null;
}

function validateAssetSpec(assetType, spec) {
  const asset = getAssetDefinition(assetType);
  if (!asset) {
    return { success: false, error: `Unknown asset_type: ${assetType}` };
  }

  const parsed = asset.schema.safeParse(spec);
  if (!parsed.success) {
    return {
      success: false,
      error: 'Invalid spec',
      details: parsed.error.issues.map((issue) => ({
        path: issue.path.join('.'),
        message: issue.message,
      })),
    };
  }

  return { success: true, data: parsed.data };
}

module.exports = {
  ASSET_TYPES,
  AssetTypeSchema,
  getAssetDefinition,
  getAssetSchemaDescription,
  listAssetTypes,
  validateAssetSpec,
};
