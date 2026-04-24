/**
 * schema.js
 * Zod schema for the document specification JSON.
 * Used for validation in both MCP tools and HTTP API.
 */

'use strict';

const { z } = require('zod');

// ─── Coerce "true"/"false" strings to booleans (LLM output guard) ─────────────
const boolField = z.union([
  z.boolean(),
  z.string().transform(v => v === 'true' ? true : v === 'false' ? false : undefined),
]).optional();

// ─── Shared color field (accepts "red", "black", "RRGGBB", or "#RRGGBB") ──────
const colorField = z.string().optional().transform(v => {
  if (!v) return v;
  if (v === 'black' || v === 'red') return v;
  const hex = v.startsWith('#') ? v.slice(1) : v;
  return /^[0-9A-Fa-f]{6}$/.test(hex) ? hex : v;
}).pipe(z.enum(['black', 'red']).or(z.string().regex(/^[0-9A-Fa-f]{6}$/)).optional());

// ─── Paragraph segment (mixed formatting within one paragraph) ────────────────
const SegmentSchema = z.object({
  text:      z.string(),
  bold:      boolField,
  color:     colorField,
  underline: boolField,
});

// ─── Single paragraph spec ────────────────────────────────────────────────────
const ParagraphSchema = z.object({
  text:      z.string().optional(),           // simple single-run text
  segments:  z.array(SegmentSchema).optional(), // OR mixed-format runs
  bold:      boolField,
  color:     colorField,
  underline: boolField,
  align:     z.enum(['left', 'center', 'justify', 'right']).optional(),
  indent:    z.number().int().optional(),     // DXA left indent
  sz:        z.number().int().optional(),     // font size override (half-points)
}).refine(d => d.text !== undefined || (d.segments && d.segments.length > 0), {
  message: 'Either text or segments must be provided',
});

// ─── Comment annotation ───────────────────────────────────────────────────────
const CommentSchema = z.object({
  anchor: z.string().describe('Exact text to anchor the comment to (must appear in old or new content)'),
  text:   z.string().describe('Comment body text'),
  // Auto-correct LLM errors: "old_paragraphs" -> "old", "new_paragraphs" -> "new"
  column: z.string().optional().default('old').transform(v => {
    if (!v || v.startsWith('old')) return 'old';
    if (v.startsWith('new')) return 'new';
    if (v === 'both') return 'both';
    return 'old';
  }),
});

// ─── History entry schemas (shared by comparison sections & manual) ───────────
const HistoryEntrySchema = z.object({
  date:   z.string(),
  reason: z.string(),
  note:   z.string().optional().default(''),
  color:  z.string().optional().describe('Row color, e.g. "red" for new/changed entries'),
});

const RevisionEntrySchema = z.object({
  number: z.string(),
  date:   z.string(),
  reason: z.string(),
  note:   z.string().optional().default(''),
  color:  z.string().optional().describe('Row color, e.g. "red" for new/changed entries'),
});

const HistoryDataSchema = z.object({
  doc_title:   z.string().optional().describe('仕様書名 shown in the header line'),
  established: z.array(HistoryEntrySchema).optional().default([]),
  revised:     z.array(RevisionEntrySchema).optional().default([]),
});

// ─── Document section ─────────────────────────────────────────────────────────
const SectionSchema = z.object({
  id:             z.string().describe('Unique section identifier, e.g. "hyoshi", "section19"'),
  title:          z.string().describe('Section header label, e.g. "【表紙】", "【制定改廃経歴表】"'),
  status:         z.enum(['changed', 'new', 'unchanged']).describe(
                    'changed = 変更あり, new = 新規追加, unchanged = 変更なし'),
  old_paragraphs: z.array(ParagraphSchema).optional().default([]),
  new_paragraphs: z.array(ParagraphSchema).optional().default([]),
  // History table mode — used instead of old/new_paragraphs for 制定改廃経歴表 sections
  old_history:    HistoryDataSchema.optional().describe(
                    '旧 column history data — renders full 制定改廃経歴表 table layout'),
  new_history:    HistoryDataSchema.optional().describe(
                    '新 column history data — renders full 制定改廃経歴表 table layout'),
  notes:          z.array(z.string()).optional().default([])
                    .describe('備考 column bullet points (will render red)'),
  comments:       z.array(CommentSchema).optional().default([])
                    .describe('Word comment annotations with blue underline anchors'),
});

// ─── Top-level document spec ──────────────────────────────────────────────────
const DocSpecSchema = z.object({
  doc_title:  z.string().default('新旧比較表')
                .describe('Document title shown at top, e.g. "通信関係請負工事共通仕様書　比較表"'),
  sections:   z.array(SectionSchema).min(1).describe('Array of comparison sections'),
});

// ─── Manual document schemas ──────────────────────────────────────────────────
const ManualSectionSchema = z.object({
  id:         z.string(),
  heading:    z.string().optional(),
  paragraphs: z.array(ParagraphSchema).optional().default([]),
});

const ManualDocSchema = z.object({
  doc_title:    z.string().default('仕様書'),
  doc_subtitle: z.string().optional().default(''),
  revision:     z.string().optional().default(''),
  company:      z.string().optional().default(''),
  department:   z.string().optional().default(''),
  history: z.object({
    established: z.array(HistoryEntrySchema).optional().default([]),
    revised:     z.array(RevisionEntrySchema).optional().default([]),
  }).optional(),
  sections: z.array(ManualSectionSchema).default([]),
});

// ─── Export ───────────────────────────────────────────────────────────────────
const SCHEMA_DESCRIPTION = {
  description: 'JSON schema for generating a 新旧比較表 (before/after comparison) Word document',
  format: {
    doc_title: 'string - document title',
    sections: [
      {
        id:     'string - unique ID',
        title:  'string - e.g. 【表紙】 or 【１９．施工方法】',
        status: '"changed" | "new" | "unchanged"',
        old_paragraphs: [
          {
            text: 'string (for simple runs) OR',
            segments: [{ text: 'string', bold: 'bool?', color: '"black"|"red"|"hex"?', underline: 'bool?' }],
            bold: 'bool?', color: '"black"|"red"|"hex"?', underline: 'bool?',
            align: '"left"|"center"|"justify"|"right"?',
          },
        ],
        new_paragraphs: ['same as old_paragraphs'],
        notes: ['string - 備考 bullet (rendered red)'],
        comments: [
          { anchor: 'string - exact text to underline+comment', text: 'string - comment body', column: '"old"|"new"|"both"?' }
        ],
      },
    ],
  },
  example: {
    doc_title: '通信関係請負工事共通仕様書　比較表',
    sections: [
      {
        id: 'hyoshi',
        title: '【表紙】',
        status: 'changed',
        old_paragraphs: [
          { text: '通信関係請負工事共通仕様書', align: 'center', sz: 23 },
          { text: '２０２３年　４月　１日（第２回改正）', color: 'red' },
        ],
        new_paragraphs: [
          { text: '通信関係請負工事共通仕様書', align: 'center', sz: 23 },
          { text: '２０２５年　２月　１日（第３回改正）', color: 'red', underline: true },
        ],
        notes: ['・改正日を修正'],
        comments: [],
      },
      {
        id: 'section19',
        title: '【１９．施工方法および工事工程】',
        status: 'changed',
        old_paragraphs: [
          { text: '１９．施工方法および工事工程', bold: true },
          {
            segments: [
              { text: '(3)受注者は，...明らかにした' },
              { text: '作業手順確認書類', underline: true },
              { text: '，施工計画書...' },
            ],
          },
        ],
        new_paragraphs: [
          { text: '２０．施工方法および工事工程', bold: true },
          {
            segments: [
              { text: '(3)受注者は，...明らかにした' },
              { text: '作業安全確認表', color: 'red', underline: true },
              { text: '，施工計画書...' },
              { text: 'なお，危険要素の...', color: 'red', underline: true },
            ],
          },
        ],
        notes: ['・名称変更に伴う見直し', '・TD再発防止対策の追加'],
        comments: [
          { anchor: '作業手順確認書類', text: '名称変更：作業手順確認書類→作業安全確認表', column: 'old' },
          { anchor: '作業安全確認表', text: '新名称：作業安全確認表（旧：作業手順確認書類）', column: 'new' },
        ],
      },
    ],
  },
};

module.exports = {
  DocSpecSchema, ManualDocSchema, ParagraphSchema, SectionSchema,
  HistoryEntrySchema, RevisionEntrySchema, HistoryDataSchema,
  SCHEMA_DESCRIPTION,
};
