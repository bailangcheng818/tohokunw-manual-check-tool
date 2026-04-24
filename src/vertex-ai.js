'use strict';

const PROJECT = process.env.GOOGLE_CLOUD_PROJECT;
const LOCATION = process.env.VERTEX_AI_LOCATION || 'us-central1';
const MODEL = 'gemini-2.5-flash';

const LABEL_PROMPT = `あなたは業務フロー図・組織図・システム構成図を分析する専門家です。

【文書の背景情報】
この画像は東北電力ネットワーク株式会社 情報通信部が管理する基準・マニュアル（通信回線運用基準等）の中に含まれる図です。
文書内に登場する主な組織・用語：
- 中央情報通信所、宮城支社通信センター、各支社通信センター
- 中央給電指令所、系統給電指令所、広域機関
- 本社回線 / 支社回線
- 業務代行、発令・受令、第２拠点

文書内に含まれる図の主な種類：
1. 業務フローチャート（泳道形式）- 非常事態発生時の業務フロー等
2. 連絡ルート図 - 組織間の矢印で示す連絡経路図
3. システム構成図 - PBX・転送設定・電話回線等の構成図

【出力指示】
以下のJSON形式のみで返してください（前置き・コードブロック不要）：
{
  "label": "図の種類と主題（50字以内）",
  "summary": "図の内容の詳細説明（300字以内）。登場する組織名・フローの流れ・判断分岐・矢印の方向を含めること",
  "figure_type": "flowchart | org_chart | system_diagram | table | other のいずれか",
  "key_elements": ["主要な構成要素を配列で列挙"],
  "mermaid": "Mermaidコードで再現可能な場合のみ記載。不可能な場合は空文字"
}`;

function isVertexEnabled() {
  return !!PROJECT;
}

function stripCodeFences(text) {
  return text
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
}

/**
 * Label an image using Gemini vision.
 *
 * @param {object} params
 * @param {Buffer} params.imageBuffer
 * @param {string} [params.mimeType='image/png']
 * @param {string} [params.contextText='']
 * @returns {Promise<{label: string, summary: string, figure_type: string, key_elements: string[], mermaid: string}>}
 */
async function labelImage({ imageBuffer, mimeType = 'image/png', contextText = '' }) {
  if (!isVertexEnabled()) {
    return { label: '（Vertex AI 未設定）', summary: '', figure_type: 'other', key_elements: [], mermaid: '' };
  }

  const { GoogleGenAI } = require('@google/genai');
  const ai = new GoogleGenAI({
    vertexai: true,
    project: PROJECT,
    location: LOCATION,
  });

  const imagePart = {
    inlineData: {
      mimeType,
      data: imageBuffer.toString('base64'),
    },
  };

  const promptText = contextText
    ? `${LABEL_PROMPT}\n\n【前後テキスト】\n${contextText}`
    : LABEL_PROMPT;

  const textPart = { text: promptText };

  const response = await ai.models.generateContent({
    model: MODEL,
    contents: [{ role: 'user', parts: [imagePart, textPart] }],
  });
  const rawText = response.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const cleaned = stripCodeFences(rawText);

  try {
    const parsed = JSON.parse(cleaned);
    return {
      label: String(parsed.label || ''),
      summary: String(parsed.summary || ''),
      figure_type: String(parsed.figure_type || 'other'),
      key_elements: Array.isArray(parsed.key_elements) ? parsed.key_elements.map(String) : [],
      mermaid: String(parsed.mermaid || ''),
    };
  } catch {
    return { label: '（解析失敗）', summary: cleaned.slice(0, 200), figure_type: 'other', key_elements: [], mermaid: '' };
  }
}

const SUMMARIZE_PROMPT = `あなたは東北電力ネットワーク株式会社の通信関係マニュアル・基準類を読む文書解析エキスパートです。
以下の文書テキスト（抜粋）を読み、以下のJSON形式のみで回答してください（前置き・コードブロック不要）：

{
  "summary": "文書の目的・内容を2〜3文で要約（日本語）",
  "key_topics": ["主要なトピックを3〜5個の配列で"],
  "effective_date": "施行日・改正日（YYYY-MM-DD形式。不明の場合はnull）",
  "document_type": "仕様書 | 運用基準 | 取扱基準 | マニュアル | その他"
}`;

/**
 * Summarize a document's text content using Gemini.
 *
 * @param {object} params
 * @param {string} params.text - full or truncated document text
 * @param {string} [params.fileName='']
 * @returns {Promise<{summary: string, key_topics: string[], effective_date: string|null, document_type: string|null}>}
 */
async function summarizeDocument({ text, fileName = '' }) {
  if (!isVertexEnabled()) {
    return { summary: '', key_topics: [], effective_date: null, document_type: null };
  }

  const truncated = text.length > 6000 ? text.slice(0, 6000) + '\n...(以下省略)' : text;
  const prompt = [
    fileName ? `ファイル名: ${fileName}` : '',
    SUMMARIZE_PROMPT,
    '\n【文書テキスト】',
    truncated,
  ].filter(Boolean).join('\n\n');

  const { GoogleGenAI } = require('@google/genai');
  const ai = new GoogleGenAI({ vertexai: true, project: PROJECT, location: LOCATION });

  const response = await ai.models.generateContent({
    model: MODEL,
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
  });

  const rawText = response.candidates?.[0]?.content?.parts?.[0]?.text || '';
  const cleaned = stripCodeFences(rawText);

  try {
    const parsed = JSON.parse(cleaned);
    return {
      summary: String(parsed.summary || ''),
      key_topics: Array.isArray(parsed.key_topics) ? parsed.key_topics.map(String) : [],
      effective_date: parsed.effective_date || null,
      document_type: parsed.document_type || null,
    };
  } catch {
    return { summary: cleaned.slice(0, 300), key_topics: [], effective_date: null, document_type: null };
  }
}

module.exports = { isVertexEnabled, labelImage, summarizeDocument };
