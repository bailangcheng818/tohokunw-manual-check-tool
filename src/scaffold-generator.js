'use strict';

const { normalizeAssetType } = require('./template-analyzer');

function toConstName(assetType) {
  return normalizeAssetType(assetType).toUpperCase();
}

function renderManifest({ formatName, assetType, analysis }) {
  return JSON.stringify({
    format_name: formatName,
    asset_type: assetType,
    version: '0.1.0',
    generator_kind: 'docx',
    source: 'template-analysis',
    paragraph_count: analysis.paragraph_count,
    schema_style: analysis.suggested_asset_blueprint.schema_style,
  }, null, 2);
}

function renderSchemaModule({ assetType, analysis }) {
  const exampleBlocks = analysis.suggested_asset_blueprint.blocks.slice(0, 12).map((block) => ({
    id: block.id,
    role: block.role,
    text: block.example.text,
    align: block.formatting.align,
    size: block.formatting.size,
  }));

  return `'use strict';

const { z } = require('zod');

const ${toConstName(assetType)}_BLOCK_SCHEMA = z.object({
  id: z.string(),
  role: z.string().optional(),
  text: z.string(),
  align: z.enum(['left', 'center', 'justify', 'right']).optional(),
  size: z.number().int().optional(),
  bold: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string().optional(),
});

const ${toConstName(assetType)}_SCHEMA = z.object({
  doc_title: z.string().optional(),
  blocks: z.array(${toConstName(assetType)}_BLOCK_SCHEMA).min(1),
});

const ${toConstName(assetType)}_SCHEMA_DESCRIPTION = {
  description: 'Generated schema scaffold for ${assetType}',
  format: {
    doc_title: 'string?',
    blocks: [
      {
        id: 'string',
        role: 'string?',
        text: 'string',
        align: '"left" | "center" | "justify" | "right"?',
        size: 'number?',
        bold: 'boolean?',
        underline: 'boolean?',
        color: 'string?',
      },
    ],
  },
  example: {
    doc_title: '${assetType}',
    blocks: ${JSON.stringify(exampleBlocks, null, 4)},
  },
};

module.exports = {
  ${toConstName(assetType)}_BLOCK_SCHEMA,
  ${toConstName(assetType)}_SCHEMA,
  ${toConstName(assetType)}_SCHEMA_DESCRIPTION,
};
`;
}

function renderRendererModule({ assetType }) {
  return `'use strict';

const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

function mapAlign(value) {
  return {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    justify: AlignmentType.JUSTIFY,
    right: AlignmentType.RIGHT,
  }[value] || AlignmentType.LEFT;
}

async function generate${assetType.replace(/(^|_)([a-z])/g, (_, a, b) => b.toUpperCase())}Doc(spec, outputPath) {
  const children = [];

  if (spec.doc_title) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: spec.doc_title, bold: true, size: 28 })],
    }));
  }

  for (const block of spec.blocks || []) {
    children.push(new Paragraph({
      alignment: mapAlign(block.align),
      spacing: { after: 60 },
      children: [new TextRun({
        text: block.text || '',
        size: block.size || 22,
        bold: block.bold || false,
        underline: block.underline ? {} : undefined,
        color: block.color,
      })],
    }));
  }

  const doc = new Document({
    sections: [{ children }],
  });

  const buffer = await Packer.toBuffer(doc);
  const result = { buffer, base64: buffer.toString('base64') };

  if (outputPath) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, buffer);
    result.path = outputPath;
  }

  return result;
}

module.exports = {
  generate${assetType.replace(/(^|_)([a-z])/g, (_, a, b) => b.toUpperCase())}Doc,
};
`;
}

function renderRegistrySnippet({ assetType }) {
  const fnName = `generate${assetType.replace(/(^|_)([a-z])/g, (_, a, b) => b.toUpperCase())}Doc`;
  return `const { ${toConstName(assetType)}_SCHEMA, ${toConstName(assetType)}_SCHEMA_DESCRIPTION } = require('./${assetType}.schema');
const { ${fnName} } = require('./${assetType}.renderer');

${assetType}: {
  asset_type: '${assetType}',
  title: '${assetType}',
  description: 'Generated asset scaffold',
  schema: ${toConstName(assetType)}_SCHEMA,
  schemaDescription: ${toConstName(assetType)}_SCHEMA_DESCRIPTION,
  defaultFileName: (spec) => spec.doc_title || '${assetType}',
  preview: (spec) => JSON.stringify(spec, null, 2),
  generate: ${fnName},
}`;
}

function renderHttpClient({ assetType, serverUrl }) {
  return `#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');

const SERVER_URL = process.env.DOCUMENT_ASSET_SERVER_URL || '${serverUrl}';

const spec = {
  doc_title: '${assetType}',
  blocks: [
    { id: 'block_1', role: 'title', text: 'Replace me', align: 'center', size: 28, bold: true }
  ]
};

async function main() {
  const response = await fetch(\`\${SERVER_URL}/assets/${assetType}/generate\`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ spec, output_filename: '${assetType}', return_base64: true }),
  });

  const payload = await response.json();
  if (!response.ok) throw new Error(JSON.stringify(payload, null, 2));

  const outDir = path.join(process.cwd(), 'output');
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, payload.filename), Buffer.from(payload.base64, 'base64'));
  console.log(payload);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
`;
}

function renderCloudScript({ assetType, serverUrl }) {
  return `export async function run({ input }) {
  const response = await fetch('${serverUrl}/assets/${assetType}/generate', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      spec: input.spec,
      output_filename: input.output_filename || '${assetType}',
      return_base64: true
    })
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(JSON.stringify(payload));
  }

  return payload;
}
`;
}

function renderMcpPrompt({ assetType }) {
  return `{
  "asset_type": "${assetType}",
  "workflow": [
    "call list_asset_types",
    "call get_asset_schema with ${assetType}",
    "edit the spec JSON",
    "call validate_asset_spec",
    "call generate_asset_document"
  ]
}`;
}

function generateAssetScaffold({ formatName, analysis, serverUrl = 'http://localhost:3456' }) {
  const assetType = normalizeAssetType(analysis.asset_type || formatName);

  return {
    asset_type: assetType,
    files: {
      'asset-manifest.json': renderManifest({ formatName, assetType, analysis }),
      [`${assetType}.schema.js`]: renderSchemaModule({ assetType, analysis }),
      [`${assetType}.renderer.js`]: renderRendererModule({ assetType }),
      'registry-entry.js': renderRegistrySnippet({ assetType }),
      'client.http.js': renderHttpClient({ assetType, serverUrl }),
      'cloudscript.js': renderCloudScript({ assetType, serverUrl }),
      'mcp-usage.json': renderMcpPrompt({ assetType }),
    },
  };
}

module.exports = {
  generateAssetScaffold,
};
