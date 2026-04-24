#!/usr/bin/env node
'use strict';

const path = require('path');

const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require('@modelcontextprotocol/sdk/types.js');

const {
  getAssetDefinition,
  getAssetSchemaDescription,
  listAssetTypes,
  validateAssetSpec,
} = require('./asset-registry');
const { EXCEL_TOOLS, handleExcelTool } = require('./excel-tools');
const { generateAssetProgram } = require('./program-generator');
const { analyzeDocxTemplateBase64 } = require('./template-analyzer');
const { generateAssetScaffold } = require('./scaffold-generator');
const { OUTPUT_DIR, PORT, SERVICE_NAME, VERSION, safeFilename } = require('./config');

async function generateAssetDocument(assetType, spec, outputFilename) {
  const validation = validateAssetSpec(assetType, spec);
  if (!validation.success) {
    return {
      content: [{
        type: 'text',
        text: `Invalid spec for ${assetType}:\n${(validation.details || []).map((i) => `- ${i.path}: ${i.message}`).join('\n') || validation.error}`,
      }],
      isError: true,
    };
  }

  const asset = getAssetDefinition(assetType);
  const filename = safeFilename(outputFilename || asset.defaultFileName(validation.data));
  const outputPath = path.join(OUTPUT_DIR, `${filename}.docx`);
  const result = await asset.generate(validation.data, outputPath);

  return {
    content: [{
      type: 'text',
      text: [
        'Document generated successfully.',
        `asset_type: ${assetType}`,
        `file: ${result.path}`,
        `base64_length: ${result.base64.length}`,
      ].join('\n'),
    }],
    _meta: {
      asset_type: assetType,
      path: result.path,
      base64: result.base64,
    },
  };
}

const DOCUMENT_TOOLS = [
  {
    name: 'list_asset_types',
    description: 'List built-in document asset types supported by this server.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'get_asset_schema',
    description: 'Return schema guidance and an example for a specific asset_type.',
    inputSchema: {
      type: 'object',
      required: ['asset_type'],
      properties: { asset_type: { type: 'string' } },
    },
  },
  {
    name: 'validate_asset_spec',
    description: 'Validate a candidate spec JSON for a given asset_type before generation.',
    inputSchema: {
      type: 'object',
      required: ['asset_type', 'spec'],
      properties: { asset_type: { type: 'string' }, spec: { type: 'object' } },
    },
  },
  {
    name: 'preview_asset',
    description: 'Preview the text structure of an asset spec without generating a file.',
    inputSchema: {
      type: 'object',
      required: ['asset_type', 'spec'],
      properties: { asset_type: { type: 'string' }, spec: { type: 'object' } },
    },
  },
  {
    name: 'generate_asset_document',
    description: 'Generate a Word document for any supported asset_type.',
    inputSchema: {
      type: 'object',
      required: ['asset_type', 'spec'],
      properties: {
        asset_type: { type: 'string' },
        spec: { type: 'object' },
        output_filename: { type: 'string' },
      },
    },
  },
  {
    name: 'generate_asset_program',
    description: 'Generate a JavaScript starter program for calling this server with a specific asset_type.',
    inputSchema: {
      type: 'object',
      required: ['asset_type'],
      properties: {
        asset_type: { type: 'string' },
        transport: { type: 'string', enum: ['http', 'mcp'] },
        server_url: { type: 'string' },
      },
    },
  },
  {
    name: 'analyze_docx_format',
    description: 'Analyze an existing Word .docx template from base64 and suggest an asset schema blueprint.',
    inputSchema: {
      type: 'object',
      required: ['base64_docx'],
      properties: {
        base64_docx: { type: 'string' },
        format_name: { type: 'string' },
      },
    },
  },
  {
    name: 'generate_asset_scaffold',
    description: 'Analyze a docx template and generate scaffold files for a new document asset type.',
    inputSchema: {
      type: 'object',
      required: ['base64_docx'],
      properties: {
        base64_docx: { type: 'string' },
        format_name: { type: 'string' },
        server_url: { type: 'string' },
      },
    },
  },
  {
    name: 'generate_comparison_doc',
    description: 'Backward-compatible alias for generate_asset_document with asset_type="comparison_doc".',
    inputSchema: {
      type: 'object',
      required: ['spec'],
      properties: { spec: { type: 'object' }, output_filename: { type: 'string' } },
    },
  },
  {
    name: 'generate_manual_doc',
    description: 'Backward-compatible alias for generate_asset_document with asset_type="manual_doc".',
    inputSchema: {
      type: 'object',
      required: ['spec'],
      properties: { spec: { type: 'object' }, output_filename: { type: 'string' } },
    },
  },
  {
    name: 'get_template_schema',
    description: 'Backward-compatible schema endpoint. Pass doc_type="manual" for manual_doc.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_type: { type: 'string', enum: ['comparison', 'manual'] },
      },
    },
  },
  {
    name: 'preview_sections',
    description: 'Backward-compatible preview endpoint for comparison_doc specs.',
    inputSchema: {
      type: 'object',
      required: ['spec'],
      properties: { spec: { type: 'object' } },
    },
  },
];

const TOOLS = [...DOCUMENT_TOOLS, ...EXCEL_TOOLS];

async function handleDocumentTool(name, args = {}) {
  if (name === 'list_asset_types') {
    return toolJson(listAssetTypes());
  }

  if (name === 'get_asset_schema') {
    const schema = getAssetSchemaDescription(args.asset_type);
    if (!schema) return toolError(`Unknown asset_type: ${args.asset_type}`);
    return toolJson(schema);
  }

  if (name === 'validate_asset_spec') {
    const validation = validateAssetSpec(args.asset_type, args.spec);
    return { ...toolJson(validation), isError: !validation.success };
  }

  if (name === 'preview_asset') {
    const validation = validateAssetSpec(args.asset_type, args.spec);
    if (!validation.success) return { ...toolJson(validation), isError: true };
    return toolText(getAssetDefinition(args.asset_type).preview(validation.data));
  }

  if (name === 'generate_asset_document') {
    return generateAssetDocument(args.asset_type, args.spec, args.output_filename);
  }

  if (name === 'generate_asset_program') {
    const schema = getAssetSchemaDescription(args.asset_type);
    if (!schema) return toolError(`Unknown asset_type: ${args.asset_type}`);
    return toolText(generateAssetProgram({
      assetType: args.asset_type,
      specExample: schema.example,
      transport: args.transport || 'http',
      serverUrl: args.server_url || `http://localhost:${PORT}`,
    }));
  }

  if (name === 'analyze_docx_format') {
    const analysis = await analyzeDocxTemplateBase64(args.base64_docx, args.format_name || 'Generated Format');
    return toolJson(analysis);
  }

  if (name === 'generate_asset_scaffold') {
    const analysis = await analyzeDocxTemplateBase64(args.base64_docx, args.format_name || 'Generated Format');
    const scaffold = generateAssetScaffold({
      formatName: args.format_name || analysis.format_name,
      analysis,
      serverUrl: args.server_url || `http://localhost:${PORT}`,
    });
    return toolJson(scaffold);
  }

  if (name === 'generate_comparison_doc') {
    return generateAssetDocument('comparison_doc', args.spec, args.output_filename);
  }

  if (name === 'generate_manual_doc') {
    return generateAssetDocument('manual_doc', args.spec, args.output_filename);
  }

  if (name === 'get_template_schema') {
    return toolJson(getAssetSchemaDescription(args.doc_type === 'manual' ? 'manual_doc' : 'comparison_doc'));
  }

  if (name === 'preview_sections') {
    return toolText(getAssetDefinition('comparison_doc').preview(args.spec || {}));
  }

  return null;
}

async function main() {
  const server = new Server(
    { name: SERVICE_NAME, version: VERSION },
    { capabilities: { tools: {} } },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args = {} } = request.params;

    try {
      const excelResult = await handleExcelTool(name, args);
      if (excelResult) return excelResult;

      const documentResult = await handleDocumentTool(name, args);
      if (documentResult) return documentResult;

      return toolError(`Unknown tool: ${name}`);
    } catch (error) {
      return toolError(error.stack || error.message);
    }
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`[${SERVICE_NAME}] Server started (stdio). Output dir: ${OUTPUT_DIR}`);
}

function toolText(text) {
  return { content: [{ type: 'text', text: String(text) }] };
}

function toolJson(value) {
  return toolText(JSON.stringify(value, null, 2));
}

function toolError(message) {
  return { content: [{ type: 'text', text: String(message) }], isError: true };
}

main().catch((error) => {
  console.error(`[${SERVICE_NAME}] Fatal:`, error);
  process.exit(1);
});
