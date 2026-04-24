'use strict';

function generateHttpProgram({ assetType, specExample, serverUrl }) {
  const safeExample = JSON.stringify(specExample, null, 2);

  return `#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');

const SERVER_URL = process.env.DOCUMENT_ASSET_SERVER_URL || '${serverUrl}';
const ASSET_TYPE = '${assetType}';

const spec = ${safeExample};

async function main() {
  const response = await fetch(\`\${SERVER_URL}/assets/\${ASSET_TYPE}/generate\`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      spec,
      output_filename: 'generated-from-program',
      return_base64: true
    })
  });

  const payload = await response.json();
  if (!response.ok) {
    throw new Error(JSON.stringify(payload, null, 2));
  }

  const outputDir = path.join(process.cwd(), 'output');
  fs.mkdirSync(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, payload.filename);
  fs.writeFileSync(outputPath, Buffer.from(payload.base64, 'base64'));

  console.log('Generated asset:', payload.asset_type);
  console.log('Saved to:', outputPath);
  console.log('Server path:', payload.path);
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
`;
}

function generateMcpProgram({ assetType, specExample }) {
  const safeExample = JSON.stringify(specExample, null, 2);

  return `/**
 * MCP client-side prompt contract example for asset_type="${assetType}".
 *
 * Typical sequence:
 * 1. Call list_asset_types
 * 2. Call get_asset_schema with "${assetType}"
 * 3. Edit the spec JSON below
 * 4. Call generate_asset_document
 */

const toolRequest = {
  asset_type: '${assetType}',
  spec: ${safeExample},
  output_filename: 'generated-from-mcp'
};

console.log(JSON.stringify(toolRequest, null, 2));
`;
}

function generateAssetProgram({ assetType, specExample, transport = 'http', serverUrl = 'http://localhost:3456' }) {
  if (transport === 'mcp') {
    return generateMcpProgram({ assetType, specExample });
  }

  return generateHttpProgram({ assetType, specExample, serverUrl });
}

module.exports = {
  generateAssetProgram,
};
