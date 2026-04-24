tohokunw-manual-check-tool（核心逻辑，只有一份）
         │
         ├── HTTP接口（src/http-server.js）
         │         └── Dify workflow → POST /ingest, /generate
         │
         ├── MCP接口（src/index.js）
         │         └── Claude Desktop / Claude Code → MCP tools
         │
         └── skill（SKILL.md，新写）
                   └── Claude 在对话中直接调 HTTP 接口
