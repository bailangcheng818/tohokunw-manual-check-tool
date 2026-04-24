# tohokunw-manual-check-tool

`excel-mcp` と `docx-mcp` の機能を、既存互換を保ちながら 1 つの MCP/HTTP サービスにまとめたツールです。

既存の生成ロジックはできるだけ薄いラッパーで呼び出しているため、元の `excel-mcp` / `docx-mcp` ディレクトリには影響しません。

## 方針

- `src/index.js`: MCP stdio entrypoint
- `src/http-server.js`: HTTP entrypoint for Dify, workflow tools, local API calls
- `src/asset-registry.js`: document asset registry. New document formats can be added here.
- `src/excel-tools.js`: Excel tool adapter. Existing `excel-writer.js` is reused.
- `src/docx-generator.js` / `src/manual-generator.js`: Word generation adapters reused from the existing docx flow.

Future expansion points are intentionally separated:

- Dify から渡されるファイルの受け取り
- package-based document parsing
- local API / local LLM calls
- Word / Excel / image extraction
- additional document asset types

Those flows should be added as new adapters or pipeline modules rather than being mixed directly into the current generators.

## Install

```bash
cd tohokunw-manual-check-tool
npm install
```

## MCP

```bash
npm start
```

Claude/Codex MCP config example:

```json
{
  "mcpServers": {
    "tohokunw-manual-check-tool": {
      "command": "node",
      "args": ["/absolute/path/to/tohokunw-manual-check-tool/src/index.js"],
      "env": {
        "OUTPUT_DIR": "/absolute/path/to/output"
      }
    }
  }
}
```

## HTTP

`.env` を使う場合（推奨）:

```bash
npm run start:http
```

スクリプトに `--env-file=.env` が含まれているため、プロジェクトルートの `.env` が自動的に読み込まれます。

`.env` の例:

```
GOOGLE_CLOUD_PROJECT=ai-business-x-dify
GOOGLE_APPLICATION_CREDENTIALS=/absolute/path/to/gcp-key.json
PORT=3456
OUTPUT_DIR=~/Desktop/tohokunw-manual-check-output
```


Core endpoints:

- `GET /health`
- `GET /schema`
- `GET /assets`
- `GET /schema/assets/:asset_type`
- `POST /generate/:asset_type`
- `POST /generate/:asset_type/download`
- `GET /schema/excel`
- `GET /schema/excel/:file`
- `POST /excel/append-row`
- `POST /excel/update-cell`
- `POST /excel/edit-record`

Backward-compatible endpoints:

- `GET /schema/comparison`
- `GET /schema/manual`
- `POST /generate`
- `POST /generate/download`
- `POST /generate/manual`
- `POST /generate/manual/download`
- `POST /append-row`
- `POST /update-cell`
- `POST /edit-record`

## MCP Tools

Document tools:

- `list_asset_types`
- `get_asset_schema`
- `validate_asset_spec`
- `preview_asset`
- `generate_asset_document`
- `generate_comparison_doc`
- `generate_manual_doc`
- `get_template_schema`
- `preview_sections`

Excel tools:

- `read_excel_headers`
- `get_excel_schema`
- `append_row`
- `update_cell`
- `append_edit_record`
- `get_schema` as a compatibility alias for Excel schema guidance

## Ingest Pipeline

`POST /ingest` accepts `{ files, dify_base_url, dify_api_key }` and processes each file into a persistent store keyed by `ref_id`.

### DOCX Processing Flow

1. **Paragraph & table extraction** — text content, headings, and tables are extracted from `word/document.xml` and stored in `scheme.json` + `content.txt`.
2. **Raster image extraction** — embedded PNG/JPG images are extracted from `word/media/` via relationship entries, analyzed by Gemini Vision, and saved as `img_NNN.{ext}` with `img_NNN_meta.json`.
3. **Drawing detection** — paragraphs containing the following are classified as `drawing` type and trigger page-level analysis:
   - `<w:drawing><wp:anchor>` — anchored shapes, SmartArt, charts
   - `<mc:AlternateContent>` — complex drawing fallback markup
   - `<w:pict>` with `<v:shape>` or `<v:imagedata>` — VML-style drawings (simple `<v:line>` horizontal rules are excluded)
   - `<w:object>` — OLE embedded objects
4. **Drawing → page images** — when drawings are detected, the DOCX is converted to per-page PNGs via LibreOffice + pdftoppm. Each drawing is mapped to its estimated page (linear interpolation); only those exact pages are analyzed (no ±1 tolerance spread). Pages whose PNG is under 150 KB are also skipped as text-only.
5. **Drawing page LLM analysis** — each drawing page is sent to Gemini 2.5 Flash Vision with surrounding paragraph context. Results include `label`, `summary`, `figure_type`, `key_elements`, and `mermaid`.
6. **Drawing pages saved as formal image entries** — analyzed drawing pages are saved as `img_NNN.png` (index continuing from raster images), with metadata written via `writeImageMeta()`. They appear in `images_summary` alongside raster images and are also persisted in `scheme.json` under `drawing_pages`.

### Ingest Response Shape

```json
{
  "manual": {
    "ref_id": "uuid",
    "content": "plain text extraction",
    "scheme": {
      "file_type": "docx",
      "image_count": 5,
      "drawing_count": 2,
      "drawing_pages": [{ "page": 3, "img_ref": "img_004.png", "label": "...", "figure_type": "flowchart" }]
    },
    "images_summary": [
      { "ref": "img_001.jpg", "label": "...", "summary": "..." },
      { "ref": "img_004.png", "source": "drawing_page", "page": 3, "label": "...", "figure_type": "flowchart" }
    ],
    "drawing_detected": true,
    "drawing_preview": [{ "page": 3, "img_ref": "img_004.png", "label": "...", "mermaid": "..." }]
  }
}
```

### System Requirements for Drawing Extraction

```bash
# Debian/Ubuntu
apt-get install -y libreoffice poppler-utils
```

If these tools are not available, raster image extraction still works; drawing conversion falls back gracefully with a `変換失敗` label.

## Adding Future Capabilities

Prefer adding new modules instead of changing generator internals:

- `src/input-*`: receive files from Dify or local uploads
- `src/extract-*`: unpack Word, Excel, PDF, or image assets
- `src/llm-*`: local API / local LLM calls
- `src/pipelines/*`: orchestration from input file to extracted content to generated check outputs
- `src/asset-registry.js`: register new output document types

This keeps current generation behavior stable while allowing the manual-check workflow to grow.
