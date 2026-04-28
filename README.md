# tohokunw-manual-check-tool

DOCX 生成・Excel 操作・ローカルファイル探索を 1 つの MCP/HTTP サービスにまとめたツールです。ポート **3456** で稼働します。

旧 `tohoku-manual-check-tool-excel`（ポート 3457）の Excel 機能はすべてここに統合済みです。Dify ワークフローのポート番号を 3457 → 3456 に変更するだけで移行できます（パスは後方互換）。

## 方針

- `src/index.js`: MCP stdio entrypoint
- `src/http-server.js`: HTTP entrypoint for Dify, workflow tools, local API calls
- `src/asset-registry.js`: document asset registry. New document formats can be added here.
- `src/excel-tools.js`: Excel tool adapter. `excel-writer.js` is reused.
- `src/file-discovery.js`: local file discovery — list manual folders and read primary DOCX + attachments.
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

## Endpoints

### Document generation

- `GET /health`
- `GET /schema`
- `GET /assets`
- `GET /schema/assets/:asset_type`
- `POST /generate/:asset_type`
- `POST /generate/:asset_type/download`
- `POST /ingest`
- `POST /generate/from-edit` — apply edit diff to a stored DOCX; supports `"paragraph"`, `"paragraph_runs"`, `"table_cell"`, `"table_row_append"`, `"header_paragraph"`, `"footer_paragraph"` edit types
- `POST /edit/new-manual` — Dify ワークフロー出力形式の修正差分を受け取り、マークアップなしのクリーンな改正 DOCX を返す。`table_row_append` を含む全 edit type に対応
- `POST /edit/comparison-test` — 元 DOCX の構造を保ったまま、旧文を取り消し線、新文を赤字段落、理由を Word コメントで可視化する検証用 DOCX を返す
- `POST /edit/comparison` — Dify ワークフロー出力形式から新旧比較表 DOCX を生成（A3 横向き、元の書式をそのまま保持）

### File discovery

- `GET /list-files` — list manual folders under `MANUAL_DATABASE_DIR`; includes manifest summary when pre-ingested
- `POST /read-file` — read a manual folder; serves full content + image summaries from cache when pre-ingested
- `POST /pre-ingest-folder` — run full ingest (text + images + summary + attachments) and persist to manifest cache
- `GET /ingest-status` — return the full manifest (which folders have been pre-ingested)
- `GET /paragraphs/:ref_id` — return ingested paragraph list (`xml_index`, `type`, `text`) for a given `ref_id`; used by Dify LLM to determine correct `xml_index` values before sending edits
- `GET /tables/:ref_id` — return ingested table list (`table_index`, `headers`, `row_count`) for a given `ref_id`; used by Dify LLM to determine correct `table_index` for `table_row_append` edits

### Excel

- `GET /schema/excel` — usage guidance
- `GET /schema/excel/:file` — read workbook headers
- `GET /schema/:file` — backward-compat alias for the above (matches files not caught by specific `/schema/*` routes)
- `POST /excel/append-row`
- `POST /excel/update-cell`
- `POST /excel/edit-record`
- `POST /excel/update-row` — update cells in a specific row by key-value
- `GET /find-row/:file?column=&value=` — search for a row by column value

### POST /generate/from-edit — edit types

| `type` | Locator | Description |
|--------|---------|-------------|
| `"paragraph"` | `xml_index` (primary) + `old_text` (fallback / verification) | Replace entire paragraph text. Preserves first run's formatting. |
| `"paragraph_runs"` | `xml_index` (primary) + `old_text` (fallback) | Replace paragraph content with a multi-run spec. Each run can independently set `bold`, `underline`, `color`. Font/size inherited from first original run. |
| `"table_cell"` | `table_index`, `row`, `col` (0-based) | Replace a single table cell. |
| `"table_row_append"` | `table_index` (0-based), `new_text` (pipe-separated cell values) | Clone the last row of the target table and append it with new cell content. `new_text` format: `"val1\|val2\|val3\|"`. If fewer values than columns are supplied, remaining cells are left empty. |
| `"header_paragraph"` | `part` (e.g. `"header1"`), `xml_index` | Edit a paragraph inside a Word header. |
| `"footer_paragraph"` | `part` (e.g. `"footer1"`), `xml_index` | Edit a paragraph inside a Word footer. |

`para_id` and `xml_index` are assigned during `/ingest` / `/pre-ingest-folder` and stored in `scheme.json` under each paragraph entry. Pass `xml_index` for deterministic addressing; `old_text` acts as a mismatch guard — if the text at `xml_index` doesn't match `old_text`, the edit falls back to full-text search. Use `GET /paragraphs/:ref_id` to obtain the correct `xml_index` values before generating edits.

### POST /edit/new-manual — Dify 形式で改正 DOCX を生成

Dify ワークフローが出力する `amendment_edits`（JSON 文字列配列）を受け取り、修正を適用したクリーンな DOCX を返します。マークアップ・コメントなし。

Request:

```json
{
  "amendment_edits": "[{\"ref_id\":\"uuid\",\"type\":\"paragraph\",\"xml_index\":241,\"old_text\":\"旧テキスト\",\"new_text\":\"新テキスト\"}]",
  "output_filename": "改正後_マニュアル",
  "return_base64": false
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `amendment_edits` | ✅ | Dify 出力の JSON 文字列配列。各要素に `ref_id`（ingested 文書の識別子）、`type`、`xml_index`、`old_text`、`new_text` を含む |
| `output_filename` | — | 出力ファイル名（拡張子なし） |
| `return_base64` | — | `true` のとき `base64` フィールドを付加 |

### POST /edit/comparison-test — 元 DOCX 上で変更箇所を可視化

`/edit/new-manual` と同じく元の DOCX を ZIP ベースとして読み込み、本文構造を再構築せずに差分をその場へ挿入します。正式な比較表ではなく、元文書のページ構成・セクション・表をできるだけ保ったまま、変更箇所レビュー用の DOCX を返す検証用エンドポイントです。

表示ルール:
- 旧テキスト: 対象段落または対象セル内段落を取り消し線で残す
- 新テキスト: 直後の新しい段落として追加し、赤字で表示する。元段落の `<w:pPr>` を継承する
- 変更理由: `rationale` がある場合、新テキスト run に Word コメントとして付与する
- ヘッダー・フッター: `header_paragraph` / `footer_paragraph` は `/edit/new-manual` と同じ単純置換。可視マークアップとコメントは付けない

Request:

```json
{
  "amendment_edits": "[{\"ref_id\":\"uuid\",\"type\":\"paragraph\",\"xml_index\":241,\"old_text\":\"旧テキスト\",\"new_text\":\"新テキスト\",\"rationale\":\"改正理由の説明\"}]",
  "output_filename": "改正後_マニュアル_比較",
  "return_base64": false
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `amendment_edits` | ✅ | Dify 出力の JSON 文字列配列。`/edit/new-manual` と同じ形式に加え、任意で `rationale` を指定できる |
| `output_filename` | — | 出力ファイル名（拡張子なし） |
| `return_base64` | — | `true` のとき `base64` フィールドを付加 |

対応 edit type:

| `type` | Behaviour |
|--------|-----------|
| `"paragraph"` | `xml_index` を優先し、なければ `old_text` で検索。旧段落を取り消し線にし、直後に赤字の新段落を追加 |
| `"table_cell"` | `table_index`, `row`, `col` でセルを指定。セル内の元段落を取り消し線にし、同じセル内へ赤字の新段落を追加 |
| `"table_row_append"` | `table_index` と `new_text`（パイプ区切り）でテーブルの最終行をクローンし、各セルを赤字テキストで新規行として追加。`rationale` は Word コメントとして付与 |
| `"header_paragraph"` / `"footer_paragraph"` | 指定 header/footer part 内の段落を単純置換 |

Response:

```json
{
  "success": true,
  "path": "/absolute/path/to/改正後_マニュアル_比較.docx",
  "filename": "改正後_マニュアル_比較.docx",
  "download_url": "http://localhost:3456/files/改正後_マニュアル_比較.docx",
  "size_kb": 75
}
```

### POST /edit/comparison — 新旧比較表 DOCX を生成

A3 横向き、3 列（旧／新／備考）の繰り返しテーブル形式の新旧比較表 DOCX を生成します。元の DOCX をベースとして使用するため、段落スタイル・インデント・タブリーダー・Word テーブルなどの書式がそのまま保持されます。

Request:

```json
{
  "filename": "通信関係請負工事共通仕様書",
  "date": "2026-04-28",
  "amendment_edits": "[{\"ref_id\":\"uuid\",\"xml_index\":241,\"old_text\":\"旧テキスト\",\"new_text\":\"新テキスト\",\"rationale\":\"変更理由\"}]",
  "output_filename": "新旧比較表",
  "return_base64": false
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `filename` | — | マニュアル名（カバーページタイトルおよびヘッダー行の「旧」「新」ラベルに使用） |
| `date` | — | カバーページに表示する作成日（例: `2026-04-28`） |
| `amendment_edits` | ✅ | Dify 出力の JSON 文字列配列。`ref_id` で元の ingested 文書を特定し、`xml_index` または `old_text` で変更箇所を指定。テーブルセル内段落の `xml_index` も指定可能 |
| `output_filename` | — | 出力ファイル名（拡張子なし） |

`amendment_edits` の各要素:

| Field | Required | Description |
|-------|----------|-------------|
| `ref_id` | ✅ | ingest 済み文書の識別子 |
| `xml_index` | 推奨 | 段落の絶対インデックス（ingest 時に付与）。テーブルセル内段落でも指定可能 |
| `old_text` | — | 変更前テキスト。`xml_index` がない場合は全文一致でフォールバック検索 |
| `new_text` | ✅ | 変更後テキスト（新列に赤字で表示） |
| `rationale` | — | 変更理由（備考列に表示） |

Response:

```json
{
  "success": true,
  "path": "/absolute/path/to/新旧比較表.docx",
  "filename": "新旧比較表.docx",
  "download_url": "http://localhost:3456/files/新旧比較表.docx",
  "size_kb": 75
}
```

比較表の構造:
- カバーページ（タイトル・新旧比較表・作成日）→ 改ページ
- 全セクションを網羅する 1 枚の 3 列テーブル（各ページでヘッダー行が繰り返される）
- 変更なし行：旧・新の両列に元の書式のまま同一テキスト（Word テーブルを含む）
- 変更あり行：旧列に取り消し線、新列に赤字かつ元の `<w:pPr>`（インデント・行間・タブストップ）を継承した新テキスト、備考列に変更理由
- テーブルセル内段落の変更：対象セルのみ取り消し線／赤字を適用し、テーブル構造はそのまま保持
- 描画オブジェクト：`<w:drawing>` / `<mc:AlternateContent>` は安全のためストリップし、テキスト部分のみ保持

### Backward-compatible aliases

- `GET /schema/comparison`
- `GET /schema/manual`
- `POST /generate`
- `POST /generate/download`
- `POST /generate/manual`
- `POST /generate/manual/download`
- `POST /append-row`
- `POST /update-cell`
- `POST /update-row`
- `POST /edit-record`

## File Discovery

`MANUAL_DATABASE_DIR`（デフォルト: `~/Desktop/tohokunw-manual-database`）の下に、マニュアルごとのフォルダを置く構造を想定しています。

```
MANUAL_DATABASE_DIR/
  保安規程_v3/
    保安規程_v3.docx          ← プライマリ文書（フォルダ名と一致するファイルが優先）
    気づき管理表(AI用).xlsx    ← 添付 Excel（pre-ingest 時に処理）
    別添1_フロー図.docx        ← 添付 Word
    別紙/
      別紙1_仕様表.xlsx        ← サブフォルダ内の添付ファイルも対象（1 段まで）
      別紙2_フロー.xlsx
  緊急対応マニュアル/
    緊急対応マニュアル.docx
```

サブフォルダは **1 段まで** 再帰的に探索されます。`list_files` / `read_file` / `pre-ingest-folder` のすべてで適用されます。

### GET /list-files

| Query param | Default | Description |
|-------------|---------|-------------|
| `folder` | — | `MANUAL_DATABASE_DIR` 下のサブフォルダ名（省略時は直下） |
| `extensions` | `.docx,.doc,.xlsx,.xls` | 対象拡張子（カンマ区切り） |

Response（`POST /pre-ingest-folder` 実行後はマニフェストデータが付加される）:

```json
{
  "folder": "/absolute/path",
  "manuals": [
    {
      "name": "保安規程_v3",
      "primary_doc": "保安規程_v3.docx",
      "type": "word",
      "ext": ".docx",
      "size_kb": 142,
      "modified": "2025-04-20T10:32:00.000Z",
      "attachments": [
        { "name": "気づき管理表(AI用).xlsx", "type": "excel", "ext": ".xlsx", "size_kb": 38 }
      ],
      "ref_id": "uuid-or-null",
      "images_analyzed": true,
      "summary": "本文書は通信工事の共通仕様を定めたものである。",
      "effective_date": "2024-04-01",
      "key_topics": ["工事仕様", "安全基準", "検査手順"],
      "document_type": "仕様書",
      "attachment_summaries": [
        {
          "name": "気づき管理表(AI用).xlsx",
          "type": "excel",
          "ref_id": "uuid2",
          "sheet_names": ["気づき管理表"],
          "sheets": { "気づき管理表": { "row_count": 50, "col_count": 10, "headers": ["入力月日", ...] } }
        }
      ]
    }
  ],
  "count": 1
}
```

`summary` / `effective_date` / `key_topics` / `attachment_summaries` は `POST /pre-ingest-folder` 実行後に付加されます。未実行の場合は省略されます。

### POST /read-file

Request:

```json
{
  "folder_name": "保安規程_v3",
  "mode": "full"
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `folder_name` | ✅ | フォルダ名のみ（`/` や `..` を含む値は 400 エラー） |
| `mode` | — | `full`（テキスト抽出、デフォルト）または `schema`（Excel ヘッダーのみ） |

Response（`pre-ingest` 済みの場合は cache から即返却）:

```json
{
  "folder_name": "保安規程_v3",
  "ref_id": "uuid",
  "primary": {
    "file": "保安規程_v3.docx",
    "type": "word",
    "content": "plain text of the document...",
    "scheme": { "file_name": "...", "file_type": "docx", "paragraphs": 120, "tables": 3 }
  },
  "attachments": [...],
  "images_analyzed": true,
  "images_summary": [
    { "ref": "img_001.jpg", "label": "接続図", "summary": "...", "figure_type": "system_diagram" }
  ]
}
```

`images_stale: true` が返る場合はプライマリ文書が更新されており、re-ingest が必要です。

### POST /pre-ingest-folder

一度だけ実行するセットアップ用エンドポイント。プライマリ DOCX + 添付 Excel をフル処理し、Gemini で要約を生成してマニフェストに永続化します。同じファイル（mtime 一致）を再度実行した場合は即座に `cached: true` を返します。

Request:

```json
{ "folder_name": "保安規程_v3" }
```

Response:

```json
{
  "success": true,
  "folder_name": "保安規程_v3",
  "ref_id": "uuid",
  "images_analyzed": true,
  "image_count": 3,
  "summary": "本文書は...",
  "effective_date": "2024-04-01",
  "key_topics": ["工事仕様", ...],
  "attachment_summaries": [...],
  "cached": false
}
```

### GET /ingest-status

マニフェスト全体を返します。どのフォルダが pre-ingest 済みかを一覧できます。

### 起動時の自動チェック

サーバー起動後、`MANUAL_DATABASE_DIR` 内のフォルダを自動スキャンし、未処理または mtime が変更されたフォルダがある場合にターミナルで確認を求めます（インタラクティブモードのみ）:

```
[pre-ingest] 2 folder(s) not yet ingested or stale:
  • 保安規程_v3
  • 緊急対応マニュアル

Run pre-ingest now? (y/N)
```

`y` を入力するとすべて順番に処理されます。Docker などの非インタラクティブ環境では自動的にスキップされます。

## MCP Tools

Document tools:

- `list_asset_types`
- `get_asset_schema`
- `validate_asset_spec`
- `preview_asset`
- `generate_asset_document`
- `generate_asset_program`
- `analyze_docx_format`
- `generate_asset_scaffold`
- `generate_comparison_doc`
- `generate_manual_doc`
- `get_template_schema`
- `preview_sections`

Excel tools:

- `read_excel_headers`
- `get_excel_schema`
- `append_row`
- `update_cell`
- `update_row`
- `find_row`
- `append_edit_record`
- `get_schema` — compatibility alias for Excel schema guidance

File discovery tools:

- `list_files` — list manual folders in MANUAL_DATABASE_DIR (includes manifest summary when pre-ingested)
- `read_file` — read a manual folder; serves from manifest cache when pre-ingested
- `pre_ingest_folder` — (HTTP only) run full ingest + summarize + cache to manifest

## Ingest Pipeline

`POST /ingest` accepts `{ files, dify_base_url, dify_api_key }` and processes each file into a persistent store keyed by `ref_id`.

### DOCX Processing Flow

1. **Paragraph & table extraction** — text content, headings, and tables are extracted from `word/document.xml` and stored in `scheme.json` + `content.txt`. Each paragraph is assigned a `para_id` (`"p_001"`, `"p_002"`, ...) and `xml_index` (absolute position in the full `<w:p>` array) for stable addressing during edits. Run-level formatting (`bold`, `underline`, `color`) is also stored per paragraph.
2. **Header / footer extraction** — `word/_rels/document.xml.rels` is parsed to discover header and footer XML parts. Text paragraphs are extracted from each part and stored in `scheme.headers` / `scheme.footers`.
3. **Text box extraction** — `<w:txbx>` blocks inside `word/document.xml` are extracted and stored in `scheme.textboxes` (read-only; edit write-back is not currently supported).
4. **Raster image extraction** — embedded PNG/JPG images are extracted from `word/media/` via relationship entries, analyzed by Gemini Vision, and saved as `img_NNN.{ext}` with `img_NNN_meta.json`.
5. **Drawing detection** — paragraphs containing the following are classified as `drawing` type and trigger page-level analysis:
   - `<w:drawing><wp:anchor>` — anchored shapes, SmartArt, charts
   - `<mc:AlternateContent>` — complex drawing fallback markup
   - `<w:pict>` with `<v:shape>` or `<v:imagedata>` — VML-style drawings (simple `<v:line>` horizontal rules are excluded)
   - `<w:object>` — OLE embedded objects
6. **Drawing → page images** — when drawings are detected, the DOCX is converted to per-page PNGs via LibreOffice + pdftoppm. Each drawing is mapped to its estimated page (linear interpolation); only those exact pages are analyzed (no ±1 tolerance spread). Pages whose PNG is under 150 KB are also skipped as text-only.
7. **Drawing page LLM analysis** — each drawing page is sent to Gemini 2.5 Flash Vision with surrounding paragraph context. Results include `label`, `summary`, `figure_type`, `key_elements`, and `mermaid`.
8. **Drawing pages saved as formal image entries** — analyzed drawing pages are saved as `img_NNN.png` (index continuing from raster images), with metadata written via `writeImageMeta()`. They appear in `images_summary` alongside raster images and are also persisted in `scheme.json` under `drawing_pages`.

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
      "paragraphs": [
        { "para_id": "p_001", "xml_index": 0, "type": "title", "text": "...", "runs": [{ "text": "...", "bold": false }], "size": 28, "align": "center" }
      ],
      "headers": [{ "part": "header1", "paragraphs": [{ "para_id": "p_001", "xml_index": 0, "text": "..." }] }],
      "footers": [{ "part": "footer1", "paragraphs": [{ "para_id": "p_001", "xml_index": 0, "text": "..." }] }],
      "textboxes": [{ "index": 0, "paragraphs": [{ "text": "..." }] }],
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
