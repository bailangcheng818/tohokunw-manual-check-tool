# Development Log — tohokunw-manual-check-tool

---

## 2026-04-27

### Subfolder 別紙 file discovery

**Context:** 別紙ファイル（Excel 附表等）が manual フォルダ直下でなくサブフォルダに置かれている場合、これまでの `readdirSync(...).filter(f => f.isFile())` パターンは 1 段目のファイルしか見ておらず、サブフォルダ内のファイルを静かに無視していた。

**Fix:** `src/file-discovery.js` に `collectFiles(dir, maxDepth=1)` ヘルパーを追加。1 段深くまでサブフォルダを再帰的に走査してファイルの flat list（`{ name, fullPath }`）を返す。

**Changes:**

- `src/file-discovery.js`: `collectFiles()` 追加・export。`listManualFolders()` と `readManualFolder()` の両方で `readdirSync().filter(f.isFile())` を `collectFiles()` に置き換え。stat や file read のパスを `f.fullPath` に統一。
- `src/http-server.js`: `collectFiles` を import。`runPreIngestFolder()` の files リスト取得と `checkPendingPreIngest()` のプライマリ検索を同様に置き換え。アタッチメントの `attPath` も `attFile.fullPath` に更新。

**Folder structure now supported:**

```
ManualA/
  ManualA.docx        ← primary（変更なし）
  別紙A.xlsx          ← attachment（変更なし）
  subfolder/
    別紙B.xlsx        ← ✅ now discovered and processed
```

---

## 2026-04-24

### Pre-ingest manifest cache (Phase 1)

**Context:** The `改正判断エージェント` was using `POST /read-file` to load manual content at query time. This is fast (jszip + XML) but text-only — images and drawings in DOCX files are ignored. Running full `POST /ingest` (Gemini Vision + LibreOffice) on every agent call is too slow (tens of seconds per call).

**Design decision:** Decouple slow image analysis from query time. A one-time `POST /pre-ingest-folder` runs the full pipeline and caches results to disk. `POST /read-file` then detects the cache and serves text + image summaries instantly.

**Changes:**

- `src/file-store.js`: Added `manifest.json` at `FILE_STORE_DIR/manifest.json`. New exports: `getManifestEntry`, `setManifestEntry`, `readImagesMeta`.
- `src/ingest.js`: Exported `processDocxFile`, `processDocFile` so they can be called directly with a local buffer (not just via Dify URL downloads).
- `src/file-discovery.js`: `readManualFolder` now checks the manifest. On cache hit (mtime match), serves `content.txt` + `img_*_meta.json` from file-store without re-processing. Returns `images_analyzed: true` and `images_stale: true` if file changed since last ingest.
- `src/http-server.js`: Added `POST /pre-ingest-folder` and `GET /ingest-status`.

**Cache key:** primary document mtime (ISO string). Same file ingested twice → reuse. File updated → `images_stale: true` flag on next `read-file` call.

---

## 2026-04-25

### Pre-ingest Phase 2: summary, attachments, startup prompt

**Context (from user feedback):**
1. The first phase didn't process Excel attachments at all.
2. Pre-ingest should also generate a document summary (purpose, effective date, key topics) so that `GET /list-files` returns enough context for Dify workflows to choose which manuals to load, without needing a full `read-file`.
3. Need an ergonomic way to trigger pre-ingest in the POC: startup terminal prompt.

**Changes:**

- `src/vertex-ai.js`: Added `summarizeDocument({ text, fileName })`. Calls Gemini 2.5 Flash to produce `{ summary, key_topics, effective_date, document_type }`. Gracefully returns empty values if Vertex AI is not configured.
- `src/ingest.js`: Exported `processExcelFile` alongside the existing exports.
- `src/file-discovery.js`: `listManualFolders` now injects manifest fields into each entry (`ref_id`, `images_analyzed`, `summary`, `effective_date`, `key_topics`, `document_type`, `attachment_summaries`). Fields are omitted when manifest entry doesn't exist.
- `src/http-server.js`:
  - Extracted `runPreIngestFolder(folderName)` as a shared async function. Handles: mtime cache check → primary doc ingest → summarization → Excel attachment ingest → manifest write.
  - `POST /pre-ingest-folder` now delegates to `runPreIngestFolder`.
  - Added `checkPendingPreIngest()`: scans `MANUAL_DATABASE_DIR` at startup, lists un-ingested folders, asks `Run pre-ingest now? (y/N)` in the terminal. Skipped automatically in non-TTY environments.
  - `setImmediate(() => checkPendingPreIngest())` called inside `app.listen` callback.

**Manifest entry structure (v2):**
```json
{
  "ref_id": "uuid",
  "primary_mtime": "2026-04-25T...",
  "ingested_at": "2026-04-25T...",
  "images_analyzed": true,
  "image_count": 3,
  "summary": "本文書は...",
  "effective_date": "2024-07-01",
  "key_topics": ["通信ケーブル", "運用基準"],
  "document_type": "運用基準",
  "attachment_summaries": [
    {
      "name": "附表.xlsx",
      "type": "excel",
      "ref_id": "uuid2",
      "sheet_names": ["Sheet1"],
      "sheets": { "Sheet1": { "row_count": 50, "col_count": 8, "headers": [...] } }
    }
  ]
}
```

**Not yet done (see todo.md):**
- MCP tool `pre_ingest_folder` (index.js) — currently HTTP only
- File watcher for production auto-re-ingest
- Non-DOCX primary documents in pre-ingest summary (currently `.doc` has no Gemini summary since content extraction is simpler)

---

### Ingest / generate/from-edit 三項目最適化

**Context:** LLM が生成する編集仕様 (`/generate/from-edit`) に以下の問題があった。
1. `old_text` 完全一致依存 — 同一テキストが複数段落に存在すると最初のものが変更される。
2. run 単位の書式変更未対応 — 段落全体を単一 run に置換するため、一部だけ太字・赤字にする指定ができない。
3. ヘッダー・フッター・テキストボックスがスキームに含まれず、LLM に内容が見えない。

**Changes:**

- `src/ingest.js`:
  - `processDocxFile()` の段落抽出ループを `.entries()` に変更し、各段落に `para_id`（`"p_001"` 形式）・`xml_index`（全 `<w:p>` 中の絶対位置）・`runs`（run 配列）を追加して `scheme.json` に保存。
  - `extractHeadersFooters(zip)` を追加。`word/_rels/document.xml.rels` を読んでヘッダー・フッターの XML パーツを取得し、`scheme.headers` / `scheme.footers` として保存。各段落にも `para_id` / `xml_index` を付与。
  - `extractTextboxes(docXml)` を追加。`<w:txbx>` 内のテキストを抽出し `scheme.textboxes` として保存（書き戻しは今回スコープ外）。

- `src/edit-applier.js`:
  - `applyParagraphEdits()` を更新。`xml_index` が指定された場合は配列直接アドレス指定を優先し、`old_text` はオプション検証用に降格。`xml_index` 未指定の場合は従来の `old_text` マッチにフォールバック（後方互換）。
  - `buildRunXml(run, baseRpr)` を追加。run spec（`text`, `bold`, `underline`, `color`）から `<w:r>` XML を生成。第 1 run の `rPr` からフォント・サイズを継承しつつ書式タグは上書き。
  - `applyRunsEdit(paraXml, runsSpec)` / `applyRunsEdits(docXml, edits)` を追加。新編集タイプ `"paragraph_runs"` を実装。
  - ヘッダー・フッター編集サポートを `applyEdits()` に追加。新編集タイプ `"header_paragraph"` / `"footer_paragraph"` で `word/headerN.xml` / `word/footerN.xml` を直接パッチ。

**scheme.json 変更（段落）:**
```json
{
  "para_id": "p_003",
  "xml_index": 5,
  "type": "body",
  "text": "...",
  "runs": [{ "text": "通常", "bold": false }, { "text": "太字", "bold": true }],
  "size": null,
  "align": "left"
}
```

**新 edit type — `"paragraph_runs"`:**
```json
{
  "type": "paragraph_runs",
  "para_id": "p_003",
  "xml_index": 5,
  "runs": [
    { "text": "通常テキスト " },
    { "text": "太字部分", "bold": true },
    { "text": "赤字下線", "color": "FF0000", "underline": true }
  ]
}
```

**新 edit type — `"header_paragraph"`:**
```json
{ "type": "header_paragraph", "part": "header1", "para_id": "p_001", "xml_index": 0, "new_text": "新ヘッダー" }
```

**Backward compatibility:** 既存の `{ "type": "paragraph", "old_text": "...", "new_text": "..." }` 形式は変更なし。`xml_index` は省略可能。
