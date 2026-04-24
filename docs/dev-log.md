# Development Log — tohokunw-manual-check-tool

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
