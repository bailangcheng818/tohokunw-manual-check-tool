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

---

# TODO

## Pre-ingest / File Discovery

- [ ] `src/index.js`: Add MCP tool `pre_ingest_folder` — currently HTTP-only; agent workflows via Claude MCP can't trigger pre-ingest
- [ ] Non-DOCX primary docs (`.doc`): `summarizeDocument` is called but `.doc` text extraction is simpler (no paragraph structure). Consider skipping summary or using a lighter prompt for legacy format.
- [ ] `checkPendingPreIngest`: add `--skip-precheck` flag or env var `SKIP_PRE_INGEST_CHECK=true` for CI/CD environments that don't want the prompt at all
- [ ] Production re-ingest trigger: file watcher (`chokidar`) on `MANUAL_DATABASE_DIR` that auto-calls `runPreIngestFolder` when a file changes. Design as an opt-in via `WATCH_MANUAL_DIR=true` env var.

## Amendment Agent Integration

- [ ] Update Dify workflow to call `GET /list-files` first and use `summary` / `effective_date` from manifest to pick relevant manuals before calling `POST /read-file`
- [ ] Pass `images_summary` from `read-file` response into the agent system prompt as structured context (not just raw text)
- [ ] Handle `images_stale: true` in the Dify workflow — surface a warning to the user that images may be outdated

## Excel Attachments

- [ ] `attachment_summaries` in manifest includes `sheets` with full headers — consider whether to expose this in `GET /list-files` response or trim it (can be large for complex workbooks)
- [ ] Support non-Excel attachments (`.doc` secondary files) in pre-ingest attachment_summaries

## General

- [ ] Add integration test: ingest a sample DOCX → check manifest → call read-file → verify `images_analyzed: true`
- [ ] Consider content-hash (sha256) in addition to mtime for cache invalidation (mtime can be reset by copy operations without content change)
