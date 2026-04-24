# XML Office Transform Tool — Skill Package

A production-ready microservice + LLM prompt for modifying Office XML files
(XLSX, DOCX, PPTX) without a code interpreter sandbox.

## Files

```
xml-tool-skill/
├── tool_server.py        ← FastAPI microservice (395 lines)
├── llm_system_prompt.md  ← System prompt for the LLM layer
├── requirements.txt      ← Python dependencies
└── README.md             ← This file
```

## Architecture

```
User natural language
        │
        ▼
┌───────────────────┐
│   LLM (any API)   │  ← uses llm_system_prompt.md
│                   │  Calls: inspect → reason → transform
└────────┬──────────┘
         │ structured JSON
         ▼
┌───────────────────┐
│  tool_server.py   │  ← FastAPI on port 8000
│  /inspect         │  Read file structure
│  /transform       │  Execute operation plan
└───────────────────┘
         │ base64 XLSX
         ▼
      Modified file
```

## Quickstart

```bash
pip install -r requirements.txt
python tool_server.py
# → http://localhost:8000
```

## API Reference

### GET /inspect

Inspect file structure. Returns headers, merges, shape count.
LLM must call this before planning any transform.

```bash
curl "http://localhost:8000/inspect?file_b64=$(base64 -w0 file.xlsx)&sheet_index=0"
```

Response:
```json
{
  "all_sheets": ["別図－１　新設工事", "別図－１　補償金", ...],
  "headers": {
    "D9": {"col_idx": 3, "col_letter": "D", "value": "通信工事\nセンター"}
  },
  "merges": ["B8:D8", "E8:F8", "G8:G10", ...],
  "shape_count": 84,
  "drawing_path": "xl/drawings/drawing1.xml"
}
```

### POST /transform

Execute a structured operation plan produced by the LLM.

```json
{
  "file_b64": "...",
  "sheet_index": 0,
  "column_ops": [
    {"operation": "delete", "col_index": 3}
  ],
  "merge_updates": [
    {"old_ref": "B8:D8", "new_ref": "B8:C8"},
    {"old_ref": "D9:D10", "new_ref": null}
  ],
  "shape_rules": {
    "delete_col_shapes": true,
    "shift_adjacent": true
  },
  "raw_xml_patches": [
    {
      "file_path": "xl/worksheets/sheet1.xml",
      "find": "old string",
      "replace": "new string"
    }
  ]
}
```

Response:
```json
{
  "file_b64": "...",
  "report": {
    "operations": [
      {
        "op": "delete_column",
        "deleted_col": 3,
        "cells_shifted": 384,
        "merges_deleted": 1,
        "merges_updated": 5,
        "shapes_deleted": 29,
        "shapes_shifted": 49
      }
    ]
  }
}
```

## LLM Integration

Use `llm_system_prompt.md` as the system prompt.

The LLM (Claude / GPT-4o / Gemini) will:
1. Receive user's natural language request
2. Call `/inspect` to understand the file
3. Reason through merge impacts using the table in the prompt
4. Call `/transform` with the resolved JSON plan

**Works with any LLM API** — no code execution capability required.
The LLM only does reasoning; all file I/O stays in the Python server.

## Supported operations

| Operation | Status |
|-----------|--------|
| Delete column | ✅ Full support (cells + merges + shapes) |
| Shift column refs | ✅ Automatic |
| Update merge ranges | ✅ Automatic + manual override |
| Delete/shift drawing shapes | ✅ Full support |
| Raw XML find/replace | ✅ Any file in ZIP |
| Multiple column deletes | ✅ Sorted high→low |

## Limitations

| Limitation | Notes |
|------------|-------|
| Cell value updates | Requires sharedStrings.xml rewrite (not yet implemented) |
| DOCX / PPTX | Architecture supports it; XML path logic differs |
| Conditional formatting | Not tracked across column shifts |
| Named ranges | Not updated on column shift |

## Design notes

The key insight enabling this without a sandbox:

Office XML files (.xlsx, .docx, .pptx) are ZIP archives containing
plain XML files. All operations — column deletion, shape repositioning,
merge range updates — reduce to:

1. Parse XML with ElementTree
2. Apply coordinate arithmetic (column index arithmetic is ~10 lines)
3. Repack ZIP

The LLM's only job is mapping natural language → {col_index, merge_updates}.
The arithmetic and XML surgery stay deterministic in Python.

## Extension points

**Add insert column:**
```python
# In _delete_column, replace "col_idx - 1" with "col_idx + 1"
# and reverse the condition
```

**Add DOCX table column delete:**
```python
# Target: word/document.xml
# Namespace: http://schemas.openxmlformats.org/wordprocessingml/2006/main
# Elements: w:tbl > w:tr > w:tc (table cells)
```

**MCP server wrapper:**
```python
# Expose /inspect and /transform as MCP tools
# Drop-in for Claude Desktop or Dify MCP plugin
```
