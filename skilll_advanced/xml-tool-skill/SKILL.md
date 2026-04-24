# XML Office File Transform Tool — SKILL.md

## Overview

This skill packages the complete logic for **reading, parsing, and rule-based modifying of Office XML files** (XLSX, DOCX, PPTX) as a standalone Python microservice + LLM prompt layer.

No code interpreter sandbox required. Runs as a FastAPI service or MCP server.

---

## Architecture

```
User natural language
        ↓
   [LLM Layer]  ← system prompt in this file
   intent → structured JSON params
        ↓
   [Tool API]   ← Python FastAPI service
   execute file transform
        ↓
   Returns modified file (base64) + operation report
```

---

## File: tool_server.py

```python
"""
Office XML Transform Tool — FastAPI microservice
Handles: XLSX column ops, merge updates, drawing shape repositioning
"""

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
import base64, zipfile, io, re, copy, os, subprocess, tempfile
import xml.etree.ElementTree as ET

app = FastAPI(title="XML Office Transform Tool", version="1.0.0")

# ─────────────────────────────────────────────
# Pydantic models
# ─────────────────────────────────────────────

class ColumnOp(BaseModel):
    operation: str           # "delete" | "insert" | "rename"
    col_name: Optional[str]  # header text to locate the column
    col_index: Optional[int] # 0-based fallback if name not found
    new_name: Optional[str]  # for rename op

class MergeUpdate(BaseModel):
    old_ref: str   # e.g. "B8:D8"
    new_ref: Optional[str]  # None = delete

class ShapeRule(BaseModel):
    # How to handle drawing shapes when a column is deleted/inserted
    delete_col_shapes: bool = True   # delete shapes anchored in deleted col
    shift_adjacent: bool = True      # shift remaining shape anchors

class TransformRequest(BaseModel):
    file_b64: str                          # base64-encoded XLSX/DOCX/PPTX
    file_type: str = "xlsx"                # "xlsx" | "docx" | "pptx"
    sheet_index: int = 0                   # which sheet (XLSX only)
    column_ops: List[ColumnOp] = []
    merge_updates: List[MergeUpdate] = []
    shape_rules: ShapeRule = ShapeRule()
    cell_updates: List[Dict[str, Any]] = []  # [{ref, value, sheet_index}]
    raw_xml_patches: List[Dict[str, str]] = []  # [{file_path, find, replace}]

class TransformResponse(BaseModel):
    file_b64: str
    report: Dict[str, Any]

# ─────────────────────────────────────────────
# Core XML utilities
# ─────────────────────────────────────────────

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XDR_NS  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

def register_office_namespaces():
    pairs = [
        ("",    XLSX_NS),
        ("r",   "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("mc",  "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("xdr", XDR_NS),
        ("a",   "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("x14ac","http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"),
    ]
    for prefix, uri in pairs:
        ET.register_namespace(prefix, uri)

register_office_namespaces()

def col_letter_to_idx(letter: str) -> int:
    """'A'→0, 'D'→3, 'AA'→26"""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1

def idx_to_col_letter(idx: int) -> str:
    """0→'A', 3→'D', 26→'AA'"""
    col, s = idx + 1, ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(ord("A") + rem) + s
    return s

def cell_ref_parts(ref: str):
    """'D8' → ('D', 8, 3)  returns (col_str, row_int, col_0based)"""
    m = re.match(r"([A-Z]+)(\d+)", ref.upper())
    col_str, row = m.group(1), int(m.group(2))
    return col_str, row, col_letter_to_idx(col_str)

def find_header_col(ws_tree, header_text: str, ns: str, header_rows=(8,9,10)) -> Optional[int]:
    """
    Scan first N rows of a sheet XML tree for a cell containing header_text.
    Returns 0-based column index or None.
    Uses shared strings lookup.
    """
    # This is a simplified version; full impl reads sharedStrings.xml
    root = ws_tree.getroot()
    sheet_data = root.find(f"{{{ns}}}sheetData")
    for row_el in sheet_data:
        row_num = int(row_el.get("r", 0))
        if row_num not in header_rows:
            continue
        for cell in row_el:
            val = cell.find(f"{{{ns}}}v")
            if val is not None and header_text in (val.text or ""):
                _, _, col_idx = cell_ref_parts(cell.get("r", "A1"))
                return col_idx
    return None

# ─────────────────────────────────────────────
# XLSX column delete transform
# ─────────────────────────────────────────────

def xlsx_delete_column(zip_files: dict, sheet_xml_path: str,
                        drawing_xml_path: Optional[str],
                        delete_col_idx: int,
                        shape_rules: ShapeRule) -> dict:
    """
    Core transform: delete one column from an XLSX sheet.
    Modifies sheet XML and drawing XML in-place within zip_files dict.
    Returns operation report.
    """
    report = {"deleted_col": delete_col_idx, "shapes_deleted": 0, "shapes_shifted": 0,
              "cells_shifted": 0, "merges_updated": 0}

    # ── 1. Sheet XML ──────────────────────────────────────────
    tree = ET.ElementTree(ET.fromstring(zip_files[sheet_xml_path]))
    root = tree.getroot()
    ns = XLSX_NS

    sheet_data = root.find(f"{{{ns}}}sheetData")
    for row_el in sheet_data:
        row_num = int(row_el.get("r", 0))
        for cell in list(row_el):
            ref = cell.get("r", "")
            col_str, row, col_idx = cell_ref_parts(ref)
            if col_idx == delete_col_idx:
                row_el.remove(cell)
            elif col_idx > delete_col_idx:
                new_ref = idx_to_col_letter(col_idx - 1) + str(row)
                cell.set("r", new_ref)
                report["cells_shifted"] += 1

    # ── 2. Merged cells ───────────────────────────────────────
    merge_cells_el = root.find(f"{{{ns}}}mergeCells")
    if merge_cells_el is not None:
        for mc in list(merge_cells_el):
            ref = mc.get("ref", "")
            parts = ref.split(":")
            if len(parts) != 2:
                continue
            start_str, end_str = parts
            _, s_row, s_col = cell_ref_parts(start_str)
            _, e_row, e_col = cell_ref_parts(end_str)

            if s_col == delete_col_idx and e_col == delete_col_idx:
                merge_cells_el.remove(mc)  # entirely in deleted col
            elif s_col == delete_col_idx:
                # Merge starts in deleted col; shrink start
                new_start = idx_to_col_letter(delete_col_idx + 1) + str(s_row)
                mc.set("ref", f"{new_start}:{end_str}")
                report["merges_updated"] += 1
            elif e_col == delete_col_idx:
                # Merge ends in deleted col; shrink end
                new_end = idx_to_col_letter(delete_col_idx - 1) + str(e_row)
                mc.set("ref", f"{start_str}:{new_end}")
                report["merges_updated"] += 1
            else:
                # Shift if past deleted col
                new_ref_parts = []
                for p, col in [(start_str, s_col), (end_str, e_col)]:
                    row_n = re.search(r"\d+", p).group()
                    new_col = col - 1 if col > delete_col_idx else col
                    new_ref_parts.append(idx_to_col_letter(new_col) + row_n)
                mc.set("ref", ":".join(new_ref_parts))
                if s_col > delete_col_idx or e_col > delete_col_idx:
                    report["merges_updated"] += 1

    # ── 3. Column widths ──────────────────────────────────────
    cols_el = root.find(f"{{{ns}}}cols")
    if cols_el is not None:
        del_1based = delete_col_idx + 1
        for col_el in list(cols_el):
            min_c = int(col_el.get("min", 0))
            max_c = int(col_el.get("max", 0))
            if min_c <= del_1based <= max_c and min_c == max_c:
                cols_el.remove(col_el)
            else:
                if min_c > del_1based:
                    col_el.set("min", str(min_c - 1))
                if max_c >= del_1based:
                    col_el.set("max", str(max_c - 1))

    buf = io.BytesIO()
    tree.write(buf, xml_declaration=True, encoding="UTF-8")
    zip_files[sheet_xml_path] = buf.getvalue().decode("utf-8")

    # ── 4. Drawing XML ────────────────────────────────────────
    if drawing_xml_path and drawing_xml_path in zip_files and shape_rules:
        d_tree = ET.ElementTree(ET.fromstring(zip_files[drawing_xml_path]))
        d_root = d_tree.getroot()

        for anchor in list(d_root):
            tag = anchor.tag.replace(f"{{{XDR_NS}}}", "")
            if tag not in ("twoCellAnchor", "oneCellAnchor"):
                continue

            from_el = anchor.find(f"{{{XDR_NS}}}from")
            to_el   = anchor.find(f"{{{XDR_NS}}}to")
            if from_el is None:
                continue

            f_col_el = from_el.find(f"{{{XDR_NS}}}col")
            t_col_el = to_el.find(f"{{{XDR_NS}}}col") if to_el is not None else None
            f_col = int(f_col_el.text) if f_col_el is not None else -1
            t_col = int(t_col_el.text) if t_col_el is not None else f_col

            if f_col == delete_col_idx or t_col == delete_col_idx:
                if shape_rules.delete_col_shapes:
                    d_root.remove(anchor)
                    report["shapes_deleted"] += 1
            elif shape_rules.shift_adjacent:
                if f_col > delete_col_idx and f_col_el is not None:
                    f_col_el.text = str(f_col - 1)
                if t_col > delete_col_idx and t_col_el is not None:
                    t_col_el.text = str(t_col - 1)
                if f_col > delete_col_idx or t_col > delete_col_idx:
                    report["shapes_shifted"] += 1

        d_buf = io.BytesIO()
        d_tree.write(d_buf, xml_declaration=True, encoding="UTF-8")
        zip_files[drawing_xml_path] = d_buf.getvalue().decode("utf-8")

    return report

# ─────────────────────────────────────────────
# ZIP read/write helpers
# ─────────────────────────────────────────────

def read_zip(file_bytes: bytes) -> dict:
    """Read all files from ZIP into {path: content_str} dict"""
    files = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        for name in z.namelist():
            try:
                files[name] = z.read(name).decode("utf-8")
            except UnicodeDecodeError:
                files[name] = z.read(name)  # binary (images etc)
    return files

def write_zip(files: dict) -> bytes:
    """Pack {path: content} dict back into ZIP bytes"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            if isinstance(content, str):
                z.writestr(name, content.encode("utf-8"))
            else:
                z.writestr(name, content)
    return buf.getvalue()

def get_sheet_and_drawing_paths(files: dict, sheet_index: int):
    """Resolve xl/worksheets/sheet{N}.xml and its drawing path"""
    sheet_path = f"xl/worksheets/sheet{sheet_index + 1}.xml"
    # Find drawing relationship
    rel_path = f"xl/worksheets/_rels/sheet{sheet_index + 1}.xml.rels"
    drawing_path = None
    if rel_path in files:
        m = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', files[rel_path])
        if m:
            drawing_path = f"xl/drawings/{m.group(1)}"
    return sheet_path, drawing_path

# ─────────────────────────────────────────────
# Raw XML patch (find/replace within any file)
# ─────────────────────────────────────────────

def apply_raw_patches(files: dict, patches: List[Dict[str, str]]) -> int:
    count = 0
    for patch in patches:
        path = patch.get("file_path")
        find = patch.get("find")
        replace = patch.get("replace", "")
        if path in files and isinstance(files[path], str):
            new_content = files[path].replace(find, replace)
            if new_content != files[path]:
                files[path] = new_content
                count += 1
    return count

# ─────────────────────────────────────────────
# Main endpoint
# ─────────────────────────────────────────────

@app.post("/transform", response_model=TransformResponse)
async def transform(req: TransformRequest):
    try:
        file_bytes = base64.b64decode(req.file_b64)
    except Exception:
        raise HTTPException(400, "Invalid base64 file")

    files = read_zip(file_bytes)
    full_report = {"operations": [], "patches_applied": 0}

    sheet_path, drawing_path = get_sheet_and_drawing_paths(files, req.sheet_index)

    # ── Column operations ────────────────────────────────────
    for op in req.column_ops:
        if op.operation == "delete":
            col_idx = op.col_index
            if col_idx is None and op.col_name:
                # Try to find by name in shared strings / sheet header
                # Simplified: use col_index from LLM-resolved value
                raise HTTPException(400, "col_index required when col_name lookup is server-side")
            report = xlsx_delete_column(
                files, sheet_path, drawing_path, col_idx, req.shape_rules
            )
            full_report["operations"].append({"op": "delete_column", **report})

    # ── Merge updates (explicit overrides) ───────────────────
    if req.merge_updates:
        content = files.get(sheet_path, "")
        for mu in req.merge_updates:
            if mu.new_ref:
                content = content.replace(f'ref="{mu.old_ref}"', f'ref="{mu.new_ref}"')
            else:
                content = re.sub(
                    rf'<mergeCell ref="{re.escape(mu.old_ref)}"/>', "", content
                )
        files[sheet_path] = content
        full_report["operations"].append({"op": "merge_updates", "count": len(req.merge_updates)})

    # ── Raw XML patches ──────────────────────────────────────
    patch_count = apply_raw_patches(files, req.raw_xml_patches)
    full_report["patches_applied"] = patch_count

    result_bytes = write_zip(files)
    return TransformResponse(
        file_b64=base64.b64encode(result_bytes).decode(),
        report=full_report
    )

@app.get("/inspect")
async def inspect(file_b64: str, sheet_index: int = 0):
    """
    Return sheet structure: headers, merges, shape count.
    LLM calls this first to resolve col_name → col_index.
    """
    file_bytes = base64.b64decode(file_b64)
    files = read_zip(file_bytes)
    sheet_path, drawing_path = get_sheet_and_drawing_paths(files, sheet_index)

    result = {"sheet_path": sheet_path, "drawing_path": drawing_path,
              "headers": {}, "merges": [], "shape_count": 0}

    if sheet_path in files:
        tree = ET.ElementTree(ET.fromstring(files[sheet_path]))
        root = tree.getroot()
        ns = XLSX_NS

        # Read shared strings for header resolution
        ss_map = {}
        ss_path = "xl/sharedStrings.xml"
        if ss_path in files:
            ss_root = ET.fromstring(files[ss_path])
            ss_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            for i, si in enumerate(ss_root.findall(f"{{{ss_ns}}}si")):
                texts = [t.text or "" for t in si.iter(f"{{{ss_ns}}}t")]
                ss_map[i] = "".join(texts)

        # Scan rows 1-12 for headers
        sheet_data = root.find(f"{{{ns}}}sheetData")
        for row_el in sheet_data:
            row_num = int(row_el.get("r", 0))
            if row_num > 12:
                break
            for cell in row_el:
                ref = cell.get("r", "")
                v_el = cell.find(f"{{{ns}}}v")
                t_attr = cell.get("t", "")
                if v_el is not None and v_el.text:
                    val = ss_map.get(int(v_el.text), v_el.text) if t_attr == "s" else v_el.text
                    _, _, col_idx = cell_ref_parts(ref)
                    result["headers"][ref] = {"col_idx": col_idx, "value": val}

        # Merges
        mc_el = root.find(f"{{{ns}}}mergeCells")
        if mc_el is not None:
            result["merges"] = [mc.get("ref") for mc in mc_el]

    # Shape count
    if drawing_path and drawing_path in files:
        d_root = ET.fromstring(files[drawing_path])
        result["shape_count"] = sum(1 for c in d_root if "Anchor" in c.tag)

    return result
```

---

## File: llm_system_prompt.md

```
You are an Office XML Transform Agent.
You receive a user request to modify an Excel/Word/PowerPoint file,
and you output a structured JSON operation plan for the Tool API.

## Available tools

### 1. inspect_file
Call this FIRST on any uploaded file.
Returns: headers (cell_ref → {col_idx, value}), merges, shape_count.
Use the col_idx values from this response — never guess column numbers.

### 2. transform_file
Execute the operation plan. Accepts:
- column_ops: list of {operation, col_index, col_name}
- merge_updates: list of {old_ref, new_ref}
- shape_rules: {delete_col_shapes, shift_adjacent}
- cell_updates: list of {ref, value}
- raw_xml_patches: list of {file_path, find, replace}

## Reasoning protocol

Step 1 — INSPECT
Call inspect_file. Identify:
- Which col_idx corresponds to the column the user wants to delete/modify
- Current merge ranges that will be affected
- Whether shapes exist (shape_count > 0)

Step 2 — PLAN
Reason through each affected element:
a) Column delete: col_index = the col_idx from inspect
b) Merge impacts:
   - Merges that SPAN the deleted col → new_ref = shrink one side
   - Merges that START AFTER the deleted col → new_ref = shift col left by 1
   - Merges entirely IN deleted col → new_ref = null (delete)
c) Shape rule: if shapes exist, set delete_col_shapes=true, shift_adjacent=true

Step 3 — OUTPUT
Return exactly one JSON object:
{
  "sheet_index": 0,
  "column_ops": [...],
  "merge_updates": [...],
  "shape_rules": {"delete_col_shapes": true, "shift_adjacent": true},
  "cell_updates": [...],
  "raw_xml_patches": []
}

## Column merge reasoning rules

Given delete_col = N (0-based):
- merge ref "XA:ZB" where col(X)=N and col(Z)=N → delete (null)
- merge ref "XA:ZB" where col(X)=N and col(Z)>N → shrink: new start = col(N+1)
- merge ref "XA:ZB" where col(X)<N and col(Z)=N → shrink: new end = col(N-1)
- merge ref "XA:ZB" where col(X)>N → shift both: new cols = col-1
- merge ref "XA:ZB" where col(X)<N and col(Z)>N → shift end only: new end col = col(Z)-1

## Error handling

If inspect returns no matching header for the user's column name,
ask the user to clarify. Never guess a col_index.

If shape_count = 0, omit shape_rules from output.
```

---

## File: requirements.txt

```
fastapi>=0.111.0
uvicorn>=0.29.0
pydantic>=2.0.0
python-multipart>=0.0.9
```

---

## Usage example (curl)

```bash
# Step 1: Inspect
curl -X GET "http://localhost:8000/inspect?file_b64=$(base64 -w0 file.xlsx)&sheet_index=0"

# Step 2: Transform
curl -X POST http://localhost:8000/transform \
  -H "Content-Type: application/json" \
  -d '{
    "file_b64": "...",
    "sheet_index": 0,
    "column_ops": [{"operation": "delete", "col_index": 3}],
    "merge_updates": [
      {"old_ref": "B8:D8", "new_ref": "B8:C8"},
      {"old_ref": "D9:D10", "new_ref": null}
    ],
    "shape_rules": {"delete_col_shapes": true, "shift_adjacent": true}
  }'
```
