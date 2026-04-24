"""
Office XML Transform Tool — FastAPI microservice
Handles: XLSX column ops, merge updates, drawing shape repositioning
No sandbox / code interpreter required. Pure file I/O + XML.
"""

from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
import base64, zipfile, io, re
import xml.etree.ElementTree as ET

app = FastAPI(title="XML Office Transform Tool", version="1.0.0")

# ─────────────────────────────────────────────────────────────────
# Constants & namespace registration
# ─────────────────────────────────────────────────────────────────

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XDR_NS  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

_NS_MAP = {
    "":      XLSX_NS,
    "r":     "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc":    "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "xdr":   XDR_NS,
    "a":     "http://schemas.openxmlformats.org/drawingml/2006/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
}
for _prefix, _uri in _NS_MAP.items():
    ET.register_namespace(_prefix, _uri)


# ─────────────────────────────────────────────────────────────────
# Pydantic models
# ─────────────────────────────────────────────────────────────────

class ColumnOp(BaseModel):
    operation: str                    # "delete" | "rename"
    col_index: int                    # 0-based, resolved by LLM via /inspect
    new_name: Optional[str] = None    # for rename

class MergeUpdate(BaseModel):
    old_ref: str
    new_ref: Optional[str] = None     # None → delete the merge

class ShapeRule(BaseModel):
    delete_col_shapes: bool = True
    shift_adjacent: bool = True

class CellUpdate(BaseModel):
    ref: str          # e.g. "D9"
    value: str        # new cell text (written to sharedStrings)
    sheet_index: int = 0

class RawPatch(BaseModel):
    file_path: str    # path inside the zip, e.g. "xl/worksheets/sheet1.xml"
    find: str
    replace: str = ""

class TransformRequest(BaseModel):
    file_b64: str
    file_type: str = "xlsx"
    sheet_index: int = 0
    column_ops: List[ColumnOp] = []
    merge_updates: List[MergeUpdate] = []
    shape_rules: ShapeRule = ShapeRule()
    cell_updates: List[CellUpdate] = []
    raw_xml_patches: List[RawPatch] = []

class TransformResponse(BaseModel):
    file_b64: str
    report: Dict[str, Any]


# ─────────────────────────────────────────────────────────────────
# Column / cell ref utilities
# ─────────────────────────────────────────────────────────────────

def col_to_idx(col_str: str) -> int:
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1

def idx_to_col(idx: int) -> str:
    col, s = idx + 1, ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(ord("A") + rem) + s
    return s

def parse_ref(ref: str):
    """'D8' → (col_str='D', row=8, col_idx=3)"""
    m = re.match(r"([A-Z]+)(\d+)", ref.upper())
    col_str, row = m.group(1), int(m.group(2))
    return col_str, row, col_to_idx(col_str)

def shift_ref(ref: str, delete_col: int) -> Optional[str]:
    """
    Returns new ref after deleting delete_col (0-based).
    Returns None if the ref is in the deleted column.
    """
    col_str, row, col_idx = parse_ref(ref)
    if col_idx == delete_col:
        return None
    if col_idx > delete_col:
        return idx_to_col(col_idx - 1) + str(row)
    return ref


# ─────────────────────────────────────────────────────────────────
# ZIP helpers
# ─────────────────────────────────────────────────────────────────

def read_zip(file_bytes: bytes) -> dict:
    files = {}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        for name in z.namelist():
            raw = z.read(name)
            try:
                files[name] = raw.decode("utf-8")
            except UnicodeDecodeError:
                files[name] = raw
    return files

def write_zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            data = content.encode("utf-8") if isinstance(content, str) else content
            z.writestr(name, data)
    return buf.getvalue()

def get_paths(files: dict, sheet_index: int):
    sheet_path = f"xl/worksheets/sheet{sheet_index + 1}.xml"
    rel_path   = f"xl/worksheets/_rels/sheet{sheet_index + 1}.xml.rels"
    drawing_path = None
    if rel_path in files:
        m = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', files[rel_path])
        if m:
            drawing_path = f"xl/drawings/{m.group(1)}"
    return sheet_path, drawing_path

def read_shared_strings(files: dict) -> dict:
    ss_map = {}
    ss_path = "xl/sharedStrings.xml"
    if ss_path not in files:
        return ss_map
    ss_ns = XLSX_NS
    root = ET.fromstring(files[ss_path])
    for i, si in enumerate(root.findall(f"{{{ss_ns}}}si")):
        texts = [t.text or "" for t in si.iter(f"{{{ss_ns}}}t")]
        ss_map[i] = "".join(texts)
    return ss_map


# ─────────────────────────────────────────────────────────────────
# Core transform: delete one column
# ─────────────────────────────────────────────────────────────────

def _delete_column(files: dict, sheet_path: str, drawing_path: Optional[str],
                   del_col: int, shape_rules: ShapeRule) -> dict:
    report = dict(deleted_col=del_col, cells_shifted=0,
                  merges_deleted=0, merges_updated=0,
                  shapes_deleted=0, shapes_shifted=0)
    ns = XLSX_NS

    # ── Sheet XML ─────────────────────────────────────────────
    root = ET.fromstring(files[sheet_path])

    for row_el in root.find(f"{{{ns}}}sheetData") or []:
        row_num = int(row_el.get("r", 0))
        for cell in list(row_el):
            _, _, col_idx = parse_ref(cell.get("r", "A1"))
            if col_idx == del_col:
                row_el.remove(cell)
            elif col_idx > del_col:
                new_ref = idx_to_col(col_idx - 1) + str(row_num)
                cell.set("r", new_ref)
                report["cells_shifted"] += 1

    mc_el = root.find(f"{{{ns}}}mergeCells")
    if mc_el is not None:
        for mc in list(mc_el):
            ref = mc.get("ref", "")
            parts = ref.split(":")
            if len(parts) != 2:
                continue
            _, s_row, s_col = parse_ref(parts[0])
            _, e_row, e_col = parse_ref(parts[1])

            # Both endpoints in deleted col → remove
            if s_col == del_col and e_col == del_col:
                mc_el.remove(mc)
                report["merges_deleted"] += 1
            # Merge entirely to the right of deleted col → shift both
            elif s_col > del_col:
                new_s = idx_to_col(s_col - 1) + str(s_row)
                new_e = idx_to_col(e_col - 1) + str(e_row)
                mc.set("ref", f"{new_s}:{new_e}")
                report["merges_updated"] += 1
            # Start in deleted col, end beyond → shrink start right
            elif s_col == del_col and e_col > del_col:
                new_s = idx_to_col(del_col + 1) + str(s_row)
                mc.set("ref", f"{new_s}:{parts[1]}")
                report["merges_updated"] += 1
            # End in deleted col, start before → shrink end left
            elif e_col == del_col and s_col < del_col:
                new_e = idx_to_col(del_col - 1) + str(e_row)
                mc.set("ref", f"{parts[0]}:{new_e}")
                report["merges_updated"] += 1
            # Straddles deleted col (start before, end after) → shift end only
            elif s_col < del_col < e_col:
                new_e = idx_to_col(e_col - 1) + str(e_row)
                mc.set("ref", f"{parts[0]}:{new_e}")
                report["merges_updated"] += 1

    # Column width entries
    cols_el = root.find(f"{{{ns}}}cols")
    if cols_el is not None:
        del_1 = del_col + 1
        for col_el in list(cols_el):
            mn, mx = int(col_el.get("min", 0)), int(col_el.get("max", 0))
            if mn == mx == del_1:
                cols_el.remove(col_el)
            else:
                if mn > del_1:
                    col_el.set("min", str(mn - 1))
                if mx >= del_1:
                    col_el.set("max", str(mx - 1))

    buf = io.BytesIO()
    ET.ElementTree(root).write(buf, xml_declaration=True, encoding="UTF-8")
    files[sheet_path] = buf.getvalue().decode("utf-8")

    # ── Drawing XML ───────────────────────────────────────────
    if drawing_path and drawing_path in files:
        d_root = ET.fromstring(files[drawing_path])
        for anchor in list(d_root):
            if "Anchor" not in anchor.tag:
                continue
            from_el = anchor.find(f"{{{XDR_NS}}}from")
            to_el   = anchor.find(f"{{{XDR_NS}}}to")
            if from_el is None:
                continue
            fc_el = from_el.find(f"{{{XDR_NS}}}col")
            tc_el = to_el.find(f"{{{XDR_NS}}}col") if to_el is not None else None
            fc = int(fc_el.text) if fc_el is not None else -1
            tc = int(tc_el.text) if tc_el is not None else fc

            if fc == del_col or tc == del_col:
                if shape_rules.delete_col_shapes:
                    d_root.remove(anchor)
                    report["shapes_deleted"] += 1
            elif shape_rules.shift_adjacent:
                changed = False
                if fc > del_col and fc_el is not None:
                    fc_el.text = str(fc - 1); changed = True
                if tc > del_col and tc_el is not None:
                    tc_el.text = str(tc - 1); changed = True
                if changed:
                    report["shapes_shifted"] += 1

        d_buf = io.BytesIO()
        ET.ElementTree(d_root).write(d_buf, xml_declaration=True, encoding="UTF-8")
        files[drawing_path] = d_buf.getvalue().decode("utf-8")

    return report


# ─────────────────────────────────────────────────────────────────
# Endpoints
# ─────────────────────────────────────────────────────────────────

@app.get("/inspect")
async def inspect(file_b64: str = Query(...), sheet_index: int = Query(0)):
    """
    Return sheet structure so LLM can resolve column names → col_index.
    Reads header rows (1-12), merges, and shape count.
    """
    try:
        file_bytes = base64.b64decode(file_b64)
    except Exception:
        raise HTTPException(400, "Invalid base64")

    files = read_zip(file_bytes)
    sheet_path, drawing_path = get_paths(files, sheet_index)
    ss_map = read_shared_strings(files)

    result = {
        "sheet_index": sheet_index,
        "sheet_path": sheet_path,
        "drawing_path": drawing_path,
        "headers": {},
        "merges": [],
        "shape_count": 0,
        "all_sheets": []
    }

    # Sheet names
    wb_path = "xl/workbook.xml"
    if wb_path in files:
        wb_root = ET.fromstring(files[wb_path])
        wb_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        for sh in wb_root.iter(f"{{{wb_ns}}}sheet"):
            result["all_sheets"].append(sh.get("name", ""))

    if sheet_path not in files:
        raise HTTPException(404, f"Sheet not found: {sheet_path}")

    ns = XLSX_NS
    root = ET.fromstring(files[sheet_path])
    sheet_data = root.find(f"{{{ns}}}sheetData")
    for row_el in sheet_data or []:
        row_num = int(row_el.get("r", 0))
        if row_num > 12:
            break
        for cell in row_el:
            ref = cell.get("r", "")
            v_el = cell.find(f"{{{ns}}}v")
            t_attr = cell.get("t", "")
            if v_el is not None and v_el.text:
                if t_attr == "s":
                    val = ss_map.get(int(v_el.text), f"ss:{v_el.text}")
                else:
                    val = v_el.text
                _, _, col_idx = parse_ref(ref)
                result["headers"][ref] = {"col_idx": col_idx, "col_letter": idx_to_col(col_idx), "value": val}

    mc_el = root.find(f"{{{ns}}}mergeCells")
    if mc_el is not None:
        result["merges"] = [mc.get("ref") for mc in mc_el]

    if drawing_path and drawing_path in files:
        d_root = ET.fromstring(files[drawing_path])
        result["shape_count"] = sum(1 for c in d_root if "Anchor" in c.tag)

    return result


@app.post("/transform", response_model=TransformResponse)
async def transform(req: TransformRequest):
    try:
        file_bytes = base64.b64decode(req.file_b64)
    except Exception:
        raise HTTPException(400, "Invalid base64 file")

    files = read_zip(file_bytes)
    sheet_path, drawing_path = get_paths(files, req.sheet_index)
    full_report: Dict[str, Any] = {"operations": []}

    # ── Column operations ────────────────────────────────────
    # Sort deletes descending so indices stay valid across multiple deletions
    delete_ops = sorted(
        [op for op in req.column_ops if op.operation == "delete"],
        key=lambda o: o.col_index, reverse=True
    )
    for op in delete_ops:
        report = _delete_column(files, sheet_path, drawing_path,
                                op.col_index, req.shape_rules)
        full_report["operations"].append({"op": "delete_column", **report})

    # ── Explicit merge overrides ─────────────────────────────
    if req.merge_updates:
        content = files.get(sheet_path, "")
        for mu in req.merge_updates:
            if mu.new_ref:
                content = content.replace(f'ref="{mu.old_ref}"', f'ref="{mu.new_ref}"')
            else:
                content = re.sub(rf'<mergeCell ref="{re.escape(mu.old_ref)}"/>', "", content)
        files[sheet_path] = content
        full_report["operations"].append({"op": "merge_overrides", "count": len(req.merge_updates)})

    # ── Raw XML patches ──────────────────────────────────────
    patch_count = 0
    for patch in req.raw_xml_patches:
        if patch.file_path in files and isinstance(files[patch.file_path], str):
            new_content = files[patch.file_path].replace(patch.find, patch.replace)
            if new_content != files[patch.file_path]:
                files[patch.file_path] = new_content
                patch_count += 1
    if patch_count:
        full_report["operations"].append({"op": "raw_patches", "applied": patch_count})

    result_bytes = write_zip(files)
    return TransformResponse(
        file_b64=base64.b64encode(result_bytes).decode(),
        report=full_report
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
