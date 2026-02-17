import time
import gspread
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2.service_account import Credentials
from pydantic import BaseModel
from typing import Optional
import re

# ======================================================
# APP SETUP
# ======================================================
app = FastAPI()

# CORS — allows your HTML frontend to call this API
# (equivalent to GAS allowing access to "Anyone")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace * with your GitHub Pages URL
    allow_methods=["*"],
    allow_headers=["*"],
)

# ======================================================
# GOOGLE SHEETS AUTH
# ======================================================
# Equivalent to GAS having automatic Google auth built-in.
# You need to set up credentials.json once (I'll guide you separately).
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_sheet(sheet_id: str):
    """Open a Google Sheet by ID — equivalent to SpreadsheetApp.openById()"""
    creds = Credentials.from_service_account_file("fastapi-test-487717-8a13c51d706e.json", scopes=SCOPES)
    client = gspread.authorize(creds)
    ss = client.open_by_key(sheet_id)
    return ss.get_worksheet(0)  # First sheet, same as ss.getSheets()[0]


# ======================================================
# HELPER — timing wrapper
# Equivalent to timed_() in GAS
# ======================================================
def timed(fn):
    t0 = time.time()
    result = fn()
    ms = int((time.time() - t0) * 1000)
    result["ms"] = ms
    return result


# ======================================================
# HELPER — column letter to index
# Equivalent to colToIndexAZ_() in GAS
# ======================================================
def col_to_index(col: str) -> int:
    col = col.upper()
    if not re.match(r'^[A-Z]$', col):
        raise HTTPException(status_code=400, detail=f"Invalid col (A-Z only): {col}")
    return ord(col) - ord('A') + 1  # A=1, B=2, etc.


# ======================================================
# HELPER — classify cell value type
# Equivalent to classify_() in GAS
# ======================================================
def classify(v) -> str:
    if v is None or v == "":
        return "blank"
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, (int, float)):
        return "number"
    return "string"


# ======================================================
# HELPER — smart parse string to correct type
# Equivalent to smartParse_() in GAS
# ======================================================
def smart_parse(raw: str):
    s = str(raw).strip()
    if not s:
        return ""
    if s.lower() in ("true", "false"):
        return s.lower() == "true"
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return s


# ======================================================
# ENDPOINT 1: GET /bounds
# Equivalent to handleGetBounds_() in GAS
# Old call: jsonp({ action: "getBounds", sheetId })
# New call: fetch(`http://localhost:8000/bounds?sheet_id=...`)
# ======================================================
@app.get("/bounds")
def get_bounds(sheet_id: str = Query(...)):
    def _run():
        sh = get_sheet(sheet_id)
        last_row = sh.row_count
        all_values = sh.get_all_values()

        # Find actual last row with data
        actual_last_row = 0
        for i, row in enumerate(all_values):
            if any(cell.strip() for cell in row):
                actual_last_row = i + 1

        headers = all_values[0] if all_values else []
        last_col = len(headers)

        return {
            "ok": True,
            "lastRow": actual_last_row,
            "lastCol": last_col,
            "headers": headers
        }

    return timed(_run)


# ======================================================
# ENDPOINT 2: GET /add-cols
# Equivalent to handleAddCols_() in GAS
# Old call: jsonp({ action: "addCols", sheetId })
# New call: fetch(`http://localhost:8000/add-cols?sheet_id=...`)
# ======================================================
@app.get("/add-cols")
def add_cols(sheet_id: str = Query(...)):
    def _run():
        sh = get_sheet(sheet_id)
        all_values = sh.get_all_values()

        if len(all_values) < 2:
            return {"ok": True, "message": "No data rows"}

        # Set header for column C
        sh.update_cell(1, 3, "sum")

        # Read col A and B (rows 2 onwards), write sum to col C
        out = []
        for row in all_values[1:]:  # skip header
            a = float(row[0]) if row[0] else 0
            b = float(row[1]) if len(row) > 1 and row[1] else 0
            out.append([a + b])

        # Write results to column C starting at row 2
        if out:
            start_cell = "C2"
            end_cell = f"C{1 + len(out)}"
            sh.update(f"{start_cell}:{end_cell}", out)

        return {"ok": True, "message": f"Wrote {len(out)} rows"}

    return timed(_run)


# ======================================================
# ENDPOINT 3: GET /column
# Equivalent to handleGetColumn_() in GAS
# Old call: jsonp({ action: "getColumn", sheetId, col })
# New call: fetch(`http://localhost:8000/column?sheet_id=...&col=A`)
# ======================================================
@app.get("/column")
def get_column(sheet_id: str = Query(...), col: str = Query(...)):
    def _run():
        col_upper = col.upper()
        col_idx = col_to_index(col_upper)
        sh = get_sheet(sheet_id)
        all_values = sh.get_all_values()

        if not all_values:
            return {"ok": True, "col": col_upper, "header": col_upper, "values": [], "n": 0}

        header = all_values[0][col_idx - 1] if len(all_values[0]) >= col_idx else col_upper

        # Data rows only (skip header row 1)
        values = []
        for row in all_values[1:]:
            values.append(row[col_idx - 1] if len(row) >= col_idx else "")

        return {
            "ok": True,
            "col": col_upper,
            "header": header,
            "values": values,
            "n": len(values)
        }

    return timed(_run)


# ======================================================
# ENDPOINT 4: GET /cell
# Equivalent to handleGetCell_() in GAS
# Old call: jsonp({ action: "getCell", sheetId, row, col })
# New call: fetch(`http://localhost:8000/cell?sheet_id=...&row=2&col=A`)
# ======================================================
@app.get("/cell")
def get_cell(sheet_id: str = Query(...), row: int = Query(...), col: str = Query(...)):
    def _run():
        col_upper = col.upper()
        col_idx = col_to_index(col_upper)
        sh = get_sheet(sheet_id)

        # Get header (row 1 of this column)
        feature_name = sh.cell(1, col_idx).value or ""

        # Get the actual cell
        cell_value = sh.cell(row, col_idx).value

        return {
            "ok": True,
            "row": row,
            "col": col_upper,
            "featureName": feature_name or "(no header)",
            "value": cell_value if cell_value is not None else "",
            "type": classify(cell_value)
        }

    return timed(_run)


# ======================================================
# ENDPOINT 5: GET /set-cell
# Equivalent to handleSetCell_() in GAS
# Old call: jsonp({ action: "setCell", sheetId, row, col, value })
# New call: fetch(`http://localhost:8000/set-cell?sheet_id=...&row=2&col=A&value=99`)
# ======================================================
@app.get("/set-cell")
def set_cell(
    sheet_id: str = Query(...),
    row: int = Query(...),
    col: str = Query(...),
    value: Optional[str] = Query(default="")
):
    def _run():
        col_upper = col.upper()
        col_idx = col_to_index(col_upper)
        sh = get_sheet(sheet_id)

        parsed = smart_parse(value or "")
        sh.update_cell(row, col_idx, parsed)

        written = sh.cell(row, col_idx).value

        return {
            "ok": True,
            "row": row,
            "col": col_upper,
            "writtenValue": written if written is not None else "",
            "writtenType": classify(written)
        }

    return timed(_run)


# ======================================================
# RUN (for VSCode Run button)
# ======================================================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)