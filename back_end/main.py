import time
import gspread
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2.service_account import Credentials
from pydantic import BaseModel
from typing import Optional
import re

# ======================================================
# Python libs
# ======================================================
import pandas as pd
import numpy as np
# import sklearn... -> this might be needed later!!!

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
# You need to set up credentials.json once.
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



class SheetOps:
    """
    A lightweight wrapper around a Google Sheet (first worksheet),
    providing a DataFrame view + common operations.
    """

    def __init__(self, sheet_id: str):
        # API request 1-1: "Load sheet" (initialize)
        self.sheet_id = sheet_id
        self.ws = get_sheet(sheet_id)  # gspread Worksheet
        self.headers: list[str] = []
        self.df: pd.DataFrame = self._load_df()

    # -------------------------
    # Internal helpers
    # -------------------------
    def _load_df(self) -> pd.DataFrame:
        # Pull all values as list[list[str]]
        vals = self.ws.get_all_values()

        if not vals:
            self.headers = []
            return pd.DataFrame()

        # Header row (row 1)
        self.headers = [str(h).strip() for h in vals[0]]

        # Data rows (row 2..)
        rows = vals[1:] if len(vals) > 1 else []

        # Ensure each row has same length as headers (pad/truncate)
        ncol = len(self.headers)
        norm_rows = []
        for r in rows:
            r = list(r)
            if len(r) < ncol:
                r = r + [""] * (ncol - len(r))
            elif len(r) > ncol:
                r = r[:ncol]
            norm_rows.append(r)

        return pd.DataFrame(norm_rows, columns=self.headers)

    def _col_to_zero_based(self, col: str) -> int:
        """
        Accept:
          - Excel letter (A..Z)
          - Header name (must match self.headers)
        Return:
          - 0-based column index
        """
        c = (col or "").strip()

        # A..Z
        if re.fullmatch(r"[A-Za-z]", c):
            idx = ord(c.upper()) - ord("A")
            if idx < 0 or idx >= len(self.headers):
                raise HTTPException(status_code=400, detail=f"Column {col} out of range.")
            return idx

        # Header name
        if c in self.headers:
            return self.headers.index(c)

        raise HTTPException(status_code=400, detail=f"Invalid column: {col}. Use A-Z or a header name.")

    def _actual_last_row_col(self) -> tuple[int, int]:
        """
        Compute "actual" last row/col that contains any non-empty cell.
        Return:
          (last_row, last_col) as 1-based counts excluding header for row logic:
          - last_row includes header row? Here we report lastRow INCLUDING header row indexing style?
            We'll report lastRow in spreadsheet row numbering (1 = header row).
        """
        if self.df.empty:
            # Only header row exists? Actually df empty means no data rows
            last_row = 1 if self.headers else 0
            last_col = len(self.headers)
            return last_row, last_col

        # Identify non-empty cells in df (strings)
        non_empty = self.df.astype(str).apply(lambda col: col.str.strip().ne(""), axis=0)

        # Last data row (0-based in df) that has any non-empty
        rows_any = non_empty.any(axis=1)
        if rows_any.any():
            last_data_row0 = rows_any[rows_any].index.max()
            # Spreadsheet row number = header row (1) + data_row_index (0-based) + 1
            last_row = 2 + int(last_data_row0)
        else:
            last_row = 1  # header only

        # Last col (0-based) that has any non-empty in header or data
        # Use header row too, because headers define "used columns" for your UI.
        header_any = [bool(str(h).strip()) for h in self.headers]
        cols_any = non_empty.any(axis=0).tolist()
        combined = [ha or ca for ha, ca in zip(header_any, cols_any)] if self.headers else cols_any

        last_col = 0
        for j in range(len(combined) - 1, -1, -1):
            if combined[j]:
                last_col = j + 1  # 1-based count
                break

        return last_row, last_col

    # -------------------------
    # Public operations (match your UI/API groups)
    # -------------------------

    def bounds(self) -> dict:
        # API request 1-2: ask for bounds
        last_row, last_col = self._actual_last_row_col()
        return {
            "lastRow": last_row,
            "lastCol": last_col,
            "headers": self.headers
        }

    def add_cols(self, colA: str = "A", colB: str = "B", outCol: str = "C", outHeader: str = "sum") -> dict:
        # API request 2: A+B -> outCol
        a_idx = self._col_to_zero_based(colA)
        b_idx = self._col_to_zero_based(colB)
        o_idx = ord(outCol.upper()) - ord("A")   # allow writing to new column  # requires outCol within current headers range

        # Convert to numeric safely; non-numeric -> NaN -> treat as 0
        a = pd.to_numeric(self.df.iloc[:, a_idx], errors="coerce").fillna(0.0)
        b = pd.to_numeric(self.df.iloc[:, b_idx], errors="coerce").fillna(0.0)
        s = (a + b)

        # Write header (row 1) for output column
        self.ws.update_cell(1, o_idx + 1, outHeader)

        # Prepare values for writing back starting at row 2
        out_vals = [[float(x)] for x in s.tolist()]  # list-of-lists for gspread

        if out_vals:
            start_row = 2
            end_row = start_row + len(out_vals) - 1
            col_letter = chr(ord("A") + o_idx)
            rng = f"{col_letter}{start_row}:{col_letter}{end_row}"
            self.ws.update(rng, out_vals)

        return {"writtenRows": len(out_vals), "outCol": outCol, "outHeader": outHeader}

    def get_column(self, col: str) -> dict:
        """
        Return a column for plotting.
        col can be 'A'..'Z' or a header name.
        """
        c0 = self._col_to_zero_based(col)  # 0-based index into df/headers
        header = self.headers[c0] if self.headers else col

        # Return data rows only (df corresponds to sheet rows starting at row 2)
        if self.df.empty:
            values = []
        else:
            values = self.df.iloc[:, c0].fillna("").astype(str).tolist()

        return {
            "col": col.upper() if re.fullmatch(r"[A-Za-z]", (col or "").strip()) else col,
            "header": header,
            "values": values,
            "n": len(values),
        }

    def get_cell_value(self, row: int, col: str) -> dict:
        # API request 3: Get specific cell value
        if row < 1:
            raise HTTPException(status_code=400, detail="Row must be >= 1.")

        c0 = self._col_to_zero_based(col)  # 0-based in df/headers
        header = self.headers[c0] if self.headers else ""

        if row == 1:
            # header row
            v = header
        else:
            r0 = row - 2  # df row index (0-based), because df starts at sheet row 2
            if r0 < 0 or r0 >= len(self.df):
                v = ""
            else:
                v = self.df.iloc[r0, c0]

        v_str = "" if v is None else str(v)
        return {
            "row": row,
            "col": col,
            "header": header,
            "value": v_str,
        }

# ======================================================
# Functions below are old funcs that are NOT using numpy or
# pandas
# ======================================================


# ======================================================
# ENDPOINT 1: GET /bounds
# Equivalent to handleGetBounds_() in GAS
# Old call: jsonp({ action: "getBounds", sheetId })
# New call: fetch(`http://localhost:8000/bounds?sheet_id=...`)
# ======================================================
@app.get("/bounds")
def get_bounds(sheet_id: str = Query(...)):
    def _run():
        ops = SheetOps(sheet_id)
        out = ops.bounds()
        return {"ok": True, **out}
    return timed(_run)

#def get_bounds(sheet_id: str = Query(...)):
#    def _run():
#        sh = get_sheet(sheet_id)
#        last_row = sh.row_count
#        all_values = sh.get_all_values()

#        # Find actual last row with data
#        actual_last_row = 0
#        for i, row in enumerate(all_values):
#            if any(cell.strip() for cell in row):
#                actual_last_row = i + 1

#        headers = all_values[0] if all_values else []
#        last_col = len(headers)

#        return {
#            "ok": True,
#            "lastRow": actual_last_row,
#            "lastCol": last_col,
#            "headers": headers
#        }

#    return timed(_run)


# ======================================================
# ENDPOINT 2: GET /add-cols
# Equivalent to handleAddCols_() in GAS
# Old call: jsonp({ action: "addCols", sheetId })
# New call: fetch(`http://localhost:8000/add-cols?sheet_id=...`)
# ======================================================
@app.get("/add-cols")
def add_cols(sheet_id: str = Query(...)):
    def _run():
        ops = SheetOps(sheet_id)
        out = ops.add_cols(colA="A", colB="B", outCol="C", outHeader="sum")
        return {"ok": True, "message": f"Wrote {out['writtenRows']} rows"}
    return timed(_run)

# @app.get("/add-cols")
# def add_cols(sheet_id: str = Query(...)):
#     def _run():
#         sh = get_sheet(sheet_id)
#         all_values = sh.get_all_values()

#         if len(all_values) < 2:
#             return {"ok": True, "message": "No data rows"}

#         # Set header for column C
#         sh.update_cell(1, 3, "sum")

#         # Read col A and B (rows 2 onwards), write sum to col C
#         out = []
#         for row in all_values[1:]:  # skip header
#             a = float(row[0]) if row[0] else 0
#             b = float(row[1]) if len(row) > 1 and row[1] else 0
#             out.append([a + b])

#         # Write results to column C starting at row 2
#         if out:
#             start_cell = "C2"
#             end_cell = f"C{1 + len(out)}"
#             sh.update(f"{start_cell}:{end_cell}", out)

#         return {"ok": True, "message": f"Wrote {len(out)} rows"}

#     return timed(_run)


# ======================================================
# ENDPOINT 3: GET /column
# Equivalent to handleGetColumn_() in GAS
# Old call: jsonp({ action: "getColumn", sheetId, col })
# New call: fetch(`http://localhost:8000/column?sheet_id=...&col=A`)
# ======================================================

@app.get("/column")
def get_column(sheet_id: str = Query(...), col: str = Query(...)):
    def _run():
        ops = SheetOps(sheet_id)
        out = ops.get_column(col)
        return {"ok": True, **out}
    return timed(_run)

# @app.get("/column")
# def get_column(sheet_id: str = Query(...), col: str = Query(...)):
#     def _run():
#         col_upper = col.upper()
#         col_idx = col_to_index(col_upper)
#         sh = get_sheet(sheet_id)
#         all_values = sh.get_all_values()

#         if not all_values:
#             return {"ok": True, "col": col_upper, "header": col_upper, "values": [], "n": 0}

#         header = all_values[0][col_idx - 1] if len(all_values[0]) >= col_idx else col_upper

#         # Data rows only (skip header row 1)
#         values = []
#         for row in all_values[1:]:
#             values.append(row[col_idx - 1] if len(row) >= col_idx else "")

#         return {
#             "ok": True,
#             "col": col_upper,
#             "header": header,
#             "values": values,
#             "n": len(values)
#         }

#     return timed(_run)


# ======================================================
# ENDPOINT 4: GET /cell
# Equivalent to handleGetCell_() in GAS
# Old call: jsonp({ action: "getCell", sheetId, row, col })
# New call: fetch(`http://localhost:8000/cell?sheet_id=...&row=2&col=A`)
# ======================================================
# @app.get("/cell")
# def get_cell(sheet_id: str = Query(...), row: int = Query(...), col: str = Query(...)):
#     def _run():
#         col_upper = col.upper()
#         col_idx = col_to_index(col_upper)
#         sh = get_sheet(sheet_id)

#         # Get header (row 1 of this column)
#         feature_name = sh.cell(1, col_idx).value or ""

#         # Get the actual cell
#         cell_value = sh.cell(row, col_idx).value

#         return {
#             "ok": True,
#             "row": row,
#             "col": col_upper,
#             "featureName": feature_name or "(no header)",
#             "value": cell_value if cell_value is not None else "",
#             "type": classify(cell_value)
#         }

#     return timed(_run)

@app.get("/cell")
def get_cell(sheet_id: str = Query(...), row: int = Query(...), col: str = Query(...)):
    def _run():
        ops = SheetOps(sheet_id)
        cell = ops.get_cell_value(row=row, col=col)  # col can be "A" or header name

        # Keep your old response keys so app.js doesn't change
        return {
            "ok": True,
            "row": row,
            "col": col.upper(),
            "featureName": cell["header"] or "(no header)",
            "value": cell["value"],
            "type": classify(cell["value"]),
        }
    return timed(_run)
# ======================================================
# ENDPOINT 5: GET /set-cell
# Equivalent to handleSetCell_() in GAS
# Old call: jsonp({ action: "setCell", sheetId, row, col, value })
# New call: fetch(`http://localhost:8000/set-cell?sheet_id=...&row=2&col=A&value=99`)
# This API is not using pandas due to its feature
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