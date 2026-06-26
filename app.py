"""
PDF-To-Excel Converter (Streamlit) — general purpose.

Concept: extract text from any PDF, reconstruct columns from the geometry of
the words on the page, and write a formatted Excel file. Two engines:

  • Layout engine (default): rebuilds columns from word x/y positions. Works on
    borderless / fixed-width reports as well as ordinary tables.
  • Ruled-table engine: uses pdfplumber's table finder, best for PDFs with
    visible cell borders.

Conversion is best-effort, not guaranteed pixel-perfect. Controls in the
sidebar let you tune it per file.

Run:  streamlit run app.py
"""

import io
import re
import traceback

import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="PDF-To-Excel Converter", page_icon="📄", layout="centered")
st.title("PDF-To-Excel Converter")
st.write(
    "Drop a PDF below (or click **Browse files**). The app rebuilds the data "
    "into columns and gives you a formatted Excel file to download."
)

# ----------------------------------------------------------------------
# Geometry helpers
# ----------------------------------------------------------------------
def cluster_rows(words, y_tol=3.0):
    """Group words into visual rows by their vertical position."""
    rows, cur, cur_top = [], [], None
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            rows.append(sorted(cur, key=lambda w: w["x0"]))
            cur, cur_top = [w], w["top"]
    if cur:
        rows.append(sorted(cur, key=lambda w: w["x0"]))
    return rows


def detect_columns(words, min_gap=10.0):
    """Return column x-ranges by merging word spans separated by < min_gap."""
    spans = sorted((w["x0"], w["x1"]) for w in words)
    merged = []
    for a, b in spans:
        if merged and a <= merged[-1][1] + min_gap:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return [tuple(m) for m in merged]


def assign_row(row, cols):
    """Place each word of a row into its column; join multiples with a space."""
    cells = [[] for _ in cols]
    centers = [(a + b) / 2 for a, b in cols]
    for w in row:
        c = (w["x0"] + w["x1"]) / 2
        idx = next((i for i, (a, b) in enumerate(cols) if a - 0.5 <= c <= b + 0.5), None)
        if idx is None:
            idx = min(range(len(cols)), key=lambda i: abs(centers[i] - c))
        cells[idx].append(w["text"])
    return [" ".join(c).strip() for c in cells]


# ----------------------------------------------------------------------
# Layout engine: reconstruct columns from word positions, combine pages
# ----------------------------------------------------------------------
def extract_layout(pdf, min_gap, combine, use_header):
    page_rows = []          # list of (page_num, [row_words, ...])
    all_words = []
    for page_num, page in enumerate(pdf.pages, start=1):
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False) or []
        page_rows.append((page_num, cluster_rows(words)))
        all_words.extend(words)

    if not all_words:
        return {"Extracted Text": pd.DataFrame({"Note": ["No selectable text found in this PDF."]})}

    if combine:
        cols = detect_columns(all_words, min_gap)
        n = len(cols)
        records = []
        for page_num, rows in page_rows:
            for row in rows:
                cells = assign_row(row, cols)
                if any(c for c in cells):
                    records.append([page_num] + cells)
        df = pd.DataFrame(records, columns=["Page"] + [f"Column {i+1}" for i in range(n)])
        df = _apply_header(df, use_header, page_aware=True)
        return {"Data": df}

    # One sheet per page, each with its own column model.
    sheets = {}
    for page_num, rows in page_rows:
        words = [w for r in rows for w in r]
        if not words:
            continue
        cols = detect_columns(words, min_gap)
        n = len(cols)
        records = [assign_row(r, cols) for r in rows]
        records = [r for r in records if any(c for c in r)]
        df = pd.DataFrame(records, columns=[f"Column {i+1}" for i in range(n)])
        df = _apply_header(df, use_header, page_aware=False)
        sheets[f"Page {page_num}"] = df
    return sheets or {"Extracted Text": pd.DataFrame({"Note": ["Nothing extracted."]})}


def _apply_header(df, use_header, page_aware):
    """Optionally promote the first row to column names and drop repeated headers."""
    if not use_header or df.empty:
        return df
    data_cols = df.columns[1:] if page_aware else df.columns
    first = df.iloc[0]
    names = []
    seen = {}
    for i, col in enumerate(df.columns):
        if page_aware and i == 0:
            names.append("Page")
            continue
        val = str(first[col]).strip() or f"Column {i+1}"
        if val in seen:
            seen[val] += 1
            val = f"{val} ({seen[val]})"
        else:
            seen[val] = 0
        names.append(val)
    body = df.iloc[1:].copy()
    body.columns = names
    # Drop rows identical to the header (repeated page headers).
    header_vals = [n for n in names if n != "Page"]
    mask = body[[c for c in body.columns if c != "Page"]].astype(str).apply(
        lambda r: list(r) != header_vals, axis=1)
    return body[mask].reset_index(drop=True)


# ----------------------------------------------------------------------
# Ruled-table engine
# ----------------------------------------------------------------------
def extract_tables(pdf, use_header):
    sheets, count = {}, 0
    for page_num, page in enumerate(pdf.pages, start=1):
        for t_idx, table in enumerate(page.extract_tables() or [], start=1):
            if not table:
                continue
            width = max(len(r) for r in table)
            norm = [list(r) + [""] * (width - len(r)) for r in table]
            df = pd.DataFrame(norm)
            if use_header and len(df) > 1:
                df.columns = [str(c) if c not in (None, "") else f"Column {i+1}"
                              for i, c in enumerate(df.iloc[0])]
                df = df.iloc[1:].reset_index(drop=True)
            sheets[f"P{page_num} T{t_idx}"[:31]] = df
            count += 1
    if count == 0:
        return None
    return sheets


# ----------------------------------------------------------------------
# Excel writing + best-effort styling
# ----------------------------------------------------------------------
def _style_sheet(ws, n_cols):
    try:
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(max(n_cols, 1))}{ws.max_row}"
        for col_cells in ws.columns:
            letter = col_cells[0].column_letter
            longest = max((len(str(c.value)) for c in col_cells if c.value is not None),
                          default=8)
            ws.column_dimensions[letter].width = min(max(longest + 2, 8), 50)
    except Exception:
        pass


def build_workbook(sheets):
    output = io.BytesIO()
    items = list(sheets.items())[:100]  # Excel/practicality cap
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in items:
            safe = (name or "Sheet")[:31]
            (df if len(df.columns) else pd.DataFrame({"Note": ["(empty)"]})).to_excel(
                writer, sheet_name=safe, index=False)
        wb = writer.book
        for name, df in items:
            ws = wb[(name or "Sheet")[:31]]
            _style_sheet(ws, max(len(df.columns), 1))
    output.seek(0)
    return output.getvalue()


def convert(file_bytes, engine, min_gap, combine, use_header):
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if engine == "Ruled tables":
            sheets = extract_tables(pdf, use_header)
            if sheets is None:
                sheets = extract_layout(pdf, min_gap, combine, use_header)
        else:
            sheets = extract_layout(pdf, min_gap, combine, use_header)
        total_rows = sum(len(df) for df in sheets.values())
        return build_workbook(sheets), len(sheets), total_rows


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("Options")
    engine = st.radio(
        "Extraction engine",
        ["Layout (any PDF)", "Ruled tables"],
        help="Layout rebuilds columns from text positions. Ruled tables uses "
             "visible cell borders (falls back to Layout if none are found).",
    )
    min_gap = st.slider(
        "Column gap sensitivity", 3, 40, 10,
        help="Smaller = split into more columns. Larger = merge into fewer. "
             "Tune this if columns look wrong.",
    )
    combine = st.checkbox("Combine all pages into one sheet", value=True)
    use_header = st.checkbox("Use first row as column titles", value=True)

uploaded_file = st.file_uploader(
    "Drag and drop a PDF here", type=["pdf"], accept_multiple_files=False
)

if uploaded_file is not None:
    excel_bytes = None
    with st.spinner("Converting…"):
        try:
            excel_bytes, n_sheets, n_rows = convert(
                uploaded_file.read(), engine, float(min_gap), combine, use_header
            )
            message = f"Done — {n_rows} rows across {n_sheets} sheet(s)."
        except Exception:
            st.error("Conversion failed. Open the details and share them so this can be fixed.")
            with st.expander("Error details"):
                st.code(traceback.format_exc())

    if excel_bytes:
        st.success(message)
        st.caption("If the columns look off, adjust **Column gap sensitivity** in the sidebar and re-upload.")
        out_name = uploaded_file.name.rsplit(".", 1)[0] + ".xlsx"
        st.download_button(
            label="⬇️  Download Excel file",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
