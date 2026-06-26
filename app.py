"""
PDF-To-Excel Converter
A lightweight Streamlit app that converts a PDF into a downloadable, formatted
Excel workbook.

It includes a dedicated parser for the "Daily Order Detail Maintenance Audit
Report" (MNTAUDDTL) that lays each value out in its own labelled column. Any
other PDF falls back to generic table/text extraction so the app still works.

Run locally / in Codespaces:   streamlit run app.py
"""

import io
import re

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# Page setup
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="PDF-To-Excel Converter",
    page_icon="📄",
    layout="centered",
)

st.title("PDF-To-Excel Converter")
st.write(
    "Drop a PDF in the box below (or click **Browse files**). "
    "Audit reports are parsed into labelled columns; other PDFs fall back to "
    "generic table/text extraction. You'll get a formatted Excel file to download."
)

# ----------------------------------------------------------------------
# Columns produced by the audit-report parser, in order.
# ----------------------------------------------------------------------
AUDIT_COLUMNS = [
    "Account Code", "Account Name", "Page",
    "P.O.#", "Reg#", "Ibm#", "Line", "Trn.Date", "Trn.Time", "User", "B/A",
    "Style", "Col", "Lb", "Tgt Style", "Tgt Col", "Tgt Lb",
    "Customer", "Customer Style", "Str.Date", "Cnx.Date",
    "Price", "Qty", "Rules", "Mode", "PTY", "Value", "Unparsed",
]
NUMERIC_COLUMNS = {"Price", "Qty", "Value"}

# A "Customer Style" is a long numeric barcode (>= 9 digits, maybe a *NN suffix).
# This is the reliable anchor that separates the style/customer fields (to its
# left) from the date/price/qty fields (to its right).
CUST_STYLE_RE = re.compile(r"^\d{9,}(\*\d+)?$")
ACCOUNT_RE = re.compile(r"Auto Maintenance Account:\s*(\S+)\s+(.+?)\s+Program:")
REPORT_SIGNATURE = "Daily Order Detail Maintenance Audit Report"


# ----------------------------------------------------------------------
# Audit-report parsing
# ----------------------------------------------------------------------
def _pad3(group):
    """Return exactly three items from a style/col/lb group (blank-padded)."""
    g = list(group) + ["", "", ""]
    return g[0], g[1], g[2]


def parse_payload(tokens):
    """Parse the part of a Before/After line that follows the keyword."""
    out = {c: "" for c in AUDIT_COLUMNS}

    if not tokens:
        return out
    if tokens == ["New"]:
        out["Style"] = "New"
        return out

    # Locate the customer-style barcode; everything keys off its position.
    idx = next((i for i, t in enumerate(tokens) if CUST_STYLE_RE.match(t)), None)

    if idx is None:
        # Deletion result lines look like "0/00/00 0/00/00".
        if len(tokens) >= 2 and tokens[0].count("/") == 2:
            out["Str.Date"], out["Cnx.Date"] = tokens[0], tokens[1]
        else:
            out["Unparsed"] = " ".join(tokens)
        return out

    out["Customer"] = tokens[idx - 1]
    out["Customer Style"] = tokens[idx]

    # Style / Col / Lb group(s) sit before the customer code. A single group is
    # the source; two groups means a source plus a target (relabelled) entry.
    group = tokens[: idx - 1]
    if len(group) <= 3:
        out["Style"], out["Col"], out["Lb"] = _pad3(group)
    else:
        if len(group) % 2:                       # odd -> first group carries a label
            g1, g2 = group[:3], group[3:]
        else:                                    # even -> split evenly
            half = len(group) // 2
            g1, g2 = group[:half], group[half:]
        out["Style"], out["Col"], out["Lb"] = _pad3(g1)
        out["Tgt Style"], out["Tgt Col"], out["Tgt Lb"] = _pad3(g2)

    # Tail after the barcode: Str.Date, Cnx.Date, Price, Qty, then optional
    # Rules (CHG/DLT-H), Mode (a number) and Value (a decimal).
    tail = tokens[idx + 1:]
    if tail:
        out["Str.Date"] = tail[0]
    if len(tail) > 1:
        out["Cnx.Date"] = tail[1]
    rem = tail[2:]
    if rem:
        out["Price"] = rem[0]
    if len(rem) > 1:
        out["Qty"] = rem[1]
    for t in rem[2:]:
        if any(ch.isalpha() for ch in t):
            out["Rules"] = t
        elif "." in t:
            out["Value"] = t
        else:
            out["Mode"] = t
    return out


def parse_audit_report(pdf):
    """Parse the whole audit report into (rows, totals)."""
    rows = []
    totals = []                      # list of (account_code, account_name, total)
    grand_total = None

    account_code = account_name = ""
    # Transaction-identity fields are only printed on the "Before:" line; we
    # carry them onto the following "After:" line so every row stands alone.
    lead = {k: "" for k in ["P.O.#", "Reg#", "Ibm#", "Line",
                            "Trn.Date", "Trn.Time", "User"]}

    for page_num, page in enumerate(pdf.pages, start=1):
        text = page.extract_text() or ""
        for raw in text.splitlines():
            line = raw.strip()
            if not line or REPORT_SIGNATURE in line or line.startswith("P.O.#") \
                    or line == "Target":
                continue

            m = ACCOUNT_RE.search(line)
            if m:
                account_code, account_name = m.group(1), m.group(2).strip()
                continue

            if line.startswith("Grand Total:"):
                nums = re.findall(r"[\d,]+\.\d+", line)
                grand_total = nums[-1] if nums else ""
                continue
            if line.startswith("Total:"):
                nums = re.findall(r"[\d,]+\.\d+", line)
                totals.append((account_code, account_name,
                               nums[-1] if nums else ""))
                continue

            # Data lines: a "Before:" record (with leading id fields) or the
            # "After:" continuation.
            if "Before:" in line:
                left, _, right = line.partition("Before:")
                lead_tokens = left.split()
                if len(lead_tokens) >= 7:
                    keys = ["P.O.#", "Reg#", "Ibm#", "Line",
                            "Trn.Date", "Trn.Time", "User"]
                    lead = dict(zip(keys, lead_tokens[:7]))
                ba = "Before"
                payload = right.split()
            elif line.startswith("After:"):
                ba = "After"
                payload = line.partition("After:")[2].split()
            else:
                continue

            row = parse_payload(payload)
            row.update(lead)
            row["Account Code"] = account_code
            row["Account Name"] = account_name
            row["Page"] = page_num
            row["B/A"] = ba
            rows.append(row)

    return rows, totals, grand_total


# ----------------------------------------------------------------------
# Generic fallback (for non-audit PDFs)
# ----------------------------------------------------------------------
def extract_generic(pdf):
    """Return {sheet_name: DataFrame} using table extraction, then text."""
    sheets, table_count, text_rows = {}, 0, []
    for page_num, page in enumerate(pdf.pages, start=1):
        tables = page.extract_tables()
        if tables:
            for t_idx, table in enumerate(tables, start=1):
                if not table:
                    continue
                df = pd.DataFrame(table)
                if len(df) > 1:
                    df.columns = [str(c) if c is not None else "" for c in df.iloc[0]]
                    df = df.iloc[1:].reset_index(drop=True)
                sheets[f"P{page_num}_T{t_idx}"[:31]] = df
                table_count += 1
        else:
            for ln in (page.extract_text() or "").splitlines():
                if ln.strip():
                    text_rows.append([page_num, ln])
    if table_count == 0:
        if text_rows:
            sheets = {"Extracted Text": pd.DataFrame(text_rows, columns=["Page", "Text"])}
        else:
            sheets = {"Extracted Text": pd.DataFrame(
                {"Note": ["No extractable text or tables found in this PDF."]})}
    return sheets


# ----------------------------------------------------------------------
# Excel writing + formatting
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
AFTER_FILL = PatternFill("solid", fgColor="F2F2F2")


def _style_sheet(ws, n_cols, after_col_idx=None):
    """Apply header styling, freeze panes, autofilter and column widths."""
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{ws.max_row}"

    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        longest = max((len(str(c.value)) for c in col_cells if c.value is not None),
                      default=8)
        ws.column_dimensions[letter].width = min(max(longest + 2, 8), 40)

    # Shade "After" rows for readability.
    if after_col_idx is not None:
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=after_col_idx).value == "After":
                for c in range(1, n_cols + 1):
                    ws.cell(row=r, column=c).fill = AFTER_FILL


def build_audit_workbook(rows, totals, grand_total):
    df = pd.DataFrame(rows, columns=AUDIT_COLUMNS)
    if not df["Unparsed"].astype(str).str.strip().any():
        df = df.drop(columns=["Unparsed"])
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    summary = pd.DataFrame(totals, columns=["Account Code", "Account Name", "Total"])
    summary["Total"] = pd.to_numeric(summary["Total"], errors="coerce")
    if grand_total:
        summary = pd.concat([summary, pd.DataFrame(
            [["", "GRAND TOTAL", pd.to_numeric(grand_total, errors="coerce")]],
            columns=summary.columns)], ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Audit Detail", index=False)
        if len(summary):
            summary.to_excel(writer, sheet_name="Totals", index=False)
        wb = writer.book
        ba_idx = list(df.columns).index("B/A") + 1
        _style_sheet(wb["Audit Detail"], len(df.columns), after_col_idx=ba_idx)
        if "Totals" in wb.sheetnames:
            _style_sheet(wb["Totals"], len(summary.columns))
    output.seek(0)
    return output.getvalue(), len(df)


def build_generic_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
        wb = writer.book
        for name, df in sheets.items():
            _style_sheet(wb[name], max(len(df.columns), 1))
    output.seek(0)
    return output.getvalue()


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Drag and drop a PDF here", type=["pdf"], accept_multiple_files=False
)

if uploaded_file is not None:
    with st.spinner("Converting…"):
        try:
            file_bytes = uploaded_file.read()
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                first_text = pdf.pages[0].extract_text() if pdf.pages else ""
                is_audit = REPORT_SIGNATURE in (first_text or "")

                if is_audit:
                    rows, totals, grand_total = parse_audit_report(pdf)
                    excel_bytes, n = build_audit_workbook(rows, totals, grand_total)
                    msg = f"Parsed audit report — {n} rows across labelled columns."
                else:
                    sheets = extract_generic(pdf)
                    excel_bytes = build_generic_workbook(sheets)
                    msg = f"Converted — {len(sheets)} sheet(s)."
        except Exception as exc:
            st.error(f"Something went wrong during conversion: {exc}")
            excel_bytes = None

    if excel_bytes:
        st.success(msg)
        out_name = uploaded_file.name.rsplit(".", 1)[0] + ".xlsx"
        st.download_button(
            label="⬇️  Download Excel file",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
