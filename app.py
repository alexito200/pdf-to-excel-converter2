"""
PDF-To-Excel Converter (Streamlit)

- Detects the "Daily Order Detail Maintenance Audit Report" and parses it into
  labelled columns.
- Falls back to generic table/text extraction for any other PDF, or if the
  audit parser hits something unexpected.
- Excel formatting is best-effort and never blocks the download.

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
    "Drop a PDF in the box below (or click **Browse files**). "
    "You'll get a formatted Excel file to download."
)

# ----------------------------------------------------------------------
# Audit-report config
# ----------------------------------------------------------------------
AUDIT_COLUMNS = [
    "Account Code", "Account Name", "Page",
    "P.O.#", "Reg#", "Ibm#", "Line", "Trn.Date", "Trn.Time", "User", "B/A",
    "Style", "Col", "Lb", "Tgt Style", "Tgt Col", "Tgt Lb",
    "Customer", "Customer Style", "Str.Date", "Cnx.Date",
    "Price", "Qty", "Rules", "Mode", "PTY", "Value", "Unparsed",
]
NUMERIC_COLUMNS = {"Price", "Qty", "Value"}
CUST_STYLE_RE = re.compile(r"^\d{9,}(\*\d+)?$")
ACCOUNT_RE = re.compile(r"Auto Maintenance Account:\s*(\S+)\s+(.+?)\s+Program:")
REPORT_SIGNATURE = "Daily Order Detail Maintenance Audit Report"
LEAD_KEYS = ["P.O.#", "Reg#", "Ibm#", "Line", "Trn.Date", "Trn.Time", "User"]


def _pad3(group):
    g = list(group) + ["", "", ""]
    return g[0], g[1], g[2]


def parse_payload(tokens):
    out = {c: "" for c in AUDIT_COLUMNS}
    if not tokens:
        return out
    if tokens == ["New"]:
        out["Style"] = "New"
        return out

    idx = next((i for i, t in enumerate(tokens) if CUST_STYLE_RE.match(t)), None)
    if idx is None:
        if len(tokens) >= 2 and tokens[0].count("/") == 2:
            out["Str.Date"], out["Cnx.Date"] = tokens[0], tokens[1]
        else:
            out["Unparsed"] = " ".join(tokens)
        return out

    out["Customer"] = tokens[idx - 1]
    out["Customer Style"] = tokens[idx]

    group = tokens[: idx - 1]
    if len(group) <= 3:
        out["Style"], out["Col"], out["Lb"] = _pad3(group)
    else:
        if len(group) % 2:
            g1, g2 = group[:3], group[3:]
        else:
            half = len(group) // 2
            g1, g2 = group[:half], group[half:]
        out["Style"], out["Col"], out["Lb"] = _pad3(g1)
        out["Tgt Style"], out["Tgt Col"], out["Tgt Lb"] = _pad3(g2)

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
    rows, totals, grand_total = [], [], None
    account_code = account_name = ""
    lead = {k: "" for k in LEAD_KEYS}

    for page_num, page in enumerate(pdf.pages, start=1):
        text = page.extract_text() or ""
        for raw in text.splitlines():
            line = raw.strip()
            if (not line or REPORT_SIGNATURE in line
                    or line.startswith("P.O.#") or line == "Target"):
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
                totals.append((account_code, account_name, nums[-1] if nums else ""))
                continue

            if "Before:" in line:
                left, _, right = line.partition("Before:")
                lead_tokens = left.split()
                if len(lead_tokens) >= 7:
                    lead = dict(zip(LEAD_KEYS, lead_tokens[:7]))
                ba, payload = "Before", right.split()
            elif line.startswith("After:"):
                ba, payload = "After", line.partition("After:")[2].split()
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
# Generic fallback
# ----------------------------------------------------------------------
def extract_generic(pdf):
    sheets, table_count, text_rows = {}, 0, []
    for page_num, page in enumerate(pdf.pages, start=1):
        tables = page.extract_tables() or []
        for t_idx, table in enumerate(tables, start=1):
            if not table:
                continue
            width = max(len(r) for r in table)
            norm = [list(r) + [""] * (width - len(r)) for r in table]
            df = pd.DataFrame(norm)
            if len(df) > 1:
                df.columns = [str(c) if c not in (None, "") else f"Col{i+1}"
                              for i, c in enumerate(df.iloc[0])]
                df = df.iloc[1:].reset_index(drop=True)
            sheets[f"P{page_num}_T{t_idx}"[:31]] = df
            table_count += 1
        if not tables:
            for ln in (page.extract_text() or "").splitlines():
                if ln.strip():
                    text_rows.append([page_num, ln])
    if table_count == 0:
        df = (pd.DataFrame(text_rows, columns=["Page", "Text"]) if text_rows
              else pd.DataFrame({"Note": ["No extractable text or tables found."]}))
        sheets = {"Extracted Text": df}
    return sheets


# ----------------------------------------------------------------------
# Formatting (best-effort; never raises)
# ----------------------------------------------------------------------
def _style_sheet(ws, n_cols, after_col_idx=None):
    try:
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(bold=True, color="FFFFFF")
        after_fill = PatternFill("solid", fgColor="F2F2F2")

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(max(n_cols, 1))}{ws.max_row}"

        for col_cells in ws.columns:
            letter = col_cells[0].column_letter
            longest = max((len(str(c.value)) for c in col_cells if c.value is not None),
                          default=8)
            ws.column_dimensions[letter].width = min(max(longest + 2, 8), 40)

        if after_col_idx is not None:
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=after_col_idx).value == "After":
                    for c in range(1, n_cols + 1):
                        ws.cell(row=r, column=c).fill = after_fill
    except Exception:
        pass  # styling is optional; keep the file usable


def build_audit_workbook(rows, totals, grand_total):
    df = pd.DataFrame(rows, columns=AUDIT_COLUMNS)
    if "Unparsed" in df.columns and not df["Unparsed"].astype(str).str.strip().any():
        df = df.drop(columns=["Unparsed"])
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    summary = pd.DataFrame(totals, columns=["Account Code", "Account Name", "Total"])
    if len(summary):
        summary["Total"] = pd.to_numeric(summary["Total"], errors="coerce")
    if grand_total:
        summary = pd.concat(
            [summary, pd.DataFrame(
                [["", "GRAND TOTAL", pd.to_numeric(grand_total, errors="coerce")]],
                columns=["Account Code", "Account Name", "Total"])],
            ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Audit Detail", index=False)
        if len(summary):
            summary.to_excel(writer, sheet_name="Totals", index=False)
        wb = writer.book
        ba_idx = list(df.columns).index("B/A") + 1 if "B/A" in df.columns else None
        _style_sheet(wb["Audit Detail"], len(df.columns), after_col_idx=ba_idx)
        if "Totals" in wb.sheetnames:
            _style_sheet(wb["Totals"], len(summary.columns))
    output.seek(0)
    return output.getvalue(), len(df)


def build_generic_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        wb = writer.book
        for name, df in sheets.items():
            _style_sheet(wb[name[:31]], max(len(df.columns), 1))
    output.seek(0)
    return output.getvalue()


def convert(file_bytes):
    """Return (excel_bytes, message). Tries audit parse, falls back to generic."""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        first_text = pdf.pages[0].extract_text() if pdf.pages else ""
        if REPORT_SIGNATURE in (first_text or ""):
            try:
                rows, totals, grand = parse_audit_report(pdf)
                if rows:
                    data, n = build_audit_workbook(rows, totals, grand)
                    return data, f"Parsed audit report — {n} rows in labelled columns."
            except Exception:
                pass  # fall through to generic
        sheets = extract_generic(pdf)
        return build_generic_workbook(sheets), f"Converted — {len(sheets)} sheet(s)."


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Drag and drop a PDF here", type=["pdf"], accept_multiple_files=False
)

if uploaded_file is not None:
    excel_bytes = None
    with st.spinner("Converting…"):
        try:
            excel_bytes, message = convert(uploaded_file.read())
        except Exception:
            st.error("Conversion failed. Open the details below and share them so this can be fixed.")
            with st.expander("Error details"):
                st.code(traceback.format_exc())

    if excel_bytes:
        st.success(message)
        out_name = uploaded_file.name.rsplit(".", 1)[0] + ".xlsx"
        st.download_button(
            label="⬇️  Download Excel file",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
