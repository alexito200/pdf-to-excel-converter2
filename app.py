"""
PDF-To-Excel Converter
A lightweight Streamlit app that extracts tables (and text as a fallback)
from a PDF and packages them into a downloadable Excel workbook.

Run locally / in Codespaces:   streamlit run app.py
"""

import io

import pandas as pd
import pdfplumber
import streamlit as st

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
    "The app extracts any tables it finds — falling back to raw text — "
    "and gives you an Excel file to download."
)


# ----------------------------------------------------------------------
# Conversion logic
# ----------------------------------------------------------------------
def extract_pdf_to_sheets(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    """Turn a PDF (as bytes) into a dict of {sheet_name: DataFrame}.

    - Every table found becomes its own sheet, named P<page>_T<table>.
    - If no tables are found anywhere, all text is dumped into a single
      "Extracted Text" sheet so nothing is lost.
    """
    sheets: dict[str, pd.DataFrame] = {}
    table_count = 0
    text_rows: list[list] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()

            if tables:
                for t_idx, table in enumerate(tables, start=1):
                    if not table:
                        continue
                    df = pd.DataFrame(table)

                    # If there's more than one row, promote the first row
                    # to column headers for a cleaner-looking sheet.
                    if len(df) > 1:
                        df.columns = [
                            str(c) if c is not None else "" for c in df.iloc[0]
                        ]
                        df = df.iloc[1:].reset_index(drop=True)

                    sheet_name = f"P{page_num}_T{t_idx}"[:31]  # Excel limit
                    sheets[sheet_name] = df
                    table_count += 1
            else:
                # No tables on this page — keep the text for the fallback.
                text = page.extract_text() or ""
                for line in text.splitlines():
                    if line.strip():
                        text_rows.append([page_num, line])

    if table_count == 0:
        if text_rows:
            sheets = {
                "Extracted Text": pd.DataFrame(text_rows, columns=["Page", "Text"])
            }
        else:
            sheets = {
                "Extracted Text": pd.DataFrame(
                    {"Note": ["No extractable text or tables found in this PDF."]}
                )
            }

    return sheets


def sheets_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Write the sheets to an in-memory .xlsx file and return the bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)
    return output.getvalue()


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Drag and drop a PDF here",
    type=["pdf"],
    accept_multiple_files=False,
)

if uploaded_file is not None:
    with st.spinner("Converting…"):
        try:
            file_bytes = uploaded_file.read()
            sheets = extract_pdf_to_sheets(file_bytes)
            excel_bytes = sheets_to_excel_bytes(sheets)
        except Exception as exc:  # surface a friendly message, not a stack trace
            st.error(f"Something went wrong during conversion: {exc}")
            excel_bytes = None

    if excel_bytes:
        st.success(
            f"Done — found {len(sheets)} sheet(s). Click below to download."
        )
        out_name = uploaded_file.name.rsplit(".", 1)[0] + ".xlsx"
        st.download_button(
            label="⬇️  Download Excel file",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
