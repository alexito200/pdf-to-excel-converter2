# PDF-To-Excel Converter

A lightweight web app that converts a PDF into a downloadable Excel (`.xlsx`) file.
Built with [Streamlit](https://streamlit.io). Drop in a PDF (or browse for one),
and the app extracts any tables it finds — falling back to raw text if there are
no tables — then hands you an Excel file.

## How it works

- Tables are extracted with `pdfplumber`. Each table becomes its own sheet,
  named `P<page>_T<table>` (e.g. `P1_T1` = page 1, first table).
- If no tables are detected anywhere in the document, all the text is placed in
  a single sheet called **Extracted Text** so nothing is lost.

## Run it — no local Python needed

You don't have to install anything on your work computer. Two options:

### Option A — Deploy to Streamlit Community Cloud (recommended)

1. Push these files (`app.py`, `requirements.txt`) to a GitHub repository.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **New app**, pick this repo, and set the main file to `app.py`.
4. Click **Deploy**. Streamlit installs the requirements and gives you a URL.

Anyone with the URL can use it in a browser — no installs on their end either.

### Option B — Run inside GitHub Codespaces

Codespaces gives you a cloud machine with Python already installed.

```bash
pip install -r requirements.txt
streamlit run app.py
```

When the server starts, Codespaces will offer to open the forwarded port in a
browser tab — that's your app.

## Files

| File               | Purpose                                  |
|--------------------|------------------------------------------|
| `app.py`           | The Streamlit application.               |
| `requirements.txt` | Python packages Streamlit Cloud installs.|

## Notes

- `pdfplumber` works best on PDFs with real, selectable text. Scanned/image-only
  PDFs would need OCR (e.g. Tesseract), which is intentionally left out to keep
  this lightweight.
- A table that spans multiple pages will come out as one sheet per page. You can
  stitch those together in Excel if needed.
