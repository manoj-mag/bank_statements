# Statement to Excel Web Tool

This project now supports two ways of working:

1. CLI extraction for the current AMEX UAE statement format.
2. A small web app for AMEX and HSBC card statements, with optional header hints when column names differ.

## Install

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
brew install tesseract
```

## Run the website

```bash
.venv/bin/python app.py
```

Then open `http://127.0.0.1:5050`.

## Run the CLI version

```bash
.venv/bin/python amex_statement_extractor.py "STATEMENT.pdf"
```

## Header hints

If another bank statement uses different column names, enter them in the web form as comma-separated aliases, for example:

- `Transaction Date`: `Txn Date, Date`
- `Details`: `Description, Narration`
- `Amount in AED`: `Billing Amount, Debit Amount`

The AMEX parser first tries table-header matching and then raw text parsing.

The HSBC parser is tuned for the statement layout with:

- `Posting Date`
- `Transaction Date`
- `Transaction Details`
- `Amount (AED)`

For some HSBC PDFs, the embedded text layer is corrupted even though the statement looks normal on screen. In those cases the app now falls back to OCR on the host machine, which requires `tesseract` to be installed on the Mac running the site.
