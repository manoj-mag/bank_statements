from __future__ import annotations

import secrets
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from amex_statement_extractor import dataframe_to_excel_bytes, run_extraction
from hsbc_statement_extractor import hsbc_dataframe_to_excel_bytes, run_hsbc_extraction

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "webapp_temp"
ALLOWED_EXTENSIONS = {".pdf"}

UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

HEADER_FIELDS = [
    ("Transaction Date", "Transaction Date, Trans Date, Date"),
    ("Posting Date", "Posting Date, Value Date"),
    ("Details", "Details, Description, Narration"),
    ("Non AED Spending", "Foreign Amount, Original Amount"),
    ("Amount in AED", "Amount, Billing Amount, Debit/Credit"),
]

CARD_OPTIONS = [
    ("amex", "AMEX"),
    ("hsbc", "HSBC"),
]


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template(
        "index.html",
        card_options=CARD_OPTIONS,
        header_fields=HEADER_FIELDS,
        form_data={"card_type": "amex"},
        result=None,
        error=None,
    )


@app.post("/extract")
def extract():
    uploaded_file = request.files.get("statement_pdf")
    form_data = {key: request.form.get(key, "") for key, _ in HEADER_FIELDS}
    form_data["card_type"] = request.form.get("card_type", "amex")
    selected_card = form_data["card_type"]

    if not uploaded_file or not uploaded_file.filename:
        return render_template(
            "index.html",
            card_options=CARD_OPTIONS,
            header_fields=HEADER_FIELDS,
            form_data=form_data,
            result=None,
            error="Please upload a PDF statement first.",
        ), 400

    if not allowed_file(uploaded_file.filename):
        return render_template(
            "index.html",
            card_options=CARD_OPTIONS,
            header_fields=HEADER_FIELDS,
            form_data=form_data,
            result=None,
            error="Only PDF files are supported right now.",
        ), 400

    safe_name = secure_filename(uploaded_file.filename) or "statement.pdf"
    token = secrets.token_hex(8)
    pdf_path = UPLOAD_DIR / f"{token}_{safe_name}"
    excel_path = UPLOAD_DIR / f"{token}_transactions.xlsx"
    uploaded_file.save(pdf_path)

    try:
        if selected_card == "hsbc":
            df, summary = run_hsbc_extraction(pdf_path)
        else:
            custom_aliases = {key: value for key, value in form_data.items() if key != "card_type"}
            df, summary = run_extraction(pdf_path, custom_aliases=custom_aliases)
    except Exception as exc:  # pragma: no cover - keeps UI resilient
        if pdf_path.exists():
            pdf_path.unlink()
        return render_template(
            "index.html",
            card_options=CARD_OPTIONS,
            header_fields=HEADER_FIELDS,
            form_data=form_data,
            result=None,
            error=f"Could not parse that statement: {exc}",
        ), 500

    if df.empty:
        if pdf_path.exists():
            pdf_path.unlink()
        return render_template(
            "index.html",
            card_options=CARD_OPTIONS,
            header_fields=HEADER_FIELDS,
            form_data=form_data,
            result=None,
            error="No transactions were found. Try a different card type or add the header names used by this statement.",
        ), 422

    if selected_card == "hsbc":
        excel_path.write_bytes(hsbc_dataframe_to_excel_bytes(df))
    else:
        excel_path.write_bytes(dataframe_to_excel_bytes(df, summary))

    preview_rows = df.to_dict(orient="records")
    result = {
        "download_url": url_for("download", token=token),
        "filename": excel_path.name,
        "card_label": dict(CARD_OPTIONS).get(selected_card, selected_card.upper()),
        "row_count": len(df),
        "credit_count": int((df["Amount (AED)" if selected_card == "hsbc" else "Amount in AED"] < 0).sum()),
        "debit_count": int((df["Amount (AED)" if selected_card == "hsbc" else "Amount in AED"] >= 0).sum()),
        "summary": summary,
        "preview_rows": preview_rows,
        "columns": list(df.columns),
    }

    return render_template(
        "index.html",
        card_options=CARD_OPTIONS,
        header_fields=HEADER_FIELDS,
        form_data=form_data,
        result=result,
        error=None,
    )


@app.get("/download/<token>")
def download(token: str):
    matches = list(UPLOAD_DIR.glob(f"{token}_transactions.xlsx"))
    if not matches:
        abort(404)
    file_path = matches[0]
    return send_file(file_path, as_attachment=True, download_name="transactions.xlsx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)