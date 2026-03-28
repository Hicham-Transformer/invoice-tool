from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import List, Optional

import pandas as pd
from flask import Flask, Response, redirect, render_template_string, request, send_file, session, url_for
from openpyxl import Workbook

try:
    import fitz
except Exception:
    fitz = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "invoice-tool")


# =====================
# MODELS
# =====================
@dataclass
class InvoiceResult:
    bestandsnaam: str
    factuurnummer: Optional[str]
    awb_nummer: Optional[str]
    totaal_kg: Optional[float]
    charges: Optional[float]
    prijs_per_kg: Optional[float]
    status: str


# =====================
# HELPERS
# =====================
def parse_decimal(value: str) -> Optional[Decimal]:
    value = value.replace(",", ".")
    try:
        return Decimal(value)
    except InvalidOperation:
        return None


def extract_text(data: bytes) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF ontbreekt")
    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(page.get_text() for page in doc)


# =====================
# CORE LOGIC
# =====================

def find_awb(text: str) -> Optional[str]:
    m = re.search(r"\b\d{3}-\d{8}\b", text)
    return m.group(0) if m else None


def find_invoice(text: str) -> Optional[str]:
    m = re.search(r"\b20\d{6}\b", text)
    return m.group(0) if m else None


# 🔥 HIER ZIT DE MAGIC (gewicht fix)
def find_weight(text: str) -> Optional[Decimal]:
    text = text.lower()

    # 1. probeer bruto direct
    m = re.search(r"bruto[^0-9]{0,20}(\d{2,5})", text)
    if m:
        return parse_decimal(m.group(1))

    # 2. fallback → pak grootste getal (werkt bij jouw PDF)
    nums = re.findall(r"\b\d{2,5}\b", text)
    nums = [int(n) for n in nums]

    if nums:
        return Decimal(max(nums))

    return None


# 🔥 charges simpel en robuust
def find_charges(text: str) -> Optional[Decimal]:
    text = text.lower()

    total = Decimal("0")
    found = False

    for line in text.splitlines():
        if "import warehouse" in line or "handling" in line:
            nums = re.findall(r"\d+[.,]\d{2}", line)
            if nums:
                val = parse_decimal(nums[-1])
                if val:
                    total += val
                    found = True

    return total if found else None


def parse_invoice(file_name: str, data: bytes) -> InvoiceResult:
    try:
        text = extract_text(data)

        invoice = find_invoice(text)
        awb = find_awb(text)
        weight = find_weight(text)
        charges = find_charges(text)

        price = None
        if weight and charges and weight != 0:
            price = charges / weight

        missing = []
        if not invoice:
            missing.append("factuurnummer")
        if not awb:
            missing.append("awb")
        if not weight:
            missing.append("gewicht")
        if not charges:
            missing.append("charges")

        status = "OK" if not missing else f"Ontbreekt: {', '.join(missing)}"

        return InvoiceResult(
            bestandsnaam=file_name,
            factuurnummer=invoice,
            awb_nummer=awb,
            totaal_kg=float(weight) if weight else None,
            charges=float(charges) if charges else None,
            prijs_per_kg=float(price) if price else None,
            status=status,
        )

    except Exception as e:
        return InvoiceResult(file_name, None, None, None, None, None, str(e))


# =====================
# DATAFRAME
# =====================
def df_from_results(results):
    return pd.DataFrame([r.__dict__ for r in results])


def build_excel(df):
    wb = Workbook()
    ws = wb.active

    ws.append(["Factuur", "AWB", "KG", "Prijs/kg"])

    for _, row in df.iterrows():
        ws.append([
            row["factuurnummer"],
            row["awb_nummer"],
            row["totaal_kg"],
            row["prijs_per_kg"],
        ])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# =====================
# ROUTES
# =====================
@app.get("/")
def index():
    df = session.get("df")
    rows = df if df else []
    return render_template_string("""
    <h2>Upload PDF</h2>
    <form method="post" action="/upload" enctype="multipart/form-data">
        <input type="file" name="files" multiple>
        <button>Upload</button>
    </form>

    <h2>Result</h2>
    {% for r in rows %}
        <p>{{ r }}</p>
    {% endfor %}
    """, rows=rows)


@app.post("/upload")
def upload():
    files = request.files.getlist("files")
    results = []

    for f in files:
        if f.filename.endswith(".pdf"):
            results.append(parse_invoice(f.filename, f.read()))

    df = df_from_results(results)
    session["df"] = df.to_dict(orient="records")

    return redirect("/")


@app.get("/excel")
def excel():
    df = pd.DataFrame(session.get("df", []))
    return send_file(
        build_excel(df),
        as_attachment=True,
        download_name="result.xlsx"
    )


@app.get("/health")
def health():
    return Response("ok")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
