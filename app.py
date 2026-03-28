from __future__ import annotations

import io
import os
import re
from dataclasses import dataclass
from decimal import Decimal
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
# MODEL
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
    except:
        return None


def extract_text(data: bytes) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF ontbreekt")
    doc = fitz.open(stream=data, filetype="pdf")
    return "\n".join(page.get_text() for page in doc)


# =====================
# PARSING
# =====================

def find_awb(text: str) -> Optional[str]:
    m = re.search(r"\b\d{3}-\d{8}\b", text)
    return m.group(0) if m else None


def find_invoice(text: str) -> Optional[str]:
    m = re.search(r"\b20\d{6}\b", text)
    return m.group(0) if m else None


# ✅ PERFECTE FIX VOOR JOUW PDF
def find_weight(text: str) -> Optional[Decimal]:
    text = text.lower()

    # zoek "bruto" en pak eerste getal daarna
    match = re.search(r"bruto[^0-9]{0,30}(\d{2,5})", text)

    if match:
        return parse_decimal(match.group(1))

    return None


def find_charges(text: str) -> Optional[Decimal]:
    total = Decimal("0")
    found = False

    for line in text.lower().splitlines():

        if "import warehouse charges" in line or "handling" in line:

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
        if weight and charges:
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
            file_name,
            invoice,
            awb,
            float(weight) if weight else None,
            float(charges) if charges else None,
            float(price) if price else None,
            status,
        )

    except Exception as e:
        return InvoiceResult(file_name, None, None, None, None, None, str(e))


# =====================
# DATA
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
# UI (FIXED)
# =====================
HTML = """
<!doctype html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body { font-family: Arial; padding:20px; background:#f5f5f5; }
.card { background:white; padding:15px; border-radius:10px; margin-bottom:10px; }
button { padding:10px; width:100%; }
</style>
</head>
<body>

<div class="card">
<h2>Upload PDF</h2>
<form method="post" action="/upload" enctype="multipart/form-data">
<input type="file" name="files" multiple>
<button>Verwerken</button>
</form>
</div>

{% for r in rows %}
<div class="card">
<b>Factuur:</b> {{ r.factuurnummer }}<br>
<b>AWB:</b> {{ r.awb_nummer }}<br>
<b>KG:</b> {{ r.totaal_kg }}<br>
<b>Prijs/kg:</b> {{ r.prijs_per_kg }}<br>
<b>Status:</b> {{ r.status }}
</div>
{% endfor %}

</body>
</html>
"""


# =====================
# ROUTES
# =====================
@app.get("/")
def index():
    data = session.get("data", [])
    return render_template_string(HTML, rows=data)


@app.post("/upload")
def upload():
    files = request.files.getlist("files")
    results = []

    for f in files:
        if f.filename.endswith(".pdf"):
            results.append(parse_invoice(f.filename, f.read()))

    session["data"] = [r.__dict__ for r in results]

    return redirect("/")


@app.get("/health")
def health():
    return Response("ok")


# =====================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
