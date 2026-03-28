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
from openpyxl.styles import Alignment, Font, PatternFill

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mission-freight-invoice-tool")

CHARGE_KEYWORDS = [
    "import warehouse charges",
    "handling",
    "handling charges",
    "handling fee",
]


@dataclass
class InvoiceResult:
    bestandsnaam: str
    factuurnummer: Optional[str]
    awb_nummer: Optional[str]
    totaal_kg: Optional[float]
    import_warehouse_charges_totaal_eur: Optional[float]
    prijs_per_kg_eur: Optional[float]
    status: str


def normalize_spaces(text: str) -> str:
    text = text.replace("\xa0", " ").replace("\u200b", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def parse_decimal_eu(value: str) -> Optional[Decimal]:
    cleaned = value.strip().replace("EUR", "").replace("€", "").replace(" ", "")
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def extract_text_from_pdf_bytes(data: bytes) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF is niet beschikbaar. Installeer pymupdf.")
    with fitz.open(stream=data, filetype="pdf") as doc:
        return "\n".join(page.get_text("text") for page in doc)


def extract_words_from_pdf_bytes(data: bytes) -> list:
    if fitz is None:
        raise RuntimeError("PyMuPDF is niet beschikbaar. Installeer pymupdf.")
    all_words = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            all_words.extend(page.get_text("words"))
    return all_words


def find_invoice_number(text: str) -> Optional[str]:
    patterns = [
        r"FACTUURNUMMER\s+([0-9]{5,})",
        r"FACTUURNUMMER\s*[:\-]?\s*([A-Z0-9\-]{5,})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def find_awb_number(text: str) -> Optional[str]:
    patterns = [
        r"AWB\s*NUMMER\s*[:\-]?\s*([0-9]{3}-[0-9]{8,})",
        r"\b([0-9]{3}-[0-9]{8,})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def find_total_weight_kg_from_words(words: list) -> Optional[Decimal]:
    """
    Pakt de waarde onder de kolom 'BRUTO' uit de PDF-layout.
    Werkt op woord-posities in plaats van platte tekstregels.
    """
    if not words:
        return None

    bruto_candidates = []
    for w in words:
        text = str(w[4]).strip().lower()
        if text == "bruto":
            bruto_candidates.append(w)

    if not bruto_candidates:
        return None

    # Neem de laatste/laagste BRUTO in de pagina
    bruto_word = sorted(bruto_candidates, key=lambda x: (x[1], x[0]))[-1]
    bx0, by0, bx1, by1 = bruto_word[:4]
    bcx = (bx0 + bx1) / 2

    numeric_candidates = []
    for w in words:
        text = str(w[4]).strip()
        if not re.fullmatch(r"\d+(?:[.,]\d+)?", text):
            continue

        x0, y0, x1, y1 = w[:4]
        cx = (x0 + x1) / 2

        # Alleen woorden onder BRUTO en ongeveer in dezelfde kolom
        if y0 <= by1:
            continue
        if abs(cx - bcx) > 80:
            continue

        value = parse_decimal_eu(text)
        if value is None:
            continue

        numeric_candidates.append((y0, value))

    if not numeric_candidates:
        return None

    # Neem de dichtstbijzijnde numerieke waarde onder BRUTO
    numeric_candidates.sort(key=lambda t: t[0])
    return numeric_candidates[0][1]


def sum_import_warehouse_charges(text: str) -> Optional[Decimal]:
    total = Decimal("0")
    found = False

    normalized_text = normalize_spaces(text).lower()
    normalized_text = re.sub(r"\s+", " ", normalized_text)

    matches = re.findall(
        r"(import warehouse charges|handling|handling charges|handling fee).*?([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{2}))",
        normalized_text,
        re.IGNORECASE,
    )

    for _, amount_str in matches:
        amount = parse_decimal_eu(amount_str)
        if amount is not None:
            total += amount
            found = True

    return total if found else None


def parse_invoice(file_name: str, data: bytes) -> InvoiceResult:
    try:
        text = extract_text_from_pdf_bytes(data)
        words = extract_words_from_pdf_bytes(data)

        factuurnummer = find_invoice_number(text)
        awb_nummer = find_awb_number(text)
        totaal_kg = find_total_weight_kg_from_words(words)
        charges = sum_import_warehouse_charges(text)

        prijs_per_kg = None
        if charges is not None and totaal_kg not in (None, Decimal("0")):
            prijs_per_kg = charges / totaal_kg

        missing = []
        if not factuurnummer:
            missing.append("factuurnummer")
        if not awb_nummer:
            missing.append("AWB nummer")
        if totaal_kg is None:
            missing.append("totaal kg")
        if charges is None:
            missing.append("Import warehouse charges")

        status = "OK" if not missing else f"Ontbreekt: {', '.join(missing)}"

        return InvoiceResult(
            bestandsnaam=file_name,
            factuurnummer=factuurnummer,
            awb_nummer=awb_nummer,
            totaal_kg=float(totaal_kg) if totaal_kg is not None else None,
            import_warehouse_charges_totaal_eur=float(charges) if charges is not None else None,
            prijs_per_kg_eur=float(prijs_per_kg) if prijs_per_kg is not None else None,
            status=status,
        )
    except Exception as exc:
        return InvoiceResult(
            bestandsnaam=file_name,
            factuurnummer=None,
            awb_nummer=None,
            totaal_kg=None,
            import_warehouse_charges_totaal_eur=None,
            prijs_per_kg_eur=None,
            status=f"Fout bij verwerken: {exc}",
        )


def dataframe_from_results(results: List[InvoiceResult]) -> pd.DataFrame:
    rows = []
    for r in results:
        rows.append(
            {
                "Factuurnummer": r.factuurnummer,
                "AWB nummer": r.awb_nummer,
                "Totaal kg": r.totaal_kg,
                "Prijs per kg (EUR)": r.prijs_per_kg_eur,
                "Warehouse charges totaal (EUR)": r.import_warehouse_charges_totaal_eur,
                "Bestandsnaam": r.bestandsnaam,
                "Status": r.status,
            }
        )
    return pd.DataFrame(rows)


def build_excel_bytes(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultaten"

    headers = ["Factuurnummer", "AWB nummer", "Totaal kg", "Prijs per kg (EUR)"]
    ws.append(headers)

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for col, _header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        ws.append([
            row.get("Factuurnummer"),
            row.get("AWB nummer"),
            row.get("Totaal kg"),
            row.get("Prijs per kg (EUR)"),
        ])

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 18

    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = "0.000"
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = "0.00000"

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()


def load_session_df() -> pd.DataFrame:
    raw = session.get("results_json")
    if not raw:
        return pd.DataFrame(
            columns=[
                "Factuurnummer",
                "AWB nummer",
                "Totaal kg",
                "Prijs per kg (EUR)",
                "Warehouse charges totaal (EUR)",
                "Bestandsnaam",
                "Status",
            ]
        )
    return pd.read_json(io.StringIO(raw))


HTML = """
<!doctype html>
<html lang="nl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>Mission Freight factuur uitlezer</title>
  <style>
    :root {
      --bg: #f5f7fb;
      --card: #ffffff;
      --text: #14213d;
      --muted: #5b6475;
      --primary: #0f62fe;
      --border: #dbe2f0;
      --ok: #117a37;
      --warn: #b26a00;
    }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: var(--bg); color: var(--text); }
    .wrap { max-width: 900px; margin: 0 auto; padding: 16px; }
    .card { background: var(--card); border: 1px solid var(--border); border-radius: 18px; padding: 18px; box-shadow: 0 8px 24px rgba(20, 33, 61, 0.06); margin-bottom: 16px; }
    h1 { font-size: 1.55rem; margin: 0 0 8px; }
    h2 { font-size: 1.2rem; margin: 0 0 12px; }
    p { margin: 0 0 10px; color: var(--muted); line-height: 1.5; }
    .dropzone { border: 2px dashed var(--primary); border-radius: 18px; padding: 22px; text-align: center; background: #f8fbff; }
    .dropzone.dragover { background: #eef5ff; }
    input[type=file] { width: 100%; margin-top: 12px; font-size: 16px; }
    .btn { display: inline-block; width: 100%; border: 0; border-radius: 14px; padding: 14px 16px; font-size: 16px; font-weight: 600; text-decoration: none; text-align: center; margin-top: 12px; cursor: pointer; }
    .btn-primary { background: var(--primary); color: white; }
    .btn-secondary { background: #eef3ff; color: var(--primary); }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
    .chip { background: #eef3ff; color: var(--primary); padding: 8px 10px; border-radius: 999px; font-size: 13px; }
    table { width: 100%; border
