#!/usr/bin/env python3
import io
import os
import re
import traceback
from datetime import datetime
from typing import Dict, Any, Optional, List

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook

# ================================================================
#  Helper: Extract invoice ID (number only from filename)
# ================================================================
def extract_invoice_id(filename: str):
    filename = filename.upper()

    # SY invoices → SY0050227 or SY0050227A
    m = re.search(r"(SY\d+[A-Z]?)", filename)
    if m:
        return m.group(1)

    # KN invoices → e.g. "DN CAD 26693.pdf" → 26693
    m = re.search(r"(\d{4,8})", filename)
    if m:
        return m.group(1)

    return filename


# ================================================================
#  Helper: Extract currency from filename (USD/CAD/EUR)
# ================================================================
def extract_currency_from_filename(filename: str):
    filename = filename.upper()

    if "USD" in filename:
        return "USD"
    if "CAD" in filename:
        return "CAD"
    if "EUR" in filename:
        return "EUR"

    return ""


# ================================================================
#  Fixed Output Columns
# ================================================================
HEADERS = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_Rate",
]

# ================================================================
#  General Helper Functions + Regex
# ================================================================
_f = lambda s: float(s.replace(",", "")) if s else None
_to_kg = lambda v, u: v if u.lower().startswith("kg") else v * 0.453592

INVOICE_DATE_PAT = re.compile(
    r"(?:INVOICE\s*DATE|DATE\s+DE\s+LA\s+FACTURE)\s*[:\-]?\s*(\d{4}-\d{2}-\d{2})",
    re.I,
)

SHIPPER_PAT = re.compile(
    r"(?:SHIPPER['S]*\s*NAME|NOM\s+DE\s+L['’]EXP[ÉE]DITEUR)\s*[:\-]?\s*(.+?)(?:\n[A-Z]|$)",
    re.I | re.S,
)

WEIGHT_PAT = re.compile(r"(\d+(?:\.\d+)?)\s*KG", re.I)
VOLUME_PAT = re.compile(r"(\d+(?:\.\d+)?)\s*M3", re.I)
CHARGEABLE_KG_PAT = re.compile(r"CHW.*?(\d+(?:\.\d+)?)", re.I)
PIECES_PAT = re.compile(r"(\d+)\s*PCS", re.I)

# Subtotal (below AMOUNT or Total)
SUBTOTAL_PAT = re.compile(
    r"Total\s*[:\-]?\s*([\d,]+\.\d{2})",
    re.I
)

# KN Air freight → USD xxxx.xx at the end of line
KN_AIR_FREIGHT = re.compile(
    r"AIRFREIGHT.*?USD\s*([\d,]+\.\d{2})",
    re.I | re.S
)


# ================================================================
#  MAIN PARSE FUNCTION
# ================================================================
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        # ----------------------
        # Invoice Date
        # ----------------------
        inv_date = None
        m = INVOICE_DATE_PAT.search(text)
        if m:
            inv_date = m.group(1).strip()

        # ----------------------
        # Currency (from filename)
        # ----------------------
        currency = extract_currency_from_filename(filename)

        # ----------------------
        # Shipper
        # ----------------------
        shipper = None
        m = SHIPPER_PAT.search(text)
        if m:
            shipper = m.group(1).strip().replace("\n", " ")

        # ----------------------
        # Weight / Volume / Pieces
        # ----------------------
        weight = None
        m = WEIGHT_PAT.search(text)
        if m:
            weight = _f(m.group(1))

        volume = None
        m = VOLUME_PAT.search(text)
        if m:
            volume = _f(m.group(1))

        chargeable_kg = None
        m = CHARGEABLE_KG_PAT.search(text)
        if m:
            chargeable_kg = _f(m.group(1))

        pieces = None
        m = PIECES_PAT.search(text)
        if m:
            pieces = int(m.group(1))

        # ----------------------
        # Subtotal (Total)
        # ----------------------
        subtotal = None
        m = SUBTOTAL_PAT.search(text)
        if m:
            subtotal = _f(m.group(1))

        # ----------------------
        # Freight Rate (KN invoices)
        # ----------------------
        freight_mode = None
        freight_rate = None

        m = KN_AIR_FREIGHT.search(text)
        if m:
            freight_mode = "Air"
            freight_rate = _f(m.group(1))

        # ----------------------
        # Return row
        # ----------------------
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": filename,
            "Invoice_Date": inv_date,
            "Currency": currency,
            "Shipper": shipper,
            "Weight_KG": weight,
            "Volume_M3": volume,
            "Chargeable_KG": chargeable_kg,
            "Chargeable_CBM": None,
            "Pieces": pieces,
            "Subtotal": subtotal,
            "Freight_Mode": freight_mode,
            "Freight_Rate": freight_rate,
        }

    except Exception:
        traceback.print_exc()
        return None


# ================================================================
#  STREAMLIT UI
# ================================================================
st.set_page_config(page_title="Invoice Extractor", page_icon="📄", layout="wide")

st.title("📄 Invoice Extractor – KN + SY Invoices")
st.caption("Upload PDF → Extract → Download Excel")

uploads = st.file_uploader(
    "Upload invoices",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("Extract") and uploads:
    rows = []
    prog = st.progress(0)
    total = len(uploads)

    for i, f in enumerate(uploads, start=1):
        data = f.read()

        # Clean filename → extract numeric ID
        invoice_id = extract_invoice_id(f.name)

        row = parse_invoice_pdf_bytes(data, invoice_id)

        if row:
            rows.append(row)
        else:
            st.warning(f"Could not parse {f.name}")

        prog.progress(i / total)

    if rows:
        df = pd.DataFrame(rows)
        df = df[HEADERS]

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)

        for _, r in df.iterrows():
            ws.append([r[h] for h in HEADERS])

        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            output,
            file_name="Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
