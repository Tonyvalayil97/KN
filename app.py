#!/usr/bin/env python3
# Streamlit – Universal Air Freight Invoice Extractor (KLN + Kuehne & Nagel)

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

# --------------------------------------------------------------
# Extract numeric invoice ID from filename (ex: 26693 or 26693A)
# --------------------------------------------------------------
def extract_invoice_id(filename: str):
    name = filename.upper()
    m = re.search(r"(\d{4,6}[A-Z]?)", name)
    if m:
        return m.group(1)
    return filename


# --------------------------------------------------------------
# Extract currency from filename only
# --------------------------------------------------------------
def extract_currency_from_filename(filename: str):
    name = filename.upper()
    if " CAD" in name:
        return "CAD"
    if " USD" in name:
        return "USD"
    if " EUR" in name:
        return "EUR"
    return None


# --------------------------------------------------------------
# Required output columns
# --------------------------------------------------------------
HEADERS = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_Rate"
]


# --------------------------------------------------------------
# REGEX patterns for multiple invoice formats
# --------------------------------------------------------------

# Invoice date (supports yyyy-mm-dd and dd.mm.yyyy)
INVOICE_DATE_PAT = re.compile(
    r"(?:INVOICE DATE|INVOICE NO\.?\/ DATE|DATE)\s*[:\-]?\s*(\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})",
    re.I
)

# Shipper Name (KLN)
SHIPPER_PAT_KLN = re.compile(
    r"SHIPPER'S NAME.*?\n(.+)",
    re.I
)

# Shipper Name (Kuehne & Nagel)
SHIPPER_PAT_KN = re.compile(
    r"SHIPPER\s*:\s*(.+)",
    re.I
)

# Packages
PIECES_PAT = re.compile(r"\b(\d+)\s+PACKAGE", re.I)

# Weight
WEIGHT_PAT = re.compile(r"Gross Weight.*?([\d.]+)\s*KG", re.I)

# Volume (CBM)
VOLUME_PAT = re.compile(r"([\d.]+)\s*CBM", re.I)

# Chargeable weight
CHARGEABLE_KG_PAT = re.compile(r"Chargeable Weight.*?([\d.]+)", re.I)

# KLN volume weight KG (convert to CBM)
VOLUME_WEIGHT_KG_PAT = re.compile(r"Volume Weight[:\s]+([\d.]+)", re.I)

# Subtotal (Total)
SUBTOTAL_PAT = re.compile(
    r"Total\s*[:\-]?\s*([\d,]+\.\d{2})",
    re.I
)

# KLN Freight Amount (last value on AIR FREIGHT line)
KLN_FREIGHT_AMOUNT_PAT = re.compile(
    r"AIR FREIGHT[^\n]*?([\d,]+\.\d{2})\s*$",
    re.I | re.M
)

# Kuehne + Nagel AIRFREIGHT USD amount (Option A)
KN_FREIGHT_USD_PAT = re.compile(
    r"AIRFREIGHT.*?USD\s*([\d,]+\.\d{2})",
    re.I
)


# --------------------------------------------------------------
# UNIVERSAL PARSER (supports KLN + Kuehne+Nagel)
# --------------------------------------------------------------
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join([p.extract_text() or "" for p in pdf.pages])

        # -------- Invoice Date --------
        inv_date = None
        m = INVOICE_DATE_PAT.search(text)
        if m:
            date_str = m.group(1).strip()
            if "." in date_str:  # dd.mm.yyyy
                d, mth, y = date_str.split(".")
                inv_date = f"{y}-{mth}-{d}"
            else:
                inv_date = date_str

        # -------- Currency (from filename only) --------
        currency = extract_currency_from_filename(filename)

        # -------- Shipper --------
        shipper = None
        m = SHIPPER_PAT_KLN.search(text)
        if not m:
            m = SHIPPER_PAT_KN.search(text)
        if m:
            shipper = m.group(1).strip()

        # -------- Pieces --------
        pieces = None
        m = PIECES_PAT.search(text)
        if m:
            pieces = int(m.group(1))

        # -------- Weight (KG) --------
        weight = None
        m = WEIGHT_PAT.search(text)
        if m:
            weight = float(m.group(1))

        # -------- Volume (CBM) --------
        volume_m3 = None
        m = VOLUME_PAT.search(text)
        if m:
            volume_m3 = float(m.group(1))
        else:
            # KLN Volume KG → convert to CBM
            m = VOLUME_WEIGHT_KG_PAT.search(text)
            if m:
                volume_weight_kg = float(m.group(1))
                volume_m3 = volume_weight_kg / 167.0

        # -------- Chargeable KG --------
        chargeable_kg = None
        m = CHARGEABLE_KG_PAT.search(text)
        if m:
            chargeable_kg = float(m.group(1))
        elif weight and volume_m3:
            chargeable_kg = max(weight, volume_m3 * 167)

        # -------- Chargeable CBM --------
        chargeable_cbm = volume_m3

        # -------- Freight Mode --------
        f_mode = "Air"

        # -------- Freight Rate (Option A for Kuehne + Nagel) --------
        f_rate = None

        # KN logic first
        m = KN_FREIGHT_USD_PAT.search(text)
        if m:
            f_rate = float(m.group(1).replace(",", ""))

        # KLN fallback
        if f_rate is None:
            m = KLN_FREIGHT_AMOUNT_PAT.search(text)
            if m:
                f_rate = float(m.group(1).replace(",", ""))

        # -------- Subtotal --------
        subtotal = None
        m = SUBTOTAL_PAT.search(text)
        if m:
            subtotal = float(m.group(1).replace(",", ""))

        # -------- Build Output Row --------
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": extract_invoice_id(filename),
            "Invoice_Date": inv_date,
            "Currency": currency,
            "Shipper": shipper,
            "Weight_KG": weight,
            "Volume_M3": volume_m3,
            "Chargeable_KG": chargeable_kg,
            "Chargeable_CBM": chargeable_cbm,
            "Pieces": pieces,
            "Subtotal": subtotal,
            "Freight_Mode": f_mode,
            "Freight_Rate": f_rate,
        }

    except Exception:
        traceback.print_exc()
        return None


# --------------------------------------------------------------
# STREAMLIT UI
# --------------------------------------------------------------
st.set_page_config(
    page_title="Air Freight Invoice Extractor",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Air Freight Invoice → Excel Extractor")
st.caption("Supports KLN Freight + Kuehne & Nagel PDFs")

uploads = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True,
)

extract_btn = st.button("Extract", type="primary", disabled=not uploads)

if extract_btn and uploads:

    rows = []
    progress = st.progress(0)
    total = len(uploads)

    for i, f in enumerate(uploads, start=1):
        data = f.read()
        row = parse_invoice_pdf_bytes(data, f.name)
        if row:
            rows.append(row)
        else:
            st.warning(f"❌ Failed to extract {f.name}")
        progress.progress(i / total)

    if rows:
        df = pd.DataFrame(rows).reindex(columns=HEADERS)

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        # Build Excel
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)

        for _, r in df.iterrows():
            ws.append([r[h] for h in HEADERS])

        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇️ Download Invoice_Summary.xlsx",
            data=output,
            file_name="Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

