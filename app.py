#!/usr/bin/env python3
# Streamlit – Kuehne + Nagel Invoice Extractor (13 fields)

import io
import re
import traceback
from datetime import datetime
from typing import Dict, Any, Optional, List

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook

# ------------------------------------------------------------
# Extract only the invoice number (Example: 2400047153)
# ------------------------------------------------------------
def extract_invoice_id(filename: str):
    name = filename.upper()
    m = re.search(r"(\d{6,12})", name)
    if m:
        return m.group(1)
    return filename


# ------------------------------------------------------------
# Column order (13 fields)
# ------------------------------------------------------------
HEADERS = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_Rate"
]

# ------------------------------------------------------------
# Main parser (KN SALES INVOICE FORMAT)
# ------------------------------------------------------------
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages)

        # ------------------ Invoice Number & Date ------------------
        inv_date = None
        m = re.search(r"INVOICE NO\.?\s*\/\s*DATE\s*\d+\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            d, mth, y = m.group(1).split(".")
            inv_date = f"{y}-{mth}-{d}"

        # ------------------ Currency (USD from subtotal) ------------------
        currency = None
        m = re.search(r"SUBTOTAL\s+(USD|CAD|EUR)\s+[\d,]+\.\d{2}", text)
        if m:
            currency = m.group(1).upper()

        # ------------------ Shipper ------------------
        shipper = None
        m = re.search(r"SHIPPER\s*:\s*(.+)", text)
        if m:
            shipper = m.group(1).strip()

        # ------------------ Pieces ------------------
        pieces = None
        m = re.search(r"(\d+)\s+PCS", text)
        if m:
            pieces = int(m.group(1))

        # ------------------ Weight ------------------
        weight = None
        m = re.search(r"GROSS WT\.?\s*\/\s*KG\s*([\d.]+)", text)
        if m:
            weight = float(m.group(1))

        # ------------------ Volume (CBM) ------------------
        volume = None
        m = re.search(r"VOLUME\s*\/\s*CBM\s*([\d.]+)", text)
        if m:
            volume = float(m.group(1))

        # ------------------ Chargeable Weight ------------------
        chargeable_kg = None
        m = re.search(r"CHG\.?\s*WT\.?\s*([\d.]+)", text)
        if m:
            chargeable_kg = float(m.group(1))

        # ------------------ Chargeable CBM ------------------
        chargeable_cbm = volume

        # ------------------ Subtotal ------------------
        subtotal = None
        m = re.search(r"SUBTOTAL\s+[A-Z]{3}\s+([\d,]+\.\d{2})", text)
        if m:
            subtotal = float(m.group(1).replace(",", ""))

        # ------------------ Freight Mode ------------------
        f_mode = "Air"  # always Air for these invoices

        # ------------------ Freight Rate (Option A) ------------------
        f_rate = None
        m = re.search(r"AIRFREIGHT.*?USD\s*([\d,]+\.\d{2})", text)
        if m:
            f_rate = float(m.group(1).replace(",", ""))

        # ------------------ Build output row ------------------
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": extract_invoice_id(filename),
            "Invoice_Date": inv_date,
            "Currency": currency,
            "Shipper": shipper,
            "Weight_KG": weight,
            "Volume_M3": volume,
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


# ------------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------------
st.set_page_config(
    page_title="KN Invoice Extractor",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Kuehne + Nagel Invoice → Excel Extractor")
st.caption("Extracts 13 required fields from KN Sales Invoice PDFs.")

uploads = st.file_uploader("Upload KN Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Extract", type="primary") and uploads:

    rows = []
    progress = st.progress(0)
    total = len(uploads)

    for i, f in enumerate(uploads, start=1):
        data = f.read()
        row = parse_invoice_pdf_bytes(data, f.name)
        if row:
            rows.append(row)
        else:
            st.warning(f"⚠ Could not extract: {f.name}")
        progress.progress(i / total)

    if rows:
        df = pd.DataFrame(rows).reindex(columns=HEADERS)
        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        # Excel export
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)

        for _, r in df.iterrows():
            ws.append([r[h] for h in HEADERS])

        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇ Download Invoice_Summary.xlsx",
            data=output,
            file_name="Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


