#!/usr/bin/env python3
# Streamlit App – KN Invoice Extractor

import io
import re
import pdfplumber
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import Workbook


# ==========================================================
#   REGEX DEFINITIONS FOR KUEHNE + NAGEL ONLY
# ==========================================================
INVOICE_DATE_PAT = re.compile(
    r"INVOICE NO\.?\s*\/\s*DATE\s*(\d+)\s+(\d{2}\.\d{2}\.\d{4})",
    re.I
)

SHIPPER_BLOCK_PAT = re.compile(
    r"SHIPPER\s+NOTIFY\s+(.+?)(?=\nCONSIGNEE)",
    re.S | re.I
)

LINE_ITEM_PAT = re.compile(
    r"(\d+)\s+ELEGANT SHOES\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)",
    re.I
)

SUBTOTAL_PAT = re.compile(
    r"SUBTOTAL\s+USD\s+([\d,]+\.\d{2})",
    re.I
)

FREIGHT_RATE_PAT = re.compile(
    r"AIRFREIGHT.*?USD\s*([\d,]+\.\d{2})",
    re.I
)


# ==========================================================
#   MAIN KN PARSER
# ==========================================================
def parse_kn_invoice(text: str, filename: str):
    inv_no = None
    inv_date = None

    # ---- invoice number + date ----
    m = INVOICE_DATE_PAT.search(text)
    if m:
        inv_no = m.group(1)
        d, mm, yy = m.group(2).split(".")
        inv_date = f"{yy}-{mm}-{d}"

    # ---- clean filename (use number only) ----
    clean_filename = inv_no if inv_no else filename

    # ---- shipper ----
    shipper = None
    m = SHIPPER_BLOCK_PAT.search(text)
    if m:
        block = m.group(1).strip()
        shipper = block.split("\n")[0].strip()

    # ---- weight, volume, chargeable, pcs ----
    pieces = weight = volume = chargeable = None
    m = LINE_ITEM_PAT.search(text)
    if m:
        pieces = int(m.group(1))
        weight = float(m.group(2))
        volume = float(m.group(3))
        chargeable = float(m.group(4))

    # ---- subtotal ----
    subtotal = None
    m = SUBTOTAL_PAT.search(text)
    if m:
        subtotal = float(m.group(1).replace(",", ""))

    # ---- freight rate ----
    freight_rate = None
    m = FREIGHT_RATE_PAT.search(text)
    if m:
        freight_rate = float(m.group(1).replace(",", ""))

    return {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Filename": clean_filename,
        "Invoice_Date": inv_date,
        "Currency": "USD",
        "Shipper": shipper,
        "Weight_KG": weight,
        "Volume_M3": volume,
        "Chargeable_KG": chargeable,
        "Chargeable_CBM": volume,
        "Pieces": pieces,
        "Subtotal": subtotal,
        "Freight_Mode": "Air",
        "Freight_Rate": freight_rate,
    }


# ==========================================================
#   STREAMLIT APP UI
# ==========================================================
st.set_page_config(page_title="KN Invoice Extractor", page_icon="📄", layout="wide")

st.title("📄 KN Invoice Extractor – Air Freight")
st.caption("Upload KN invoices → Extract values → Download Excel.")

uploads = st.file_uploader(
    "Upload KN invoice PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

extract = st.button("Extract Invoices", type="primary", disabled=not uploads)


if extract and uploads:
    rows = []
    progress = st.progress(0)
    status = st.empty()
    total = len(uploads)

    for i, file in enumerate(uploads, start=1):
        status.write(f"Processing **{file.name}**")
        pdf_bytes = file.read()

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        row = parse_kn_invoice(text, file.name)

        if row:
            rows.append(row)
        else:
            st.warning(f"⚠️ Could not extract data from {file.name}")

        progress.progress(i / total)

    if not rows:
        st.error("❌ No data extracted.")
    else:
        df = pd.DataFrame(rows)

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        # --- Build Excel file ---
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "KN_Summary"

        headers = list(df.columns)
        ws.append(headers)

        for _, r in df.iterrows():
            ws.append([r[h] for h in headers])

        wb.save(output)
        output.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            output,
            file_name="KN_Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

