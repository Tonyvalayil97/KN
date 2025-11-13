import re
from datetime import datetime

# ─────────────────────────────────────────────────────────────
# REGEX DEFINITIONS FOR KUEHNE + NAGEL ONLY
# ─────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────
# MAIN PARSER — KN INVOICE ONLY
# ─────────────────────────────────────────────────────────────

def parse_kn_invoice(text: str, filename: str):
    
    # ---------- Invoice Number & Date ----------
    inv_no = None
    inv_date = None

    m = INVOICE_DATE_PAT.search(text)
    if m:
        inv_no = m.group(1)
        d, mm, yy = m.group(2).split(".")
        inv_date = f"{yy}-{mm}-{d}"

    # ---------- Clean Filename (use invoice number ONLY) ----------
    clean_filename = inv_no if inv_no else filename

    # ---------- Shipper ----------
    shipper = None
    m = SHIPPER_BLOCK_PAT.search(text)
    if m:
        block = m.group(1).strip()
        first_line = block.split("\n")[0].strip()
        shipper = first_line

    # ---------- Pieces, Weight, Volume, Chargeable ----------
    pieces = weight = volume = chargeable = None
    m = LINE_ITEM_PAT.search(text)
    if m:
        pieces = int(m.group(1))
        weight = float(m.group(2))
        volume = float(m.group(3))
        chargeable = float(m.group(4))

    # ---------- Subtotal ----------
    subtotal = None
    m = SUBTOTAL_PAT.search(text)
    if m:
        subtotal = float(m.group(1).replace(",", ""))

    # ---------- Freight Rate (Option A USD) ----------
    freight_rate = None
    m = FREIGHT_RATE_PAT.search(text)
    if m:
        freight_rate = float(m.group(1).replace(",", ""))

    # ───────── Return final KN row ────────────

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



