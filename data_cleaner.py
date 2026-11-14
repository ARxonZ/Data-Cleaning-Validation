"""
Data Cleaning & Structural Validation (Excel)
---------------------------------------------
Reads the raw sales workbook (Data_Cleaning_Start.xlsx) and produces
Cleaned_Data.xlsx with three sheets:
  - Cleaned Data : the cleaned, analysis-ready table (Profit Margin is a live formula)
  - Change Log   : every change / removed row / flagged value, with the reason
  - Raw Data     : untouched copy of the original data for comparison

Usage:  python data_cleaner.py [input.xlsx]
"""
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

INPUT_FILE = sys.argv[1] if len(sys.argv) > 1 else "Data_Cleaning_Start.xlsx"
RAW_SHEET = "Sheet1 (2)"
OUTPUT_FILE = "Cleaned_Data.xlsx"

# Company names that plain Title Case would get wrong
NAME_FIXES = {
    "AMZN": "Amazon.com, Inc.", "JPM": "JPMorgan Chase & Co.", "MCD": "McDonald's Corporation",
    "T": "AT&T Inc.", "FDX": "FedEx Corporation", "PG": "The Procter & Gamble Company",
    "GS": "The Goldman Sachs Group, Inc.", "HD": "The Home Depot, Inc.", "MS": "Morgan Stanley", "BAC": "Bank of America Corporation",
}

log = []  # (excel row in raw sheet, column, action, detail)


def add_log(row, column, action, detail):
    log.append((row, column, action, detail))


# ---------- load ----------
def load():
    df = pd.read_excel(INPUT_FILE, sheet_name=RAW_SHEET, header=1, usecols="B:I")
    df = df.drop(columns=["Profit Margin"], errors="ignore")  # recomputed as a formula later
    df = df.dropna(how="all").copy()
    df["raw_row"] = df.index + 3  # header is on Excel row 2, data starts on row 3
    return df.reset_index(drop=True)


# ---------- cleaning helpers ----------
def clean_person(name):
    name = re.sub(r"\s+", " ", str(name)).strip()   # trim + collapse extra spaces
    name = name.title()
    return re.sub(r"\bMc([a-z])", lambda m: "Mc" + m.group(1).upper(), name)  # Mckay -> McKay


def split_client(raw):
    """'AMAZON.COM, INC. (XNAS:AMZN)' -> company, exchange, ticker"""
    m = re.match(r"^(.*?)\s*\((\w+):(\w+)\)\s*$", str(raw).strip())
    if not m:
        return str(raw).strip().title(), None, None
    company, exchange, ticker = m.groups()
    company = NAME_FIXES.get(ticker, company.title())
    return company, exchange, ticker


def main():
    df = load()
    raw = df.copy()
    n_raw = len(df)

    # 1. remove exact duplicates (keep first occurrence)
    key = ["Date", "Client", "Contact", "Department", "Payment", "Revenue", "Profit"]
    dup_mask = df.duplicated(subset=key, keep="first")
    for _, r in df[dup_mask].iterrows():
        add_log(r["raw_row"], "Entire row", "Removed duplicate", f"Identical to an earlier row: {r['Client']}")
    df = df[~dup_mask].copy()

    # 2. clean text columns
    cleaned = []
    for _, r in df.iterrows():
        row = r["raw_row"]
        contact = clean_person(r["Contact"])
        if contact != r["Contact"]:
            add_log(row, "Contact", "Standardised", f"'{r['Contact']}' -> '{contact}'")

        company, exchange, ticker = split_client(r["Client"])
        add_log(row, "Client", "Split / standardised", f"'{r['Client']}' -> '{company}' | {exchange} | {ticker}")

        dept, _, loc = str(r["Department"]).partition("_")
        add_log(row, "Department", "Split", f"'{r['Department']}' -> '{dept.strip()}' | '{loc.strip()}'")

        payment = r["Payment"]
        if pd.isna(payment):
            payment = "Unknown"
            add_log(row, "Payment", "Missing value", "Blank -> 'Unknown'")

        revenue = r["Revenue"]
        if pd.isna(revenue):
            add_log(row, "Revenue", "Missing value flagged",
                    "Revenue is blank; left empty (not guessed). Profit Margin shows blank. Verify with source.")

        cleaned.append({
            "Date": r["Date"].to_pydatetime(), "Company": company, "Exchange": exchange,
            "Ticker": ticker, "Contact": contact, "Department": dept.strip(),
            "Location": loc.strip(), "Payment": payment,
            "Revenue": None if pd.isna(revenue) else float(revenue), "Profit": float(r["Profit"]),
        })

    out = pd.DataFrame(cleaned).sort_values(["Date", "Company"], kind="stable").reset_index(drop=True)

    # ---------- write workbook ----------
    arial = Font(name="Arial", size=10)
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F3864")
    flag_fill = PatternFill("solid", start_color="FFF2CC")

    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned Data"
    cols = list(out.columns) + ["Profit Margin"]
    ws.append(cols)
    for i, r in out.iterrows():
        ws.append(list(r) + [None])
        xl = i + 2
        ws.cell(xl, len(cols)).value = f'=IF(OR(I{xl}="",I{xl}=0),"",J{xl}/I{xl})'

    for c in range(1, len(cols) + 1):
        cell = ws.cell(1, c)
        cell.font, cell.fill = head_font, head_fill
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = arial
        row[0].number_format = "yyyy-mm-dd"
        row[8].number_format = "#,##0.00"
        row[9].number_format = "#,##0.00"
        row[10].number_format = "0.0%"
        if row[8].value is None:
            row[8].fill = flag_fill
        if row[7].value == "Unknown":
            row[7].fill = flag_fill
    for i, w in enumerate([12, 34, 10, 8, 20, 18, 12, 10, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    n = len(out) + 3
    ws.cell(n, 1, "Note: yellow cells = missing values that were flagged, not guessed. "
                  "Profit Margin = Profit / Revenue (formula).").font = Font(name="Arial", size=9, italic=True)

    # change log
    cl = wb.create_sheet("Change Log")
    cl.append(["Raw Sheet Row", "Column", "Action", "Detail"])
    for entry in sorted(log, key=lambda x: (x[0], x[1])):
        cl.append(list(entry))
    for c in range(1, 5):
        cl.cell(1, c).font, cl.cell(1, c).fill = head_font, head_fill
    for row in cl.iter_rows(min_row=2):
        for cell in row:
            cell.font = arial
    for i, w in enumerate([14, 14, 24, 90], 1):
        cl.column_dimensions[get_column_letter(i)].width = w
    cl.freeze_panes = "A2"

    # raw data copy
    rw = wb.create_sheet("Raw Data")
    rcols = ["raw_row"] + [c for c in raw.columns if c != "raw_row"]
    rw.append(["Original Row"] + rcols[1:])
    for _, r in raw[rcols].iterrows():
        rw.append([None if pd.isna(v) else (v.to_pydatetime() if hasattr(v, "to_pydatetime") else v) for v in r])
    for c in range(1, len(rcols) + 1):
        rw.cell(1, c).font, rw.cell(1, c).fill = head_font, head_fill
    for row in rw.iter_rows(min_row=2):
        for cell in row:
            cell.font = arial
        row[1].number_format = "yyyy-mm-dd"
    for i, w in enumerate([13, 12, 42, 24, 24, 12, 12, 10], 1):
        rw.column_dimensions[get_column_letter(i)].width = w

    wb.save(OUTPUT_FILE)

    removed = n_raw - len(out)
    print(f"Raw rows: {n_raw} | Duplicates removed: {removed} | Clean rows: {len(out)}")
    print(f"Missing Revenue flagged: {int(out['Revenue'].isna().sum())} | "
          f"Missing Payment set to Unknown: {int((out['Payment'] == 'Unknown').sum())}")
    print(f"Saved '{OUTPUT_FILE}'")


if __name__ == "__main__":
    main()
