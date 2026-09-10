"""
Export the "EIMS Inventory" tab from the RIMS & EIMS Central Portal.

The portal (https://vnatmfg.intel.com/vf-rims-portal/home/) is a single-page app.
The "EIMS Inventory" tab just fetches a tab-delimited text file:
    {site}/EIMS_Inventory_Report.txt   (default site = VNAT)
i.e. https://vnatmfg.intel.com/vf-rims-portal/home/vnat/EIMS_Inventory_Report.txt

That endpoint requires Windows Integrated Authentication (Negotiate/NTLM).
Python `requests` can't do SSPI here, so we download with Windows' built-in
curl.exe using --negotiate (current logged-in credentials), then convert the
tab file to Excel (.xlsx) and CSV.

Run:
    C:\\Users\\ngocluup\\AppData\\Local\\miniforge3\\envs\\ngocluup\\python.exe export_eims_inventory.py
"""
import os
import subprocess
import sys

from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# This script lives in scripts/; data files live in the project's data/ folder.
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "data")
os.makedirs(DATA_DIR, exist_ok=True)

SITE = "vnat"  # change to another site code if needed (e.g. produces {site}/EIMS_Inventory_Report.txt)
URL = "https://vnatmfg.intel.com/vf-rims-portal/home/%s/EIMS_Inventory_Report.txt" % SITE

TXT_PATH = os.path.join(DATA_DIR, "EIMS_Inventory_Report.txt")
XLSX_PATH = os.path.join(DATA_DIR, "EIMS_Inventory_Report.xlsx")
CSV_PATH = os.path.join(DATA_DIR, "EIMS_Inventory_Report.csv")

# --- Row filter: keep only rows where the column value is in the given set. ---
# Set to None or {} to disable filtering. Column names must match the file headers.
# These reproduce the AutoFilter previously applied in Excel.
FILTERS = {
    "Operation": {"7000"},
    "Mgr_Name": {
        "Bui Anh Huy", "Dong Dang Phu", "Le Thanh Loc",
        "Nguyen Dang Hai", "Nguyen Duc Loc", "Sky Tran",
    },
    "Department": {"A/T MFG", "VNAT TEG-T"},
    "Prodgroup3": {
        "ADLN", "ARLR816B", "ARLS816B", "ARLU281", "ASL", "BMG21",
        "BTL12P", "GRR", "MTLU281", "RPLP282", "RPLP282IOTG", "RPRP282", "TWL",
    },
}


def download():
    """Download the tab file using Windows curl.exe with integrated auth."""
    print("Downloading:", URL)
    result = subprocess.run(
        [
            "curl.exe",
            "-s", "-S",           # silent but show errors
            "--negotiate",        # use Windows Integrated Authentication
            "-u", ":",            # use current logged-in credentials
            "--fail",             # non-zero exit on HTTP >= 400
            "-o", TXT_PATH,
            URL,
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        sys.exit("curl failed (exit %d): %s" % (result.returncode, result.stderr.strip()))
    size = os.path.getsize(TXT_PATH)
    print("Saved raw file: %s (%d bytes)" % (TXT_PATH, size))


def parse_tab(path):
    """Parse the tab-delimited file into (headers, rows)."""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        lines = f.read().split("\n")
    lines = [ln for ln in lines if ln.strip() != ""]
    if not lines:
        return [], []
    headers = [h.strip() for h in lines[0].split("\t")]
    rows = []
    for line in lines[1:]:
        values = line.split("\t")
        row = [values[i].strip() if i < len(values) else "" for i in range(len(headers))]
        rows.append(row)
    return headers, rows


def apply_filters(headers, rows):
    """Keep only rows matching every active filter in FILTERS."""
    if not FILTERS:
        return rows
    idx = {}
    for col in FILTERS:
        if col not in headers:
            sys.exit("Filter column not found in data: %s" % col)
        idx[col] = headers.index(col)
    kept = []
    for row in rows:
        if all(row[idx[col]] in allowed for col, allowed in FILTERS.items()):
            kept.append(row)
    print("Filtered: %d -> %d row(s)" % (len(rows), len(kept)))
    return kept


def export_excel(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "EIMS Inventory"
    ws.append(headers)
    for row in rows:
        ws.append(row)
    try:
        wb.save(XLSX_PATH)
        print("Exported %d row(s) to: %s" % (len(rows), XLSX_PATH))
    except PermissionError:
        import time
        alt = os.path.join(DATA_DIR, "EIMS_Inventory_Report_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S"))
        wb.save(alt)
        print("Target .xlsx was locked (open in Excel?). Exported %d row(s) to: %s" % (len(rows), alt))


def export_csv(headers, rows):
    import csv
    try:
        f = open(CSV_PATH, "w", newline="", encoding="utf-8-sig")
        path = CSV_PATH
    except PermissionError:
        import time
        path = os.path.join(DATA_DIR, "EIMS_Inventory_Report_%s.csv" % time.strftime("%Y%m%d_%H%M%S"))
        f = open(path, "w", newline="", encoding="utf-8-sig")
    with f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    print("Exported %d row(s) to: %s" % (len(rows), path))


def main():
    download()
    headers, rows = parse_tab(TXT_PATH)
    if not headers:
        sys.exit("No data parsed from the downloaded file.")
    print("Columns:", ", ".join(headers))
    rows = apply_filters(headers, rows)
    export_excel(headers, rows)
    export_csv(headers, rows)


if __name__ == "__main__":
    main()
