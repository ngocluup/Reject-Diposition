import requests
import json
import os

# Credentials come from env vars or the git-ignored config.local.json in the
# project root. Never hard-code them here.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_CONFIG_PATH = os.path.join(_ROOT, "config.local.json")
_LOCAL_CONFIG = {}
if os.path.exists(_CONFIG_PATH):
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8-sig") as _f:
            _LOCAL_CONFIG = json.load(_f) or {}
    except (ValueError, OSError):
        _LOCAL_CONFIG = {}


def _secret(name, default=""):
    return os.environ.get(name) or _LOCAL_CONFIG.get(name) or default


header = {
    "SITE": _secret("RUPS_SITE", "VNAT"),
    "TOKEN": _secret("RUPS_TOKEN"),
    "WWID": _secret("RUPS_WWID"),
}

# --- Input: put the unit list in "units.txt" (one unit per line) ---
base_dir = os.path.dirname(os.path.abspath(__file__))
units_file = os.path.join(base_dir, "units.txt")

if not os.path.exists(units_file):
    with open(units_file, "w", encoding="utf-8") as f:
        f.write("# One unit per line. Lines starting with # are ignored.\n")
        f.write("U6NN769103285\n")
    print("Template file created:", units_file, "-> fill in the units and re-run.")

with open(units_file, "r", encoding="utf-8") as f:
    units = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]

unit_info = {
    "api": "CUSTOMIZED_API_SEARCH_UNIT_INFO",
    "unit_list": ",".join(units),
}

url = 'https://rups.intel.com/RUPS_api'
d = json.dumps(unit_info)
res = requests.post(url, data=d, headers=header, verify=False)

print(res.status_code)
print(res.text)

# --- Export the returned unit info to an Excel file ---
from openpyxl import Workbook

output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "RUPS_data.xlsx")

records = []
try:
    data = res.json().get("return", {}).get("data", [])
    for group in data:
        if isinstance(group, list):
            records.extend(item for item in group if isinstance(item, dict))
        elif isinstance(group, dict):
            records.append(group)
except ValueError:
    records = []

if records:
    fieldnames = []
    for rec in records:
        for key in rec:
            if key not in fieldnames:
                fieldnames.append(key)

    wb = Workbook()
    ws = wb.active
    ws.append(fieldnames)
    for rec in records:
        ws.append([
            json.dumps(rec.get(k, ""), ensure_ascii=False)
            if isinstance(rec.get(k), (dict, list)) else rec.get(k, "")
            for k in fieldnames
        ])
    wb.save(output_path)
    print("Exported %d record(s) to: %s" % (len(records), output_path))
else:
    print("No unit records found to export.")